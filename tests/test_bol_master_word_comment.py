from pathlib import Path
from zipfile import ZipFile

import pytest

from app.services.bol_multistop_docx_generator import generate_multistop_docx_set
from app.utils.bol_facilities import BOL_FACILITY_LOOKUP
from app.services.bol_pdf_template_stamper import stamp_bol_pdf_set
from pypdf import PdfReader
from app.services import bol_pdf_template_stamper as stamper
from lxml import etree
from tests.test_bol_multistop_grouping import _three_stop_record


@pytest.mark.parametrize('mode,template', [
    ('Standard', 'standard_bol_template.docx'),
    ('No Recourse', 'no_recourse_bol_template.docx'),
])
@pytest.mark.parametrize('enabled', [True, False])
@pytest.mark.parametrize('scope', ['Master only', 'All BOLs', 'Children only'])
def test_master_word_comment_is_toggleable_and_does_not_reach_stops(tmp_path, mode, template, enabled, scope, monkeypatch):
    records = _three_stop_record()
    result = generate_multistop_docx_set(
        records, selected_facility=BOL_FACILITY_LOOKUP['Green Bay Packaging'],
        template_path=Path('app/templates') / template,
        individual_stop_template_path=Path('app/templates') / template,
        master_template_mode=mode, output_dir=tmp_path,
        master_comment_enabled=enabled,
        master_comment_text='Master only: call dispatch & confirm <arrival>.',
        comment_scope=scope,
    )
    assert not result.failed_records
    assert len(result.generated_files) == 4
    for file in result.generated_files:
        with ZipFile(file.file_path) as archive:
            xml = archive.read('word/document.xml').decode()
        expected = enabled and (scope == 'All BOLs' or
                    (scope == 'Master only' and file.document_type == 'combined') or
                    (scope == 'Children only' and file.document_type == 'stop'))
        assert ('MasterBOLCommentBox' in xml) == expected
        assert ('Master only:' in xml) == expected
        if expected:
            assert '&amp;' in xml and '&lt;arrival&gt;' in xml
            root = etree.fromstring(xml.encode())
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            ns['wp'] = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
            marker = root.xpath('//wp:anchor[wp:docPr[@name="MasterBOLCommentBox"]]', namespaces=ns)[0]
            paragraph = marker.xpath('.//w:txbxContent/w:p', namespaces=ns)[0]
            assert marker.xpath('./wp:positionH[@relativeFrom="page"]', namespaces=ns)
            assert marker.xpath('./wp:positionV[@relativeFrom="page"]', namespaces=ns)
            assert len(paragraph.xpath('./w:r/w:t', namespaces=ns)) == 1
            assert paragraph.xpath('./w:r/w:rPr/w:b', namespaces=ns)
            assert paragraph.xpath('./w:r/w:rPr/w:highlight[@w:val="yellow"]', namespaces=ns)
            assert 'Comments:' not in ''.join(paragraph.itertext())
    drawn_boxes = []
    original_draw = stamper._draw_master_comment_box
    def capture_comment(canv, text, box):
        drawn_boxes.append(box)
        original_draw(canv, text, box)
    monkeypatch.setattr(stamper, '_draw_master_comment_box', capture_comment)
    pdfs = stamp_bol_pdf_set(records, BOL_FACILITY_LOOKUP['Green Bay Packaging'],
        result.generated_files, mode='Multistop', multistop_no_recourse=mode == 'No Recourse',
        master_comment_enabled=enabled, master_comment_text='SCOPE-CHECK',
        comment_scope=scope, output_dir=tmp_path / 'pdf')
    assert not pdfs.failed_conversions
    assert len(pdfs.converted_files) == 4
    if enabled and mode == 'Standard':
        if scope == 'Children only':
            assert drawn_boxes == [stamper.CHILD_COMMENT_BOX] * 3
        elif scope == 'Master only':
            assert drawn_boxes == [stamper.MASTER_COMMENT_BOX]
        else:
            assert drawn_boxes.count(stamper.CHILD_COMMENT_BOX) == 3
            assert drawn_boxes.count(stamper.MASTER_COMMENT_BOX) == 1
    for file in pdfs.converted_files:
        expected = enabled and (scope == 'All BOLs' or
                    (scope == 'Master only' and file.document_type == 'combined') or
                    (scope == 'Children only' and file.document_type == 'stop'))
        text = '\n'.join(page.extract_text() for page in PdfReader(file.file_path).pages)
        assert ('SCOPE-CHECK' in text) == expected
