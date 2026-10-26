from io import BytesIO
from zipfile import ZipFile

from docx import Document
from pypdf import PdfReader
import pytest

from app.models.skid_tag import SkidTag
from app.services.skid_tag_batches import generate_skid_tag_batches, zip_skid_tag_batches


def tag(dc, po, number=1, total=1):
    return SkidTag(dc, po, "123456789012", "24", number, total, 2)


def test_grouped_documents_contain_only_their_dc_and_preserve_pallets():
    tags = [tag("020", "111", 1, 2), tag("010", "222"), tag("020", "111", 2, 2)]
    batches = generate_skid_tag_batches(tags, group_by_dc=True)
    assert [b.dc for b in batches] == ["020", "010"]
    assert [b.tag_count for b in batches] == [2, 1]
    for batch in batches:
        pages = PdfReader(BytesIO(batch.pdf)).pages
        assert len(pages) == batch.tag_count
        assert all(f"DC {batch.dc}" in page.extract_text() for page in pages)
        text = "\n".join(p.text for p in Document(BytesIO(batch.docx)).paragraphs)
        assert text.count(f"DC {batch.dc}") == batch.tag_count
        assert ("DC 010" if batch.dc == "020" else "DC 020") not in text
    assert "Pallet 2 of 2" in PdfReader(BytesIO(batches[0].pdf)).pages[1].extract_text()
    with ZipFile(BytesIO(zip_skid_tag_batches(batches))) as archive:
        assert set(archive.namelist()) == {
            "skid_tags_DC_020.docx", "skid_tags_DC_020.pdf",
            "skid_tags_DC_010.docx", "skid_tags_DC_010.pdf",
        }
    with ZipFile(BytesIO(zip_skid_tag_batches(batches, include_docx=False))) as archive:
        assert set(archive.namelist()) == {"skid_tags_DC_020.pdf", "skid_tags_DC_010.pdf"}
        for batch in batches:
            assert archive.read(f"{batch.file_stem}.pdf") == batch.pdf


def test_ungrouped_keeps_input_order():
    batches = generate_skid_tag_batches([tag("20", "111"), tag("10", "222"), tag("20", "333")])
    assert len(batches) == 1
    assert batches[0].file_stem == "skid_tags"
    pages = PdfReader(BytesIO(batches[0].pdf)).pages
    assert [page.extract_text().splitlines()[0] for page in pages] == ["DC 20", "DC 10", "DC 20"]


def test_dc_filenames_are_safe_and_unique():
    batches = generate_skid_tag_batches([tag("a/b", "1"), tag("a_b", "2"), tag("A_B", "3")], group_by_dc=True)
    assert len({b.file_stem.casefold() for b in batches}) == 3
    assert all("/" not in b.file_stem for b in batches)


def test_empty_input_is_rejected():
    with pytest.raises(ValueError, match="No SKID tags"):
        generate_skid_tag_batches([], group_by_dc=True)
