"""Review and validate manual pallet assignments without re-optimizing them."""
from collections import defaultdict
import math

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import _evaluate_truck_groups


def assignment_rows(trucks):
    return [{"Pallet ID": f"{truck.truck_id}:{box.box_id}", "Truck #": truck.truck_id,
             "Order": index, "Item #": ' + '.join(dict.fromkeys(m.item_number for m in box.allocations)), "Qty": box.represented_qty,
             "Store #": ', '.join(dict.fromkeys(m.store_number for m in box.allocations)),
             "DC #": ', '.join(dict.fromkeys(m.dc_number for m in box.allocations)),
             "Load": ', '.join(dict.fromkeys(m.kkg_load_number for m in box.allocations)),
             "PO": ', '.join(dict.fromkeys(m.retailer_po_number for m in box.allocations))}
            for truck in trucks for index, box in enumerate(truck.boxes, start=1)]


def apply_manual_assignments(trucks, rows):
    """Return a validated replacement plan or raise without modifying the input."""
    if not trucks or any(t.validation_status != 'valid' for t in trucks):
        raise ValueError('Resolve unassigned pallets before editing truck assignments.')
    sources = {f'{t.truck_id}:{b.box_id}': b for t in trucks for b in t.boxes}
    ids = [row.get('Pallet ID') for row in rows]
    if len(ids) != len(sources) or set(ids) != set(sources):
        raise ValueError('Every pallet must appear exactly once.')
    groups = defaultdict(list)
    templates = {truck.truck_id: truck for truck in trucks}
    for index, row in enumerate(rows):
        try:
            target = float(row['Truck #'])
            order = float(row['Order'])
            if not math.isfinite(target) or target < 1 or not target.is_integer() or not math.isfinite(order) or order < 1:
                raise ValueError()
        except (ValueError, TypeError, KeyError):
            raise ValueError('Truck numbers must be positive whole numbers and order must be numeric.') from None
        box = sources[row['Pallet ID']]
        template = templates.get(int(target), trucks[0])
        for member in box.allocations:
            record = TruckInventoryRecord(
                source_type='manual', source_file='', po_number=member.retailer_po_number,
                retailer_po_number=member.retailer_po_number, kkg_load_number=member.kkg_load_number,
                item_number=member.item_number, store_number=member.store_number, dc_number=member.dc_number,
                qty=member.represented_qty, item_length=member.item_length, item_width=member.item_width,
                item_height=member.item_height, item_weight=member.item_weight,
                is_stackable=box.represented_qty > 1, stack_qty=box.represented_qty,
                stack_by_dimension=bool(box.stack_members), manual_stack_id=row['Pallet ID'],
                pallet_type=member.item_type, truck_length=template.truck_length,
                truck_width=template.truck_width, truck_height=template.truck_height,
                truck_max_weight=template.truck_max_weight, truck_type=template.truck_preset)
            record.truck_roof_clearance = template.roof_clearance_in
            groups[int(target)].append((order, index, record, member.color))
    result = []
    for truck_id, entries in sorted(groups.items()):
        entries.sort(key=lambda entry: (entry[0], entry[1]))
        records = [entry[2] for entry in entries]
        truck = _evaluate_truck_groups(records, grouping_rule='combined')[0]
        if truck.validation_status != 'valid':
            raise ValueError(f"Truck {truck_id}: " + '; '.join(truck.validation_notes))
        truck.truck_id = truck_id
        colors = {entry[2].item_number: entry[3] for entry in entries}
        for box in truck.boxes:
            box.color = colors[box.item_number]
            for member in box.stack_members:
                member.color = colors[member.item_number]
        result.append(truck)
    return result
