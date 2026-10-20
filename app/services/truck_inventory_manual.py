"""Review and validate manual pallet assignments without re-optimizing them."""
from collections import defaultdict
import math

from app.models.truck_inventory_record import TruckInventoryRecord
from app.services.truck_inventory_truck_assigner import _evaluate_truck_groups


def assignment_rows(trucks):
    return [{"Pallet ID": f"{truck.truck_id}:{box.box_id}", "Truck #": truck.truck_id,
             "Order": index, "Item #": box.item_number, "Qty": box.represented_qty,
             "Store #": box.store_number, "DC #": box.dc_number,
             "Load": box.kkg_load_number, "PO": box.retailer_po_number}
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
        record = TruckInventoryRecord(
            source_type='manual', source_file='', po_number=box.retailer_po_number,
            retailer_po_number=box.retailer_po_number, kkg_load_number=box.kkg_load_number,
            item_number=box.item_number, store_number=box.store_number, dc_number=box.dc_number,
            qty=box.represented_qty, item_length=box.item_length, item_width=box.item_width,
            item_height=box.item_height, item_weight=box.item_weight,
            is_stackable=box.represented_qty > 1, stack_qty=box.represented_qty,
            pallet_type=box.item_type, truck_length=template.truck_length,
            truck_width=template.truck_width, truck_height=template.truck_height,
            truck_max_weight=template.truck_max_weight, truck_type=template.truck_preset)
        groups[int(target)].append((order, index, record, box.color))
    result = []
    for truck_id, entries in sorted(groups.items()):
        entries.sort(key=lambda entry: (entry[0], entry[1]))
        records = [entry[2] for entry in entries]
        truck = _evaluate_truck_groups(records, grouping_rule='combined')[0]
        if truck.validation_status != 'valid':
            raise ValueError(f"Truck {truck_id}: " + '; '.join(truck.validation_notes))
        truck.truck_id = truck_id
        for box, entry in zip(truck.boxes, entries):
            box.color = entry[3]
        result.append(truck)
    return result
