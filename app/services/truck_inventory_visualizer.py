"""Visualize item-based truck assignments using matplotlib and Streamlit."""

from __future__ import annotations

import os
import tempfile
from html import escape

os.environ.setdefault("MPLCONFIGDIR", tempfile.gettempdir())

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib.colors import to_rgb
from matplotlib.patches import Polygon, Rectangle

from app.models.truck_summary import TruckSummary


FIT_COLOR = "#1B7F3A"
FAIL_COLOR = "#B42318"
TRUCK_EDGE_COLOR = "#202124"
TRUCK_FILL_COLOR = "#F8FAFC"
TRUCK_SIDE_COLOR = "#E4E7EC"
TRUCK_FRONT_COLOR = "#D0D5DD"
PROJECTION_SKEW = 0.35
PROJECTION_Y_SCALE = 0.55


def visualize_truck(truck: TruckSummary, fig_width: float = 12) -> plt.Figure:
    """Draw a proportional overhead plan with compact item references."""
    length = max(truck.truck_length, 1.0)
    width = max(truck.truck_width, 1.0)
    fig, ax = plt.subplots(figsize=(fig_width, max(3.6, fig_width * width / length + 1.1)))
    fig.patch.set_facecolor('#FFFFFF')
    ax.add_patch(Rectangle((0, 0), length, width, facecolor='#F1F5F9', edgecolor='#94A3B8', linewidth=1.5))
    # A faint centerline makes the two trailer lanes easy to follow.
    ax.plot([0, length], [width / 2, width / 2], color='#CBD5E1', linewidth=0.8, linestyle='--', zorder=1)
    references = {item: index + 1 for index, item in enumerate(sorted({b.item_number or 'UNKNOWN' for b in truck.boxes}))}
    for box in truck.boxes:
        x, y = box.x * length, box.y * width
        w, h = box.width * length, box.height * width
        ax.add_patch(Rectangle((x, y), w, h, facecolor=box.color, edgecolor='#FFFFFF', linewidth=1.6, zorder=2))
        red, green, blue = to_rgb(box.color)
        ink = '#FFFFFF' if 0.2126 * red + 0.7152 * green + 0.0722 * blue < 0.48 else '#172B4D'
        if box.represented_qty > 1:
            # Two short horizontal bars mark a double stack, even on small blocks.
            for offset in (0.73, 0.81):
                ax.plot([x + w * 0.70, x + w * 0.88],
                        [y + h * offset, y + h * offset],
                        color=ink, linewidth=1.6, solid_capstyle='round', zorder=4)
        # Only show labels that fit comfortably at the rendered figure size.
        if w / length * fig_width >= 0.30 and h / width >= 0.14:
            label = f"{references[box.item_number or 'UNKNOWN']:02d}"
            ax.text(x + w / 2, y + h / 2, label, ha='center', va='center', fontsize=8, color=ink, fontweight='bold', zorder=3)
    ax.text(0, -width * 0.14, 'REAR / DOORS', fontsize=8, color='#64748B', ha='left')
    ax.text(length, -width * 0.14, 'FRONT / CAB', fontsize=8, color='#64748B', ha='right')
    ax.text(length / 2, width * 1.12, f'{length:g} in long × {width:g} in wide', ha='center', color='#64748B', fontsize=9)
    if not truck.boxes:
        ax.text(length / 2, width / 2, 'No pallets placed — review fit notes', ha='center', color='#64748B', fontsize=10)
    ax.set_xlim(-length * 0.015, length * 1.015)
    ax.set_ylim(-width * 0.25, width * 1.25)
    ax.set_aspect('equal')
    ax.axis('off')
    fig.tight_layout(pad=0.6)
    return fig


def visualize_truck_3d(truck: TruckSummary, fig_width: float = 9) -> plt.Figure:
    """Draw actual pallet heights inside a dimensioned, open trailer frame."""
    length, width, height = (max(value, 1.0) for value in
                             (truck.truck_length, truck.truck_width, truck.truck_height))
    fig = plt.figure(figsize=(fig_width, 4.8), facecolor='white')
    ax = fig.add_subplot(111, projection='3d', proj_type='ortho')
    for box in truck.boxes:
        x, y = box.x * length, box.y * width
        dx, dy = box.width * length, box.height * width
        # Separate blocks show the seam between the lower and upper pallets.
        for level in range(max(1, box.represented_qty)):
            ax.bar3d(x, y, level * box.item_height, dx, dy, box.item_height,
                     color=box.color, edgecolor='#FFFFFF', linewidth=0.35, shade=True)
        if box.represented_qty > 1:
            ax.text(x + dx / 2, y + dy / 2, box.grouped_height + 1,
                    '=', ha='center', va='bottom', fontsize=8, color='#172B4D')
    frame = '#B5A800'
    # All twelve edges preserve the empty space above and beside the cargo.
    for z in (0, height):
        for y in (0, width):
            ax.plot([0, length], [y, y], [z, z], color=frame, linewidth=0.85)
        for x in (0, length):
            ax.plot([x, x], [0, width], [z, z], color=frame, linewidth=0.85)
    for x in (0, length):
        for y in (0, width):
            ax.plot([x, x], [y, y], [0, height], color=frame, linewidth=0.85)
    ink = '#475467'
    ax.text2D(0.5, 0.96, f'Truck {truck.truck_id} · {truck.truck_preset}',
              transform=ax.transAxes, ha='center', fontsize=11, fontweight='bold', color=ink)
    # Screen-aligned measurements stay readable beside the 2D view and cannot
    # overlap projected cargo, regardless of the trailer proportions.
    for x, label, value in ((0.22, 'INTERIOR LENGTH', length),
                             (0.50, 'INTERIOR WIDTH', width),
                             (0.78, 'INTERIOR HEIGHT', height)):
        fig.text(x, 0.145, label, ha='center', va='center', fontsize=9, color='#64748B')
        fig.text(x, 0.09, f'{value:g} in', ha='center', va='center',
                 fontsize=14, fontweight='bold', color='#172B4D')
    ax.set_xlim(0, length)
    ax.set_ylim(-width * 0.12, width)
    ax.set_zlim(0, height * 1.08)
    ax.set_box_aspect((length, width, height))
    ax.view_init(elev=22, azim=-65)
    ax.set_axis_off()
    fig.subplots_adjust(left=0, right=1, bottom=0.21, top=1)
    return fig


def _visualize_projected_truck(truck: TruckSummary, fig_width: float = 12) -> plt.Figure:
    """Create a scaled projected item layout for one truck."""
    truck_length = max(truck.truck_length, 1.0)
    truck_width = max(truck.truck_width, 1.0)
    aspect = truck_width / truck_length
    fig_height = max(3.8, min(6.5, fig_width * aspect + 2.2))
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))

    status_color = FIT_COLOR if truck.validation_status == "valid" else FAIL_COLOR
    shell_drop = max(truck_width * 0.16, 14)
    truck_poly = _project_rect(0, 0, truck_length, truck_width)
    truck_front = [truck_poly[1], truck_poly[2], (truck_poly[2][0], truck_poly[2][1] - shell_drop), (truck_poly[1][0], truck_poly[1][1] - shell_drop)]
    truck_side = [truck_poly[2], truck_poly[3], (truck_poly[3][0], truck_poly[3][1] - shell_drop), (truck_poly[2][0], truck_poly[2][1] - shell_drop)]
    ax.add_patch(Polygon(truck_front, closed=True, facecolor=TRUCK_FRONT_COLOR, edgecolor=status_color, linewidth=1.6))
    ax.add_patch(Polygon(truck_side, closed=True, facecolor=TRUCK_SIDE_COLOR, edgecolor=status_color, linewidth=1.6))
    ax.add_patch(Polygon(truck_poly, closed=True, facecolor=TRUCK_FILL_COLOR, edgecolor=status_color, linewidth=2.4))

    _draw_truck_direction_labels(ax, truck_length, truck_width)
    _draw_item_boxes(ax, truck, truck_length, truck_width)

    title = f"{truck.truck_preset} | KKG Load {truck.kkg_load_number}"
    subtitle = (
        f"{truck_length:g} x {truck_width:g} x {truck.truck_height:g} in | "
        f"Floor {truck.utilization_display} | "
        f"Weight {truck.total_weight:,.0f}/{truck.truck_max_weight:,.0f} lb | "
        f"{truck.validation_status.upper()}"
    )
    ax.text(
        _project_point(truck_length / 2, truck_width)[0],
        _project_point(truck_length / 2, truck_width)[1] + max(truck_width * 0.22, 26),
        title,
        ha="center",
        fontsize=12,
        fontweight="bold",
        color=TRUCK_EDGE_COLOR,
    )
    ax.text(
        _project_point(truck_length / 2, truck_width)[0],
        _project_point(truck_length / 2, truck_width)[1] + max(truck_width * 0.14, 16),
        subtitle,
        ha="center",
        fontsize=9,
        color=status_color,
    )

    projected_width = truck_width * PROJECTION_SKEW
    projected_height = truck_width * PROJECTION_Y_SCALE
    ax.set_xlim(-max(truck_length * 0.02, 10), truck_length + projected_width + max(truck_length * 0.02, 10))
    ax.set_ylim(-shell_drop - max(truck_width * 0.12, 12), projected_height + max(truck_width * 0.34, 34))
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    return fig


def render_truck_visualization(trucks: list[TruckSummary]) -> None:
    """Render truck visualizations, fit status, and item legend in Streamlit."""
    if not trucks:
        st.info("No evaluated truck plan to visualize. Build the plan first.")
        return

    _render_selected_load_summary(trucks)
    st.caption("3D shows pallet heights and the truck frame. The 2D overhead view shows pallet positions and item key numbers. Both share item colors; = marks a double stack. Dimensions are inches.")

    for truck in trucks:
        _render_truck_header(truck)
        view_3d, view_2d = st.columns(2)
        with view_3d:
            st.markdown("**3D load view**")
            fig = visualize_truck_3d(truck)
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)
        with view_2d:
            st.markdown("**2D overhead view**")
            fig = visualize_truck(truck, fig_width=9)
            st.pyplot(fig, use_container_width=True)
            plt.close(fig)
        with st.expander("Item key", expanded=len({box.item_number for box in truck.boxes}) <= 8):
            _render_item_legend([truck])
        with st.expander("Pallet dimensions and quantities"):
            _render_item_counts(truck)
        with st.expander("Store / DC / PO details"):
            st.dataframe(pd.DataFrame([{
                "Position": index, "Store #": box.store_number, "DC #": box.dc_number,
                "Load": box.kkg_load_number, "PO": box.retailer_po_number,
                "Item #": box.item_number, "Qty": box.represented_qty,
            } for index, box in enumerate(truck.boxes, start=1)]), hide_index=True, use_container_width=True)
        st.divider()


def _draw_truck_direction_labels(ax, truck_length: float, truck_width: float) -> None:
    rear_x, rear_y = _project_point(0, 0)
    front_x, front_y = _project_point(truck_length, 0)
    mid_x, mid_y = _project_point(truck_length / 2, 0)
    width_x, width_y = _project_point(0, truck_width / 2)
    ax.text(rear_x, rear_y - max(truck_width * 0.12, 10), "Rear", ha="left", va="center", fontsize=8, color="#475467")
    ax.text(
        front_x,
        front_y - max(truck_width * 0.12, 10),
        "Front",
        ha="right",
        va="center",
        fontsize=8,
        color="#475467",
    )
    ax.text(
        mid_x,
        mid_y - max(truck_width * 0.18, 16),
        f"Length {truck_length:g} in",
        ha="center",
        va="center",
        fontsize=8,
        color="#475467",
    )
    ax.text(
        width_x - max(truck_length * 0.012, 8),
        width_y,
        f"Width {truck_width:g} in",
        ha="center",
        va="center",
        rotation=28,
        fontsize=8,
        color="#475467",
    )


def _draw_item_boxes(ax, truck: TruckSummary, truck_length: float, truck_width: float) -> None:
    min_label_area = truck_length * truck_width * 0.012
    for box in truck.boxes:
        x = box.x * truck_length
        y = box.y * truck_width
        width = box.width * truck_length
        depth = box.height * truck_width
        grouped_height = box.grouped_height or (box.item_height * max(box.represented_qty, 1))
        height_drop = _visual_height_drop(grouped_height, truck.truck_height, truck_width)
        top = _project_rect(x, y, width, depth)
        front = [top[0], top[1], (top[1][0], top[1][1] - height_drop), (top[0][0], top[0][1] - height_drop)]
        side = [top[1], top[2], (top[2][0], top[2][1] - height_drop), (top[1][0], top[1][1] - height_drop)]
        ax.add_patch(Polygon(front, closed=True, facecolor=_shade_color(box.color, 0.72), edgecolor="#111827", linewidth=0.7, alpha=0.95))
        ax.add_patch(Polygon(side, closed=True, facecolor=_shade_color(box.color, 0.58), edgecolor="#111827", linewidth=0.7, alpha=0.95))
        ax.add_patch(Polygon(top, closed=True, facecolor=box.color, edgecolor="#111827", linewidth=0.9, alpha=0.9))

        if width * depth >= min_label_area:
            type_label = f"{box.item_type}\n" if box.item_type else ""
            label = f"{type_label}{box.item_number}\nx{box.represented_qty}"
            center_x = sum(point[0] for point in top) / len(top)
            center_y = sum(point[1] for point in top) / len(top)
            ax.text(
                center_x,
                center_y,
                label,
                ha="center",
                va="center",
                fontsize=7,
                fontweight="bold",
                color="white",
            )


def _project_point(x: float, y: float) -> tuple[float, float]:
    return x + y * PROJECTION_SKEW, y * PROJECTION_Y_SCALE


def _project_rect(x: float, y: float, width: float, depth: float) -> list[tuple[float, float]]:
    return [
        _project_point(x, y),
        _project_point(x + width, y),
        _project_point(x + width, y + depth),
        _project_point(x, y + depth),
    ]


def _visual_height_drop(grouped_height: float, truck_height: float, truck_width: float) -> float:
    if truck_height <= 0:
        return max(truck_width * 0.05, 4)
    ratio = min(max(grouped_height / truck_height, 0.08), 1.0)
    return max(truck_width * 0.06, 5) + ratio * max(truck_width * 0.12, 12)


def _shade_color(color: str, factor: float) -> str:
    try:
        red, green, blue = to_rgb(color)
    except ValueError:
        red, green, blue = to_rgb("#98A2B3")
    return "#{:02X}{:02X}{:02X}".format(
        int(red * 255 * factor),
        int(green * 255 * factor),
        int(blue * 255 * factor),
    )


def _render_selected_load_summary(trucks: list[TruckSummary]) -> None:
    failed = [truck for truck in trucks if truck.validation_status == "error"]
    load_numbers = sorted({truck.kkg_load_number for truck in trucks if truck.kkg_load_number})
    truck_types = sorted({truck.truck_preset for truck in trucks if truck.truck_preset})
    unique_items = {
        box.item_number
        for truck in trucks
        for box in truck.boxes
        if box.item_number
    }
    total_qty = sum(truck.items_count for truck in trucks)
    total_render_units = sum(len(truck.boxes) for truck in trucks)
    total_weight = sum(truck.total_weight for truck in trucks)

    st.caption("Truck: " + ", ".join(truck_types))
    with st.expander("Selected loads / DC groups"):
        st.write(", ".join(load_numbers) if load_numbers else "No loads selected")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Item Types Placed", len(unique_items))
    with col2:
        st.metric("Total Qty", total_qty)
    with col3:
        st.metric("Total Weight", f"{total_weight:,.0f} lb")
    with col4:
        st.metric("Fit Status", "Fail" if failed else "Fit")

    if failed:
        st.error("This load does not fit the selected truck rules. Review the fit notes below the visualization.")
    st.caption(f"Pallet positions placed: {total_render_units}")


def _render_truck_header(truck: TruckSummary) -> None:
    status_label = "Fits" if truck.validation_status == "valid" else "Does Not Fit"
    st.markdown(f"**Truck {truck.truck_id} · {escape(truck.truck_preset)} · {status_label}**")
    st.caption(f"Load capacity {truck.load_capacity_display} · Floor area {truck.utilization_display} · Interior height {truck.truck_height:g} in · Weight {truck.total_weight:,.0f} / {truck.truck_max_weight:,.0f} lb")
    if truck.pallet_positions_capacity:
        full = (truck.pallet_positions_used or 0) >= truck.pallet_positions_capacity
        st.caption(f"{'FULL · ' if full else ''}{truck.pallet_positions_used:g} / {truck.pallet_positions_capacity:g} pallet positions. One full pallet or two half pallets uses one position.")

    if truck.validation_notes:
        with st.expander("Fit issues", expanded=True):
            st.markdown("\n".join(f"- {note}" for note in truck.validation_notes))


def _render_item_counts(truck: TruckSummary) -> None:
    if not truck.boxes:
        st.caption("No item boxes were placed for this load.")
        return

    rows = []
    item_counts: dict[str, dict] = {}
    for box in truck.boxes:
        if box.item_number not in item_counts:
            item_counts[box.item_number] = {
                "Item #": box.item_number,
                "Render Units": 0,
                "Represented Qty": 0,
                "Color": box.color,
                "Type": box.item_type,
                "Length": box.item_length,
                "Width": box.item_width,
                "Height": box.item_height,
                "Grouped Height": box.grouped_height,
                "Weight": box.item_weight,
                "Stack Qty": box.stack_qty,
            }
        item_counts[box.item_number]["Render Units"] += 1
        item_counts[box.item_number]["Represented Qty"] += box.represented_qty
    rows.extend(item_counts.values())
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _render_item_legend(trucks: list[TruckSummary]) -> None:
    item_rows = _item_legend_rows(trucks)

    st.markdown("**Item Color Key**")
    if not item_rows:
        st.caption("No placed grouped units are available for a color key.")
        return

    legend_html = (
        "<table style='width:100%;border-collapse:collapse;font-size:0.92rem;'>"
        "<thead>"
        "<tr style='border-bottom:1px solid #D0D5DD;text-align:left;'>"
        "<th style='padding:6px 8px;'>Key</th>"
        "<th style='padding:6px 8px;'>Color</th>"
        "<th style='padding:6px 8px;'>Item #</th>"
        "<th style='padding:6px 8px;'>Type</th>"
        "<th style='padding:6px 8px;text-align:right;'>Qty</th>"
        "<th style='padding:6px 8px;text-align:right;'>Grouped Units</th>"
        "</tr>"
        "</thead><tbody>"
        + "".join(_legend_row_html(row) for row in item_rows)
        + "</tbody></table>"
    )
    st.markdown(legend_html, unsafe_allow_html=True)
    st.caption("Each item has its own color. Key numbers match the layout. Quantities include placed pallets only.")


def _item_legend_rows(trucks: list[TruckSummary]) -> list[dict]:
    items: dict[str, dict] = {}
    for truck in trucks:
        for box in truck.boxes:
            item_number = box.item_number or "UNKNOWN"
            if item_number not in items:
                items[item_number] = {
                    "item_number": item_number,
                    "color": box.color,
                    "item_type": box.item_type or "Custom",
                    "qty": 0,
                    "render_units": 0,
                }
            items[item_number]["qty"] += box.represented_qty
            items[item_number]["render_units"] += 1

    return [dict(items[item_number], reference=f"{index:02d}") for index, item_number in enumerate(sorted(items), start=1)]


def _legend_row_html(row: dict) -> str:
    color = escape(str(row["color"]))
    return (
        "<tr style='border-bottom:1px solid #EAECF0;'>"
        f"<td style='padding:7px 8px;font-weight:600;'>{row['reference']}</td>"
        "<td style='padding:7px 8px;'>"
        f"<span style='display:inline-block;width:18px;height:18px;background:{color};"
        "border:1px solid #111827;vertical-align:middle;'></span>"
        "</td>"
        f"<td style='padding:7px 8px;font-weight:600;'>{escape(str(row['item_number']))}</td>"
        f"<td style='padding:7px 8px;'>{escape(str(row['item_type']))}</td>"
        f"<td style='padding:7px 8px;text-align:right;'>{row['qty']}</td>"
        f"<td style='padding:7px 8px;text-align:right;'>{row['render_units']}</td>"
        "</tr>"
    )
