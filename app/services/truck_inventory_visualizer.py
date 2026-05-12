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
from matplotlib.patches import Polygon

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
    st.caption(
        "Truck view shows the current evaluated KKG Load # only. "
        "Each colored block is one grouped stack unit; the x-count on a block is the item quantity represented by that unit."
    )
    _render_item_legend(trucks)
    st.divider()

    for truck in trucks:
        _render_truck_header(truck)
        fig = visualize_truck(truck, fig_width=12)
        st.pyplot(fig, use_container_width=True)
        _render_item_counts(truck)
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

    st.markdown("**Selected Load Summary**")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric("KKG Load #", ", ".join(load_numbers) if load_numbers else "-")
    with col2:
        st.metric("Truck Type", ", ".join(truck_types) if truck_types else "-")
    with col3:
        st.metric("Unique Items", len(unique_items))
    with col4:
        st.metric("Total Qty", total_qty)
    with col5:
        st.metric("Total Weight", f"{total_weight:,.0f} lb")
    with col6:
        st.metric("Fit Status", "Fail" if failed else "Fit")

    if failed:
        st.error("This load does not fit the selected truck rules. Review the fit notes below the visualization.")
    else:
        st.success("This load fits the selected truck rules.")

    st.caption(f"Grouped render units placed: {total_render_units}")


def _render_truck_header(truck: TruckSummary) -> None:
    status_label = "Fits" if truck.validation_status == "valid" else "Does Not Fit"
    if truck.validation_status == "valid":
        st.success(f"Visualization: KKG Load {truck.kkg_load_number} on {truck.truck_preset} | {status_label}")
    else:
        st.error(f"Visualization: KKG Load {truck.kkg_load_number} on {truck.truck_preset} | {status_label}")

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Truck Interior", f"{truck.truck_length:g} x {truck.truck_width:g} in")
    with col2:
        st.metric("Interior Height", f"{truck.truck_height:g} in")
    with col3:
        st.metric("Floor Use", truck.utilization_display)
    with col4:
        st.metric("Weight", f"{truck.total_weight:,.0f}/{truck.truck_max_weight:,.0f} lb")

    if truck.validation_notes:
        st.markdown("**Why this load does not fit:**")
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
    st.caption("Colors identify Item # values in the selected load; PURE is teal and CDW is pink for the current sample mapping.")


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

    return [items[item_number] for item_number in sorted(items)]


def _legend_row_html(row: dict) -> str:
    color = escape(str(row["color"]))
    return (
        "<tr style='border-bottom:1px solid #EAECF0;'>"
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
