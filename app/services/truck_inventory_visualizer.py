"""Visualize item-based truck assignments using matplotlib and Streamlit."""

from __future__ import annotations

import os
import tempfile

os.environ.setdefault("MPLCONFIGDIR", tempfile.gettempdir())

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from matplotlib.patches import Rectangle

from app.models.truck_summary import TruckSummary


FIT_COLOR = "#1B7F3A"
FAIL_COLOR = "#B42318"
TRUCK_EDGE_COLOR = "#202124"
TRUCK_FILL_COLOR = "#F8FAFC"


def visualize_truck(truck: TruckSummary, fig_width: float = 12) -> plt.Figure:
    """Create a scaled 2D top-down item layout for one truck."""
    truck_length = max(truck.truck_length, 1.0)
    truck_width = max(truck.truck_width, 1.0)
    aspect = truck_width / truck_length
    fig_height = max(3.2, min(6.0, fig_width * aspect + 1.5))
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))

    status_color = FIT_COLOR if truck.validation_status == "valid" else FAIL_COLOR
    truck_rect = Rectangle(
        (0, 0),
        truck_length,
        truck_width,
        linewidth=2.4,
        edgecolor=status_color,
        facecolor=TRUCK_FILL_COLOR,
    )
    ax.add_patch(truck_rect)

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
        truck_length / 2,
        truck_width + max(truck_width * 0.09, 8),
        title,
        ha="center",
        fontsize=12,
        fontweight="bold",
        color=TRUCK_EDGE_COLOR,
    )
    ax.text(
        truck_length / 2,
        truck_width + max(truck_width * 0.045, 4),
        subtitle,
        ha="center",
        fontsize=9,
        color=status_color,
    )

    ax.set_xlim(-max(truck_length * 0.02, 10), truck_length + max(truck_length * 0.02, 10))
    ax.set_ylim(-max(truck_width * 0.18, 16), truck_width + max(truck_width * 0.18, 16))
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    return fig


def render_truck_visualization(trucks: list[TruckSummary]) -> None:
    """Render truck visualizations, fit status, and item legend in Streamlit."""
    if not trucks:
        st.info("No evaluated truck plan to visualize. Build the plan first.")
        return

    _render_overall_status(trucks)
    _render_item_legend(trucks)
    st.divider()

    for truck in trucks:
        _render_truck_header(truck)
        fig = visualize_truck(truck, fig_width=12)
        st.pyplot(fig, use_container_width=True)
        _render_item_counts(truck)
        st.divider()


def _draw_truck_direction_labels(ax, truck_length: float, truck_width: float) -> None:
    ax.text(0, -max(truck_width * 0.07, 6), "Rear", ha="left", va="center", fontsize=8, color="#475467")
    ax.text(
        truck_length,
        -max(truck_width * 0.07, 6),
        "Front",
        ha="right",
        va="center",
        fontsize=8,
        color="#475467",
    )
    ax.text(
        truck_length / 2,
        -max(truck_width * 0.12, 11),
        f"Length {truck_length:g} in",
        ha="center",
        va="center",
        fontsize=8,
        color="#475467",
    )
    ax.text(
        -max(truck_length * 0.012, 8),
        truck_width / 2,
        f"Width {truck_width:g} in",
        ha="center",
        va="center",
        rotation=90,
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
        stack_offset = (box.stack_level - 1) * min(max(truck_width * 0.006, 0.5), 2.0)
        rect = Rectangle(
            (x + stack_offset, y + stack_offset),
            width,
            depth,
            linewidth=0.8,
            edgecolor="#111827",
            facecolor=box.color,
            alpha=0.82,
        )
        ax.add_patch(rect)

        if width * depth >= min_label_area:
            label = box.item_number if box.stack_qty <= 1 else f"{box.item_number}\nS{box.stack_level}"
            ax.text(
                x + stack_offset + width / 2,
                y + stack_offset + depth / 2,
                label,
                ha="center",
                va="center",
                fontsize=7,
                fontweight="bold",
                color="white",
            )


def _render_overall_status(trucks: list[TruckSummary]) -> None:
    failed = [truck for truck in trucks if truck.validation_status == "error"]
    total_items = sum(truck.items_count for truck in trucks)
    total_weight = sum(truck.total_weight for truck in trucks)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("KKG Loads", len(trucks))
    with col2:
        st.metric("Item Boxes", total_items)
    with col3:
        st.metric("Total Weight", f"{total_weight:,.0f} lb")
    with col4:
        st.metric("Fit Status", "Fail" if failed else "Fit")

    if failed:
        st.error(f"{len(failed)} load(s) need attention before the plan is fit-ready.")
    else:
        st.success("All evaluated KKG loads fit the selected truck rules.")


def _render_truck_header(truck: TruckSummary) -> None:
    status_label = "Fits" if truck.validation_status == "valid" else "Does Not Fit"
    if truck.validation_status == "valid":
        st.success(f"{truck.truck_preset} | KKG Load {truck.kkg_load_number} | {status_label}")
    else:
        st.error(f"{truck.truck_preset} | KKG Load {truck.kkg_load_number} | {status_label}")

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
                "Placed Boxes": 0,
                "Color": box.color,
                "Length": box.item_length,
                "Width": box.item_width,
                "Height": box.item_height,
                "Weight": box.item_weight,
                "Stack Qty": box.stack_qty,
            }
        item_counts[box.item_number]["Placed Boxes"] += 1
    rows.extend(item_counts.values())
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)


def _render_item_legend(trucks: list[TruckSummary]) -> None:
    item_rows = []
    item_colors = {}
    for truck in trucks:
        for box in truck.boxes:
            item_colors.setdefault(box.item_number, box.color)

    for item_number, color in sorted(item_colors.items()):
        item_rows.append(
            f"<span style='display:inline-block;width:14px;height:14px;"
            f"background:{color};border:1px solid #111827;margin-right:6px;'></span>"
            f"<strong>{item_number}</strong>"
        )

    st.markdown("**Item # Color Key**")
    if not item_rows:
        st.caption("No placed item boxes available for a legend.")
        return

    legend_html = (
        "<div style='display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));"
        "gap:6px 14px;align-items:center;'>"
        + "".join(f"<div>{row}</div>" for row in item_rows)
        + "</div>"
    )
    st.markdown(legend_html, unsafe_allow_html=True)
