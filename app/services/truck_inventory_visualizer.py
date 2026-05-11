"""Visualize item-based truck assignments using matplotlib and Streamlit."""

from __future__ import annotations

import os
import tempfile

os.environ.setdefault("MPLCONFIGDIR", tempfile.gettempdir())

import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import streamlit as st
from matplotlib.patches import Rectangle

from app.models.truck_summary import TruckSummary


def visualize_truck(truck: TruckSummary, fig_width: float = 12, fig_height: float = 4) -> plt.Figure:
    """Create a 2D top-down item layout for one truck."""
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    truck_width = 1.0
    truck_height = 0.4

    truck_rect = Rectangle((0, 0), truck_width, truck_height, linewidth=2, edgecolor="black", facecolor="white")
    ax.add_patch(truck_rect)

    title = f"{truck.truck_preset} #{truck.truck_id} | KKG Load {truck.kkg_load_number}"
    ax.text(truck_width / 2, truck_height + 0.05, title, ha="center", fontsize=12, fontweight="bold")

    util_text = (
        f"Floor: {truck.utilization_display} | "
        f"Weight: {truck.total_weight:.0f}/{truck.truck_max_weight:.0f} lb | "
        f"Status: {truck.validation_status.upper()}"
    )
    ax.text(truck_width / 2, -0.05, util_text, ha="center", fontsize=9, style="italic")

    for box in truck.boxes:
        rect = Rectangle(
            (box.x, box.y * truck_height),
            box.width,
            box.height * truck_height,
            linewidth=0.8,
            edgecolor="black",
            facecolor=box.color,
            alpha=0.75,
        )
        ax.add_patch(rect)

        if box.width > 0.05 and box.height > 0.08:
            ax.text(
                box.x + box.width / 2,
                (box.y + box.height / 2) * truck_height,
                box.item_number,
                ha="center",
                va="center",
                fontsize=7,
                fontweight="bold",
                color="white",
            )

    ax.set_xlim(-0.02, truck_width + 0.02)
    ax.set_ylim(-0.12, truck_height + 0.12)
    ax.set_aspect("equal")
    ax.axis("off")

    plt.tight_layout()
    return fig


def create_legend(trucks: list[TruckSummary]) -> plt.Figure:
    """Create an item-number color legend."""
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.axis("off")

    item_colors = {}
    for truck in trucks:
        for box in truck.boxes:
            item_colors.setdefault(box.item_number, box.color)

    patches = [
        mpatches.Patch(facecolor=color, edgecolor="black", label=item)
        for item, color in sorted(item_colors.items())
    ]

    if patches:
        ax.legend(handles=patches, loc="center", ncol=min(4, len(patches)), fontsize=10, frameon=True)
        ax.text(0.5, -0.1, "Item # Colors", ha="center", fontsize=12, fontweight="bold", transform=ax.transAxes)
    else:
        ax.text(0.5, 0.5, "No item boxes to display", ha="center", va="center", fontsize=12)

    plt.tight_layout()
    return fig


def render_truck_visualization(trucks: list[TruckSummary]) -> None:
    """Render truck visualizations and item legend in Streamlit."""
    if not trucks:
        st.info("No truck plans available. Upload files and process input to begin.")
        return

    for truck in trucks:
        st.subheader(f"{truck.truck_preset} #{truck.truck_id}")

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Floor Use", truck.utilization_display)
        with col2:
            st.metric("Weight", f"{truck.total_weight:.0f} lb")
        with col3:
            st.metric("Item Qty", truck.items_count)
        with col4:
            st.metric("Status", truck.validation_status.title())

        if truck.validation_notes:
            st.warning("; ".join(truck.validation_notes))

        fig = visualize_truck(truck, fig_width=12, fig_height=4)
        st.pyplot(fig, use_container_width=True)

        if truck.boxes:
            st.markdown("**Items in Truck:**")
            item_counts = {}
            for box in truck.boxes:
                item_counts[box.item_number] = item_counts.get(box.item_number, 0) + 1
            for item_number, count in sorted(item_counts.items()):
                st.caption(f"- {item_number}: {count} boxes")

        st.divider()

    st.subheader("Item Color Legend")
    fig_legend = create_legend(trucks)
    st.pyplot(fig_legend, use_container_width=True)
