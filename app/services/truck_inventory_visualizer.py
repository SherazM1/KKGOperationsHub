"""Visualize truck assignments using matplotlib and Streamlit."""

from __future__ import annotations

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Rectangle
import streamlit as st

from app.models.truck_summary import TruckSummary


def visualize_truck(truck: TruckSummary, fig_width: float = 12, fig_height: float = 4) -> plt.Figure:
    """
    Create a 2D top-down visualization of a truck and its load.
    
    Args:
        truck: TruckSummary with boxes to visualize
        fig_width: Figure width in inches
        fig_height: Figure height in inches
        
    Returns:
        Matplotlib figure object
    """
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    
    # Draw truck boundary (simple rectangle)
    truck_width = 1.0
    truck_height = 0.4
    truck_rect = Rectangle((0, 0), truck_width, truck_height, 
                            linewidth=2, edgecolor="black", facecolor="white")
    ax.add_patch(truck_rect)
    
    # Add truck label
    ax.text(truck_width / 2, truck_height + 0.05, 
            f"{truck.truck_preset} #{truck.truck_id}", 
            ha="center", fontsize=12, fontweight="bold")
    
    # Add utilization info
    util_text = f"Utilization: {truck.utilization_display} ({truck.total_used_pallets:.0f}/{truck.total_capacity_pallets:.0f} pallets)"
    ax.text(truck_width / 2, -0.05, util_text, 
            ha="center", fontsize=9, style="italic")
    
    # Draw boxes (loads) inside truck
    for box in truck.boxes:
        rect = Rectangle(
            (box.x, box.y),
            box.width,
            box.height,
            linewidth=1,
            edgecolor="black",
            facecolor=box.color,
            alpha=0.7,
        )
        ax.add_patch(rect)
        
        # Add label inside box
        label = f"{box.load_group}\n{box.pallet_count:.1f}p"
        ax.text(
            box.x + box.width / 2,
            box.y + box.height / 2,
            label,
            ha="center",
            va="center",
            fontsize=8,
            fontweight="bold",
            color="white",
        )
    
    # Set axis properties
    ax.set_xlim(-0.2, truck_width + 0.2)
    ax.set_ylim(-0.15, truck_height + 0.15)
    ax.set_aspect("equal")
    ax.axis("off")
    
    plt.tight_layout()
    return fig


def visualize_all_trucks(trucks: list[TruckSummary]) -> list[plt.Figure]:
    """Create visualizations for all trucks."""
    figures = []
    for truck in trucks:
        fig = visualize_truck(truck)
        figures.append(fig)
    return figures


def create_legend(trucks: list[TruckSummary]) -> plt.Figure:
    """
    Create a legend showing all load groups and their colors.
    
    Args:
        trucks: List of trucks with load group assignments
        
    Returns:
        Matplotlib figure with legend
    """
    fig, ax = plt.subplots(figsize=(8, 3))
    ax.axis("off")
    
    # Collect all unique load groups and their colors
    group_colors = {}
    for truck in trucks:
        for box in truck.boxes:
            if box.load_group not in group_colors:
                group_colors[box.load_group] = box.color
    
    # Create legend patches
    patches = [
        mpatches.Patch(facecolor=color, edgecolor="black", label=group)
        for group, color in sorted(group_colors.items())
    ]
    
    if patches:
        ax.legend(handles=patches, loc="center", ncol=min(4, len(patches)), 
                  fontsize=10, frameon=True, fancybox=True)
        ax.text(0.5, -0.1, "Load Groups", ha="center", fontsize=12, 
               fontweight="bold", transform=ax.transAxes)
    else:
        ax.text(0.5, 0.5, "No load groups to display", ha="center", va="center",
               fontsize=12)
    
    plt.tight_layout()
    return fig


def render_truck_visualization(trucks: list[TruckSummary]) -> None:
    """
    Render truck visualizations and legend in Streamlit.
    
    Args:
        trucks: List of trucks to visualize
    """
    if not trucks:
        st.info("No trucks assigned. Upload files and configure settings to begin.")
        return
    
    # Display each truck
    for truck in trucks:
        st.subheader(f"{truck.truck_preset} #{truck.truck_id}")
        
        # Truck stats in columns
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Utilization", truck.utilization_display)
        with col2:
            st.metric("Used Pallets", f"{truck.total_used_pallets:.0f}")
        with col3:
            st.metric("Capacity", f"{truck.total_capacity_pallets:.0f}")
        with col4:
            st.metric("Remaining", f"{truck.remaining_capacity:.0f}")
        
        # Truck visualization
        fig = visualize_truck(truck, fig_width=12, fig_height=4)
        st.pyplot(fig, use_container_width=True)
        
        # Load details
        if truck.boxes:
            st.markdown("**Load Groups in Truck:**")
            for box in truck.boxes:
                st.caption(f"• {box.load_group}: {box.pallet_count:.1f} pallets")
        
        st.divider()
    
    # Render legend
    st.subheader("Load Group Legend")
    fig_legend = create_legend(trucks)
    st.pyplot(fig_legend, use_container_width=True)
