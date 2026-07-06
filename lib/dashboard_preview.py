"""
Draws an on-screen mirror of the dashboard's header banner + sample charts
using matplotlib, so changes show up live as the user moves a slider or
picks a color -- before anything is written into a real Excel file.

Placeholder numbers here are illustrative only (they are not read from any
organization's real data). The real numbers only ever come from the org's
own workbook, via excel_writer.py, on "Apply to Excel".
"""

from __future__ import annotations

from io import BytesIO

import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

from lib.org_profile import OrgProfile

# Placeholder data, shaped like the real dashboard (4 fiscal-year lines,
# cumulative-YTD look; one Budget vs Actual bar) purely to preview styling.
_MONTHS = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May"]
_SAMPLE_LINES = [
    np.linspace(0.5, 8.5, len(_MONTHS)),
    np.linspace(0.6, 9.8, len(_MONTHS)),
    np.linspace(0.7, 11.2, len(_MONTHS)),
    np.linspace(0.8, 12.5, len(_MONTHS)),
]
_SAMPLE_BUDGET = 3.15
_SAMPLE_ACTUAL = 3.60


def render_preview(profile: OrgProfile, logo_bytes: bytes | None = None):
    """Return a matplotlib Figure mirroring the dashboard header + two sample charts."""
    fig = plt.figure(figsize=(9, 6), dpi=120)
    fig.patch.set_facecolor("white")

    # --- Header banner (mirrors Dashboard!A1:A3 fills) ---
    banner_ax = fig.add_axes([0.0, 0.85, 1.0, 0.15])
    banner_ax.set_xlim(0, 1)
    banner_ax.set_ylim(0, 1)
    banner_ax.axis("off")
    banner_ax.add_patch(plt.Rectangle((0, 0.55), 1, 0.45, color=profile.banner_primary))
    banner_ax.add_patch(plt.Rectangle((0, 0.15), 1, 0.40, color=profile.banner_secondary))
    banner_ax.add_patch(plt.Rectangle((0, 0.0), 1, 0.15, color=profile.accent))
    banner_ax.text(0.02, 0.77, profile.org_name, color="white", fontsize=16,
                   fontweight="bold", va="center", ha="left")
    banner_ax.text(0.02, 0.35, profile.header_subtitle, color="white", fontsize=10,
                   va="center", ha="left")

    if logo_bytes:
        try:
            logo_img = Image.open(BytesIO(logo_bytes))
            logo_ax = fig.add_axes([0.85, 0.87, 0.13, 0.11])
            logo_ax.imshow(logo_img)
            logo_ax.axis("off")
        except Exception:
            pass  # bad/unsupported image -- preview just skips it, doesn't crash

    # --- Sample bar chart (Contributions: Budget vs Actual) ---
    bar_ax = fig.add_axes([0.08, 0.46, 0.38, 0.32])
    bar_ax.bar(["Actual", "Budgeted"], [_SAMPLE_ACTUAL, _SAMPLE_BUDGET],
               color=[profile.bar_actual_color, profile.bar_budget_color])
    bar_ax.set_title("Contributions (sample)", fontsize=9)
    bar_ax.set_ylabel("$M", fontsize=8)
    bar_ax.tick_params(labelsize=8)

    # --- Sample line chart (Income/Expense-style, 4 series) ---
    line_ax = fig.add_axes([0.54, 0.46, 0.42, 0.32])
    for i, series in enumerate(_SAMPLE_LINES):
        color = profile.line_colors[i] if i < len(profile.line_colors) else "#888888"
        line_ax.plot(_MONTHS, series, color=color, linewidth=2)
    line_ax.set_title("Income (sample, 4 fiscal years)", fontsize=9)
    line_ax.tick_params(labelsize=7, rotation=45)

    # --- Footer ---
    footer_ax = fig.add_axes([0.0, 0.0, 1.0, 0.05])
    footer_ax.axis("off")
    footer_ax.text(0.02, 0.5, f"{profile.org_name}", fontsize=8, color="gray", va="center")

    return fig
