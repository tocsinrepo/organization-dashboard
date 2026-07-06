"""
Draws an on-screen mirror of the dashboard -- header banner, the Income and
Expense line charts, the Data (Budget vs Actual) bar chart, and a small
at-a-glance summary -- using matplotlib, so branding changes show up live as
the user picks a color or adjusts an axis, before anything is written into a
real Excel file.

If a `data` (DashboardData from lib/dashboard_data.py) is passed in, the preview
renders the organization's REAL Income/Expense/Data numbers read from their
uploaded workbook. If it isn't (no workbook uploaded yet), the preview falls
back to clearly-labeled placeholder numbers purely to show the styling. The
real numbers still only ever come from the org's own workbook; this module
reads them for display but never writes anything.
"""
from __future__ import annotations

from io import BytesIO

import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

from lib.org_profile import OrgProfile

# Placeholder data, shaped like the real dashboard (4 fiscal-year lines,
# cumulative-YTD look; one Budget vs Actual bar) purely to preview styling
# when no real workbook has been uploaded yet.
_MONTHS = ["Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May"]
_SAMPLE_INCOME = [
    np.linspace(0.5, 8.5, len(_MONTHS)),
    np.linspace(0.6, 9.8, len(_MONTHS)),
    np.linspace(0.7, 11.2, len(_MONTHS)),
    np.linspace(0.8, 12.5, len(_MONTHS)),
]
_SAMPLE_EXPENSE = [
    np.linspace(0.4, 7.2, len(_MONTHS)),
    np.linspace(0.5, 8.1, len(_MONTHS)),
    np.linspace(0.6, 9.0, len(_MONTHS)),
    np.linspace(0.7, 10.1, len(_MONTHS)),
]
_SAMPLE_BUDGET = 3.15
_SAMPLE_ACTUAL = 3.60


def _clean(values):
    """Drop trailing Nones and coerce to floats for plotting; keep alignment."""
    return [float(v) if isinstance(v, (int, float)) else np.nan for v in values]


def _draw_line_chart(ax, series_list, categories, display_order, line_colors,
                     axis_min, title, is_sample):
    """Draw one Income/Expense-style multi-series line chart onto `ax`."""
    order = display_order or list(range(len(series_list)))
    x = categories if categories else list(range(1, len(series_list[0].values) + 1 if series_list else 1))
    any_plotted = False
    for slot, src_idx in enumerate(order):
        if src_idx >= len(series_list):
            continue
        s = series_list[src_idx]
        vals = _clean(s.values)
        if not vals or all(np.isnan(vals)):
            continue
        xx = x[:len(vals)] if len(x) >= len(vals) else list(range(len(vals)))
        color = line_colors[slot] if slot < len(line_colors) else "#888888"
        label = getattr(s, "name", None) or f"Series {slot + 1}"
        ax.plot(xx, vals[:len(xx)], color=color, linewidth=2, label=label)
        any_plotted = True
    ax.set_title(title + (" (sample)" if is_sample else ""), fontsize=10, fontweight="bold")
    ax.tick_params(labelsize=7)
    for lbl in ax.get_xticklabels():
        lbl.set_rotation(45)
        lbl.set_ha("right")
    ax.set_ylim(bottom=axis_min)
    ax.grid(True, axis="y", alpha=0.25, linewidth=0.6)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    if any_plotted:
        ax.legend(fontsize=6, loc="upper left", frameon=False, ncol=2)


def _sample_series(sample_arrays, prefix):
    from lib.dashboard_data import Series
    return [Series(name=f"{prefix} FY{i + 1}", values=list(a)) for i, a in enumerate(sample_arrays)]


def render_preview(profile: OrgProfile, logo_bytes: bytes | None = None, data=None):
    """Return a matplotlib Figure mirroring the full dashboard.

    Args:
        profile: the branding to apply (colors, text, axis mins, display order).
        logo_bytes: optional logo image bytes to show in the banner.
        data: optional DashboardData with the org's real numbers. When None or
            data.ok is False, clearly-labeled placeholder numbers are used.
    """
    use_real = data is not None and getattr(data, "ok", False)

    if use_real:
        income = data.income or _sample_series(_SAMPLE_INCOME, "Income")
        expense = data.expense or _sample_series(_SAMPLE_EXPENSE, "Expense")
        categories = data.categories or _MONTHS
        bar_cats = data.bar_categories or ["Actual", "Budgeted"]
        bar_vals = [v if isinstance(v, (int, float)) else 0.0 for v in (data.bar_values or [_SAMPLE_ACTUAL, _SAMPLE_BUDGET])]
    else:
        income = _sample_series(_SAMPLE_INCOME, "Income")
        expense = _sample_series(_SAMPLE_EXPENSE, "Expense")
        categories = _MONTHS
        bar_cats = ["Actual", "Budgeted"]
        bar_vals = [_SAMPLE_ACTUAL, _SAMPLE_BUDGET]

    is_sample = not use_real

    fig = plt.figure(figsize=(10, 7.4), dpi=120)
    fig.patch.set_facecolor("white")

    gs = fig.add_gridspec(
        nrows=3, ncols=2,
        height_ratios=[0.85, 2.3, 2.0],
        hspace=0.55, wspace=0.2,
        left=0.06, right=0.97, top=0.98, bottom=0.09,
    )

    # --- Header banner (mirrors Dashboard!A1:A3 fills) ---
    banner_ax = fig.add_subplot(gs[0, :])
    banner_ax.set_xlim(0, 1)
    banner_ax.set_ylim(0, 1)
    banner_ax.axis("off")
    banner_ax.add_patch(plt.Rectangle((0, 0.55), 1, 0.45, color=profile.banner_primary))
    banner_ax.add_patch(plt.Rectangle((0, 0.15), 1, 0.40, color=profile.banner_secondary))
    banner_ax.add_patch(plt.Rectangle((0, 0.0), 1, 0.15, color=profile.accent))
    banner_ax.text(0.02, 0.77, profile.org_name, color="white", fontsize=17,
                   fontweight="bold", va="center", ha="left")
    banner_ax.text(0.02, 0.34, profile.header_subtitle, color="white", fontsize=11,
                   va="center", ha="left")
    if logo_bytes:
        try:
            logo_img = Image.open(BytesIO(logo_bytes))
            logo_ax = fig.add_axes([0.85, 0.885, 0.12, 0.09])
            logo_ax.imshow(logo_img)
            logo_ax.axis("off")
        except Exception:
            pass  # bad/unsupported image -- preview just skips it, doesn't crash

    # --- Income + Expense line charts (real data when available) ---
    income_ax = fig.add_subplot(gs[1, 0])
    expense_ax = fig.add_subplot(gs[1, 1])
    _draw_line_chart(income_ax, income, categories, profile.display_order,
                     profile.line_colors, profile.line_axis_min, "Income", is_sample)
    _draw_line_chart(expense_ax, expense, categories, profile.display_order,
                     profile.line_colors, profile.line_axis_min, "Expense", is_sample)

    # --- Data (Budget vs Actual) bar chart ---
    bar_ax = fig.add_subplot(gs[2, 0])
    colors = [profile.bar_actual_color, profile.bar_budget_color]
    bar_ax.bar(bar_cats[:len(bar_vals)], bar_vals,
               color=colors[:len(bar_vals)], width=0.55)
    bar_ax.set_title("Data" + (" (sample)" if is_sample else ""), fontsize=10, fontweight="bold")
    bar_ax.tick_params(labelsize=8)
    bar_ax.set_ylim(bottom=profile.bar_axis_min)
    bar_ax.grid(True, axis="y", alpha=0.25, linewidth=0.6)
    for spine in ("top", "right"):
        bar_ax.spines[spine].set_visible(False)
    for i, v in enumerate(bar_vals):
        bar_ax.text(i, v, f"{v:,.1f}", ha="center", va="bottom", fontsize=8)

    # --- At-a-glance summary panel (fills what used to be blank space) ---
    kpi_ax = fig.add_subplot(gs[2, 1])
    kpi_ax.axis("off")
    _draw_summary(kpi_ax, income, expense, is_sample)

    # --- Footer ---
    fig.text(0.06, 0.03, profile.org_name, fontsize=8, color="gray", va="center")
    if is_sample:
        fig.text(0.97, 0.03,
                 "Preview shows sample numbers — upload your workbook to see real data",
                 fontsize=7.5, color="#B00020", va="center", ha="right", style="italic")

    return fig


def _latest_total(series_list):
    """Sum of the latest (last non-empty) series' values -- a simple headline number."""
    for s in reversed(series_list):
        vals = [v for v in s.values if isinstance(v, (int, float))]
        if vals:
            return vals[-1] if len(vals) == 1 else sum(vals), s.name
    return None, None


def _draw_summary(ax, income, expense, is_sample):
    inc_total, inc_name = _latest_total(income)
    exp_total, exp_name = _latest_total(expense)
    ax.text(0.0, 0.92, "At a glance" + (" (sample)" if is_sample else ""),
            fontsize=10, fontweight="bold", va="top")
    rows = []
    if inc_total is not None:
        rows.append((f"Income · {inc_name}", inc_total))
    if exp_total is not None:
        rows.append((f"Expense · {exp_name}", exp_total))
    if inc_total is not None and exp_total is not None:
        rows.append(("Net", inc_total - exp_total))
    y = 0.68
    for label, value in rows:
        ax.text(0.0, y, label, fontsize=9, va="center")
        ax.text(1.0, y, f"{value:,.1f}", fontsize=11, fontweight="bold",
                va="center", ha="right")
        y -= 0.26
    if not rows:
        ax.text(0.0, 0.5, "Upload a workbook to see totals here.",
                fontsize=8.5, color="gray", va="center")
