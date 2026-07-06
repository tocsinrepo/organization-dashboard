"""
Draws an on-screen mirror of the dashboard using matplotlib, so branding
changes show up live before anything is written into a real Excel file.

Default layout (below the header banner), a 2x2 grid:
    top-left  = Data (Budget vs Actual)      top-right = At a glance (summary)
    bottom-left = Income                      bottom-right = Expense

Each of the Income, Expense, and Data charts can be drawn as a line or a bar
chart, driven by the profile's *_chart_type fields (the same switch is applied
to the real workbook by excel_writer.py).

If a `data` (DashboardData from lib/dashboard_data.py) is passed in, the preview
renders the organization's REAL numbers read from their uploaded workbook. If
it isn't (no workbook uploaded yet), the preview falls back to clearly-labeled
placeholder numbers purely to show the styling. This module reads numbers for
display but never writes anything.
"""
from __future__ import annotations

from io import BytesIO

import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

from lib.org_profile import OrgProfile

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
    return [float(v) if isinstance(v, (int, float)) else np.nan for v in values]


def _tidy(ax, axis_min, title, is_sample):
    ax.set_title(title + (" (sample)" if is_sample else ""), fontsize=10, fontweight="bold")
    ax.tick_params(labelsize=7)
    for lbl in ax.get_xticklabels():
        lbl.set_rotation(45)
        lbl.set_ha("right")
    ax.set_ylim(bottom=axis_min)
    ax.grid(True, axis="y", alpha=0.25, linewidth=0.6)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


def _draw_multiseries(ax, series_list, categories, display_order, line_colors,
                      axis_min, title, is_sample, chart_type):
    """Draw the Income/Expense-style multi-series data as lines or grouped bars."""
    order = display_order or list(range(len(series_list)))
    slots = [(slot, src) for slot, src in enumerate(order) if src < len(series_list)]
    # Filter to series that actually have values.
    plotted = []
    for slot, src in slots:
        vals = _clean(series_list[src].values)
        if vals and not all(np.isnan(vals)):
            plotted.append((slot, src, vals))
    if not plotted:
        _tidy(ax, axis_min, title, is_sample)
        return

    max_len = max(len(v) for _, _, v in plotted)
    x_labels = categories[:max_len] if categories and len(categories) >= max_len else list(range(max_len))
    x = np.arange(max_len)

    if chart_type == "bar":
        n = len(plotted)
        width = 0.8 / n
        for i, (slot, src, vals) in enumerate(plotted):
            color = line_colors[slot] if slot < len(line_colors) else "#888888"
            label = getattr(series_list[src], "name", None) or f"Series {slot + 1}"
            offset = (i - (n - 1) / 2) * width
            ax.bar(x + offset, vals + [np.nan] * (max_len - len(vals)) if len(vals) < max_len else vals,
                   width=width, color=color, label=label)
        ax.set_xticks(x)
        ax.set_xticklabels(x_labels)
    else:  # line
        for slot, src, vals in plotted:
            color = line_colors[slot] if slot < len(line_colors) else "#888888"
            label = getattr(series_list[src], "name", None) or f"Series {slot + 1}"
            ax.plot(x[:len(vals)], vals, color=color, linewidth=2, label=label)
        ax.set_xticks(x)
        ax.set_xticklabels(x_labels)

    _tidy(ax, axis_min, title, is_sample)
    ax.legend(fontsize=6, loc="upper left", frameon=False, ncol=2)


def _draw_data(ax, bar_cats, bar_vals, actual_color, budget_color, axis_min,
               title, is_sample, chart_type):
    """Draw the Data (Actual vs Budgeted) chart as bars or a line."""
    colors = [actual_color, budget_color]
    x = np.arange(len(bar_vals))
    if chart_type == "line":
        ax.plot(x, bar_vals, color=actual_color, linewidth=2, marker="o")
    else:
        ax.bar(bar_cats[:len(bar_vals)], bar_vals, color=colors[:len(bar_vals)], width=0.55)
    ax.set_xticks(x)
    ax.set_xticklabels(bar_cats[:len(bar_vals)])
    for i, v in enumerate(bar_vals):
        ax.text(i, v, f"{v:,.1f}", ha="center", va="bottom", fontsize=8)
    _tidy(ax, axis_min, title, is_sample)


def _sample_series(sample_arrays, prefix):
    from lib.dashboard_data import Series
    return [Series(name=f"{prefix} FY{i + 1}", values=list(a)) for i, a in enumerate(sample_arrays)]


def render_preview(profile: OrgProfile, logo_bytes: bytes | None = None, data=None):
    """Return a matplotlib Figure mirroring the full dashboard."""
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
        height_ratios=[0.85, 2.15, 2.15],
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

    # --- Top-left: Data (Actual vs Budgeted) ---
    data_ax = fig.add_subplot(gs[1, 0])
    _draw_data(data_ax, bar_cats, bar_vals, profile.bar_actual_color,
               profile.bar_budget_color, profile.bar_axis_min, "Data", is_sample,
               profile.data_chart_type)

    # --- Top-right: At a glance (analysis box) ---
    kpi_ax = fig.add_subplot(gs[1, 1])
    kpi_ax.axis("off")
    _draw_summary(kpi_ax, income, expense, is_sample)

    # --- Bottom-left: Income ---
    income_ax = fig.add_subplot(gs[2, 0])
    _draw_multiseries(income_ax, income, categories, profile.display_order,
                      profile.line_colors, profile.line_axis_min, "Income", is_sample,
                      profile.income_chart_type)

    # --- Bottom-right: Expense ---
    expense_ax = fig.add_subplot(gs[2, 1])
    _draw_multiseries(expense_ax, expense, categories, profile.display_order,
                      profile.line_colors, profile.line_axis_min, "Expense", is_sample,
                      profile.expense_chart_type)

    # --- Footer ---
    fig.text(0.06, 0.03, profile.org_name, fontsize=8, color="gray", va="center")
    if is_sample:
        fig.text(0.97, 0.03,
                 "Preview shows sample numbers — upload your workbook to see real data",
                 fontsize=7.5, color="#B00020", va="center", ha="right", style="italic")

    return fig


def _latest_total(series_list):
    """Headline number for the most recent fiscal-year series: its final data
    point. For the dashboard's cumulative-YTD lines that final point is the
    year-to-date total; summing every month would double-count, so we don't."""
    for s in reversed(series_list):
        vals = [v for v in s.values if isinstance(v, (int, float))]
        if vals:
            return vals[-1], s.name
    return None, None


def _draw_summary(ax, income, expense, is_sample):
    inc_total, inc_name = _latest_total(income)
    exp_total, exp_name = _latest_total(expense)
    ax.text(0.0, 0.98, "At a glance" + (" (sample)" if is_sample else ""),
            fontsize=11, fontweight="bold", va="top")
    rows = []
    if inc_total is not None:
        rows.append((f"Income · {inc_name}", inc_total))
    if exp_total is not None:
        rows.append((f"Expense · {exp_name}", exp_total))
    if inc_total is not None and exp_total is not None:
        rows.append(("Net", inc_total - exp_total))
    y = 0.72
    for label, value in rows:
        ax.text(0.0, y, label, fontsize=9.5, va="center")
        ax.text(1.0, y, f"{value:,.1f}", fontsize=12, fontweight="bold",
                va="center", ha="right")
        y -= 0.28
    if not rows:
        ax.text(0.0, 0.5, "Upload a workbook to see totals here.",
                fontsize=8.5, color="gray", va="center")
