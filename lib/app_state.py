"""
Pure (no-Streamlit) helpers for keeping the web app's widgets, an OrgProfile,
and a chosen color scheme all in sync.

Streamlit's rule that trips people up: once a widget with a given `key` has
a value in `st.session_state`, passing a different `value=` to that widget
on a later rerun does NOT change what it displays -- the widget keeps
whatever's already in session_state. That's what caused a real bug reported
by Jon (2026-07-07): uploading a settings form updated `profile.*` in
Python, but the color pickers/text inputs below immediately overwrote it
right back with their own stale values, because they'd already rendered
once with their own session_state entries and a changed `value=` argument
doesn't move them.

The fix used throughout app.py: every profile-driven widget has an explicit
key (WIDGET_KEYS / LINE_COLOR_KEYS / ORDER_POSITION_KEYS below). Whenever a
new profile should take effect (switching org, uploading a settings form,
picking a color scheme), the app writes directly into
`st.session_state[key] = value` for every affected key BEFORE those widgets
are instantiated on the next run, then calls `st.rerun()` so the widgets
pick up the fresh values as their *initial* session_state, not as an
ignored `value=` argument.

Kept here, not inline in app.py, so the pure merge/format logic (no
Streamlit calls) can be unit tested directly with plain dicts standing in
for `st.session_state`.
"""

from __future__ import annotations

from lib.org_profile import OrgProfile
from lib.settings_form import SCHEMES

CUSTOM_SCHEME_LABEL = "Custom"

# OrgProfile field name -> widget/session-state key. banner_primary/secondary
# have no picker widget anymore (removed 2026-07-07 per Jon's request) --
# they're driven only by the chosen color scheme (or a settings-form
# upload's custom values), so they live in plain hidden session_state
# entries, not a widget's own key.
WIDGET_KEYS = {
    "org_name": "w_org_name",
    "header_subtitle": "w_header_subtitle",
    "banner_primary": "_banner_primary",
    "banner_secondary": "_banner_secondary",
    "accent": "w_accent",
    "bar_actual_color": "w_bar_actual",
    "bar_budget_color": "w_bar_budget",
    "bar_axis_min": "w_bar_axis_min",
    "line_axis_min": "w_line_axis_min",
    "income_chart_type": "w_income_type",
    "expense_chart_type": "w_expense_type",
    "data_chart_type": "w_data_type",
}
LINE_COLOR_KEYS = ["w_line0", "w_line1", "w_line2", "w_line3"]
ORDER_POSITION_KEYS = ["w_order0", "w_order1", "w_order2", "w_order3"]  # 1-4, indexed by series i

SCHEME_KEY = "w_scheme"
LAST_APPLIED_SCHEME_KEY = "_last_applied_scheme"
LOADED_SLUG_KEY = "_loaded_slug"
ORG_CHOICE_KEY = "w_org_choice"

# The 6 fields that make up a "color scheme" -- used to detect whether the
# current colors match one of the 5 presets exactly.
COLOR_SCHEME_FIELDS = ("banner_primary", "banner_secondary", "accent",
                       "bar_actual_color", "bar_budget_color", "line_colors")


def field_updates_to_widget_state(updates: dict) -> dict:
    """Map a dict keyed by OrgProfile field names (a full profile.to_dict(),
    a parse_settings_form() result, or a SCHEMES[name] entry) to a flat dict
    of {widget/session-state key: value}, ready to merge into
    st.session_state. Only includes keys actually present in `updates`, so
    a partial dict (e.g. a settings form where the org name was left blank)
    doesn't stomp on fields it didn't touch."""
    state: dict = {}
    for field, key in WIDGET_KEYS.items():
        if field in updates:
            state[key] = updates[field]

    if "line_colors" in updates:
        colors = list(updates["line_colors"])
        for i, key in enumerate(LINE_COLOR_KEYS):
            if i < len(colors):
                state[key] = colors[i]

    if "display_order" in updates:
        positions_by_series = [0, 0, 0, 0]
        for slot, series_idx in enumerate(updates["display_order"]):
            if 0 <= series_idx < 4:
                positions_by_series[series_idx] = slot + 1
        for i, key in enumerate(ORDER_POSITION_KEYS):
            state[key] = positions_by_series[i] or (i + 1)

    return state


def widget_state_to_field_updates(state, defaults: OrgProfile) -> dict:
    """Read the current widget/session-state values back into a dict keyed
    by OrgProfile field names, falling back to `defaults` for anything not
    yet present in `state`. `state` only needs to support `.get`, so a
    plain dict works for tests -- it doesn't have to be a real
    st.session_state."""
    updates: dict = {}
    for field, key in WIDGET_KEYS.items():
        updates[field] = state.get(key, getattr(defaults, field))

    updates["line_colors"] = [
        state.get(key, defaults.line_colors[i] if i < len(defaults.line_colors) else "#888888")
        for i, key in enumerate(LINE_COLOR_KEYS)
    ]

    positions = [state.get(key, i + 1) for i, key in enumerate(ORDER_POSITION_KEYS)]
    if sorted(positions) == [1, 2, 3, 4]:
        display_order = [0, 0, 0, 0]
        for series_idx, pos in enumerate(positions):
            display_order[pos - 1] = series_idx
        updates["display_order"] = display_order
    else:
        # Not a clean 1-4 permutation (e.g. mid-edit) -- keep whatever the
        # profile already had rather than writing something invalid.
        updates["display_order"] = list(defaults.display_order)

    updates["bar_axis_min"] = float(updates["bar_axis_min"])
    updates["line_axis_min"] = float(updates["line_axis_min"])

    return updates


def detect_scheme(colors: dict) -> str:
    """Given a dict with (at least) the 6 color fields in
    COLOR_SCHEME_FIELDS, return the name of the matching preset in SCHEMES,
    or CUSTOM_SCHEME_LABEL if the colors don't match any preset exactly."""
    for name, scheme in SCHEMES.items():
        if all(colors.get(f) == scheme[f] for f in COLOR_SCHEME_FIELDS):
            return name
    return CUSTOM_SCHEME_LABEL


def scheme_options(current_colors: dict) -> list[str]:
    """Options for the scheme selectbox: the 5 real presets, plus "Custom"
    only when the current colors don't already match one of them -- so the
    dropdown never lies about what's actually applied."""
    names = list(SCHEMES.keys())
    if detect_scheme(current_colors) == CUSTOM_SCHEME_LABEL:
        names = names + [CUSTOM_SCHEME_LABEL]
    return names
