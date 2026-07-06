"""
Organization Dashboard Styler
------------------------------
Streamlit app: pick or create your organization, upload your dashboard
workbook to preview your real numbers, choose a color scheme (or fine-tune
individual chart colors), adjust its logo and header text, watch the preview
update live, save the profile so it's remembered next time, and apply it into
a real copy of your dashboard workbook.

Run it with:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

import streamlit as st

from lib import org_profile as op
from lib import storage
from lib.app_state import (
    WIDGET_KEYS, LINE_COLOR_KEYS, ORDER_POSITION_KEYS, SCHEME_KEY,
    LAST_APPLIED_SCHEME_KEY, LOADED_SLUG_KEY, ORG_CHOICE_KEY,
    CUSTOM_SCHEME_LABEL,
    field_updates_to_widget_state, widget_state_to_field_updates,
    detect_scheme, scheme_options,
)
from lib.dashboard_data import extract_dashboard_data
from lib.dashboard_preview import render_preview
from lib.excel_writer import TemplateMismatchError, apply_profile_to_workbook
from lib.settings_form import (
    SettingsFormError, build_settings_form, parse_settings_form, SCHEMES,
)

st.set_page_config(page_title="Organization Dashboard Styler", layout="wide")

# Make the control tabs (Text & logo / Colors / Chart options) bolder and a
# few points larger than Streamlit's default, per Jon's request -- they're the
# main way to move between control groups, so they should read as headings.
st.markdown(
    """
    <style>
      .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p,
      .stTabs [data-baseweb="tab"] p {
          font-size: 1.18rem;
          font-weight: 700;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def _seed_widget_state(profile: op.OrgProfile) -> None:
    """Push every widget key from `profile` into st.session_state,
    overwriting whatever was there, and set the scheme dropdown to whichever
    preset (or "Custom") those colors actually match."""
    for k, v in field_updates_to_widget_state(profile.to_dict()).items():
        st.session_state[k] = v
    detected = detect_scheme(profile.to_dict())
    st.session_state[SCHEME_KEY] = detected
    st.session_state[LAST_APPLIED_SCHEME_KEY] = detected


# ---------------------------------------------------------------------------
# 1. Choose or create an organization (sidebar)
# ---------------------------------------------------------------------------
# Streamlit forbids writing to st.session_state[key] for a widget that has
# already been instantiated *this run* -- so "select this org after saving"
# (below, in section 3) can't set ORG_CHOICE_KEY directly. Instead it leaves
# a plain (non-widget) "_pending_org_choice" flag, applied here, before the
# selectbox is created, on the very next run.
if "_pending_org_choice" in st.session_state:
    st.session_state[ORG_CHOICE_KEY] = st.session_state.pop("_pending_org_choice")

st.sidebar.title("Organization Dashboard Styler")

existing_orgs = storage.list_orgs()
choice = st.sidebar.selectbox(
    "Organization",
    options=["+ New organization"] + existing_orgs,
    index=0,
    key=ORG_CHOICE_KEY,
)

# If we just switched to a different org (or this is the very first run ever
# -- LOADED_SLUG_KEY won't exist yet), seed every widget's session_state
# from that org's saved profile (or blank defaults for a new org) and rerun
# immediately. This is required, not optional: Streamlit widgets that already
# rendered once keep whatever's in their own session_state and silently
# ignore a changed `value=` argument on later reruns -- see app_state.py's
# module docstring for the full explanation (this is what caused the
# settings-form-upload bug Jon reported).
if st.session_state.get(LOADED_SLUG_KEY) != choice:
    fresh_profile = op.OrgProfile() if choice == "+ New organization" else storage.load_profile(choice)
    _seed_widget_state(fresh_profile)
    st.session_state[LOADED_SLUG_KEY] = choice
    st.rerun()

# By this point every widget key above is guaranteed already seeded (see the
# block above), so reading st.session_state directly here reflects the true
# current state even though most of those widgets haven't been drawn yet.
current_profile = op.OrgProfile(
    **widget_state_to_field_updates(st.session_state, defaults=op.OrgProfile())
)

# ---------------------------------------------------------------------------
# 1b. Dashboard workbook -- upload once, used for BOTH the live preview (to
# show real numbers) and the "Apply branding" step lower down. Optional: with
# no workbook the preview falls back to clearly-labeled sample numbers.
# ---------------------------------------------------------------------------
st.sidebar.divider()
st.sidebar.subheader("Your dashboard workbook")
workbook_upload = st.sidebar.file_uploader(
    "Upload your dashboard (.xlsx) -- shows your real numbers in the preview "
    "and is what branding gets applied to.",
    type=["xlsx"], key="template",
)
dashboard_data = extract_dashboard_data(workbook_upload) if workbook_upload is not None else None

# ---------------------------------------------------------------------------
# 1c. Settings form -- an alternative to the color controls. Download a copy
# of the current settings as an Excel form, edit it there, then upload it back.
# ---------------------------------------------------------------------------
with st.sidebar.expander("Prefer Excel? Use a settings form instead", expanded=False):
    st.caption(
        "Download a form pre-filled with the values below, edit it in Excel, "
        "then upload it back here -- same result as using the controls directly."
    )
    st.download_button(
        "Download settings form",
        data=build_settings_form(current_profile),
        file_name=f"{(choice if choice != '+ New organization' else 'organization')}-settings-form.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    uploaded_form = st.file_uploader("Upload a filled-in settings form", type=["xlsx"], key="settings_form")
    if uploaded_form is not None and uploaded_form.file_id != st.session_state.get("_processed_upload_id"):
        try:
            updates = parse_settings_form(uploaded_form)
            for k, v in field_updates_to_widget_state(updates).items():
                st.session_state[k] = v
            merged = widget_state_to_field_updates(st.session_state, defaults=op.OrgProfile())
            detected = detect_scheme(merged)
            st.session_state[SCHEME_KEY] = detected
            st.session_state[LAST_APPLIED_SCHEME_KEY] = detected
            st.session_state["_processed_upload_id"] = uploaded_form.file_id
            st.success("Settings form applied below and in the preview.")
            st.rerun()
        except SettingsFormError as e:
            st.session_state["_processed_upload_id"] = uploaded_form.file_id
            st.error(str(e))

# ---------------------------------------------------------------------------
# 2. Branding controls (left, grouped into tabs) + live preview (right)
# ---------------------------------------------------------------------------
col_controls, col_preview = st.columns([1, 1.55], gap="large")

with col_controls:
    tab_text, tab_colors, tab_charts = st.tabs(["Text & logo", "Colors", "Chart options"])

    with tab_text:
        st.text_input("Organization name (top banner)", key=WIDGET_KEYS["org_name"])
        st.text_input("Header subtitle", key=WIDGET_KEYS["header_subtitle"])
        uploaded_logo = st.file_uploader("Upload a logo (PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo_upload")

    logo_bytes = None
    if uploaded_logo is not None:
        logo_bytes = uploaded_logo.getvalue()
    elif choice != "+ New organization":
        logo_bytes = storage.logo_bytes(choice)

    with tab_colors:
        current_colors = widget_state_to_field_updates(st.session_state, defaults=op.OrgProfile())
        options = scheme_options(current_colors)
        if st.session_state.get(SCHEME_KEY) not in options:
            st.session_state[SCHEME_KEY] = options[0]
        st.selectbox(
            "Color scheme -- sets banner, accent, bar, and line colors together",
            options=options,
            key=SCHEME_KEY,
        )

        chosen_scheme = st.session_state[SCHEME_KEY]
        if chosen_scheme != CUSTOM_SCHEME_LABEL and chosen_scheme != st.session_state.get(LAST_APPLIED_SCHEME_KEY):
            # User just picked a different real preset -- cascade its 6 color
            # fields (including the now-hidden banner colors) into every
            # relevant widget key, then rerun so the pickers show the new values.
            for k, v in field_updates_to_widget_state(SCHEMES[chosen_scheme]).items():
                st.session_state[k] = v
            st.session_state[LAST_APPLIED_SCHEME_KEY] = chosen_scheme
            st.rerun()

        st.caption(
            "Fine-tune any single color below -- picking a different scheme "
            "above resets all of these again."
        )
        st.markdown("**Data bar chart**")
        bc1, bc2 = st.columns(2)
        bc1.color_picker("Actual", key=WIDGET_KEYS["bar_actual_color"])
        bc2.color_picker("Budgeted", key=WIDGET_KEYS["bar_budget_color"])
        st.color_picker("Accent strip", key=WIDGET_KEYS["accent"])

        st.markdown("**Income / Expense line colors** (4 fiscal years)")
        line_cols = st.columns(4)
        for i, c in enumerate(line_cols):
            c.color_picker(f"Series {i + 1}", key=LINE_COLOR_KEYS[i])

    with tab_charts:
        st.markdown("**Fiscal year display order** (1 = first in legend and drawn on top)")
        order_cols = st.columns(4)
        for i, c in enumerate(order_cols):
            c.number_input(f"Series {i + 1}", min_value=1, max_value=4, step=1, key=ORDER_POSITION_KEYS[i])

        positions = [st.session_state[k] for k in ORDER_POSITION_KEYS]
        if sorted(positions) != [1, 2, 3, 4]:
            st.warning("Each series needs a different position (1, 2, 3, 4, no repeats) -- keeping the last valid order.")

        st.markdown("**Axis minimums**")
        ac1, ac2 = st.columns(2)
        ac1.number_input("Bar chart axis min ($)", key=WIDGET_KEYS["bar_axis_min"], step=1000.0)
        ac2.number_input("Line chart axis min ($)", key=WIDGET_KEYS["line_axis_min"], step=1000.0)

        st.markdown("**Chart type** (line or bar, per chart)")
        tc1, tc2, tc3 = st.columns(3)
        _type_opts = ["line", "bar"]
        tc1.selectbox("Income", _type_opts, key=WIDGET_KEYS["income_chart_type"],
                      format_func=str.capitalize)
        tc2.selectbox("Expense", _type_opts, key=WIDGET_KEYS["expense_chart_type"],
                      format_func=str.capitalize)
        tc3.selectbox("Data", _type_opts, key=WIDGET_KEYS["data_chart_type"],
                      format_func=str.capitalize)
        st.caption(
            "Switching a chart's type also changes it in the workbook you apply "
            "to — open that file in Excel to confirm it looks right, as this "
            "app can't fully verify Excel chart conversions on its own."
        )

profile = op.OrgProfile(**widget_state_to_field_updates(st.session_state, defaults=op.OrgProfile()))

if choice == "+ New organization":
    org_name_now = st.session_state.get(WIDGET_KEYS["org_name"], "").strip()
    slug = op.slugify(org_name_now) if org_name_now else None
else:
    slug = choice

with col_preview:
    st.subheader("Live preview")
    if dashboard_data is not None and dashboard_data.ok:
        st.caption("✅ Showing your real Income / Expense / Data numbers from the uploaded workbook.")
    elif dashboard_data is not None:
        st.caption(f"⚠️ Using sample numbers — {dashboard_data.note}")
    else:
        st.caption("Showing sample numbers. Upload your workbook in the sidebar to preview your real data.")
    fig = render_preview(profile, logo_bytes=logo_bytes, data=dashboard_data)
    st.pyplot(fig, width="stretch")

st.divider()

# ---------------------------------------------------------------------------
# 3. Save the profile / 4. Apply branding to the uploaded workbook
# ---------------------------------------------------------------------------
save_col, apply_col = st.columns(2, gap="large")

with save_col:
    st.subheader("Save")
    st.caption("Remember this organization's branding for next time.")
    if not slug:
        st.info("Enter an organization name (Text & logo tab) to enable saving.")
    elif st.button("Save organization profile", type="primary"):
        ext = "png"
        if uploaded_logo is not None:
            ext = uploaded_logo.name.rsplit(".", 1)[-1].lower()
        storage.save_profile(
            slug,
            profile,
            logo_bytes=uploaded_logo.getvalue() if uploaded_logo is not None else None,
            logo_ext=ext,
        )
        st.session_state["_pending_org_choice"] = slug
        st.success(f"Saved. '{slug}' will be remembered next time you open this app.")
        st.rerun()

with apply_col:
    st.subheader("Apply to Excel")
    st.caption(
        "Applies your branding into a new copy of the workbook you uploaded in "
        "the sidebar. Your original file is never modified."
    )
    if workbook_upload is None:
        st.info("Upload your dashboard workbook in the sidebar to enable this.")
    elif st.button("Apply branding"):
        with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
            tmp_in.write(workbook_upload.getvalue())
            tmp_in_path = Path(tmp_in.name)

        tmp_logo_path = None
        if logo_bytes:
            with NamedTemporaryFile(suffix=".png", delete=False) as tmp_logo:
                tmp_logo.write(logo_bytes)
                tmp_logo_path = Path(tmp_logo.name)

        out_path = tmp_in_path.with_name(tmp_in_path.stem + "_branded.xlsx")
        try:
            apply_profile_to_workbook(tmp_in_path, profile, out_path, logo_path=tmp_logo_path)
            st.success("Branding applied.")
            st.download_button(
                "Download branded workbook",
                data=out_path.read_bytes(),
                file_name=f"{(slug or 'organization')}-dashboard.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except TemplateMismatchError as e:
            st.error(
                "This file doesn't look like the standard dashboard template, "
                f"so nothing was changed. Details: {e}"
            )
