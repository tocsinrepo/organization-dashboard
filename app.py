"""
Organization Dashboard Styler
------------------------------
Streamlit app: pick or create your organization, adjust its logo, header
text, and color scheme, watch the preview update live, save the profile so
it's remembered next time, and (optionally) apply it into a real copy of
your dashboard workbook.

Run it with:
    streamlit run app.py
"""

from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

import streamlit as st

from lib import org_profile as op
from lib.dashboard_preview import render_preview
from lib.excel_writer import TemplateMismatchError, apply_profile_to_workbook

st.set_page_config(page_title="Organization Dashboard Styler", layout="wide")
st.title("Organization Dashboard Styler")
st.caption(
    "Set up your organization's logo, header text, and colors. The preview "
    "on the right updates instantly. Save when you're happy, then apply it "
    "to your real dashboard file whenever you're ready."
)

# ---------------------------------------------------------------------------
# 1. Choose or create an organization
# ---------------------------------------------------------------------------
existing_orgs = op.list_orgs()
choice = st.sidebar.selectbox(
    "Organization",
    options=["+ New organization"] + existing_orgs,
    index=0,
)

if choice == "+ New organization":
    new_name = st.sidebar.text_input("New organization name", value="")
    slug = op.slugify(new_name) if new_name else None
    profile = op.OrgProfile(org_name=new_name or "Your Organization")
else:
    slug = choice
    profile = op.load_profile(slug)

st.sidebar.divider()

# ---------------------------------------------------------------------------
# 2. Branding controls
# ---------------------------------------------------------------------------
col_controls, col_preview = st.columns([1, 1.3])

with col_controls:
    st.subheader("Header text")
    profile.org_name = st.text_input("Organization name (top banner)", value=profile.org_name)
    profile.header_subtitle = st.text_input("Header subtitle", value=profile.header_subtitle)

    st.subheader("Logo")
    uploaded_logo = st.file_uploader("Upload a logo (PNG/JPG)", type=["png", "jpg", "jpeg"])
    logo_bytes = None
    if uploaded_logo is not None:
        logo_bytes = uploaded_logo.getvalue()
    elif slug and op.logo_path(slug):
        logo_bytes = op.logo_path(slug).read_bytes()

    st.subheader("Colors")
    profile.banner_primary = st.color_picker("Banner - primary", profile.banner_primary)
    profile.banner_secondary = st.color_picker("Banner - secondary", profile.banner_secondary)
    profile.accent = st.color_picker("Accent strip", profile.accent)

    st.markdown("**Contributions bar chart**")
    profile.bar_actual_color = st.color_picker("Actual", profile.bar_actual_color)
    profile.bar_budget_color = st.color_picker("Budgeted", profile.bar_budget_color)

    st.markdown("**Income / Expense line colors** (4 fiscal years)")
    line_cols = st.columns(4)
    new_line_colors = []
    for i, c in enumerate(line_cols):
        default = profile.line_colors[i] if i < len(profile.line_colors) else "#888888"
        new_line_colors.append(c.color_picker(f"Series {i + 1}", default, key=f"line{i}"))
    profile.line_colors = new_line_colors

with col_preview:
    st.subheader("Live preview")
    fig = render_preview(profile, logo_bytes=logo_bytes)
    st.pyplot(fig, use_container_width=True)

st.divider()

# ---------------------------------------------------------------------------
# 3. Save the organization profile (persists for next time)
# ---------------------------------------------------------------------------
save_col, apply_col = st.columns(2)

with save_col:
    st.subheader("Save")
    if not slug:
        st.info("Enter an organization name above to enable saving.")
    elif st.button("Save organization profile", type="primary"):
        ext = "png"
        if uploaded_logo is not None:
            ext = uploaded_logo.name.rsplit(".", 1)[-1].lower()
        op.save_profile(
            slug,
            profile,
            uploaded_logo_bytes=uploaded_logo.getvalue() if uploaded_logo is not None else None,
            uploaded_logo_ext=ext,
        )
        st.success(f"Saved. '{slug}' will be remembered next time you open this app.")
        st.rerun()

# ---------------------------------------------------------------------------
# 4. Apply to a real dashboard workbook
# ---------------------------------------------------------------------------
with apply_col:
    st.subheader("Apply to Excel")
    st.caption(
        "Upload a copy of your dashboard workbook (built from the standard "
        "template). This never touches your original file -- you'll download "
        "a new file with your branding applied."
    )
    template_upload = st.file_uploader("Your dashboard workbook (.xlsx)", type=["xlsx"], key="template")

    if template_upload is not None and st.button("Apply branding"):
        with NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
            tmp_in.write(template_upload.getvalue())
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
