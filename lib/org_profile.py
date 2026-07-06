"""
Load, save, and list per-organization branding profiles.

Each organization gets its own folder under orgs/<slug>/ containing:
  - profile.json   (org name, header text, colors)
  - logo.<ext>     (their uploaded logo image, any format Pillow can read)

Nothing here touches the dashboard workbook itself — that's excel_writer.py.
This module only knows about the small JSON "branding" record.
"""

from __future__ import annotations

import json
import re
import shutil
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Optional

ORGS_DIR = Path(__file__).resolve().parent.parent / "orgs"

# Defaults match the current Cornerstones REV10 dashboard, confirmed against
# the real workbook (Dashboard!A1:A3 fills, chart series colors) on 2026-07-06.
DEFAULT_BANNER_PRIMARY = "#14495A"     # Dashboard!A1:K1 fill
DEFAULT_BANNER_SECONDARY = "#0C3540"  # Dashboard!A2:K2 fill
DEFAULT_ACCENT = "#F2A900"            # Dashboard!A3:M3 fill
DEFAULT_LINE_COLORS = ["#B8D4D9", "#5FA8B8", "#1F5B6B", "#F2A900"]  # series 0-3, both line charts
DEFAULT_BAR_ACTUAL = "#F2A900"        # bar chart series "Actual"
DEFAULT_BAR_BUDGET = "#D2D2D7"        # bar chart series "Budgeted"
DEFAULT_DISPLAY_ORDER = [0, 1, 2, 3]  # identity: series 0..3 shown/drawn in original order
DEFAULT_AXIS_MIN = 0.0


@dataclass
class OrgProfile:
    org_name: str = "Your Organization"
    header_subtitle: str = "Executive Director's Report"
    banner_primary: str = DEFAULT_BANNER_PRIMARY
    banner_secondary: str = DEFAULT_BANNER_SECONDARY
    accent: str = DEFAULT_ACCENT
    line_colors: list = field(default_factory=lambda: list(DEFAULT_LINE_COLORS))
    bar_actual_color: str = DEFAULT_BAR_ACTUAL
    bar_budget_color: str = DEFAULT_BAR_BUDGET
    logo_filename: Optional[str] = None  # e.g. "logo.png", relative to the org's folder

    # Display order for the 4 line-chart series: display_order[i] = which original
    # series index (0-3) appears in slot i. This controls BOTH legend order and
    # z-order/draw order together -- openpyxl re-sorts series by their `order`
    # value on save, so the two can't be set independently (confirmed 2026-07-07
    # by a direct save/reload test). Slot 0's color is line_colors[0], etc.
    display_order: list = field(default_factory=lambda: list(DEFAULT_DISPLAY_ORDER))

    # Value-axis minimums. bar_axis_min is the Contributions chart's value axis
    # (drawn horizontally, so it reads as the "X axis" visually). line_axis_min
    # is the Income/Expense charts' value axis (drawn vertically, the "Y axis").
    bar_axis_min: float = DEFAULT_AXIS_MIN
    line_axis_min: float = DEFAULT_AXIS_MIN

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "OrgProfile":
        """
        Build an OrgProfile from a saved dict, tolerating older profile.json
        files saved before newer fields (like display_order) existed.

        Deliberately does NOT do `{k: d.get(k, getattr(cls, k, None)) ...}` --
        that looked reasonable but breaks for fields declared with
        `field(default_factory=...)` (list-valued fields): those have no
        class-level attribute to fall back to, so a missing key silently
        became None instead of the real default. Confirmed as the cause of a
        production AttributeError on Streamlit Cloud (2026-07-07) when an org
        profile saved before this field existed was reloaded after the field
        was added. Instantiating a real default object first and overlaying
        only the keys actually present sidesteps this entirely.
        """
        profile = cls()
        for k in cls.__dataclass_fields__:
            if k in d and d[k] is not None:
                setattr(profile, k, d[k])
        return profile


def slugify(name: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", name.strip().lower()).strip("-")
    return slug or "organization"


def list_orgs() -> list[str]:
    """Return org slugs that have a saved profile.json, sorted alphabetically."""
    if not ORGS_DIR.exists():
        return []
    return sorted(
        p.parent.name for p in ORGS_DIR.glob("*/profile.json") if p.is_file()
    )


def org_dir(slug: str) -> Path:
    return ORGS_DIR / slug


def load_profile(slug: str) -> OrgProfile:
    path = org_dir(slug) / "profile.json"
    with open(path, "r", encoding="utf-8") as f:
        return OrgProfile.from_dict(json.load(f))


def logo_path(slug: str) -> Optional[Path]:
    profile_path = org_dir(slug) / "profile.json"
    if not profile_path.exists():
        return None
    profile = load_profile(slug)
    if not profile.logo_filename:
        return None
    p = org_dir(slug) / profile.logo_filename
    return p if p.exists() else None


def save_profile(slug: str, profile: OrgProfile, uploaded_logo_bytes: bytes | None = None,
                  uploaded_logo_ext: str = "png") -> Path:
    """Save profile.json (and, if provided, a new logo file) into orgs/<slug>/."""
    d = org_dir(slug)
    d.mkdir(parents=True, exist_ok=True)

    if uploaded_logo_bytes is not None:
        logo_name = f"logo.{uploaded_logo_ext.lstrip('.')}"
        with open(d / logo_name, "wb") as f:
            f.write(uploaded_logo_bytes)
        profile.logo_filename = logo_name

    with open(d / "profile.json", "w", encoding="utf-8") as f:
        json.dump(profile.to_dict(), f, indent=2)

    return d / "profile.json"


def delete_org(slug: str) -> None:
    d = org_dir(slug)
    if d.exists():
        shutil.rmtree(d)
