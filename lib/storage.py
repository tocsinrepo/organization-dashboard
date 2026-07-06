"""
Where organization profiles live -- with a switch between two backends so the
same app is safe both locally and when hosted for multiple visitors.

Why this exists: the app was designed for each organization to run its own
local copy, saving profiles to the on-disk `orgs/` folder. But when the app is
hosted as ONE shared Streamlit Cloud URL, that on-disk folder is shared by
every visitor and wiped on redeploy -- so visitor A would see (and overwrite)
visitor B's organizations. That's the multi-tenant bug flagged in the handoff.

Two backends:
  - **local (default):** persist to disk via lib/org_profile.py, exactly as
    before -- one machine, one org owner, profiles survive restarts.
  - **multi-tenant:** keep each browser session's organizations in that
    session's own `st.session_state`, isolated from every other visitor. These
    are in-memory and ephemeral (they last for the session), which is the right
    trade-off for a shared hosted demo with no login: no cross-visitor bleed.

Turn on multi-tenant mode by setting the env var `MULTI_TENANT=1` (or a
Streamlit secret `multi_tenant = true`) on the hosted deployment. Locally,
leave it unset and behavior is unchanged.

The functions accept an optional `session` mapping so the multi-tenant path is
unit-testable with a plain dict standing in for st.session_state.
"""
from __future__ import annotations

import os

from lib import org_profile as op
from lib.org_profile import OrgProfile

SESSION_ORGS_KEY = "_session_orgs"


def multitenant() -> bool:
    """True when the app should isolate organizations per browser session."""
    if os.environ.get("MULTI_TENANT", "").strip().lower() in ("1", "true", "yes", "on"):
        return True
    try:  # a Streamlit secret is the other way to switch it on when hosted
        import streamlit as st
        return bool(st.secrets.get("multi_tenant", False))
    except Exception:
        return False


def _store(session):
    """The per-session dict of {slug: {profile: dict, logo_bytes, logo_ext}}."""
    if session is None:
        import streamlit as st
        session = st.session_state
    if SESSION_ORGS_KEY not in session:
        session[SESSION_ORGS_KEY] = {}
    return session[SESSION_ORGS_KEY]


def list_orgs(session=None) -> list[str]:
    if multitenant():
        return sorted(_store(session).keys())
    return op.list_orgs()


def load_profile(slug: str, session=None) -> OrgProfile:
    if multitenant():
        rec = _store(session).get(slug)
        return OrgProfile.from_dict(rec["profile"]) if rec else OrgProfile()
    return op.load_profile(slug)


def save_profile(slug: str, profile: OrgProfile, logo_bytes: bytes | None = None,
                 logo_ext: str = "png", session=None) -> None:
    if multitenant():
        store = _store(session)
        rec = store.get(slug, {})
        pdict = profile.to_dict()
        if logo_bytes is not None:
            rec["logo_bytes"] = logo_bytes
            rec["logo_ext"] = logo_ext.lstrip(".")
            pdict["logo_filename"] = f"logo.{rec['logo_ext']}"
        else:
            # keep any logo already stored for this org
            pdict["logo_filename"] = rec.get("profile", {}).get("logo_filename")
        rec["profile"] = pdict
        store[slug] = rec
        return
    op.save_profile(slug, profile, uploaded_logo_bytes=logo_bytes, uploaded_logo_ext=logo_ext)


def logo_bytes(slug: str, session=None) -> bytes | None:
    if multitenant():
        rec = _store(session).get(slug)
        return rec.get("logo_bytes") if rec else None
    p = op.logo_path(slug)
    return p.read_bytes() if p else None
