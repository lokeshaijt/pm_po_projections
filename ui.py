"""Presentation helpers (branding, headers, CSS) for the Streamlit app.
Display only: nothing here reads or changes app data."""

import base64
from pathlib import Path

import streamlit as st

ASSETS = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS / "jay_logo.jpg"

BLACK = "#111111"
GOLD = "#F1BA19"
GOLD_DEEP = "#C8930C"
GOLD_LIGHT = "#FEE68C"


def _logo_data_uri() -> str:
    return "data:image/jpeg;base64," + base64.b64encode(LOGO_PATH.read_bytes()).decode()


CSS = f"""
<style>
/* Primary buttons: gold with black text (white on gold is unreadable). */
[data-testid^="stBaseButton-primary"] {{
    background: linear-gradient(180deg, {GOLD} 0%, {GOLD_DEEP} 100%) !important;
    color: {BLACK} !important; border: 1px solid {GOLD_DEEP} !important;
    font-weight: 600 !important;
}}
[data-testid^="stBaseButton-primary"]:hover {{ filter: brightness(1.06); }}
[data-testid^="stBaseButton-primary"] p {{ color: {BLACK} !important; }}
[data-testid^="stBaseButton-secondary"]:hover {{ border-color: {GOLD_DEEP} !important; color: {BLACK} !important; }}

/* Sidebar: black panel with gold accents. */
section[data-testid="stSidebar"] {{ border-right: 3px solid {GOLD}; }}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
    background: #1C1C1C; border: 1px dashed #5A4A1A;
}}
.jay-side-title {{ color: {GOLD}; font-weight: 700; font-size: 1.05rem; letter-spacing: .02em;
    margin: .2rem 0 .6rem; display: flex; align-items: center; gap: .5rem; }}
.jay-side-note {{ color: #B8B2A3; font-size: .8rem; margin-top: .4rem; }}

/* Header banner. */
.jay-hero {{ display: flex; align-items: center; gap: 1.1rem; padding: 1rem 1.4rem;
    background: radial-gradient(120% 140% at 0% 0%, #2A2A2A 0%, {BLACK} 60%);
    border-radius: 16px; border-bottom: 4px solid {GOLD}; margin-bottom: 1.2rem;
    box-shadow: 0 6px 18px rgba(0,0,0,.18); }}
.jay-hero img {{ width: 74px; height: 74px; border-radius: 50%; background: #fff; object-fit: contain; }}
.jay-hero h1 {{ color: {GOLD} !important; font-size: 1.75rem !important; margin: 0 !important; padding: 0 !important;
    font-weight: 800 !important; letter-spacing: .01em; }}
.jay-hero p {{ color: #E9E3D3; margin: .2rem 0 0; font-size: .95rem; }}

/* Section headers. */
.jay-section {{ display: flex; align-items: center; gap: .75rem; margin: 1.6rem 0 .35rem;
    padding-bottom: .5rem; border-bottom: 2px solid {GOLD_LIGHT}; }}
.jay-step {{ flex: 0 0 auto; width: 2rem; height: 2rem; border-radius: 50%;
    background: {BLACK}; color: {GOLD}; border: 2px solid {GOLD};
    display: flex; align-items: center; justify-content: center; font-weight: 800; }}
.jay-section h3 {{ margin: 0 !important; padding: 0 !important; font-size: 1.3rem !important; color: {BLACK}; }}
.jay-sub {{ color: #6B6457; font-size: .9rem; margin: 0 0 .8rem; }}

/* Summary cards. */
.jay-cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: .8rem; margin: .4rem 0 .6rem; }}
.jay-card {{ background: #fff; border: 1px solid #E4DCC8; border-left: 5px solid {GOLD};
    border-radius: 12px; padding: .75rem 1rem; }}
.jay-card .v {{ font-size: 1.5rem; font-weight: 800; color: {BLACK}; line-height: 1.2; }}
.jay-card .l {{ font-size: .8rem; color: #6B6457; text-transform: uppercase; letter-spacing: .04em; }}
.jay-card.alert {{ border-left-color: #D64545; }}

/* Getting-started steps. */
.jay-steps {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: .8rem; margin-top: .6rem; }}
.jay-steps div {{ background: #fff; border: 1px solid #E4DCC8; border-top: 4px solid {GOLD}; border-radius: 12px; padding: .9rem 1rem; }}
.jay-steps b {{ display: block; color: {BLACK}; margin-bottom: .25rem; }}
.jay-steps span {{ color: #6B6457; font-size: .9rem; }}

/* Tabs: gold underline on the active tab. */
[data-testid="stTab"][aria-selected="true"], [data-testid="stTab"][aria-selected="true"] p,
.stTabs [data-baseweb="tab"][aria-selected="true"],
.stTabs [data-baseweb="tab"][aria-selected="true"] p {{ color: {BLACK} !important; font-weight: 700; }}
@media (max-width: 640px) {{ .jay-hero img {{ width: 52px; height: 52px; }} .jay-hero h1 {{ font-size: 1.35rem !important; }} }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {GOLD}; height: 3px; }}

/* Expanders and bordered containers as soft cards. */
[data-testid="stExpander"] details {{ background: #fff; border-radius: 12px; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ background: #fff; }}

/* Preview tables: a little breathing room and horizontal scroll on small screens. */
[data-testid="stMarkdownContainer"] table {{ display: block; overflow-x: auto; max-width: 100%; }}

footer {{ visibility: hidden; }}
</style>
"""


def inject_css() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def hero(title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="jay-hero"><img src="{_logo_data_uri()}" alt="JAY logo">'
        f"<div><h1>{title}</h1><p>{subtitle}</p></div></div>",
        unsafe_allow_html=True,
    )


def section(step, title: str, subtitle: str = "") -> None:
    st.markdown(
        f'<div class="jay-section"><div class="jay-step">{step}</div><h3>{title}</h3></div>'
        + (f'<p class="jay-sub">{subtitle}</p>' if subtitle else ""),
        unsafe_allow_html=True,
    )


def sidebar_title(text: str) -> None:
    st.markdown(f'<div class="jay-side-title">{text}</div>', unsafe_allow_html=True)


def sidebar_note(text: str) -> None:
    st.markdown(f'<div class="jay-side-note">{text}</div>', unsafe_allow_html=True)


def cards(items) -> None:
    """items: [(label, value, alert_bool), ...]"""
    html = "".join(
        f'<div class="jay-card{" alert" if alert else ""}"><div class="v">{value}</div><div class="l">{label}</div></div>'
        for label, value, alert in items
    )
    st.markdown(f'<div class="jay-cards">{html}</div>', unsafe_allow_html=True)


def getting_started() -> None:
    st.markdown(
        '<div class="jay-steps">'
        "<div><b>1 · Upload</b><span>Add the Pending PO, PM Requirement and Item Category Master files in the sidebar.</span></div>"
        "<div><b>2 · Generate</b><span>Pick the as-of date and click <i>Generate report</i>.</span></div>"
        "<div><b>3 · Assign &amp; send</b><span>Assign suppliers to each week's projection, preview, download and email.</span></div>"
        "</div>",
        unsafe_allow_html=True,
    )
