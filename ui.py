"""Presentation helpers (branding, headers, CSS) for the Streamlit app.
Display only: nothing here reads or changes app data."""

import base64
from pathlib import Path

import streamlit as st

ASSETS = Path(__file__).parent / "assets"
LOGO_PATH = ASSETS / "jay_logo.jpg"

INK = "#1B1915"        # warm off-black
MUTED = "#6E675A"      # warm grey, same family as the background
GOLD = "#D6A21E"       # the one accent (saturation kept under 80%)
GOLD_DEEP = "#B88813"
GOLD_WASH = "#F6E7B8"
GOLD_LIGHT = GOLD_WASH  # kept for callers of the old name


def _logo_data_uri() -> str:
    return "data:image/jpeg;base64," + base64.b64encode(LOGO_PATH.read_bytes()).decode()


CSS = f"""
<style>
/* Page frame: a max width so wide screens don't stretch every line. */
.block-container {{ max-width: 1280px; padding-top: 4.75rem; padding-bottom: 4rem; }}
html, body, [data-testid="stAppViewContainer"] {{ font-variant-numeric: tabular-nums; scroll-behavior: smooth; }}
h1, h2, h3 {{ text-wrap: balance; }}

/* Buttons: flat gold primary, quiet secondary, real hover / press / focus. */
[data-testid^="stBaseButton-"] {{ transition: background-color .2s ease, transform .12s ease, border-color .2s ease; }}
[data-testid^="stBaseButton-"]:active {{ transform: scale(.98); }}
[data-testid^="stBaseButton-"]:focus-visible {{ outline: 2px solid {GOLD}; outline-offset: 2px; }}
[data-testid^="stBaseButton-primary"] {{
    background: {GOLD} !important; color: {INK} !important;
    border: 1px solid {GOLD_DEEP} !important; font-weight: 600 !important; border-radius: 10px !important;
}}
[data-testid^="stBaseButton-primary"]:hover {{ background: {GOLD_DEEP} !important; }}
[data-testid^="stBaseButton-primary"] p {{ color: {INK} !important; font-weight: 600; }}
[data-testid^="stBaseButton-secondary"] {{ border-radius: 10px !important; }}
[data-testid^="stBaseButton-secondary"]:hover {{ background: {GOLD_WASH} !important; border-color: {INK} !important; color: {INK} !important; }}

/* Sidebar: one dark panel, hairline edge, gold used only for the step label. */
section[data-testid="stSidebar"] {{ border-right: 1px solid #3B382F; }}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {{
    background: #1D1B16; border: 1px dashed #5C5340; border-radius: 10px;
}}
.jay-side-title {{ display: flex; align-items: baseline; gap: .6rem; margin: .4rem 0 .9rem; }}
.jay-side-title .n {{ font-family: 'Geist Mono', ui-monospace, monospace; font-size: .78rem; color: {GOLD}; letter-spacing: .04em; }}
.jay-side-title .t {{ color: #F3EFE4; font-weight: 600; font-size: 1.05rem; letter-spacing: -.01em; }}
.jay-side-note {{ color: #A9A291; font-size: .8rem; line-height: 1.5; margin-top: .6rem; }}

/* Header: on the page background, not a dark slab; the gold rule carries the brand. */
.jay-hero {{ display: flex; align-items: center; gap: 1.1rem; padding: 0 0 1.25rem; margin-bottom: 1.6rem;
    border-bottom: 2px solid {GOLD}; }}
.jay-hero img {{ width: 64px; height: 64px; border-radius: 16px; background: #fff; object-fit: contain;
    border: 1px solid #000; padding: 4px; }}
.jay-hero .eyebrow {{ font-family: 'Geist Mono', ui-monospace, monospace; font-size: .74rem; letter-spacing: .08em;
    color: {MUTED}; margin: 0 0 .2rem; }}
.jay-hero h1 {{ color: {INK} !important; font-size: 2.1rem !important; line-height: 1.1 !important; margin: 0 !important;
    padding: 0 !important; font-weight: 700 !important; letter-spacing: -.025em; }}
.jay-hero p.sub {{ color: {MUTED}; margin: .35rem 0 0; font-size: .98rem; max-width: 62ch; line-height: 1.5; }}
@media (max-width: 640px) {{
    .jay-hero img {{ width: 48px; height: 48px; border-radius: 12px; }}
    .jay-hero h1 {{ font-size: 1.55rem !important; }}
}}

/* Section headings: mono step number + title on a hairline, no badges. */
.jay-section {{ display: flex; align-items: baseline; gap: .8rem; margin: 2.6rem 0 .4rem;
    padding-bottom: .55rem; border-bottom: 1px solid #000; }}
.jay-step {{ font-family: 'Geist Mono', ui-monospace, monospace; font-size: .85rem; font-weight: 500;
    color: {GOLD_DEEP}; letter-spacing: .04em; min-width: 1.6rem; }}
.jay-section h3 {{ margin: 0 !important; padding: 0 !important; font-size: 1.35rem !important; font-weight: 600 !important;
    letter-spacing: -.015em; color: {INK}; }}
.jay-sub {{ color: {MUTED}; font-size: .92rem; line-height: 1.55; margin: .5rem 0 1rem; max-width: 70ch; }}

/* Getting started: a numbered run down a rule, not three equal cards. */
.jay-steps {{ list-style: none; margin: 1.2rem 0 .4rem; padding: 0 0 0 1.1rem; border-left: 2px solid {GOLD};
    max-width: 68ch; counter-reset: s; }}
.jay-steps li {{ counter-increment: s; padding: .15rem 0 1.05rem; }}
.jay-steps li:last-child {{ padding-bottom: .2rem; }}
.jay-steps li::before {{ content: counter(s, decimal-leading-zero); font-family: 'Geist Mono', ui-monospace, monospace;
    font-size: .78rem; color: {GOLD_DEEP}; margin-right: .6rem; }}
.jay-steps b {{ color: {INK}; font-weight: 600; }}
.jay-steps span {{ display: block; color: {MUTED}; font-size: .92rem; line-height: 1.5; margin: .15rem 0 0 2.05rem; }}

/* Tabs: ink label + gold underline on the active tab. */
[data-testid="stTab"] {{ transition: color .2s ease; }}
[data-testid="stTab"][aria-selected="true"], [data-testid="stTab"][aria-selected="true"] p,
.stTabs [data-baseweb="tab"][aria-selected="true"],
.stTabs [data-baseweb="tab"][aria-selected="true"] p {{ color: {INK} !important; font-weight: 600; }}
.stTabs [data-baseweb="tab-highlight"] {{ background-color: {GOLD}; height: 3px; }}

/* Containers: white surfaces, softer outer radius than the 10px controls inside. */
[data-testid="stExpander"] details {{ background: #fff; border-radius: 14px; transition: border-color .2s ease; }}
[data-testid="stExpander"] summary:hover {{ background: #FBF6E9; }}
[data-testid="stVerticalBlockBorderWrapper"] {{ background: #fff; border-radius: 14px; }}
[data-testid="stAlert"] {{ border-radius: 12px; }}
/* Info notes in the warm palette instead of Streamlit's cool blue. */
[data-testid="stAlert"]:has([data-testid="stAlertContentInfo"]) > div {{ background: #F2EDE1 !important; border: 1px solid #000; }}
[data-testid="stAlertContentInfo"], [data-testid="stAlertContentInfo"] p {{ color: {INK} !important; }}

/* Preview tables: horizontal scroll on small screens. */
[data-testid="stMarkdownContainer"] table {{ display: block; overflow-x: auto; max-width: 100%; }}

footer {{ visibility: hidden; }}
</style>
"""


def inject_css() -> None:
    st.markdown(CSS, unsafe_allow_html=True)


def hero(title: str, subtitle: str) -> None:
    st.markdown(
        f'<header class="jay-hero"><img src="{_logo_data_uri()}" alt="JAY logo">'
        f'<div><p class="eyebrow">MJIL · Packing materials</p><h1>{title}</h1><p class="sub">{subtitle}</p></div></header>',
        unsafe_allow_html=True,
    )


def section(step, title: str, subtitle: str = "") -> None:
    label = f"{step:02d}" if isinstance(step, int) else step
    st.markdown(
        f'<div class="jay-section"><span class="jay-step">{label}</span><h3>{title}</h3></div>'
        + (f'<p class="jay-sub">{subtitle}</p>' if subtitle else ""),
        unsafe_allow_html=True,
    )


def sidebar_title(text: str, step: str = "01") -> None:
    st.markdown(f'<div class="jay-side-title"><span class="n">{step}</span><span class="t">{text}</span></div>',
                unsafe_allow_html=True)


def sidebar_note(text: str) -> None:
    st.markdown(f'<div class="jay-side-note">{text}</div>', unsafe_allow_html=True)


def getting_started() -> None:
    st.markdown(
        '<ol class="jay-steps">'
        "<li><b>Upload the three files</b><span>Pending PO, PM Requirement and Item Category Master, in the sidebar.</span></li>"
        "<li><b>Generate the report</b><span>Check the as-of date, then click <i>Generate report</i>.</span></li>"
        "<li><b>Assign, review and send</b><span>Pick suppliers for each week's projection, check the supplier-wise table, download and email.</span></li>"
        "</ol>",
        unsafe_allow_html=True,
    )
