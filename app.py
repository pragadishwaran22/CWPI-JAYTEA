from datetime import date
from hashlib import sha256
from pathlib import Path

import streamlit as st

import streamlit.components.v1 as components

from classifier_engine import (
    create_all_classifier_excel_report,
    list_classifier_contractors,
    parse_classifier_workbook,
)
from report_engine import (
    build_contractor_print_html,
    build_report,
    create_excel_report,
)
from loading_overlay import printing_overlay
from version import APP_VERSION
from mapping_store import fetch_name_mappings


LOGO_PATH = Path(__file__).parent / "assets" / "glossy_gold_jay_oval_emblem.png"


@st.cache_data(show_spinner=False)
def read_classifier_shift(workbook_bytes: bytes, shift: str):
    return parse_classifier_workbook(workbook_bytes, shift)


@st.cache_data(ttl="15m", max_entries=4, show_spinner=False)
def read_name_mappings(project_url: str, publishable_key: str):
    return fetch_name_mappings(project_url, publishable_key)


def supabase_mapping_config() -> tuple[str, str] | None:
    try:
        config = st.secrets.get("supabase", {})
    except FileNotFoundError:
        return None
    project_url = str(config.get("url", "")).strip()
    publishable_key = str(config.get("publishable_key", "")).strip()
    return (project_url, publishable_key) if project_url and publishable_key else None


st.set_page_config(
    page_title="JAY Contractor wise Report",
    page_icon=str(LOGO_PATH),
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": None},
)
st.set_option("client.toolbarMode", "minimal")
st.logo(str(LOGO_PATH), size="large")

st.markdown(
    """
    <style>
    :root {
        --jay-gold: #D7A919;
        --jay-gold-light: #F4D66A;
        --jay-ink: #08090D;
        --jay-panel: rgba(28, 29, 34, .62);
        --jay-line: rgba(255, 235, 174, .19);
        --jay-silver: #C6C9CE;
        --jay-warm: #F7F2E6;
        --jay-muted: #AAA79E;
    }
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background:
            radial-gradient(circle at 8% 5%, rgba(215, 169, 25, .15), transparent 26rem),
            radial-gradient(circle at 88% 12%, rgba(198, 201, 206, .10), transparent 28rem),
            radial-gradient(circle at 55% 92%, rgba(215, 169, 25, .08), transparent 32rem),
            linear-gradient(145deg, #08090D 0%, #111218 48%, #090A0E 100%);
        background-attachment: fixed;
    }
    .stApp::before, .stApp::after {
        content: "";
        position: fixed;
        z-index: 0;
        width: 18rem;
        height: 18rem;
        border-radius: 50%;
        pointer-events: none;
        filter: blur(70px);
        opacity: .22;
    }
    .stApp::before { top: 18%; left: 30%; background: #D7A919; }
    .stApp::after { right: 8%; bottom: 8%; background: #757982; }
    [data-testid="stHeader"] {
        background: transparent;
        border-bottom: 0;
        box-shadow: none;
    }
    .block-container { max-width: 1320px; padding-top: 3rem; padding-bottom: 4rem; }
    [data-testid="stSidebar"] {
        background: linear-gradient(165deg, rgba(14, 15, 19, .94), rgba(22, 21, 18, .86));
        border-right: 1px solid rgba(244, 214, 106, .14);
        box-shadow: 18px 0 50px rgba(0, 0, 0, .22);
        backdrop-filter: blur(32px) saturate(135%);
        -webkit-backdrop-filter: blur(32px) saturate(135%);
    }
    [data-testid="stSidebar"] * { color: var(--jay-warm); }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color: var(--jay-muted); }
    [data-testid="stSidebar"] h2 { color: var(--jay-gold-light); letter-spacing: -.02em; }
    [data-testid="stSidebar"] hr { border-color: rgba(244, 214, 106, .13); }
    [data-testid="stSidebar"] input { color: var(--jay-warm); }
    [data-testid="stSidebar"] [data-testid="stDateInput"] [role="group"],
    [data-testid="stSidebar"] [data-testid="stDateInput"] [role="group"] *,
    [data-testid="stSidebar"] [data-testid="stDateInput"] [role="spinbutton"] {
        color: var(--jay-warm) !important;
        -webkit-text-fill-color: var(--jay-warm) !important;
        opacity: 1 !important;
    }
    .st-key-hero_glass_generate, .st-key-hero_glass_classifier {
        position: relative;
        overflow: hidden;
        padding: clamp(1rem, 2vw, 1.5rem);
        margin: .35rem auto 1.1rem;
        max-width: 850px;
        border: 1px solid rgba(255, 239, 187, .22);
        border-radius: 26px;
        background: linear-gradient(135deg, rgba(255, 255, 255, .13), rgba(255, 255, 255, .045));
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .24), inset 0 -1px 0 rgba(255, 211, 76, .06), 0 24px 70px rgba(0, 0, 0, .38);
        backdrop-filter: blur(30px) saturate(165%);
        -webkit-backdrop-filter: blur(30px) saturate(165%);
    }
    .st-key-hero_glass_generate::before, .st-key-hero_glass_classifier::before {
        content: "";
        position: absolute;
        width: 22rem;
        height: 22rem;
        top: -14rem;
        right: -5rem;
        border-radius: 50%;
        background: radial-gradient(circle, rgba(244, 214, 106, .38), transparent 68%);
        pointer-events: none;
    }
    .st-key-hero_logo_generate, .st-key-hero_logo_classifier {
        padding: .2rem;
        border: 1px solid rgba(255, 235, 174, .22);
        border-radius: 18px;
        background: rgba(2, 2, 3, .42);
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .14), 0 18px 38px rgba(0, 0, 0, .34);
    }
    .hero-title { margin: .2rem 0 .35rem; color: var(--jay-warm); font-size: clamp(1.85rem, 3vw, 2.7rem); line-height: 1.08; letter-spacing: -.04em; font-weight: 780; }
    .hero-copy { margin: 0; max-width: 680px; color: #D7D2C8; font-size: .94rem; line-height: 1.45; }
    .eyebrow { color: var(--jay-gold-light); font-size: .68rem; text-transform: uppercase; letter-spacing: .15em; font-weight: 800; }
    .hero-chip { display: inline-flex; margin-top: .65rem; padding: .32rem .65rem; border: 1px solid rgba(244, 214, 106, .24); border-radius: 999px; background: rgba(215, 169, 25, .10); color: #F4D66A; font-size: .69rem; font-weight: 700; letter-spacing: .03em; }
    div[data-testid="stMetric"],
    [data-testid="stVerticalBlockBorderWrapper"],
    [data-testid="stExpander"],
    [data-testid="stAlert"] {
        border: 1px solid var(--jay-line) !important;
        border-radius: 22px !important;
        background: linear-gradient(145deg, rgba(255, 255, 255, .105), rgba(255, 255, 255, .035)) !important;
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .14), 0 14px 38px rgba(0, 0, 0, .20) !important;
        backdrop-filter: blur(22px) saturate(145%);
        -webkit-backdrop-filter: blur(22px) saturate(145%);
    }
    div[data-testid="stMetric"] { padding: 18px 20px; }
    div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: var(--jay-gold-light); letter-spacing: -.025em; }
    div[data-testid="stFileUploader"] {
        padding: .55rem .75rem;
        border: 1px solid rgba(255, 235, 174, .16);
        border-radius: 22px;
        background: rgba(255, 255, 255, .055);
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .10);
        backdrop-filter: blur(18px);
        -webkit-backdrop-filter: blur(18px);
    }
    div[data-testid="stFileUploader"] section { border-color: rgba(244, 214, 106, .22); background: rgba(8, 9, 13, .24); }
    [data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="base-input"] {
        background: rgba(255, 255, 255, .07) !important;
        border-color: rgba(255, 235, 174, .19) !important;
        backdrop-filter: blur(16px);
    }
    .stButton > button, .stDownloadButton > button { min-height: 2.8rem; border-radius: 999px; font-weight: 800; letter-spacing: .01em; transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease; }
    .stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-1px); border-color: var(--jay-gold-light); box-shadow: 0 12px 26px rgba(215, 169, 25, .18); }
    .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { color: #17130A; background: linear-gradient(135deg, #F4D66A, #C69712); border: 1px solid #F7DF82; }
    .st-key-main_tabs [data-baseweb="tab-list"] {
        display: inline-flex;
        width: max-content;
        max-width: 100%;
        gap: .35rem;
        padding: .35rem;
        border: 1px solid rgba(255, 235, 174, .20);
        border-radius: 999px;
        background: rgba(255, 255, 255, .07);
        box-shadow: inset 0 1px 0 rgba(255,255,255,.12), 0 12px 30px rgba(0,0,0,.20);
        backdrop-filter: blur(18px) saturate(145%);
        overflow-x: auto;
    }
    .st-key-main_tabs [data-baseweb="tab-border"], .st-key-main_tabs [data-baseweb="tab-highlight"] { display: none; }
    .st-key-main_tabs button[data-baseweb="tab"] {
        flex: 0 0 auto;
        min-height: 2.65rem;
        border-radius: 999px;
        padding: .65rem 1.05rem;
        color: #F7F2E6;
        font-size: .94rem;
        font-weight: 750;
        white-space: nowrap;
        opacity: 1;
        transition: background .22s ease, color .22s ease, box-shadow .22s ease, transform .22s ease;
    }
    .st-key-main_tabs button[data-baseweb="tab"]:hover { background: rgba(244, 214, 106, .13); color: #F4D66A; transform: translateY(-1px); }
    .st-key-main_tabs button[data-baseweb="tab"][aria-selected="true"] { color: #17130A; background: linear-gradient(135deg, #F4D66A, #C69712); box-shadow: 0 8px 20px rgba(215,169,25,.18); }
    .st-key-main_tabs [role="tabpanel"]:not([hidden]) { animation: tab-reveal .28s ease-out both; }
    @keyframes tab-reveal {
        from { opacity: .35; transform: translateY(8px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @media (prefers-reduced-motion: reduce) {
        .st-key-main_tabs button[data-baseweb="tab"] { transition: none; }
        .st-key-main_tabs [role="tabpanel"] { animation: none !important; }
    }
    [data-testid="stProgress"] > div > div { background: linear-gradient(90deg, #B9890B, #F4D66A); }
    [data-testid="stDataFrame"] { overflow: hidden; border: 1px solid rgba(255, 235, 174, .16); border-radius: 20px; box-shadow: 0 18px 44px rgba(0,0,0,.20); }
    h1, h2, h3, h4 { letter-spacing: -.025em; }
    h2, h3 { color: var(--jay-warm); }
    hr { border-color: rgba(244, 214, 106, .12); }
    .app-version-badge {
        position: fixed;
        top: .95rem;
        left: 10rem;
        right: auto;
        z-index: 999999;
        padding: .28rem .72rem;
        border: 1px solid rgba(244, 214, 106, .30);
        border-radius: 999px;
        background: rgba(16, 16, 18, .66);
        color: var(--jay-gold-light);
        font-size: .78rem;
        font-weight: 700;
        letter-spacing: .02em;
        line-height: 1.25;
        box-shadow: inset 0 1px 0 rgba(255,255,255,.12), 0 6px 20px rgba(0,0,0,.26);
        backdrop-filter: blur(18px) saturate(150%);
        -webkit-backdrop-filter: blur(18px) saturate(150%);
    }
    @media (max-width: 640px) {
        .app-version-badge { top: .9rem; left: 9.25rem; font-size: .72rem; }
        .block-container { padding-top: 3rem; }
        .st-key-hero_glass_generate, .st-key-hero_glass_classifier { padding: 1rem; border-radius: 22px; }
        .hero-title { font-size: 1.9rem; }
        .st-key-main_tabs button[data-baseweb="tab"] { padding: .55rem .8rem; font-size: .86rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def render_hero(scope: str):
    if scope == "classifier":
        hero_title = "Contractor-wise Overall Classifier Report"
        hero_copy = "Explore each contractor's production plans and material requirements across Shift A and Shift B."
    else:
        hero_title = "Contractor-wise Printing Issue Report"
        hero_copy = "Turn Shift A and Shift B workbooks into a validated, contractor-ready PCS and KG issue report."
    with st.container(key=f"hero_glass_{scope}"):
        hero_logo, hero_content = st.columns([1, 4.8], vertical_alignment="center", gap="medium")
        with hero_logo:
            with st.container(key=f"hero_logo_{scope}"):
                st.image(str(LOGO_PATH), width="stretch")
        with hero_content:
            st.markdown(
                f"""
                <div class="eyebrow">JAY · Production intelligence</div>
                <div class="hero-title">{hero_title}</div>
                <p class="hero-copy">{hero_copy}</p>
                <div class="hero-chip">PRECISION · CONTROL · AUDIT READY</div>
                """,
                unsafe_allow_html=True,
            )

generate_tab, classifier_tab = st.tabs([
    "📊 Printing issue report",
    "🗂️ Overall classifier report",
], key="main_tabs", on_change="rerun")

st.session_state.setdefault("report_date", date.today())
st.session_state.setdefault("default_allowance", 3.0)
report_date = st.session_state["report_date"]
default_allowance = st.session_state["default_allowance"]

with st.sidebar:
    st.markdown(f'<div class="app-version-badge">Version {APP_VERSION}</div>', unsafe_allow_html=True)
    if classifier_tab.open:
        st.markdown("## Classifier report")
        st.caption("Export all contractors from both shift workbooks.")
        st.markdown("**Source workbooks**")
        st.caption("Upload Shift A and Shift B in this tab, or use the files already uploaded in Printing issue report.")
        st.markdown("**Included sections**")
        st.caption("Production plans, printing, packing, speciality tea and black tea.")
        st.markdown("---")
        st.markdown("**Data handling**")
        st.caption("Quantities stay as supplied. No printing master, report date, allowance or KG conversion is used.")
    else:
        st.markdown("## Report controls")
        st.caption("These values are shown in the generated report.")
        report_date = st.date_input("Report date", key="report_date", persist_state="session", format="DD-MM-YYYY")
        default_allowance = st.number_input("Default allowance (%)", key="default_allowance", persist_state="session", min_value=0.0, max_value=99.99, step=0.1, help="Applied only when the master Allowance field is blank.")
        st.markdown("---")
        st.markdown("**Calculation source**")
        st.caption("Request Qty → Item master → Machine weights → KG → allowance")
        st.markdown("**Safety rule**")
        st.caption("Incomplete reports cannot be downloaded.")

with generate_tab:
    render_hero("generate")
    st.subheader("Upload the three source workbooks")
    st.caption("1. Upload the printing master and both shift workbooks. 2. Generate the report. 3. Download it.")
    master_file = st.file_uploader("Permanent printing master", type=["xlsx"], key="master", help="Contains item name, machine line, PCS/KG, roll weight and core/tare weight.")
    col_a, col_b = st.columns(2)
    with col_a:
        shift_a_file = st.file_uploader("Shift A workbook", type=["xlsx"], key="shift_a")
    with col_b:
        shift_b_file = st.file_uploader("Shift B workbook", type=["xlsx"], key="shift_b")

    ready_count = sum(file is not None for file in [master_file, shift_a_file, shift_b_file])
    st.progress(ready_count / 3, text=f"{ready_count} of 3 required workbooks selected")
    if st.button("Validate and generate report", type="primary", width="stretch"):
        if ready_count < 3:
            st.error("Select the master, Shift A and Shift B workbooks.")
        else:
            with printing_overlay("Reading workbooks and validating the printing issue report"):
                try:
                    mapping_config = supabase_mapping_config()
                    mappings = read_name_mappings(*mapping_config) if mapping_config else None
                    result = build_report(master_file, shift_a_file, shift_b_file, default_allowance, mappings)
                    st.session_state["result"] = result
                except Exception as exc:
                    st.session_state.pop("result", None)
                    st.error(f"The report could not be generated: {exc}")

    result = st.session_state.get("result")
    if result:
        st.markdown("---")
        if result.errors:
            st.error("Validation failed. The report download is blocked until these issues are corrected.")
            for error in result.errors:
                st.write(f"• {error}")
        else:
            st.success("Report ready to download.")
        if result.warnings:
            with st.expander(f"Review {len(result.warnings)} warning(s)", expanded=True):
                for warning in result.warnings:
                    st.write(f"• {warning}")

        stats = result.source_stats
        with st.container(horizontal=True, horizontal_alignment="center"):
            st.metric("Report rows", f"{stats['Output rows']:,}", border=True)

        shift_a_summary, shift_b_summary = st.columns(2, gap="large", border=True)
        with shift_a_summary:
            st.markdown("#### Shift A issue KG")
            st.metric("Total issue KG", f"{stats['Shift A KG']:,.0f}")
            st.metric("Total requested TAG (KG)", f"{stats['Shift A TAG KG']:,.0f}")
            st.metric("Total requested ENV (KG)", f"{stats['Shift A ENV KG']:,.0f}")
        with shift_b_summary:
            st.markdown("#### Shift B issue KG")
            st.metric("Total issue KG", f"{stats['Shift B KG']:,.0f}")
            st.metric("Total requested TAG (KG)", f"{stats['Shift B TAG KG']:,.0f}")
            st.metric("Total requested ENV (KG)", f"{stats['Shift B ENV KG']:,.0f}")

        overall_display = result.report.copy()
        overall_display["Allowance (%)"] = overall_display["Allowance"] * 100
        overall_display = overall_display.drop(columns=["Allowance"])

        st.subheader("Overall contractor report (main report)")
        st.caption("This is the complete report for every contractor and every calculated field.")
        kg_column_config = {
            column: st.column_config.NumberColumn(format="%,d")
            for column in ["Shift A KG", "Shift B KG", "Total KG"]
        }
        st.dataframe(
            overall_display,
            column_config=kg_column_config,
            width="stretch",
            hide_index=True,
            height=460,
        )

        if result.valid:
            excel = create_excel_report(result, report_date)
            st.download_button(
                "Download overall report",
                data=excel,
                file_name=f"Contractor_Printing_Issue_Report_{report_date:%d-%m-%Y}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )

        st.subheader("Contractor-specific report")
        st.caption(
            "Every contractor's issue sheet, grouped and totalled, ready to print on A4. "
            "Issued in PCS and Issued in KG are blank entry columns for the actual quantity issued. "
            "Change the report date in the sidebar to update the date shown here."
        )
        print_html = build_contractor_print_html(result, report_date)
        components.html(print_html, height=700, scrolling=True)

with classifier_tab:
    render_hero("classifier")
    st.subheader("Create the contractor workbook")
    st.caption("1. Upload the Shift A and Shift B workbooks. 2. Generate the report. 3. Download it.")
    classifier_upload_a, classifier_upload_b = st.columns(2)
    with classifier_upload_a:
        classifier_a_file = st.file_uploader("Shift A workbook for classifier", type=["xlsx"], key="classifier_shift_a")
    with classifier_upload_b:
        classifier_b_file = st.file_uploader("Shift B workbook for classifier", type=["xlsx"], key="classifier_shift_b")

    source_a = classifier_a_file or shift_a_file
    source_b = classifier_b_file or shift_b_file
    classifier_ready = source_a is not None and source_b is not None
    if not classifier_ready:
        st.info("Upload both Shift A and Shift B workbooks to create the contractor workbook.")
    generate_classifier = st.button(
        "Generate overall classifier report",
        type="primary",
        width="stretch",
        disabled=not classifier_ready,
        key="generate_classifier_report",
    )
    if classifier_ready:
        source_fingerprint = (
            sha256(source_a.getvalue()).hexdigest(),
            sha256(source_b.getvalue()).hexdigest(),
        )
        classifier_result = st.session_state.get("classifier_result")
        if classifier_result and classifier_result["source_fingerprint"] != source_fingerprint:
            st.session_state.pop("classifier_result", None)
            classifier_result = None
    else:
        classifier_result = None

    if generate_classifier:
        try:
            with printing_overlay("Preparing every contractor worksheet from both shifts"):
                classifier_shifts = [
                    read_classifier_shift(source_a.getvalue(), "A"),
                    read_classifier_shift(source_b.getvalue(), "B"),
                ]
                contractor_names = list_classifier_contractors(classifier_shifts)
                classifier_excel = (
                    create_all_classifier_excel_report(classifier_shifts)
                    if contractor_names else None
                )
                section_counts = {
                    section_name: sum(len(shift.sections[section_name]) for shift in classifier_shifts)
                    for section_name in classifier_shifts[0].sections
                }
                classifier_result = {
                    "source_fingerprint": source_fingerprint,
                    "contractor_names": contractor_names,
                    "section_counts": section_counts,
                    "excel": classifier_excel,
                }
                st.session_state["classifier_result"] = classifier_result
        except Exception as exc:
            st.session_state.pop("classifier_result", None)
            classifier_result = None
            st.error(f"The classifier report could not be prepared: {exc}")

    if classifier_result:
        contractor_names = classifier_result["contractor_names"]
        if not contractor_names:
            st.warning("No contractor sections were found in the uploaded workbooks.")
        else:
            section_counts = classifier_result["section_counts"]
            with st.container(border=True):
                st.caption("WORKBOOK OVERVIEW")
                st.metric("Contractor worksheets", f"{len(contractor_names):,}")
                st.metric("Total source rows", f"{sum(section_counts.values()):,}")
                for section_name, count in section_counts.items():
                    st.markdown(
                        f"**{section_name}** · {count:,} rows across Shift A and Shift B"
                    )
                st.caption(
                    "Download includes one worksheet per contractor, each with all five types and both shifts. "
                    "In Excel, use the + / − controls above the columns to collapse production or material details."
                )

            st.download_button(
                "Download overall classifier report",
                data=classifier_result["excel"],
                file_name="Overall_Contractor_Classifier.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )
