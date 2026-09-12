from datetime import date
from pathlib import Path

import streamlit as st

from report_engine import (
    build_report,
    contractor_report_frame,
    create_contractor_excel_report,
    create_excel_report,
)
from version import APP_VERSION


LOGO_PATH = Path(__file__).parent / "assets" / "jay-logo.png"


st.set_page_config(
    page_title="JAY Printing Issue Report",
    page_icon=str(LOGO_PATH),
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={"About": "### JAY Printing Issue Report\nValidated contractor-wise PCS and KG reporting."},
)
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
        background: rgba(8, 9, 13, .68);
        border-bottom: 1px solid rgba(255, 235, 174, .09);
        backdrop-filter: blur(24px) saturate(145%);
        -webkit-backdrop-filter: blur(24px) saturate(145%);
    }
    .block-container { max-width: 1320px; padding-top: 2rem; padding-bottom: 4rem; }
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
    .st-key-hero_glass {
        position: relative;
        overflow: hidden;
        padding: clamp(1.5rem, 4vw, 3rem);
        margin-bottom: 1.5rem;
        border: 1px solid rgba(255, 239, 187, .22);
        border-radius: 32px;
        background: linear-gradient(135deg, rgba(255, 255, 255, .13), rgba(255, 255, 255, .045));
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .24), inset 0 -1px 0 rgba(255, 211, 76, .06), 0 24px 70px rgba(0, 0, 0, .38);
        backdrop-filter: blur(30px) saturate(165%);
        -webkit-backdrop-filter: blur(30px) saturate(165%);
    }
    .st-key-hero_glass::before {
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
    .st-key-hero_logo {
        padding: 1.15rem;
        border: 1px solid rgba(255, 235, 174, .22);
        border-radius: 28px;
        background: rgba(2, 2, 3, .42);
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .14), 0 18px 38px rgba(0, 0, 0, .34);
    }
    .hero-title { margin: .35rem 0 .55rem; color: var(--jay-warm); font-size: clamp(2rem, 4.4vw, 3.65rem); line-height: 1.04; letter-spacing: -.045em; font-weight: 780; }
    .hero-copy { margin: 0; max-width: 760px; color: #D7D2C8; font-size: clamp(1rem, 1.5vw, 1.15rem); line-height: 1.65; }
    .eyebrow { color: var(--jay-gold-light); font-size: .76rem; text-transform: uppercase; letter-spacing: .18em; font-weight: 800; }
    .hero-chip { display: inline-flex; margin-top: 1rem; padding: .4rem .78rem; border: 1px solid rgba(244, 214, 106, .24); border-radius: 999px; background: rgba(215, 169, 25, .10); color: #F4D66A; font-size: .78rem; font-weight: 700; letter-spacing: .04em; }
    .step-card, .info-card, .formula {
        border: 1px solid var(--jay-line);
        border-radius: 22px;
        background: linear-gradient(145deg, rgba(255, 255, 255, .105), rgba(255, 255, 255, .035));
        box-shadow: inset 0 1px 0 rgba(255, 255, 255, .16), 0 16px 40px rgba(0, 0, 0, .22);
        backdrop-filter: blur(22px) saturate(145%);
        -webkit-backdrop-filter: blur(22px) saturate(145%);
    }
    .step-card, .info-card { padding: 22px; min-height: 168px; }
    .step-number { width: 38px; height: 38px; border-radius: 13px; display: inline-flex; align-items: center; justify-content: center; background: linear-gradient(145deg, #F4D66A, #B9890B); color: #17130A; font-weight: 900; margin-bottom: 14px; box-shadow: 0 8px 22px rgba(215,169,25,.24); }
    .step-card h4, .info-card h4 { color: var(--jay-warm); margin: 0 0 8px; }
    .step-card p, .info-card p { color: #BDB9B0; margin: 0; line-height: 1.55; }
    .formula { padding: 20px 24px; border-left: 3px solid var(--jay-gold); color: #E9E3D7; font-size: 1rem; line-height: 1.8; }
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
    button[data-baseweb="tab"] { border-radius: 999px; padding: .55rem 1rem; color: #BDB9B0; }
    button[data-baseweb="tab"][aria-selected="true"] { color: #17130A; background: linear-gradient(135deg, #F4D66A, #C69712); box-shadow: 0 10px 24px rgba(215,169,25,.18); }
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
        .block-container { padding-top: 1.2rem; }
        .st-key-hero_glass { padding: 1.25rem; border-radius: 24px; }
        .hero-title { font-size: 2.15rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

with st.container(key="hero_glass"):
    hero_logo, hero_content = st.columns([1, 4.2], vertical_alignment="center", gap="large")
    with hero_logo:
        with st.container(key="hero_logo"):
            st.image(str(LOGO_PATH), width="stretch")
    with hero_content:
        st.markdown(
            """
            <div class="eyebrow">JAY · Production intelligence</div>
            <div class="hero-title">Contractor-wise<br>Printing Issue Report</div>
            <p class="hero-copy">Transform daily Shift A and Shift B workbooks into a validated, contractor-ready PCS and KG issue report—accurately and in one refined workflow.</p>
            <div class="hero-chip">PRECISION · CONTROL · AUDIT READY</div>
            """,
            unsafe_allow_html=True,
        )

with st.sidebar:
    st.markdown(f'<div class="app-version-badge">Version {APP_VERSION}</div>', unsafe_allow_html=True)
    st.markdown("## Report controls")
    st.caption("These values are shown in the generated report.")
    report_date = st.date_input("Report date", value=date.today(), format="DD-MM-YYYY")
    default_allowance = st.number_input("Default allowance (%)", min_value=0.0, max_value=99.99, value=3.0, step=0.1, help="Applied only when the master Allowance field is blank.")
    st.markdown("---")
    st.markdown("**Calculation source**")
    st.caption("Request Qty → Item master → Machine weights → KG → allowance")
    st.markdown("**Safety rule**")
    st.caption("Missing or conflicting mappings block the Excel download.")

generate_tab, process_tab, explain_tab = st.tabs(["📊 Generate report", "🔄 How it works", "🗣️ Explain the project"])

with generate_tab:
    st.subheader("Upload the three source workbooks")
    st.caption("The daily workbooks must contain PRINTING MATL REQ. Matching uses the exact cleaned Item Name because the daily Item IDs and master item codes are different systems.")
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
            with st.spinner("Reading sections, matching items and validating calculations…"):
                try:
                    result = build_report(master_file, shift_a_file, shift_b_file, default_allowance)
                    st.session_state["result"] = result
                    st.session_state["report_date"] = report_date
                except Exception as exc:
                    st.session_state.pop("result", None)
                    st.error(f"The workbooks could not be processed: {exc}")

    result = st.session_state.get("result")
    if result:
        st.markdown("---")
        if result.errors:
            st.error("Validation failed. The report download is blocked until these issues are corrected.")
            for error in result.errors:
                st.write(f"• {error}")
        else:
            st.success("Validation passed. Every report item has a usable master mapping.")
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
            excel = create_excel_report(result, st.session_state["report_date"])
            st.download_button(
                "Download overall report",
                data=excel,
                file_name=f"Contractor_Printing_Issue_Report_{st.session_state['report_date']:%d-%m-%Y}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )

        st.subheader("Contractor-specific report")
        contractor_options = ["All contractors"] + sorted(
            overall_display["Contractor Name"].dropna().unique().tolist()
        )
        contractor_filter = st.selectbox("Select contractor for preview and download", contractor_options)
        if contractor_filter == "All contractors":
            st.info("Select a contractor to create its concise issue report.")
        else:
            contractor_display = contractor_report_frame(result, contractor_filter)
            st.caption(
                "Issued in PCS and Issued in KG are blank entry columns for the actual quantity issued."
            )
            st.dataframe(
                contractor_display,
                column_config=kg_column_config,
                width="stretch",
                hide_index=True,
                height=420,
            )
            if result.valid:
                contractor_excel = create_contractor_excel_report(
                    result,
                    contractor_filter,
                    st.session_state["report_date"],
                )
                safe_contractor = "_".join(contractor_filter.split())
                st.download_button(
                    f"Download {contractor_filter} report",
                    data=contractor_excel,
                    file_name=f"{safe_contractor}_Printing_Issue_Report_{st.session_state['report_date']:%d-%m-%Y}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    width="stretch",
                )

with process_tab:
    st.subheader("What happens after you upload the files")
    cards = [
        ("1", "Read daily blocks", "Find each contractor section and read Item Name and Request Qty from PRINTING MATL REQ."),
        ("2", "Match the master", "Join the exact cleaned Item Name to Machine Line, PCS/KG, Roll Weight and Core/Tare."),
        ("3", "Calculate by shift", "Calculate Shift A and Shift B KG independently, then round each result to the nearest whole KG."),
        ("4", "Validate and export", "Stop on missing mappings; otherwise generate Excel with a validation sheet."),
    ]
    cols = st.columns(4)
    for col, (number, title, body) in zip(cols, cards):
        col.markdown(f'<div class="step-card"><div class="step-number">{number}</div><h4>{title}</h4><p>{body}</p></div>', unsafe_allow_html=True)
    st.markdown("### Calculation used")
    st.markdown("""<div class="formula"><b>Base KG</b> = Request PCS ÷ PCS per KG<br><b>Core/Tare KG</b> = (Base KG ÷ Roll Weight) × Core/Tare Weight<br><b>Final KG</b> = round((Base KG + Core/Tare KG) × (1 + Allowance))<br><b>Total KG</b> = Rounded Shift A KG + Rounded Shift B KG</div>""", unsafe_allow_html=True)
    st.markdown("### Example: Africa Choice C250")
    st.dataframe([{"Request PCS": "244,800", "PCS per KG": "1,740", "Roll Weight": "13 kg", "Core/Tare": "0.25 kg", "Allowance": "3%", "Final issue": "148 kg"}], hide_index=True, width="stretch")

with explain_tab:
    st.subheader("How to present this project to your teammate")
    left, right = st.columns(2)
    with left:
        st.markdown('<div class="info-card"><h4>Business problem</h4><p>The contractor-wise issue report is created every day from repeated Excel sections. Manual lookup and PCS-to-KG conversion take time and can introduce wrong machine or weight values.</p></div>', unsafe_allow_html=True)
    with right:
        st.markdown('<div class="info-card"><h4>Proposed solution</h4><p>A local Streamlit app reads both shifts, validates every exact Item Name against a permanent master, calculates issue weight and produces an auditable Excel report.</p></div>', unsafe_allow_html=True)
    st.markdown("### Technology and responsibility")
    st.dataframe([
        {"Technology": "Streamlit", "Role": "Upload screen, status messages, preview, filters and download"},
        {"Technology": "Python + Pandas", "Role": "Read repeated contractor blocks, clean text, group quantities and join data"},
        {"Technology": "OpenPyXL", "Role": "Create the formatted Excel report and validation sheet"},
        {"Data source": "Permanent master", "Role": "Control Machine Line, PCS/KG, roll weight, core/tare and allowance"},
    ], hide_index=True, width="stretch")
    st.markdown("### Key reliability decisions")
    st.markdown("""
    - Daily Item IDs and master codes belong to different numbering systems, so the app matches an **exact cleaned Item Name**.
    - No fuzzy match is automatically accepted.
    - Missing machine, conversion or weight values block the download.
    - Contractor and shift quantities remain separate until the final totals.
    - The generated workbook includes a validation sheet for audit.
    """)
