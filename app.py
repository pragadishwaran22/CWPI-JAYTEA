from datetime import date

import streamlit as st

from report_engine import (
    build_report,
    contractor_report_frame,
    create_contractor_excel_report,
    create_excel_report,
)
from version import APP_VERSION


st.set_page_config(page_title="MJIPL Printing Issue Report", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

st.markdown(
    """
    <style>
    .stApp { background: #f5f7fb; }
    .block-container { max-width: 1280px; padding-top: 1.6rem; padding-bottom: 3rem; }
    [data-testid="stSidebar"] { background: #102a43; }
    [data-testid="stSidebar"] * { color: #f8fafc; }
    [data-testid="stSidebar"] input { color: #102a43; }
    .hero { padding: 28px 32px; border-radius: 20px; background: linear-gradient(125deg, #12355b 0%, #176b87 62%, #2a9d8f 100%); color: white; box-shadow: 0 14px 35px rgba(18,53,91,.18); margin-bottom: 22px; }
    .hero h1 { margin: 0 0 8px; font-size: 2.15rem; letter-spacing: -.02em; }
    .hero p { margin: 0; opacity: .88; font-size: 1.02rem; }
    .eyebrow { font-size: .76rem; text-transform: uppercase; letter-spacing: .14em; font-weight: 700; opacity: .75; }
    .step-card, .info-card { background: white; border: 1px solid #e3e8ef; border-radius: 16px; padding: 20px; min-height: 150px; box-shadow: 0 5px 18px rgba(15,23,42,.045); }
    .step-number { width: 34px; height: 34px; border-radius: 10px; display: inline-flex; align-items: center; justify-content: center; background: #dff4f1; color: #146b62; font-weight: 800; margin-bottom: 12px; }
    .step-card h4, .info-card h4 { color: #17365d; margin: 0 0 7px; }
    .step-card p, .info-card p { color: #526273; margin: 0; line-height: 1.52; }
    .formula { padding: 18px 22px; border-radius: 14px; background: #eef5ff; border-left: 5px solid #2f6fb0; color: #17365d; font-size: 1.03rem; }
    div[data-testid="stMetric"] { background: white; border: 1px solid #e3e8ef; padding: 15px 18px; border-radius: 15px; box-shadow: 0 4px 14px rgba(15,23,42,.04); }
    div[data-testid="stFileUploader"] { background: white; border-radius: 14px; padding: 4px 12px; }
    .stButton > button, .stDownloadButton > button { border-radius: 11px; font-weight: 700; }
    .app-version-badge {
        position: fixed;
        top: .48rem;
        right: 7.4rem;
        z-index: 999999;
        padding: .28rem .72rem;
        border: 1px solid #b8c7d9;
        border-radius: 999px;
        background: rgba(255, 255, 255, .96);
        color: #17365d;
        font-size: .78rem;
        font-weight: 700;
        letter-spacing: .02em;
        line-height: 1.25;
        box-shadow: 0 2px 8px rgba(15, 23, 42, .08);
    }
    @media (max-width: 640px) {
        .app-version-badge { right: 5rem; font-size: .72rem; }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(f'<div class="app-version-badge">Version {APP_VERSION}</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="hero">
      <div class="eyebrow">Production planning automation prototype</div>
      <h1>Contractor-wise Printing Issue Report</h1>
      <p>From daily Shift A/B workbooks to a validated PCS and KG issue report in one workflow.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
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
