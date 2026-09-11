from pathlib import Path

from report_engine import build_report, create_excel_report


ROOT = Path(__file__).resolve().parents[1]
result = build_report(
    ROOT / "upload/01-Printing_Master_With_Machine_Codes.xlsx",
    ROOT / "upload/AutomaterialSlip_MDK CBE_05-09-2026A.xlsx",
    ROOT / "upload/AutomaterialSlip_MDK CBE_04-09-2026B.xlsx",
    3.0,
)
print("VALID:", result.valid)
print("ERRORS:", result.errors)
print("WARNINGS:", result.warnings)
print("STATS:", result.source_stats)
if result.valid:
    out = ROOT / "outputs/streamlit_prototype_test_report.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(create_excel_report(result))
    print("OUTPUT:", out)
