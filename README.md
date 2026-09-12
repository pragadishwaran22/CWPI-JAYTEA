# MJIPL Contractor-wise Printing Issue Report Prototype

This Streamlit application reads the daily Shift A and Shift B Auto Material Slip workbooks, matches each exact cleaned Printing Item Name to the permanent master, calculates PCS and KG by contractor and shift, validates the data, and creates an Excel report.

The prototype also contains **How it works** and **Explain the project** screens so the workflow can be demonstrated to teammates without opening the source code.

## Run on Windows

1. Install Python 3.11 or newer.
2. Open this folder and double-click `start_app.bat`.

The first start installs the required packages. Later starts open much faster. The launcher now opens `http://localhost:8501` explicitly. Keep the command window open while using the app.

If the browser does not open, enter `http://localhost:8501` directly in Chrome or Edge. If that address does not load, read the error shown in the command window.

If you prefer to run the commands manually:

3. Create an environment:

   ```bat
   python -m venv .venv
   .venv\Scripts\activate
   ```

4. Install the required packages:

   ```bat
   pip install -r requirements.txt
   ```

5. Start the app:

   ```bat
   streamlit run app.py
   ```

6. Open `http://localhost:8501` if it does not open automatically.

## Daily use

Upload:

- `Printing_Master_With_Machine_Codes.xlsx`
- the complete Shift A Auto Material Slip workbook
- the complete Shift B Auto Material Slip workbook

Choose the report date and allowance, then click **Validate and generate report**. Serious mapping or weight errors block the download. Warnings remain visible for review. Rows whose Request Qty is zero in both shifts are omitted because no material needs to be issued.

The summary separates requested KG within each shift by Printing Item Name prefix: names beginning with `TAG` contribute to that shift's **Total requested TAG (KG)**, and names beginning with `ENV` contribute to that shift's **Total requested ENV (KG)**.

Daily Auto Material Slip Item IDs and permanent-master item codes come from different numbering systems. The application therefore matches the exact normalized Printing Item Name. It never silently accepts a fuzzy match. Both the daily ID and master code are included in the output for auditing.

## Calculation

For each contractor, printing item and shift:

```text
Base KG = Request PCS / PCS per KG
Core/Tare KG = (Base KG / Roll Weight) × Core/Tare Weight
Shift KG = round((Base KG + Core/Tare KG) × (1 + Allowance))
Total KG = Rounded Shift A KG + Rounded Shift B KG
```

If Roll Weight and Core/Tare are both explicitly zero, Core/Tare KG is treated as zero and the app shows a warning. Blank weights block report generation.
