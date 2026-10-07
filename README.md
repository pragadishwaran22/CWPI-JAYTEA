# MJIPL Contractor-wise Printing Issue Report Prototype

This Streamlit application reads the daily Shift A and Shift B Auto Material Slip workbooks, matches each exact cleaned Printing Item Name to the permanent master, calculates PCS and KG by contractor and shift, validates the data, and creates an Excel report.

The Printing issue report tab includes a short explanation of the workflow. The Overall classifier report tab shows the original shift data by contractor.

## Run on Windows

1. Install Python 3.11 or newer.
2. Open this folder and double-click `start_app.bat`.

The first start installs the required packages. Later starts open much faster. The launcher opens `http://127.0.0.1:8501`. Keep the command window open while using the app.

If the browser does not open, enter `http://127.0.0.1:8501` directly in Chrome or Edge. If that address does not load, read the error shown in the command window.

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
   streamlit run app.py --server.address 127.0.0.1 --server.port 8501
   ```

6. Open `http://127.0.0.1:8501` if it does not open automatically.

## Daily use

Upload:

- `Printing_Master_With_Machine_Codes.xlsx`
- the complete Shift A Auto Material Slip workbook
- the complete Shift B Auto Material Slip workbook

Choose the report date and allowance, then click **Validate and generate report**. Serious mapping or weight errors block the download. Warnings remain visible for review. Rows whose Request Qty is zero in both shifts are omitted because no material needs to be issued.

The summary separates requested KG within each shift by Printing Item Name prefix: names beginning with `TAG` contribute to that shift's **Total requested TAG (KG)**, and names beginning with `ENV` contribute to that shift's **Total requested ENV (KG)**.

Daily Auto Material Slip Item IDs and permanent-master item codes come from different numbering systems. Matching uses only normalized item names, never IDs or fuzzy similarity. An exact name in the master wins. Otherwise, a single M4-name-to-MJP-name pair from Supabase is used automatically. If an M4 name has multiple MJP targets, exactly one pair must be approved; otherwise the report is blocked. Missing or non-master targets also block the download. Both ID systems remain in the output for auditing, but do not drive matching.

## M4/MJP name mapping setup

1. Run [the table setup SQL](supabase/printing_item_name_mapping.sql) in your Supabase project's SQL editor.
2. Run `python prepare_mapping_import.py "path/to/MJPvsM4ItemMappingDtls.xls" "outputs/printing_name_mapping_import.csv"`. Import the resulting CSV into the Supabase table. It contains only the two names and an `approved` flag—no IDs. Single-target names work automatically regardless of the flag. For each M4 name with multiple possible targets, approve exactly one intended pair. The selected target name must exist in the printing master.
3. Add this to a local `.streamlit/secrets.toml` file (already gitignored), and enter the same values in Streamlit Cloud's app secrets for deployment:

   ```toml
   [supabase]
   url = "https://YOUR-PROJECT.supabase.co"
   publishable_key = "sb_publishable_YOUR_KEY"
   ```

4. Restart the local Streamlit server after adding the secrets file. Generate a report and check the displayed mapping status and validation messages.

All name pairs and approval flags are readable through the app's publishable key so it can detect ambiguous names. This is suitable only if these names may be read by anyone with that public key; use authenticated access instead if the mapping is confidential. The app reads from Supabase but never writes workbook contents or mapping rows to it. Mapping reads are cached for 15 minutes; wait or restart the app after edits if you need an immediate refresh. If the table already exists, rerun the setup SQL to replace its old approved-only read policy before using this version.

## Overall classifier report

Open **Overall classifier report** and upload the Shift A and Shift B workbooks. If those files are already selected in **Printing issue report**, the classifier uses them automatically. The printing master is not needed for this view.

The app shows a boxed overview of all contractors and source-row counts, with no contractor selector, table preview, or row-search control. Combined generic material is excluded. Quantities and other source values are exported without conversion or aggregation. **Download overall classifier report** saves one workbook with a separate filterable worksheet for every contractor. Each sheet contains all five types and both shifts; Type, Shift, and Date columns identify each row, and the contractor name is highlighted in a gold header. Excel's outline controls let you manually collapse the production-plan columns or the material columns independently; filtering Type does not collapse columns automatically.

## Calculation

For each contractor, printing item and shift:

```text
Base KG = Request PCS / PCS per KG
Core/Tare KG = (Base KG / Roll Weight) × Core/Tare Weight
Shift KG = round((Base KG + Core/Tare KG) × (1 + Allowance))
Total KG = Rounded Shift A KG + Rounded Shift B KG
```

If Roll Weight and Core/Tare are both explicitly zero, Core/Tare KG is treated as zero and the app shows a warning. Blank weights block report generation.

## Version

The version badge beside the Streamlit toolbar is controlled by `APP_VERSION` in `version.py`. Update this value whenever a feature release should receive a new version so the local app and the GitHub-backed Streamlit Cloud app can be compared easily.
