"""Prepare a names-only Supabase CSV from the supplied M4/MJP workbook.

Usage: python prepare_mapping_import.py input.xls output.csv
The source file may have a .xls suffix while actually containing XLSX data.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


def clean_name(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u00a0", " ")).strip()


def prepare(source: Path, destination: Path) -> tuple[int, int]:
    with source.open("rb") as workbook:
        frame = pd.read_excel(workbook, sheet_name="MJPvsM4Mapping", engine="openpyxl", dtype=str)
    required = {"M4 Item Name", "MJP Item Name"}
    if not required.issubset(frame.columns):
        raise ValueError("Mapping sheet must include M4 Item Name and MJP Item Name.")
    pairs = frame[["M4 Item Name", "MJP Item Name"]].rename(columns={
        "M4 Item Name": "m4_item_name", "MJP Item Name": "mjp_item_name",
    })
    pairs = pairs.map(clean_name)
    pairs = pairs[(pairs["m4_item_name"] != "") & (pairs["mjp_item_name"] != "")].copy()
    pairs["source_key"] = pairs["m4_item_name"].str.upper()
    pairs["target_key"] = pairs["mjp_item_name"].str.upper()
    pairs = pairs.drop_duplicates(["source_key", "target_key"])
    target_counts = pairs.groupby("source_key")["target_key"].transform("nunique")
    pairs["approved"] = target_counts.eq(1)
    ambiguous_source_count = int(pairs.loc[~pairs["approved"], "source_key"].nunique())
    pairs = pairs[["m4_item_name", "mjp_item_name", "approved"]].sort_values("m4_item_name")
    destination.parent.mkdir(parents=True, exist_ok=True)
    pairs.to_csv(destination, index=False, encoding="utf-8-sig")
    return len(pairs), ambiguous_source_count


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    args = parser.parse_args()
    count, ambiguous = prepare(args.source, args.destination)
    print(f"Prepared {count} name pairs; {ambiguous} M4 names have multiple targets and remain unapproved.")
