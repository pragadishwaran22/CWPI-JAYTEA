"""Read approved M4-to-MJP name pairs from Supabase's REST API."""

from __future__ import annotations

import json
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlparse, urlunparse
from urllib.request import Request, urlopen

import pandas as pd


PAGE_SIZE = 500


def fetch_approved_name_mappings(project_url: str, publishable_key: str) -> pd.DataFrame:
    """Fetch every approved pair; never infer a match from an ID or partial name."""
    parsed = urlparse(project_url)
    if (parsed.scheme != "https" or not parsed.netloc or parsed.username or parsed.password
            or parsed.path.rstrip("/") not in ("", "/rest/v1") or parsed.query or parsed.fragment):
        raise ValueError("The Supabase project URL must be a valid HTTPS URL.")
    if not publishable_key:
        raise ValueError("A Supabase publishable key is required.")

    rows: list[dict[str, str]] = []
    project_base = urlunparse((parsed.scheme, parsed.netloc, "", "", "", ""))
    endpoint = project_base + "/rest/v1/printing_item_name_mapping"
    query = urlencode({"select": "m4_item_name,mjp_item_name", "approved": "eq.true", "order": "m4_item_name,mjp_item_name"})
    while True:
        start = len(rows)
        request = Request(
            f"{endpoint}?{query}",
            headers={
                "apikey": publishable_key,
                "Range-Unit": "items",
                "Range": f"{start}-{start + PAGE_SIZE - 1}",
                "Accept": "application/json",
                "Prefer": "count=exact",
            },
        )
        try:
            with urlopen(request, timeout=15) as response:
                page = json.load(response)
                content_range = response.headers.get("Content-Range", "")
        except (HTTPError, URLError, TimeoutError) as exc:
            raise RuntimeError("Could not read the approved name mappings from Supabase. Check the URL, key, table and read policy.") from exc
        if not isinstance(page, list):
            raise RuntimeError("Supabase returned an unexpected name-mapping response.")
        if not page:
            break
        rows.extend(page)
        total_text = content_range.rsplit("/", 1)[-1]
        if total_text.isdigit() and len(rows) >= int(total_text):
            break
        if not content_range and len(page) < PAGE_SIZE:
            break
    return pd.DataFrame(rows, columns=["m4_item_name", "mjp_item_name"])
