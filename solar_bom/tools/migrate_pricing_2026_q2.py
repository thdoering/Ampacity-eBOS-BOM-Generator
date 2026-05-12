"""
migrate_pricing_2026_q2.py

One-shot 2026-Q2 pricing migration. Safe to re-run (idempotent).

Reads  : Ampacity Standard Items CSV 2.xlsx (repo root)
Writes : data/pricing_data.json
         data/harness_library.json

Summary of changes:
  - Settings: copper price -> 6.2, tiers extended to 6 (add 6.5)
  - Extenders / whips: full 6-tier prices from spreadsheet
  - First Solar harnesses: full 6-tier prices + 22 new library entries
  - Standard harnesses: full 6-tier prices (2-3 STRING sheet)
  - Fuses: flat prices refreshed from Fuses sheet
"""

import json
import re
import sys
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
REPO_ROOT = Path(__file__).resolve().parent.parent
XLSX_PATH = REPO_ROOT / "Ampacity Standard Items CSV 2.xlsx"
PRICING_JSON = REPO_ROOT / "data" / "pricing_data.json"
HARNESS_JSON = REPO_ROOT / "data" / "harness_library.json"

TIER_COLS = [7, 8, 9, 10, 11, 12]
TIER_KEYS = ["4.0", "4.5", "5.0", "5.5", "6.0", "6.5"]

# Pattern for a harness PN — used to detect paste errors in the Fuses sheet
HARNESS_PN_RE = re.compile(r"^\d+[PN]-\d+D\d+T-\d+$")

# Maps word-form string counts to integers (2-3 STRING sheet uses these)
WORD_COUNT = {
    "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
}

# Parses harness description into (n_strings_raw, polarity, drop_awg, trunk_awg, spacing_ft)
DESC_PATTERN = re.compile(
    r"(\d+|two|three|four|five|six|seven|eight|nine|ten)\s+String,"
    r"\s+(Positive|Negative).*?(\d+)AWG\s+Drops\s+w/(\d+)AWG\s+Trunk,\s+(\d+)'",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def load_json(path: Path) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path: Path, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"  Wrote {path}")


def is_item_number(val) -> bool:
    """Return True when val is a real numeric ITEM row number (not NaN, not a string)."""
    if pd.isna(val):
        return False
    try:
        float(val)
        return True
    except (TypeError, ValueError):
        return False


def get_tier_prices(row) -> dict:
    """Extract up to 6 tier prices from cols 7-12."""
    prices = {}
    for col, key in zip(TIER_COLS, TIER_KEYS):
        try:
            val = row[col]
            if not pd.isna(val) and str(val).strip() != "":
                prices[key] = round(float(val), 2)
        except (KeyError, ValueError, TypeError):
            pass
    return prices


def data_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Filter to rows where column 0 is a real numeric ITEM number."""
    return df[df[0].apply(is_item_number)].copy()


# ---------------------------------------------------------------------------
# Step 1a -- Settings
# ---------------------------------------------------------------------------

def update_settings(pricing: dict) -> None:
    pricing["settings"]["current_copper_price"] = 6.2
    pricing["settings"]["copper_price_tiers"] = [4.0, 4.5, 5.0, 5.5, 6.0, 6.5]
    print("  current_copper_price = 6.2")
    print("  copper_price_tiers = [4.0, 4.5, 5.0, 5.5, 6.0, 6.5]")


# ---------------------------------------------------------------------------
# Step 1b -- Backfill 6.5 tier (and 6.0 for standard harnesses)
# ---------------------------------------------------------------------------

def _add_65(tier: dict) -> bool:
    if "6.5" not in tier and "6.0" in tier and "5.5" in tier:
        delta = round(tier["6.0"] - tier["5.5"], 10)
        tier["6.5"] = round(tier["6.0"] + delta, 2)
        return True
    return False


def _add_60_and_65_standard(tier: dict) -> int:
    added = 0
    if "6.0" not in tier and "5.5" in tier and "5.0" in tier:
        delta = round(tier["5.5"] - tier["5.0"], 10)
        tier["6.0"] = round(tier["5.5"] + delta, 2)
        added += 1
    if _add_65(tier):
        added += 1
    return added


def backfill_tiers(pricing: dict) -> None:
    n_cable = 0
    for section in ("extenders", "whips"):
        for parts in pricing.get(section, {}).values():
            for tier in parts.values():
                if _add_65(tier):
                    n_cable += 1

    n_fs = sum(
        1 for tier in pricing.get("harnesses", {}).get("first_solar", {}).values()
        if _add_65(tier)
    )

    std_tiers_added = sum(
        _add_60_and_65_standard(tier)
        for tier in pricing.get("harnesses", {}).get("standard", {}).values()
    )
    n_std = std_tiers_added // 2

    print(f"  Extrapolated '6.5' on {n_cable} extender/whip entries")
    print(f"  Extrapolated '6.5' on {n_fs} first-solar harness entries")
    print(f"  Extrapolated '6.0'+'6.5' on {n_std} standard harness entries")


# ---------------------------------------------------------------------------
# Step 1d -- Cable sheets (extenders & whips)
# ---------------------------------------------------------------------------

CABLE_SHEETS = {
    "8EXT-Q":     ("extenders", "8_awg",  "EXT", "8"),
    "8EXT 300+":  ("extenders", "8_awg",  "EXT", "8"),
    "10EXT-Q":    ("extenders", "10_awg", "EXT", "10"),
    "10EXT 300+": ("extenders", "10_awg", "EXT", "10"),
    "8WHI-Q":     ("whips",     "8_awg",  "WHI", "8"),
    "8WHI 300+":  ("whips",     "8_awg",  "WHI", "8"),
    "10WHI-Q":    ("whips",     "10_awg", "WHI", "10"),
    "10WHI 300+": ("whips",     "10_awg", "WHI", "10"),
}


def process_cable_sheets(xlsx: pd.ExcelFile, pricing: dict, available: list) -> int:
    total = 0
    for sheet, (section, gauge, pn_type, pn_gauge) in CABLE_SHEETS.items():
        if sheet not in available:
            print(f"  WARNING: sheet '{sheet}' not found -- skipping")
            continue
        df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
        rows = data_rows(df)
        bucket = pricing[section][gauge]
        updated = 0
        for _, row in rows.iterrows():
            try:
                length = int(float(row[4]))
            except (TypeError, ValueError):
                continue
            prices = get_tier_prices(row)
            if len(prices) < 6:
                print(f"    WARNING: '{sheet}' length={length}: only {len(prices)} tier prices")
                continue
            for pol in ("P", "N"):
                pn = f"{pn_type}-{pn_gauge}-{pol}-{length}"
                if pn not in bucket:
                    bucket[pn] = {}
                bucket[pn].update(prices)
                updated += 1
        print(f"  '{sheet}': {updated // 2} lengths, {updated} PN entries")
        total += updated
    return total


# ---------------------------------------------------------------------------
# Step 1e -- Harness sheets (pricing update + data collection for library)
# ---------------------------------------------------------------------------

def derive_pn_from_description(description: str):
    """
    Parse a harness description into a canonical PN like '8P-10D8T-26'.
    Handles both numeric ('8 String') and word-form ('Eight String') counts.
    Returns None if the description does not match expected format.
    """
    if not description.strip():
        return None
    m = DESC_PATTERN.search(description)
    if not m:
        return None
    n_raw, polarity, drop_awg, trunk_awg, spacing = m.groups()
    if n_raw.isdigit():
        n_str = n_raw
    else:
        n_int = WORD_COUNT.get(n_raw.lower())
        if not n_int:
            return None
        n_str = str(n_int)
    pol = "P" if polarity.lower() == "positive" else "N"
    return f"{n_str}{pol}-{drop_awg}D{trunk_awg}T-{spacing}"


def process_harness_sheets(xlsx: pd.ExcelFile, pricing: dict, available: list) -> dict:
    """
    Updates pricing["harnesses"][sub_key] from the spreadsheet.
    Returns harness_data: {derived_pn: {atpi_pn, description, sheet, sub_key}}
    FS entries in harness_data are never overwritten by standard-sheet entries.
    """
    harness_data = {}
    warnings = []

    for sheet, sub_key in [("FS", "first_solar"), ("2-3 STRING", "standard")]:
        if sheet not in available:
            print(f"  WARNING: sheet '{sheet}' not found -- skipping")
            continue
        df = pd.read_excel(xlsx, sheet_name=sheet, header=None)
        rows = data_rows(df)
        bucket = pricing["harnesses"][sub_key]
        count = 0
        for _, row in rows.iterrows():
            col1 = str(row[1]).strip() if pd.notna(row[1]) else ""
            atpi_pn = str(row[2]).strip() if pd.notna(row[2]) else ""
            description = str(row[3]).strip() if pd.notna(row[3]) else ""

            if not description:
                continue  # blank divider rows

            derived = derive_pn_from_description(description)
            if not derived:
                print(f"  WARNING: could not parse PN from '{description[:60]}' in '{sheet}'")
                continue

            # Skip rows that belong to the first_solar bucket when processing standard.
            # Check the live dict (not a precomputed set) so newly-added FS entries are caught.
            if sub_key == "standard" and derived in pricing["harnesses"]["first_solar"]:
                print(f"  INFO: skipping FS harness '{derived}' found as paste error in '{sheet}'")
                continue

            # Log paste-error mismatches between col 1 PN and derived PN
            if col1 and col1 != "nan" and col1 != derived:
                warnings.append(
                    f"Sheet '{sheet}': col1='{col1}' != derived='{derived}' -- using derived"
                )

            prices = get_tier_prices(row)
            if len(prices) < 6:
                print(f"  WARNING: '{derived}' in '{sheet}': only {len(prices)} tier prices")

            if derived not in bucket:
                bucket[derived] = {}
            bucket[derived].update(prices)
            count += 1

            # Preserve FS entries -- never let a standard-sheet row overwrite them
            if derived not in harness_data or harness_data[derived]["sheet"] != "FS":
                harness_data[derived] = {
                    "atpi_pn": atpi_pn,
                    "description": description,
                    "sheet": sheet,
                    "sub_key": sub_key,
                }

        print(f"  '{sheet}': {count} harness pricing entries updated")

    # Remove any first_solar PNs that were erroneously written to the standard bucket
    # (can happen when a paste-error row in 2-3 STRING sheet carried a FS harness PN)
    to_clean = [
        k for k in pricing["harnesses"]["standard"]
        if k in pricing["harnesses"]["first_solar"]
    ]
    for k in to_clean:
        del pricing["harnesses"]["standard"][k]
        print(f"  Cleaned up '{k}' from standard bucket (belongs to first_solar)")

    if warnings:
        print(f"\n  PN mismatch warnings ({len(warnings)}):")
        for w in warnings:
            print(f"    {w}")

    return harness_data


# ---------------------------------------------------------------------------
# Step 1f -- New harness library entries
# ---------------------------------------------------------------------------

def _build_library_entry(pn: str, atpi_pn: str, description: str) -> dict | None:
    m = re.match(r"(\d+)([PN])-(\d+)D(\d+)T-(\d+)$", pn)
    if not m:
        return None
    n_str, pol_char, drop_awg, trunk_awg, spacing = m.groups()
    is_pos = pol_char == "P"
    return {
        "part_number": pn,
        "atpi_part_number": atpi_pn,
        "description": description,
        "num_strings": int(n_str),
        "polarity": "positive" if is_pos else "negative",
        "string_spacing_ft": float(spacing),
        "category": "First Solar",
        "drop_wire_gauge": f"{drop_awg} AWG",
        "trunk_wire_gauge": f"{trunk_awg} AWG",
        "connector_type": "PV4S/MC4",
        "fused": is_pos,
        "fuse_rating": "5A" if is_pos else None,
    }


def add_new_harness_entries(harness_lib: dict, harness_data: dict) -> int:
    added = 0
    for pn, data in harness_data.items():
        if data["sheet"] != "FS":
            continue
        if pn in harness_lib:
            continue
        entry = _build_library_entry(pn, data["atpi_pn"], data["description"])
        if entry:
            harness_lib[pn] = entry
            print(f"    + {pn}")
            added += 1
    return added


# ---------------------------------------------------------------------------
# Step 1g -- Fuse sheet
# ---------------------------------------------------------------------------

def process_fuse_sheet(xlsx: pd.ExcelFile, pricing: dict, available: list) -> int:
    if "Fuses" not in available:
        print("  WARNING: sheet 'Fuses' not found -- skipping")
        return 0
    df = pd.read_excel(xlsx, sheet_name="Fuses", header=None)
    updated = 0
    for _, row in df.iterrows():
        # col 0 must be a real ITEM number
        if not is_item_number(row[0]):
            continue
        # col 1 must be present (MOQ 100 rows only; MOQ 500 have NaN)
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ""
        if not col1 or col1 == "nan":
            continue
        pn = col1

        # Paste-error fix: FS harness PN in the 45A MC4 row
        if HARNESS_PN_RE.match(pn):
            print(f"    Paste error: '{pn}' -> '24F0030' (45A MC4 row)")
            pn = "24F0030"

        try:
            price = round(float(row[7]), 2)
        except (KeyError, ValueError, TypeError):
            print(f"    WARNING: no price for fuse '{pn}' -- skipped")
            continue

        pricing["fuses"][pn] = price
        updated += 1

    # Remove the old 45A H4 PLUS key (24F0050) — in the new spreadsheet the
    # 45A H4 PLUS uses key 24F0030 (data anomaly), so 24F0050 is now orphaned.
    if "24F0050" in pricing["fuses"]:
        del pricing["fuses"]["24F0050"]
        print("    Removed orphaned '24F0050' (45A H4 PLUS now keyed as '24F0030')")

    print(f"  'Fuses': {updated} entries written ({len(pricing['fuses'])} unique keys)")
    return updated


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    if not XLSX_PATH.exists():
        print(f"ERROR: Spreadsheet not found: {XLSX_PATH}")
        sys.exit(1)

    print(f"\nWorkbook: {XLSX_PATH.name}")
    xlsx = pd.ExcelFile(XLSX_PATH)
    available = xlsx.sheet_names
    print(f"  Sheets: {available}")

    print("\nLoading JSON files...")
    pricing = load_json(PRICING_JSON)
    harness_lib = load_json(HARNESS_JSON)

    print("\n[1a] Settings...")
    update_settings(pricing)

    print("\n[1b] Backfilling 6.5 tier on existing entries...")
    backfill_tiers(pricing)

    print("\n[1d] Cable sheets (extenders & whips)...")
    cable_count = process_cable_sheets(xlsx, pricing, available)

    print("\n[1e] Harness sheets (pricing)...")
    harness_data = process_harness_sheets(xlsx, pricing, available)

    print("\n[1f] New harness library entries...")
    new_entries = add_new_harness_entries(harness_lib, harness_data)
    print(f"  {new_entries} new entries added")

    print("\n[1g] Fuse sheet...")
    fuse_count = process_fuse_sheet(xlsx, pricing, available)

    print("\n[1h] Writing files...")
    save_json(PRICING_JSON, pricing)
    save_json(HARNESS_JSON, harness_lib)

    print("\n=== Summary ===")
    print(f"  Cable pricing entries written  : {cable_count}")
    print(f"  Harness pricing entries written: {len(harness_data)}")
    print(f"  Fuse entries written           : {fuse_count}")
    print(f"  New harness library entries    : {new_entries}")
    print("Done.")


if __name__ == "__main__":
    main()
