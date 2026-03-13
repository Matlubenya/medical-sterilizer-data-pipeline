#!/usr/bin/env python3
import sys
from pathlib import Path

# Simple path definitions
CURRENT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = CURRENT_DIR / "data"
RESULTS_DIR = CURRENT_DIR / "results"
PARSING_RESULTS_DIR = RESULTS_DIR / "Parsing_results"

# Ritter paths
RITTER1_PDF_DIR = DATA_DIR / "RitterPDF" / "Ritter1"
RITTER2_PDF_DIR = DATA_DIR / "RitterPDF" / "Ritter2"
RITTER1_TXT_DIR = DATA_DIR / "RitterTXT" / "Ritter1"
RITTER2_TXT_DIR = DATA_DIR / "RitterTXT" / "Ritter2"

# Statim paths
STATIMA_DIR = DATA_DIR / "Statim" / "StatimA"
STATIMB_DIR = DATA_DIR / "Statim" / "StatimB"

# Create directories
PARSING_RESULTS_DIR.mkdir(parents=True, exist_ok=True)
RITTER1_TXT_DIR.mkdir(parents=True, exist_ok=True)
RITTER2_TXT_DIR.mkdir(parents=True, exist_ok=True)
import sys
from pathlib import Path

# Add config to Python path
config_path = Path(__file__).resolve().parents[2] / 'config'
sys.path.insert(0, str(config_path))

from paths import *


"""
Ritter TXT -> CSV parser (Option 1: Statim-compatible cycle summary)
- Converts °F -> °C
- Converts psi -> kPa (kPa = psi * 6.89476)
- Walks BASE_DIR/Ritter1 and /Ritter2 for .txt files
- Produces a single CSV with one row per cycle
"""

from pathlib import Path
import re
import csv
from datetime import datetime, timedelta

# === CONFIG ===
BASE_DIR = Path("RITTER_TXT_DIR")  # <- change if needed
RITTER_FOLDERS = ["Ritter1", "Ritter2"]
OUTPUT_CSV = PARSING_RESULTS_DIR / "ritter_combined_summary.csv"

# === Helpers ===
def f_to_c(f):
    try:
        return (float(f) - 32.0) * 5.0/9.0
    except Exception:
        return None

def psi_to_kpa(psi):
    try:
        return float(psi) * 6.89476
    except Exception:
        return None

def mmss_to_seconds(mmss):
    """Convert mm:ss or hh:mm:ss or m:ss to seconds."""
    if mmss is None:
        return None
    parts = mmss.strip().split(':')
    try:
        parts = [int(p) for p in parts]
    except:
        return None
    if len(parts) == 2:
        m, s = parts
        return m*60 + s
    elif len(parts) == 3:
        h, m, s = parts
        return h*3600 + m*60 + s
    else:
        return None

def parse_datetime_line(line):
    """Try parse a line like '12/06/2022 10:36 AM' trying multiple formats."""
    line = line.strip()
    formats = [
        "%m/%d/%Y %I:%M %p",
        "%d/%m/%Y %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %I:%M %p",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(line, fmt)
        except Exception:
            continue
    # try date only
    date_only_formats = ["%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"]
    for fmt in date_only_formats:
        try:
            d = datetime.strptime(line, fmt)
            return d
        except:
            continue
    return None

def safe_float_from_str(s):
    try:
        return float(s)
    except:
        return None

# === Parsing functions ===
def extract_metadata(lines):
    """Extract top-level metadata from file lines (first 30 lines or so)."""
    meta = {
        "device_model": None,
        "firmware": None,
        "total_cycles": None,
        "unit_name": None,
        "sterilizer_id": None,
        "cycle_start": None,
        "cycle_end": None,
        "cycle_type": None,
        "set_temperature_f": None,
        "set_time_min": None,
        "dry_time_min": None,
        "accepted": None,
        "rejected": None,
        "operator": None,
    }

    # scan through first N lines for key fields
    header_block = "\n".join(lines[:120])  # look wide but limited

    # model + firmware (e.g., "Midmark M11 - v1.0.5")
    m = re.search(r"(?P<model>Midmark\s*\w+\s*(?:\-.*)?)", header_block, re.IGNORECASE)
    if m:
        meta["device_model"] = m.group("model").strip()

    # firmware version in same line
    m = re.search(r"v?(\d+\.\d+(?:\.\d+)*)", header_block)
    if m:
        meta["firmware"] = m.group(1)

    m = re.search(r"Total Cycles:\s*(\d+)", header_block, re.IGNORECASE)
    if m:
        meta["total_cycles"] = int(m.group(1))

    m = re.search(r"Name:\s*(\S+)", header_block, re.IGNORECASE)
    if m:
        meta["unit_name"] = m.group(1).strip()

    m = re.search(r"Sterilizer ID:\s*([A-Za-z0-9\-]+)", header_block, re.IGNORECASE)
    if m:
        meta["sterilizer_id"] = m.group(1).strip()

    # cycle type - look for lines like "BEGIN POUCHES CYCLE" or "BEGIN UNWRAPPED CYCLE"
    m = re.search(r"BEGIN\s+([A-Z ]+?)\s+CYCLE", header_block, re.IGNORECASE)
    if m:
        meta["cycle_type"] = m.group(1).strip().title()

    # set temperature (e.g., "Temp: 270 Degrees F")
    m = re.search(r"Temp:\s*([0-9]+(?:\.[0-9]+)?)\s*Degrees\s*F", header_block, re.IGNORECASE)
    if m:
        meta["set_temperature_f"] = float(m.group(1))

    # set time (Time: 4 Minutes)
    m = re.search(r"Time:\s*([0-9]+(?:\.[0-9]+)?)\s*Minutes?", header_block, re.IGNORECASE)
    if m:
        meta["set_time_min"] = float(m.group(1))

    # Dry: 30 Minutes or Drying: 30:00
    m = re.search(r"Dry(?:ing)?:\s*([0-9]+)(?:\s*Minutes?)", header_block, re.IGNORECASE)
    if m:
        meta["dry_time_min"] = int(m.group(1))
    else:
        m = re.search(r"DRYING:\s*([0-9:]+)", header_block, re.IGNORECASE)
        if m:
            secs = mmss_to_seconds(m.group(1))
            if secs is not None:
                meta["dry_time_min"] = secs // 60

    # acceptance / rejection
    m = re.search(r"ACCEPTED:\s*(X|x|\*)?", header_block)
    if m:
        meta["accepted"] = bool(m.group(1))
    m = re.search(r"REJECTED:\s*(X|x|\*)?", header_block)
    if m:
        meta["rejected"] = bool(m.group(1))

    # operator line (common beginnings)
    m = re.search(r"Operat(?:or|o)\w*\s*[:\-]?\s*([A-Za-z0-9]{1,6})", header_block, re.IGNORECASE)
    if m:
        meta["operator"] = m.group(1).strip()

    # find start and end datetimes - sometimes two separate date lines
    # Find first date/time-looking line
    dt_candidates = []
    for l in lines[:120]:
        l_str = l.strip()
        if re.search(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}", l_str):
            dt = parse_datetime_line(l_str)
            if dt:
                dt_candidates.append(dt)
    if dt_candidates:
        meta["cycle_start"] = dt_candidates[0]
        if len(dt_candidates) > 1:
            meta["cycle_end"] = dt_candidates[-1]

    return meta

def extract_phase_durations(lines):
    """Extract phase durations like FILLING: 1:02 etc. returns dict of seconds"""
    res = {
        "filling_sec": None,
        "heating_sec": None,
        "sterilizing_sec": None,
        "venting_sec": None,
        "drying_sec": None,
        "total_cycle_sec": None
    }
    # patterns like "FILLING: 1:02" or "HEATING: 20:46" or "TOTAL CYCLE: 00:58:57"
    text = "\n".join(lines)
    for key in ["FILLING", "HEATING", "STERILIZING", "VENTING", "DRYING", "TOTAL CYCLE"]:
        m = re.search(rf"{key}\s*:\s*([0-9:\s]+)", text, re.IGNORECASE)
        if m:
            secs = mmss_to_seconds(m.group(1).strip())
            if secs is not None:
                k = key.lower().replace(" ", "_")
                if k == "total_cycle":
                    res["total_cycle_sec"] = secs
                else:
                    res[f"{k}_sec"] = secs
    return res

def extract_sterilizing_min_max(lines):
    """
    Look for Min/Max lines in sterilizing block like:
    Min 271.7 F 29.2
    Max 273.0 F 30.2
    If not found, compute from sterilizing telemetry lines.
    """
    min_temp_f = None
    max_temp_f = None
    min_psi = None
    max_psi = None

    text = "\n".join(lines)
    m_min = re.search(r"Min\s+([0-9]+(?:\.[0-9]+)?)\s*F\s*([0-9]+(?:\.[0-9]+)?)", text, re.IGNORECASE)
    m_max = re.search(r"Max\s+([0-9]+(?:\.[0-9]+)?)\s*F\s*([0-9]+(?:\.[0-9]+)?)", text, re.IGNORECASE)
    if m_min:
        min_temp_f = float(m_min.group(1))
        min_psi = float(m_min.group(2))
    if m_max:
        max_temp_f = float(m_max.group(1))
        max_psi = float(m_max.group(2))

    # If not found, parse lines under STERILIZING telemetry block (mm:ss Degrees PSI)
    if min_temp_f is None or max_temp_f is None:
        # find the sterilizing block start index
        steril_start = None
        steril_end = None
        for i, l in enumerate(lines):
            if re.match(r"^\s*STERILIZING\s*$", l, re.IGNORECASE):
                steril_start = i
                # block continues until a blank line or next ALL CAPS phase header
                for j in range(i+1, min(i+200, len(lines))):
                    if re.match(r"^[A-Z ]{3,}$", lines[j].strip()) and lines[j].strip().upper() not in ("MM:SS DEGREES PSI",):
                        steril_end = j
                        break
                break
        if steril_start is not None:
            block_lines = []
            # collect lines after the "mm:ss Degrees PSI" header if present
            for j in range(steril_start, min(steril_start+200, len(lines))):
                l = lines[j].strip()
                # telemetry lines often look like "0:00 271.8 F 29.4"
                if re.match(r"^\d+:\d+\s+[\d\.]+\s*F\s+[\d\.]+", l):
                    block_lines.append(l)
            # parse block_lines for min/max
            temps = []
            psis = []
            for l in block_lines:
                parts = re.split(r"\s+", l)
                # [mm:ss, temp, F, psi]
                if len(parts) >= 4:
                    try:
                        tempf = float(parts[1])
                        psi = float(parts[3])
                        temps.append(tempf)
                        psis.append(psi)
                    except:
                        continue
            if temps:
                min_temp_f = min(temps)
                max_temp_f = max(temps)
            if psis:
                min_psi = min(psis)
                max_psi = max(psis)

    return min_temp_f, max_temp_f, min_psi, max_psi

# === Main processing ===
def parse_txt_file(path: Path):
    """Parse a single converted TXT file into a dict for CSV row"""
    text = path.read_text(encoding="utf-8", errors="ignore")
    lines = [ln.rstrip() for ln in text.splitlines()]

    meta = extract_metadata(lines)
    phase_secs = extract_phase_durations(lines)
    min_temp_f, max_temp_f, min_psi, max_psi = extract_sterilizing_min_max(lines)

    # Convert units
    set_temp_c = f_to_c(meta["set_temperature_f"]) if meta["set_temperature_f"] is not None else None
    min_temp_c = f_to_c(min_temp_f) if min_temp_f is not None else None
    max_temp_c = f_to_c(max_temp_f) if max_temp_f is not None else None
    min_pressure_kpa = psi_to_kpa(min_psi) if min_psi is not None else None
    max_pressure_kpa = psi_to_kpa(max_psi) if max_psi is not None else None

    # Compose CSV-ready row dict (Statim-compatible fields)
    row = {
        "source_file": str(path),
        "device_model": meta.get("device_model"),
        "firmware": meta.get("firmware"),
        "total_cycles_reported": meta.get("total_cycles"),
        "unit_name": meta.get("unit_name"),
        "sterilizer_id": meta.get("sterilizer_id"),
        "cycle_start": meta.get("cycle_start").isoformat() if meta.get("cycle_start") else None,
        "cycle_end": meta.get("cycle_end").isoformat() if meta.get("cycle_end") else None,
        "cycle_type": meta.get("cycle_type"),
        "set_temp_c": round(set_temp_c, 2) if set_temp_c is not None else None,
        "set_time_min": meta.get("set_time_min"),
        "dry_time_min": meta.get("dry_time_min"),
        "filling_sec": phase_secs.get("filling_sec"),
        "heating_sec": phase_secs.get("heating_sec"),
        "sterilizing_sec": phase_secs.get("sterilizing_sec"),
        "venting_sec": phase_secs.get("venting_sec"),
        "drying_sec": phase_secs.get("drying_sec"),
        "total_cycle_sec": phase_secs.get("total_cycle_sec"),
        "min_temp_c": round(min_temp_c, 2) if min_temp_c is not None else None,
        "max_temp_c": round(max_temp_c, 2) if max_temp_c is not None else None,
        "min_pressure_kpa": round(min_pressure_kpa, 2) if min_pressure_kpa is not None else None,
        "max_pressure_kpa": round(max_pressure_kpa, 2) if max_pressure_kpa is not None else None,
        "accepted": bool(meta.get("accepted")),
        "rejected": bool(meta.get("rejected")),
        "operator": meta.get("operator"),
    }

    return row

def main():
    csv_fields = [
        "source_file",
        "device_model",
        "firmware",
        "total_cycles_reported",
        "unit_name",
        "sterilizer_id",
        "cycle_start",
        "cycle_end",
        "cycle_type",
        "set_temp_c",
        "set_time_min",
        "dry_time_min",
        "filling_sec",
        "heating_sec",
        "sterilizing_sec",
        "venting_sec",
        "drying_sec",
        "total_cycle_sec",
        "min_temp_c",
        "max_temp_c",
        "min_pressure_kpa",
        "max_pressure_kpa",
        "accepted",
        "rejected",
        "operator"
    ]

    total_files = 0
    success = 0
    failed = 0
    per_unit_counts = {}

    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
# FIXME:     with open(OUTPUT_CSV, 'w', newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=csv_fields)
        writer.writeheader()

        for folder in RITTER_FOLDERS:
            root = BASE_DIR / folder
            if not root.exists():
                print(f"Warning: {root} does not exist. Skipping.")
                continue
            for p in root.rglob("*.txt"):
                total_files += 1
                try:
                    row = parse_txt_file(p)
                    writer.writerow(row)
                    success += 1
                    per_unit_counts.setdefault(folder, 0)
                    per_unit_counts[folder] += 1
                except Exception as e:
                    print(f"Failed to parse {p}: {e}")
                    failed += 1

    print("\n=== Parsing complete ===")
    print(f"Total TXT files discovered: {total_files}")
    print(f"Successfully parsed: {success}")
    print(f"Failed to parse: {failed}")
    print("Per-unit parsed counts:")
    for unit, cnt in per_unit_counts.items():
        print(f"  {unit}: {cnt}")

    print(f"\nOutput CSV: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
