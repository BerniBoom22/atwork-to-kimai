#!/usr/bin/env python3
"""
AtWork to Kimai Migration Tool
==============================
A universal, zero-dependency Python tool to migrate timesheets, vacations,
sick leaves, and holidays from atWork (iOS) CSV exports (both Excel UTF-16
and Numbers UTF-8 formats) directly into a self-hosted Kimai 2 instance via REST API.

Author: BerniBoom22 (https://github.com/BerniBoom22)
Repository: https://github.com/BerniBoom22/atwork-to-kimai
License: MIT
"""

import argparse
import csv
import datetime
import getpass
import glob
import json
import os
import sys
import urllib.error
import urllib.request

# Multi-language section mappings (Spanish, English, German, French)
SECTIONS_WORK = {"entradas", "entries", "einträge", "entrées"}
SECTIONS_VACATION = {"días de vacaciones", "vacation days", "urlaubstage", "jours de congé", "vacaciones", "vacation"}
SECTIONS_SICK = {"días de baja", "sick days", "krankheitstage", "jours de maladie", "bajas", "sick leave"}
SECTIONS_HOLIDAY = {"festivos", "holidays", "feiertage", "jours fériés"}
SECTIONS_STOP = {
    "total", "clientes", "clients", "kunden",
    "proyectos", "projects", "projekte",
    "tareas", "tasks", "aufgaben",
    "horas extra", "overtime", "überstunden", "heures supplémentaires"
}

DATE_TIME_FORMATS = [
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y %I:%M %p",
    "%d.%m.%Y %H:%M",
    "%d-%m-%Y %H:%M"
]

DATE_ONLY_FORMATS = [
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%m/%d/%Y",
    "%d.%m.%Y",
    "%d-%m-%Y"
]


def detect_encoding_and_delimiter(filepath):
    """Automatically detects file encoding (UTF-16 vs UTF-8) and CSV delimiter."""
    with open(filepath, "rb") as f:
        raw = f.read(4096)

    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        encoding = "utf-16"
    elif b"\x00" in raw[:200]:
        encoding = "utf-16"
    else:
        encoding = "utf-8-sig"

    try:
        sample_text = raw.decode(encoding, errors="ignore")
    except Exception:
        encoding = "utf-8"
        sample_text = raw.decode(encoding, errors="ignore")

    lines = [line for line in sample_text.splitlines() if line.strip()]
    sample_lines = "\n".join(lines[:10])

    counts = {
        "\t": sample_lines.count("\t"),
        ";": sample_lines.count(";"),
        ",": sample_lines.count(",")
    }
    delimiter = max(counts, key=counts.get)
    if counts[delimiter] == 0:
        delimiter = ";"

    return encoding, delimiter


def parse_datetime(val):
    val = val.strip()
    for fmt in DATE_TIME_FORMATS:
        try:
            return datetime.datetime.strptime(val, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unknown datetime format: {val}")


def parse_date(val):
    val = val.strip().split(" ")[0]
    for fmt in DATE_ONLY_FORMATS:
        try:
            return datetime.datetime.strptime(val, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unknown date format: {val}")


def parse_atwork_csv(filepath, absence_start_str="06:30", absence_hours=8.0, default_names=None):
    if not os.path.exists(filepath):
        print(f"[!] Error: File '{filepath}' not found.")
        sys.exit(1)

    if default_names is None:
        default_names = {
            "work": "Regular Work",
            "vacation": "Vacation",
            "sick": "Sick Leave",
            "holiday": "Public Holiday"
        }

    encoding, delimiter = detect_encoding_and_delimiter(filepath)
    delim_name = "TAB" if delimiter == "\t" else delimiter
    print(f"[*] Detected format -> Encoding: {encoding.upper()} | Delimiter: '{delim_name}'")

    abs_h, abs_m = map(int, absence_start_str.split(":"))
    records = []

    with open(filepath, "r", encoding=encoding, errors="ignore") as f:
        reader = csv.reader(f, delimiter=delimiter)
        current_section = None
        has_index_col = False

        for row in reader:
            if not row or not any(cell.strip() for cell in row):
                continue

            first_cell = row[0].strip()
            first_lower = first_cell.lower()

            # Check section switches
            if first_lower in SECTIONS_WORK:
                current_section = "work"
                continue
            elif first_lower in SECTIONS_VACATION:
                current_section = "vacation"
                continue
            elif first_lower in SECTIONS_SICK:
                current_section = "sick"
                continue
            elif first_lower in SECTIONS_HOLIDAY:
                current_section = "holiday"
                continue
            elif any(first_lower.startswith(stop) for stop in SECTIONS_STOP):
                current_section = None
                continue

            # Check if it's a column header row
            if first_cell == "#" or first_lower in {"inicio", "start", "día", "day", "date", "fecha", "beginn", "datum"}:
                has_index_col = (first_cell == "#")
                continue

            if not current_section or not first_cell:
                continue

            offset = 1 if has_index_col else 0

            # 1. Process regular work entries
            if current_section == "work" and len(row) >= (2 + offset):
                start_str = row[0 + offset].strip()
                end_str = row[1 + offset].strip()
                if not start_str or not end_str:
                    continue

                # Extract optional fields if user filled them in atWork
                client_col = row[7 + offset].strip() if len(row) > (7 + offset) else ""
                project_col = row[8 + offset].strip() if len(row) > (8 + offset) else ""
                task_col = row[9 + offset].strip() if len(row) > (9 + offset) else ""
                note_col = row[10 + offset].strip() if len(row) > (10 + offset) else ""

                try:
                    dt_start = parse_datetime(start_str)
                    dt_end = parse_datetime(end_str)
                    records.append({
                        "type": "work",
                        "start": dt_start,
                        "end": dt_end,
                        "customer": client_col,
                        "project": project_col,
                        "activity": task_col or default_names["work"],
                        "description": note_col
                    })
                except ValueError:
                    continue

            # 2. Process absences (Vacations, Sick leave, Holidays)
            elif current_section in {"vacation", "sick", "holiday"} and len(row) >= (3 + offset):
                day_str = row[0 + offset].strip()
                duration_str = row[1 + offset].strip().replace(",", ".")
                name_str = row[2 + offset].strip()

                if not day_str:
                    continue

                try:
                    dt_base = parse_date(day_str)
                    try:
                        mult = float(duration_str) if duration_str else 1.0
                    except ValueError:
                        mult = 1.0

                    dt_start = dt_base.replace(hour=abs_h, minute=abs_m, second=0)
                    dt_end = dt_start + datetime.timedelta(hours=absence_hours * mult)

                    records.append({
                        "type": current_section,
                        "start": dt_start,
                        "end": dt_end,
                        "customer": "",
                        "project": "",
                        "activity": default_names[current_section],
                        "description": name_str
                    })
                except ValueError:
                    continue

    return records


class KimaiSynchronizer:
    def __init__(self, base_url, api_token, legacy_user=None):
        self.base_url = base_url.rstrip("/")
        self.api_token = api_token
        self.legacy_user = legacy_user
        self.headers = None

    def authenticate(self):
        print("\n[*] Validating connection and authentication with Kimai...")
        test_url = f"{self.base_url}/api/customers"

        # 1. Try modern Bearer Token authentication
        headers_bearer = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        req = urllib.request.Request(test_url, headers=headers_bearer)
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                if resp.status == 200:
                    print("[+] Authenticated successfully using Bearer Token (Kimai 2 Modern API).")
                    self.headers = headers_bearer
                    return True
        except urllib.error.HTTPError as e:
            if e.code not in (401, 403):
                print(f"[!] HTTP Error {e.code} on {test_url}: {e.read().decode('utf-8', errors='ignore')}")
        except urllib.error.URLError as e:
            print(f"[!] Network error connecting to {self.base_url}: {e.reason}")
            return False

        # 2. Try Legacy X-AUTH-TOKEN if user was provided
        candidates = [self.legacy_user] if self.legacy_user else ["admin"]
        for user in candidates:
            if not user:
                continue
            headers_legacy = {
                "X-AUTH-USER": user,
                "X-AUTH-TOKEN": self.api_token,
                "Content-Type": "application/json",
                "Accept": "application/json"
            }
            req = urllib.request.Request(test_url, headers=headers_legacy)
            try:
                with urllib.request.urlopen(req, timeout=15) as resp:
                    if resp.status == 200:
                        print(f"[+] Authenticated successfully using Legacy Headers (User: '{user}').")
                        self.headers = headers_legacy
                        return True
            except urllib.error.HTTPError:
                pass
            except urllib.error.URLError:
                pass

        print("[!] Authentication failed (403 Invalid Credentials). Check your API Token.")
        return False

    def api_req(self, endpoint, data=None, method=None, show_error=True):
        url = f"{self.base_url}/api/{endpoint.lstrip('/')}"
        req_data = json.dumps(data).encode("utf-8") if data is not None else None
        if method is None:
            method = "POST" if req_data is not None else "GET"

        req = urllib.request.Request(url, headers=self.headers, data=req_data, method=method)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                raw = resp.read().decode("utf-8")
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as e:
            if show_error:
                err_body = e.read().decode("utf-8", errors="ignore")
                print(f"[!] HTTP Error {e.code} on {method} {endpoint}: {err_body}")
            return None
        except urllib.error.URLError as e:
            if show_error:
                print(f"[!] Network error on {endpoint}: {e.reason}")
            return None

    def ensure_customer(self, name, currency="EUR", timezone="Europe/Madrid", country="ES"):
        customers = self.api_req("customers?visible=3", show_error=False) or self.api_req("customers") or []
        cust_id = next((c["id"] for c in customers if c.get("name") == name), None)

        if not cust_id:
            res = self.api_req("customers", {
                "name": name,
                "currency": currency,
                "timezone": timezone,
                "country": country,
                "visible": True,
                "billable": True
            })
            if res and "id" in res:
                cust_id = res["id"]
                print(f"[+] Created Customer: '{name}' (ID: {cust_id})")
        else:
            self.api_req(f"customers/{cust_id}", {"visible": True}, method="PATCH", show_error=False)
            print(f"[*] Found & verified Customer: '{name}' (ID: {cust_id})")

        return cust_id

    def ensure_project(self, name, customer_id):
        projects = self.api_req("projects?visible=3", show_error=False) or self.api_req("projects") or []
        proj_id = next((p["id"] for p in projects if p.get("name") == name), None)

        if not proj_id:
            res = self.api_req("projects", {
                "name": name,
                "customer": int(customer_id),
                "visible": True,
                "billable": True,
                "globalActivities": True
            })
            if res and "id" in res:
                proj_id = res["id"]
                print(f"[+] Created Project: '{name}' (ID: {proj_id})")
        else:
            self.api_req(f"projects/{proj_id}", {"visible": True}, method="PATCH", show_error=False)
            print(f"[*] Found & verified Project: '{name}' (ID: {proj_id})")

        return proj_id

    def ensure_activity(self, name, project_id):
        activities = self.api_req("activities?visible=3", show_error=False) or self.api_req("activities") or []
        act_id = next((a["id"] for a in activities if a.get("name") == name), None)

        if not act_id:
            res = self.api_req("activities", {
                "name": name,
                "project": int(project_id),
                "visible": True,
                "billable": True
            })
            if res and "id" in res:
                act_id = res["id"]
                print(f"[+] Created Activity: '{name}' (ID: {act_id})")
        else:
            self.api_req(f"activities/{act_id}", {"visible": True}, method="PATCH", show_error=False)
            print(f"[*] Found & verified Activity: '{name}' (ID: {act_id})")

        return act_id


def auto_discover_csv():
    """Finds the first atWork CSV file in the current directory if none is specified."""
    patterns = ["*atwork*.csv", "*.csv"]
    for pat in patterns:
        matches = [f for f in glob.glob(pat) if not f.startswith("kimai_")]
        if matches:
            return matches[0]
    return None


def main():
    parser = argparse.ArgumentParser(
        description="Universal migration tool from atWork (iOS) CSV exports to Self-Hosted Kimai 2."
    )
    parser.add_argument("-f", "--file", help="Path to the exported atWork CSV file")
    parser.add_argument("-u", "--url", help="Kimai server URL (e.g., https://kimai.example.com)")
    parser.add_argument("-t", "--token", help="Kimai API Token")
    parser.add_argument("--user", help="Username (only required for legacy Kimai versions)")
    parser.add_argument("--customer", default="My Company", help="Default Customer name (default: 'My Company')")
    parser.add_argument("--project", default="Time Tracking", help="Default Project name (default: 'Time Tracking')")
    parser.add_argument("--work-activity", default="Regular Work", help="Activity name for work entries")
    parser.add_argument("--vacation-activity", default="Vacation", help="Activity name for vacation days")
    parser.add_argument("--sick-activity", default="Sick Leave", help="Activity name for sick leave days")
    parser.add_argument("--holiday-activity", default="Public Holiday", help="Activity name for public holidays")
    parser.add_argument("--absence-start", default="06:30", help="Default start time for full-day absences (HH:MM)")
    parser.add_argument("--absence-hours", type=float, default=8.0, help="Standard daily hours for absences (default: 8.0)")
    parser.add_argument("--currency", default="EUR", help="Default currency code (default: EUR)")
    parser.add_argument("--timezone", default="Europe/Madrid", help="Default timezone (default: Europe/Madrid)")
    parser.add_argument("--country", default="ES", help="Default 2-letter country code (default: ES)")
    parser.add_argument("--dry-run", action="store_true", help="Parse file and show summary without uploading to Kimai")

    args = parser.parse_args()

    print("=" * 60)
    print("         AtWork to Kimai 2 Universal Migration Tool")
    print("                 by @BerniBoom22 on GitHub")
    print("=" * 60)

    # 1. Resolve CSV File
    csv_file = args.file or os.environ.get("ATWORK_CSV_FILE")
    if not csv_file:
        discovered = auto_discover_csv()
        if discovered:
            ans = input(f"[?] Found CSV file '{discovered}'. Use this file? [Y/n]: ").strip().lower()
            if ans in ("", "y", "yes", "s", "si", "sí"):
                csv_file = discovered
        if not csv_file:
            csv_file = input("[?] Enter the path to your atWork CSV file: ").strip().strip('"')

    default_names = {
        "work": args.work_activity,
        "vacation": args.vacation_activity,
        "sick": args.sick_activity,
        "holiday": args.holiday_activity
    }

    # 2. Parse File
    records = parse_atwork_csv(
        csv_file,
        absence_start_str=args.absence_start,
        absence_hours=args.absence_hours,
        default_names=default_names
    )

    if not records:
        print("[!] No valid records were found in the file. Exiting.")
        sys.exit(1)

    c_work = sum(1 for r in records if r["type"] == "work")
    c_vac = sum(1 for r in records if r["type"] == "vacation")
    c_sick = sum(1 for r in records if r["type"] == "sick")
    c_hol = sum(1 for r in records if r["type"] == "holiday")

    print(f"\n[+] Successfully extracted {len(records)} total records:")
    print(f"    - Regular work shifts : {c_work}")
    print(f"    - Vacation days       : {c_vac}")
    print(f"    - Sick leave days     : {c_sick}")
    print(f"    - Public holidays     : {c_hol}")

    if args.dry_run:
        print("\n[*] Dry-run mode enabled. No data was sent to Kimai.")
        sys.exit(0)

    # 3. Resolve Kimai Credentials interactively if not passed via CLI/Env
    kimai_url = args.url or os.environ.get("KIMAI_URL")
    if not kimai_url:
        kimai_url = input("\n[?] Enter your Kimai URL (e.g. http://localhost:8001): ").strip()

    kimai_token = args.token or os.environ.get("KIMAI_API_TOKEN")
    if not kimai_token:
        kimai_token = getpass.getpass("[?] Enter your Kimai API Token (input hidden): ").strip()
        if not kimai_token:
            kimai_token = input("[?] Enter your Kimai API Token: ").strip()

    # 4. Connect and Upload
    sync = KimaiSynchronizer(kimai_url, kimai_token, legacy_user=args.user)
    if not sync.authenticate():
        sys.exit(1)

    cust_id = sync.ensure_customer(args.customer, args.currency, args.timezone, args.country)
    if not cust_id:
        print("[!] Could not create or find default Customer. Aborting.")
        sys.exit(1)

    proj_id = sync.ensure_project(args.project, cust_id)
    if not proj_id:
        print("[!] Could not create or find default Project. Aborting.")
        sys.exit(1)

    # Cache activities
    activity_cache = {}
    unique_activities = set(r["activity"] for r in records)
    for act_name in unique_activities:
        act_id = sync.ensure_activity(act_name, proj_id)
        if not act_id:
            print(f"[!] Could not create or find Activity '{act_name}'. Aborting.")
            sys.exit(1)
        activity_cache[act_name] = act_id

    print(f"\n[*] Uploading {len(records)} timesheet entries to Kimai...")
    errors = 0
    for i, r in enumerate(records, 1):
        payload = {
            "begin": r["start"].strftime("%Y-%m-%dT%H:%M:%S"),
            "end": r["end"].strftime("%Y-%m-%dT%H:%M:%S"),
            "project": int(proj_id),
            "activity": int(activity_cache[r["activity"]]),
            "description": r["description"]
        }
        res = sync.api_req("timesheets", payload, show_error=(errors < 3))
        if not res:
            errors += 1

        if i % 100 == 0 or i == len(records):
            print(f"    -> Progress: {i}/{len(records)} processed (Errors: {errors})")

    if errors == 0:
        print("\n[+] Migration completed 100% successfully!")
    else:
        print(f"\n[!] Migration finished with {errors} errors.")


if __name__ == "__main__":
    main()
