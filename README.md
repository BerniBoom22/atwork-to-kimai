# AtWork to Kimai Migration Tool ⏱️➡️📊

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![Author](https://img.shields.io/badge/author-BerniBoom22-181717.svg?logo=github)](https://github.com/BerniBoom22)

A lightweight, **zero-dependency** Python tool to migrate your complete time-tracking history (work shifts, vacations, sick leave, and public holidays) from **atWork (iOS)** directly into a self-hosted **Kimai 2** instance via its REST API.

---

## ✨ Features

- **Zero External Dependencies:** Uses only Python 3 standard libraries (`urllib`, `csv`, `argparse`, `getpass`). No `pip install` or virtual environments required.
- **Smart Auto-Detection:** Automatically handles both **atWork "Excel CSV" exports** (`UTF-16LE`, tab-separated, indexed) and **Numbers / Standard CSV exports** (`UTF-8`, semicolon/comma-separated).
- **Full Overtime & Balance Preservation:** Imports not only regular work shifts (`Entradas` / `Entries`), but also **Vacations**, **Sick Leaves**, and **Public Holidays** as configurable full-day shifts (default `8.0h`) so your weekly/monthly hour balances match 100%.
- **Multi-Language Support:** Recognizes atWork section headers exported in English, Spanish, German, and French.
- **Kimai 2 API Workarounds Built-in:** Supports modern `Bearer` token authentication (with legacy `X-AUTH-TOKEN` fallback) and automatically enforces `"visible": true` entity activation to prevent Symfony API `400 Validation Failed` errors.

---

## 🚀 Quick Start

### 1. Clone the repository
```bash
git clone https://github.com/BerniBoom22/atwork-to-kimai.git
cd atwork-to-kimai
```

### 2. Export your data from atWork (iOS)
1. Open **atWork** on your iPhone/iPad.
2. Go to **Export** -> Select **CSV** (either Excel or Numbers format) covering your full date range.
3. Make sure **Vacations**, **Sick Days**, and **Holidays** are enabled in the export options if you want to keep your hour balance intact.
4. Place the exported `.csv` file in the same folder as `atwork_to_kimai.py`.

### 3. Generate a Kimai API Token
1. Log into your self-hosted **Kimai** web interface.
2. Click your avatar (top right) -> **My Profile** -> **API**.
3. Create a new **API Token** (API Password) and copy it.

### 4. Run the Migration
Simply run the script with Python 3. It will auto-discover your CSV file and prompt you for your Kimai URL and API token interactively:

```bash
python atwork_to_kimai.py
```

---

## ⚙️ Advanced CLI Usage

You can also customize the hierarchy names, daily absence hours, or run a simulation (`--dry-run`) via command-line arguments:

```bash
# Test parsing without uploading anything to Kimai
python atwork_to_kimai.py --file my_export.csv --dry-run

# Full automated execution with custom Spanish labels and 8h absences
python atwork_to_kimai.py \
  --file my_export.csv \
  --url "http://192.168.1.50:8001" \
  --token "YOUR_API_TOKEN" \
  --customer "Mi Empresa" \
  --project "Control Horario" \
  --work-activity "Trabajo ordinario" \
  --vacation-activity "Vacaciones" \
  --sick-activity "Baja médica" \
  --holiday-activity "Festivo" \
  --absence-start "06:30" \
  --absence-hours 8.0
```

### All Available Options

| Flag | Default | Description |
| :--- | :--- | :--- |
| `-f`, `--file` | *(Auto-detected)* | Path to the exported atWork CSV file |
| `-u`, `--url` | *(Prompted)* | Kimai base URL (e.g., `https://kimai.domain.com`) |
| `-t`, `--token` | *(Prompted)* | Kimai API Token |
| `--customer` | `My Company` | Default Customer name created in Kimai |
| `--project` | `Time Tracking` | Default Project name created in Kimai |
| `--work-activity` | `Regular Work` | Activity name for regular clock-in/out entries |
| `--vacation-activity` | `Vacation` | Activity name for vacation days |
| `--sick-activity` | `Sick Leave` | Activity name for sick leave days |
| `--holiday-activity` | `Public Holiday` | Activity name for public holidays |
| `--absence-start` | `06:30` | Default start time for full-day absences (`HH:MM`) |
| `--absence-hours` | `8.0` | Hours credited per full absence day |
| `--currency` | `EUR` | Default Customer currency (`EUR`, `USD`, etc.) |
| `--timezone` | `Europe/Madrid` | Default Customer timezone |
| `--country` | `ES` | Default 2-letter country code |
| `--dry-run` | `False` | Parse CSV and display summary without uploading |

---

## 👤 Author

**BerniBoom22**
- GitHub: [@BerniBoom22](https://github.com/BerniBoom22)

## 📄 License

This project is licensed under the **MIT License** — free to use, modify, and distribute.
