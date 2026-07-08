# Job Crawler Suite

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-Educational-lightgrey.svg)](#license)

A Python toolkit for crawling job listings from multiple Iranian job platforms and exporting structured data to **Excel (`.xlsx`)**.

Each source has its own standalone script and produces a **separate output file**, so you can run crawlers independently without mixing data.

---

## Table of Contents

- [Supported Sources](#supported-sources)
- [Features](#features)
- [How It Works](#how-it-works)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage](#usage)
  - [Jobinja](#jobinja-crawler)
  - [Jobvision](#jobvision-crawler)
  - [Snappfood Careers](#snappfood-careers-crawler)
- [Output Files](#output-files)
- [Output Schema](#output-schema)
- [Authentication](#authentication)
- [Configuration Tips](#configuration-tips)
- [Troubleshooting](#troubleshooting)
- [Project Structure](#project-structure)
- [Limitations & Responsible Use](#limitations--responsible-use)
- [License](#license)

---

## Supported Sources

| Platform | Script | Default Target | Auth Required |
|----------|--------|----------------|---------------|
| [Jobinja](https://jobinja.ir) | `jobinja_crawler.py` | Product Manager keyword search | Optional (recommended for salary on list page) |
| [Jobvision](https://jobvision.ir) | `jobvision_crawler.py` | Product Manager keyword search | No |
| [Snappfood Careers](https://careers.snappfood.ir) | `snappfood_crawler.py` | All open Snappfood positions | No |

---

## Features

- **List crawling** with automatic pagination
- **PDP (job detail page) enrichment** for full job metadata
- **Separate Excel export** per platform
- **CLI-based configuration** (URL, keyword, page limits, delays, output path)
- **Optional login support** for Jobinja (`.env`, CLI flags, or browser cookies)
- **Polite crawling** with configurable request delays
- **Persian-friendly output columns** in Excel

---

## How It Works

### Jobinja
1. Crawls search result pages (HTML)
2. Extracts basic listing data (title, company, location, contract type, salary, publish date, URL)
3. Optionally logs in to read salary from list pages
4. Visits each job detail page (PDP) unless `--skip-details` is used
5. Exports all fields to Excel

### Jobvision
1. Calls Jobvision public API for paginated job search results
2. Fetches full PDP data per job via detail API
3. Exports enriched records to Excel

### Snappfood Careers
1. Calls HRCando career API used by [careers.snappfood.ir](https://careers.snappfood.ir/)
2. Fetches full PDP data per job (description, requirements, department, etc.)
3. Exports enriched records to Excel

---

## Requirements

- **Python 3.10+**
- Internet connection
- Windows / macOS / Linux

### Python Dependencies

| Package | Purpose |
|---------|---------|
| `requests` | HTTP requests |
| `beautifulsoup4` | HTML parsing |
| `lxml` | Fast HTML parser backend |
| `pandas` | Data processing |
| `openpyxl` | Excel export |
| `python-dotenv` | Load `.env` credentials (Jobinja) |

---

## Installation

### 1) Clone the repository

```bash
git clone <your-repo-url>
cd crawl
```

### 2) Install dependencies

```bash
pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m pip install -r requirements.txt
```

### 3) (Optional) Configure Jobinja credentials

Copy `.env.example` to `.env` and fill in your account:

```env
JOBINJA_EMAIL=your_email@example.com
JOBINJA_PASSWORD=your_password
```

> Never commit `.env` to GitHub. It is already ignored by `.gitignore`.

---

## Quick Start

Run each crawler independently:

```powershell
# Jobinja (list + PDP)
python jobinja_crawler.py

# Jobvision (list + PDP)
python jobvision_crawler.py

# Snappfood careers (list + PDP)
python snappfood_crawler.py
```

Output files are written to the `output/` directory.

---

## Usage

## Jobinja Crawler

### Recommended commands

**Full crawl (list + PDP):**
```powershell
python jobinja_crawler.py
```

**Faster crawl (login + list salary, no PDP):**
```powershell
python jobinja_crawler.py --skip-details
```

**Quick test (first page only):**
```powershell
python jobinja_crawler.py --max-pages 1
```

**Custom search URL:**
```powershell
python jobinja_crawler.py --url "https://jobinja.ir/jobs?filters%5Bkeywords%5D%5B0%5D=product%20manager&sort_by=published_at_desc"
```

**Custom output path:**
```powershell
python jobinja_crawler.py --output output/jobinja_custom.xlsx
```

### CLI options

| Flag | Default | Description |
|------|---------|-------------|
| `--url` | Product Manager search URL | Jobinja search results URL |
| `--output` | `output/jobinja_product_manager_<timestamp>.xlsx` | Output Excel path |
| `--max-pages` | all pages | Limit number of result pages |
| `--delay` | `1.0` | Delay between requests (seconds) |
| `--skip-details` | off | Skip PDP crawling |
| `--email` | from `.env` | Jobinja login email |
| `--password` | from `.env` | Jobinja login password |
| `--cookies-file` | none | Netscape cookie file for login |

---

## Jobvision Crawler

### Recommended commands

**Full crawl (list + PDP):**
```powershell
python jobvision_crawler.py
```

**Quick test (first page only):**
```powershell
python jobvision_crawler.py --max-pages 1 --output output/jobvision_test.xlsx
```

**Custom keyword URL:**
```powershell
python jobvision_crawler.py --url "https://jobvision.ir/jobs/keyword/Product%20Manager?page=1&sort=0"
```

**Persian keyword example:**
```powershell
python jobvision_crawler.py --url "https://jobvision.ir/jobs/keyword/%D9%85%D8%AF%DB%8C%D8%B1%20%D9%85%D8%AD%D8%B5%D9%88%D9%84?page=1&sort=0"
```

### CLI options

| Flag | Default | Description |
|------|---------|-------------|
| `--url` | Product Manager search URL | Jobvision search URL |
| `--output` | `output/jobvision_product_manager_<timestamp>.xlsx` | Output Excel path |
| `--max-pages` | all pages | Limit number of result pages |
| `--page-size` | `10` | API page size |
| `--delay` | `0.5` | Delay between requests (seconds) |

---

## Snappfood Careers Crawler

### Recommended commands

**Full crawl (all open jobs + PDP):**
```powershell
python snappfood_crawler.py
```

**Quick test (first page only):**
```powershell
python snappfood_crawler.py --max-pages 1 --output output/snappfood_test.xlsx
```

**Filter by title keyword:**
```powershell
python snappfood_crawler.py --keyword "Product Manager"
```

**Persian keyword example:**
```powershell
python snappfood_crawler.py --keyword "مدیر محصول"
```

### CLI options

| Flag | Default | Description |
|------|---------|-------------|
| `--keyword` | empty (all jobs) | Filter jobs by title |
| `--output` | `output/snappfood_careers_<timestamp>.xlsx` | Output Excel path |
| `--max-pages` | all pages | Limit number of result pages |
| `--page-size` | `50` | Jobs per list page |
| `--delay` | `0.5` | Delay between requests (seconds) |

---

## Output Files

All crawlers save Excel files in:

```text
output/
```

Example filenames:

- `jobinja_product_manager_20260707_091938.xlsx`
- `jobvision_product_manager_20260707_104249.xlsx`
- `snappfood_careers_20260707_153012.xlsx`

Each file is standalone and source-specific.

---

## Output Schema

### Jobinja columns

| Column | Description |
|--------|-------------|
| عنوان | Job title |
| شرکت | Company name |
| موقعیت | Location |
| نوع قرارداد | Contract type |
| حقوق | Salary |
| تاریخ انتشار | Relative publish time (e.g. today, 2 days ago) |
| لینک | Job URL |
| دسته‌بندی شغلی | Job category |
| نوع همکاری | Employment type |
| حداقل سابقه کار | Minimum experience |
| مهارت‌ها | Required skills |
| جنسیت | Gender requirement |
| نظام وظیفه | Military service status |
| حداقل مدرک تحصیلی | Minimum education |
| صنعت شرکت | Company industry |
| اندازه شرکت | Company size |
| وب‌سایت شرکت | Company website |
| لینک شرکت | Company profile URL |
| شرح موقعیت شغلی | Full job description |
| معرفی شرکت | Company description |
| تاریخ انتشار آگهی | Exact publish date (from PDP/JSON-LD) |

### Jobvision columns

| Column | Description |
|--------|-------------|
| شناسه آگهی | Job ID |
| عنوان | Job title |
| شرکت | Company |
| صنعت | Industry |
| اندازه شرکت | Company size |
| وب‌سایت شرکت | Company website |
| لینک شرکت | Company profile URL |
| موقعیت | Location |
| نوع همکاری | Work type |
| سطح ارشدیت | Seniority level |
| حقوق | Salary |
| جنسیت | Gender requirement |
| امکان دورکاری | Remote availability |
| کارآموزی | Internship flag |
| مناسب معلولین | Suitable for disabled candidates |
| سابقه کار (سال) | Required experience years |
| بازه سنی | Age range |
| نیاز به خدمت سربازی | Military service requirement |
| روزها/ساعات کاری | Working days/hours |
| مزایا | Benefits |
| مهارت‌ها | Skills |
| دسته‌بندی شغلی | Job categories |
| برچسب‌ها | Labels/tags |
| شرح شغل | Full job description |
| تاریخ انتشار | Publish date |
| تاریخ انقضا | Expiration date |
| لینک آگهی | Job URL |

### Snappfood columns

| Column | Description |
|--------|-------------|
| شناسه آگهی | Numeric job ID |
| شناسه GUID | Job GUID |
| عنوان | Job title |
| شرکت | Company name |
| وب‌سایت شرکت | Company website |
| آدرس شرکت | Company address |
| شعبه | Branch title |
| دپارتمان | Department |
| دسته‌بندی شغلی | Job category |
| شهر | City |
| منطقه | Region |
| نوع همکاری | Work type |
| سطح ارشدیت | Seniority level |
| جنسیت | Preferred gender |
| حقوق | Salary range (if available) |
| امکان دورکاری | Remote availability |
| کارآموزی | Internship flag |
| سابقه کار (سال) | Required experience years |
| بازه سنی | Age range |
| روزها/ساعات کاری | Working days/hours |
| شرایط | Conditions |
| نیازمندی‌ها | Requirements |
| شرح شغل | Full job description |
| تاریخ انتشار | Publish datetime |
| لینک آگهی | Job detail URL |

---

## Authentication

Only **Jobinja** supports authenticated crawling.

### Option 1: `.env` file (recommended)

```env
JOBINJA_EMAIL=your_email@example.com
JOBINJA_PASSWORD=your_password
```

Then run:

```powershell
python jobinja_crawler.py --skip-details
```

### Option 2: CLI flags

```powershell
python jobinja_crawler.py --email "your@email.com" --password "your_password" --skip-details
```

### Option 3: Browser cookies

Export cookies from your browser (Netscape format) and run:

```powershell
python jobinja_crawler.py --cookies-file cookies.txt
```

### Salary behavior on Jobinja

- Without login: salary may be hidden on list pages
- With login: salary is visible on list pages (faster)
- With PDP mode: description and extra metadata are fetched from each job page

---

## Configuration Tips

### Change target role/keyword

- **Jobinja / Jobvision:** pass a custom search URL via `--url`
- **Snappfood:** pass `--keyword "your keyword"`

### Speed vs completeness

| Goal | Suggested command |
|------|-------------------|
| Fastest Jobinja export | `python jobinja_crawler.py --skip-details` (with login) |
| Richest Jobinja data | `python jobinja_crawler.py` (with PDP) |
| Test before full run | add `--max-pages 1` |
| Lower server load | increase `--delay` |

### Estimated runtime (approx.)

| Crawler | Typical volume | Estimated time |
|---------|----------------|----------------|
| Jobinja (skip details) | ~200 jobs | ~30 seconds |
| Jobinja (with PDP) | ~200 jobs | ~5–10 minutes |
| Jobvision (with PDP) | ~200 jobs | ~3–6 minutes |
| Snappfood (with PDP) | ~70 jobs | ~1–2 minutes |

---

## Troubleshooting

### `python` is not recognized

Try:

```powershell
py jobinja_crawler.py
```

### Unicode/console encoding issues on Windows

This does not affect Excel output. Files are saved as UTF-8 Excel sheets.

### Jobinja login failed

- Verify email/password in `.env`
- Check if account is active
- Try `--cookies-file` instead

### Empty results

- Verify search URL/keyword
- Site structure or API may have changed
- Try `--max-pages 1` first for debugging

### Missing salary values

- Some employers publish "negotiable" (`توافقی`) instead of numeric salary
- Jobinja salary on list page usually requires login

---

## Project Structure

```text
crawl/
├── jobinja_crawler.py        # Jobinja crawler (HTML + optional login)
├── jobvision_crawler.py      # Jobvision crawler (API-based)
├── snappfood_crawler.py      # Snappfood careers crawler (API-based)
├── requirements.txt          # Python dependencies
├── .env.example              # Sample environment file
├── .gitignore                # Ignores .env and output/
├── README.md                 # Project documentation
└── output/                   # Generated Excel files (ignored by git)
```

---

## Limitations & Responsible Use

- Website HTML structures and APIs can change; crawlers may require maintenance.
- This project is intended for **personal, educational, and research use**.
- Respect each platform's terms of service and robots/access policies.
- Use reasonable delays (`--delay`) and avoid aggressive scraping.
- Do not commit credentials (`.env`) or sensitive account data to GitHub.

---

## License

Educational/personal use.

Before commercial or high-volume usage, review and comply with the terms of:

- [Jobinja](https://jobinja.ir)
- [Jobvision](https://jobvision.ir)
- [Snappfood Careers](https://careers.snappfood.ir)

---

## Example Workflow

```powershell
# 1) Install
python -m pip install -r requirements.txt

# 2) Configure Jobinja auth (optional)
copy .env.example .env

# 3) Run all three crawlers
python jobinja_crawler.py
python jobvision_crawler.py
python snappfood_crawler.py

# 4) Check output folder
dir output
```

If you find this project useful, consider starring the repository on GitHub.
