# Job Crawler (Jobinja + Jobvision)

A Python project for crawling job listings from:

- **Jobinja**
- **Jobvision**

and exporting the results to **Excel (.xlsx)** with full PDP (job detail page) data.

---

## Features

- Crawl paginated job listing pages
- Crawl each job detail page (PDP) for full metadata
- Export separate Excel files per source
- Configurable via CLI arguments
- Optional `.env`-based authentication for Jobinja

---

## Project Structure

```text
crawl/
├── jobinja_crawler.py       # Jobinja crawler
├── jobvision_crawler.py     # Jobvision crawler
├── requirements.txt
├── .env.example
├── .gitignore
└── output/                  # Generated Excel outputs
```

---

## Requirements

- Python 3.10+
- Internet access

Install dependencies:

```bash
pip install -r requirements.txt
```

Windows alternative:

```powershell
python -m pip install -r requirements.txt
```

---

## Authentication (Jobinja only)

Create a `.env` file if you want authenticated crawling on Jobinja:

```env
JOBINJA_EMAIL=your_email@example.com
JOBINJA_PASSWORD=your_password
```

A sample file is included as `.env.example`.

> **Security note:** Never commit `.env` to GitHub (it is ignored by `.gitignore`).

---

## Run Jobinja Crawler

### Full crawl (list + PDP)

```powershell
python jobinja_crawler.py
```

### Quick test (first page only)

```powershell
python jobinja_crawler.py --max-pages 1
```

### List-only mode (skip PDP)

```powershell
python jobinja_crawler.py --skip-details
```

### Custom output path

```powershell
python jobinja_crawler.py --output output/jobinja_custom.xlsx
```

---

## Run Jobvision Crawler

### Full crawl (list + PDP)

```powershell
python jobvision_crawler.py
```

### Quick test (first page only)

```powershell
python jobvision_crawler.py --max-pages 1 --output output/jobvision_test.xlsx
```

### Custom source URL

```powershell
python jobvision_crawler.py --url "https://jobvision.ir/jobs/keyword/Product%20Manager?page=1&sort=0"
```

---

## Output

Generated Excel files are saved in `output/`.

Example filenames:

- `jobinja_product_manager_YYYYMMDD_HHMMSS.xlsx`
- `jobvision_product_manager_YYYYMMDD_HHMMSS.xlsx`

---

## Example Output Fields

### Jobinja

- Title
- Company
- Location
- Contract/Employment Type
- Salary
- Skills
- Minimum Experience
- Minimum Education
- Job Description
- Company Description
- Job URL

### Jobvision

- Job ID
- Title
- Company
- Industry
- Company Size
- Salary
- Work Type
- Seniority Level
- Skills
- Benefits
- Job Description
- Publish/Expire Dates
- Job URL

---

## Notes

- Website structure/API may change over time; selectors or API mappings might need updates.
- A request delay is used to reduce pressure on source websites.
- If `python` is not available in your shell, use `py` instead.

---

## License

Intended for personal/educational use. Review target website terms before high-volume or commercial use.
