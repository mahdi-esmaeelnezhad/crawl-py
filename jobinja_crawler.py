"""Crawl Jobinja job listings and export to Excel."""

from __future__ import annotations

import argparse
import json
import os
import re
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from http.cookiejar import MozillaCookieJar
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

DEFAULT_URL = (
    "https://jobinja.ir/jobs?"
    "filters%5Bkeywords%5D%5B0%5D=product%20manger&"
    "sort_by=published_at_desc"
)
LOGIN_URL = "https://jobinja.ir/login/user"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)
REQUEST_DELAY_SECONDS = 1.0
LOGIN_REQUIRED_TEXT = "برای مشاهده حقوق وارد شوید"

EXCEL_COLUMNS = [
    "عنوان",
    "شرکت",
    "موقعیت",
    "نوع قرارداد",
    "حقوق",
    "تاریخ انتشار",
    "لینک",
    "دسته‌بندی شغلی",
    "نوع همکاری",
    "حداقل سابقه کار",
    "مهارت‌ها",
    "جنسیت",
    "نظام وظیفه",
    "حداقل مدرک تحصیلی",
    "صنعت شرکت",
    "اندازه شرکت",
    "وب‌سایت شرکت",
    "لینک شرکت",
    "شرح موقعیت شغلی",
    "معرفی شرکت",
    "تاریخ انتشار آگهی",
]


@dataclass
class JobListing:
    title: str = ""
    company: str = ""
    location: str = ""
    contract_type: str = ""
    salary: str = ""
    published_at: str = ""
    url: str = ""
    category: str = ""
    employment_type: str = ""
    min_experience: str = ""
    skills: str = ""
    gender: str = ""
    military_status: str = ""
    min_education: str = ""
    company_industry: str = ""
    company_size: str = ""
    company_website: str = ""
    company_url: str = ""
    job_description: str = ""
    company_description: str = ""
    date_posted: str = ""


def build_page_url(base_url: str, page: int) -> str:
    parsed = urlparse(base_url)
    query = parse_qs(parsed.query, keep_blank_values=True)
    if page > 1:
        query["page"] = [str(page)]
    else:
        query.pop("page", None)
    new_query = urlencode(query, doseq=True)
    return urlunparse(parsed._replace(query=new_query))


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def create_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT})
    return session


def fetch_page(session: requests.Session, url: str) -> BeautifulSoup:
    response = session.get(url, timeout=30)
    response.raise_for_status()
    response.encoding = "utf-8"
    return BeautifulSoup(response.text, "lxml")


def login(session: requests.Session, email: str, password: str) -> None:
    login_page = fetch_page(session, LOGIN_URL)
    form = login_page.select_one("form")
    if not form:
        raise RuntimeError("Could not find Jobinja login form.")

    token_input = form.select_one('input[name="_token"]')
    if not token_input or not token_input.get("value"):
        raise RuntimeError("Could not find CSRF token on login page.")

    payload = {
        "_token": token_input["value"],
        "redirect_url": "",
        "identifier": email,
        "password": password,
        "remember_me": "1",
    }
    response = session.post(LOGIN_URL, data=payload, timeout=30, allow_redirects=True)
    response.raise_for_status()

    if "login/user" in response.url and "identifier" in response.text:
        raise RuntimeError("Login failed. Check your email and password.")

    print("Logged in successfully.")


def load_cookies(session: requests.Session, cookies_file: Path) -> None:
    jar = MozillaCookieJar(str(cookies_file))
    jar.load(ignore_discard=True, ignore_expires=True)
    session.cookies.update(jar)
    print(f"Loaded cookies from: {cookies_file}")


def parse_total_jobs(soup: BeautifulSoup) -> int | None:
    heading = soup.select_one("h3")
    if not heading:
        return None
    match = re.search(r"(\d+)", heading.get_text(strip=True).replace(",", ""))
    return int(match.group(1)) if match else None


def parse_max_page(soup: BeautifulSoup) -> int:
    max_page = 1
    for link in soup.select('a[href*="page="]'):
        href = link.get("href", "")
        match = re.search(r"page=(\d+)", href)
        if match:
            max_page = max(max_page, int(match.group(1)))
    return max_page


def normalize_salary_text(salary: str) -> str:
    salary = salary.strip().strip("()")
    salary = re.sub(r"^حقوق\s+", "", salary)
    return clean_text(salary)


def parse_contract_and_salary(meta_span) -> tuple[str, str]:
    login_link = meta_span.select_one("a")
    if login_link and LOGIN_REQUIRED_TEXT in login_link.get_text():
        return "", ""

    inner_spans = meta_span.select(":scope > span")
    if len(inner_spans) >= 2:
        contract_type = clean_text(inner_spans[0].get_text())
        salary = normalize_salary_text(inner_spans[1].get_text())
        return contract_type, salary

    contract_span = meta_span.select_one("span")
    contract_type = clean_text(contract_span.get_text()) if contract_span else ""
    salary = normalize_salary_text(meta_span.get_text().replace(contract_type, "", 1))
    return contract_type, salary


def parse_info_box(soup: BeautifulSoup) -> dict[str, str]:
    info: dict[str, str] = {}
    for item in soup.select("li.c-infoBox__item"):
        title_el = item.select_one("h4.c-infoBox__itemTitle")
        if not title_el:
            continue
        title = clean_text(title_el.get_text())
        values = [clean_text(span.get_text()) for span in item.select("span.black")]
        info[title] = "، ".join(value for value in values if value)
    return info


def parse_section_text(soup: BeautifulSoup, section_title: str) -> str:
    for heading in soup.select("h4.o-box__title"):
        if section_title not in heading.get_text():
            continue
        section = heading.find_next_sibling("div", class_=re.compile(r"o-box__text"))
        if not section:
            continue
        for unwanted in section.select(".no-desc-display"):
            unwanted.decompose()
        return clean_text(section.get_text(separator=" ", strip=True))
    return ""


def parse_company_header(soup: BeautifulSoup) -> dict[str, str]:
    header = {
        "company_name": "",
        "company_industry": "",
        "company_size": "",
        "company_website": "",
        "company_url": "",
    }

    name_el = soup.select_one("h2.c-companyHeader__name")
    if name_el:
        header["company_name"] = clean_text(name_el.get_text())

    logo_link = soup.select_one("a.c-companyHeader__logoLink")
    if logo_link and logo_link.get("href"):
        header["company_url"] = logo_link["href"].split("?")[0]

    for meta_item in soup.select(".c-companyHeader__metaItem"):
        link = meta_item.select_one("a.c-companyHeader__metaLink")
        text = clean_text(meta_item.get_text())
        if not text:
            continue
        if link and "jobinja.ir" not in link.get("href", ""):
            header["company_website"] = text
        elif link:
            header["company_industry"] = clean_text(link.get_text())
        elif "نفر" in text:
            header["company_size"] = text

    return header


def parse_json_ld(soup: BeautifulSoup) -> dict:
    for script in soup.select('script[type="application/ld+json"]'):
        if not script.string:
            continue
        try:
            data = json.loads(script.string)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("@type") == "JobPosting":
            return data
    return {}


def parse_salary_from_json_ld(data: dict) -> str:
    base_salary = data.get("baseSalary")
    if not isinstance(base_salary, dict):
        return ""
    value = base_salary.get("value")
    if value in (None, "", 0, "0"):
        return ""
    if isinstance(value, (int, float)) and value == 5_000_000:
        return ""
    unit = base_salary.get("unitText", "MONTH")
    currency = base_salary.get("currency", "IRT")
    return f"{int(value):,} {currency} / {unit}"


def parse_pdp_page(soup: BeautifulSoup) -> dict[str, str]:
    info = parse_info_box(soup)
    company = parse_company_header(soup)
    json_ld = parse_json_ld(soup)

    salary = info.get("حقوق", "")
    if not salary or salary == "توافقی":
        salary = parse_salary_from_json_ld(json_ld) or salary

    title_el = soup.select_one("h1")
    pdp_title = clean_text(title_el.get_text()) if title_el else ""

    return {
        "title": pdp_title,
        "company": company.get("company_name", ""),
        "location": info.get("موقعیت مکانی", ""),
        "salary": salary,
        "category": info.get("دسته‌بندی شغلی", ""),
        "employment_type": info.get("نوع همکاری", ""),
        "min_experience": info.get("حداقل سابقه کار", ""),
        "skills": info.get("مهارت‌های مورد نیاز", ""),
        "gender": info.get("جنسیت", ""),
        "military_status": info.get("وضعیت نظام وظیفه", ""),
        "min_education": info.get("حداقل مدرک تحصیلی", ""),
        "company_industry": company.get("company_industry", ""),
        "company_size": company.get("company_size", ""),
        "company_website": company.get("company_website", ""),
        "company_url": company.get("company_url", ""),
        "job_description": parse_section_text(soup, "شرح موقعیت شغلی"),
        "company_description": parse_section_text(soup, "معرفی شرکت"),
        "date_posted": str(json_ld.get("datePosted", "")),
    }


def apply_pdp_details(job: JobListing, details: dict[str, str]) -> None:
    if details.get("title"):
        job.title = details["title"]
    if details.get("company"):
        job.company = details["company"]
    if details.get("location"):
        job.location = details["location"]

    pdp_salary = details.get("salary", "")
    if pdp_salary:
        if not job.salary or job.salary == "توافقی":
            job.salary = pdp_salary
        elif pdp_salary != "توافقی":
            job.salary = pdp_salary

    if details.get("employment_type") and not job.contract_type:
        job.contract_type = f"قرارداد {details['employment_type']}"

    job.category = details.get("category", "")
    job.employment_type = details.get("employment_type", "")
    job.min_experience = details.get("min_experience", "")
    job.skills = details.get("skills", "")
    job.gender = details.get("gender", "")
    job.military_status = details.get("military_status", "")
    job.min_education = details.get("min_education", "")
    job.company_industry = details.get("company_industry", "")
    job.company_size = details.get("company_size", "")
    job.company_website = details.get("company_website", "")
    job.company_url = details.get("company_url", "")
    job.job_description = details.get("job_description", "")
    job.company_description = details.get("company_description", "")
    job.date_posted = details.get("date_posted", "")


def parse_job_item(item) -> JobListing | None:
    title_link = item.select_one("h2.c-jobListView__title a.c-jobListView__titleLink")
    if not title_link:
        return None

    title = clean_text(title_link.get_text())
    url = title_link.get("href", "").split("?")[0]

    published_el = item.select_one(".c-jobListView__passedDays")
    published_at = clean_text(published_el.get_text()) if published_el else ""

    meta_items = item.select("ul.c-jobListView__meta > li.c-jobListView__metaItem > span")
    company = clean_text(meta_items[0].get_text()) if len(meta_items) > 0 else ""
    location = clean_text(meta_items[1].get_text()) if len(meta_items) > 1 else ""

    contract_type = ""
    salary = ""
    if len(meta_items) > 2:
        contract_type, salary = parse_contract_and_salary(meta_items[2])

    return JobListing(
        title=title,
        company=company,
        location=location,
        contract_type=contract_type,
        salary=salary,
        published_at=published_at,
        url=url,
    )


def parse_jobs(soup: BeautifulSoup) -> list[JobListing]:
    jobs: list[JobListing] = []
    for item in soup.select("ul.c-jobListView__list > li.c-jobListView__item"):
        job = parse_job_item(item)
        if job:
            jobs.append(job)
    return jobs


def enrich_jobs_with_pdp(
    session: requests.Session,
    jobs: list[JobListing],
    delay: float,
) -> None:
    total = len(jobs)
    for index, job in enumerate(jobs, start=1):
        if index > 1:
            time.sleep(delay)
        soup = fetch_page(session, job.url)
        apply_pdp_details(job, parse_pdp_page(soup))
        print(f"PDP {index}/{total} fetched")


def crawl_jobs(
    session: requests.Session,
    base_url: str,
    max_pages: int | None = None,
    delay: float = REQUEST_DELAY_SECONDS,
    fetch_details: bool = True,
) -> list[JobListing]:
    first_url = build_page_url(base_url, 1)
    first_soup = fetch_page(session, first_url)

    total_jobs = parse_total_jobs(first_soup)
    last_page = parse_max_page(first_soup)
    if max_pages is not None:
        last_page = min(last_page, max_pages)

    all_jobs: list[JobListing] = parse_jobs(first_soup)
    print(f"Page 1/{last_page}: {len(all_jobs)} jobs")

    for page in range(2, last_page + 1):
        time.sleep(delay)
        page_url = build_page_url(base_url, page)
        soup = fetch_page(session, page_url)
        page_jobs = parse_jobs(soup)
        all_jobs.extend(page_jobs)
        print(f"Page {page}/{last_page}: {len(page_jobs)} jobs")

    if total_jobs is not None:
        print(f"Expected ~{total_jobs} jobs, collected {len(all_jobs)}")

    if fetch_details:
        print("Fetching PDP details for each job...")
        enrich_jobs_with_pdp(session, all_jobs, delay)

    return all_jobs


def export_to_excel(jobs: list[JobListing], output_path: Path) -> None:
    df = pd.DataFrame([asdict(job) for job in jobs])
    df.columns = EXCEL_COLUMNS
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")


def resolve_auth(session: requests.Session, email: str, password: str, cookies_file: str) -> None:
    if cookies_file:
        load_cookies(session, Path(cookies_file))
        return

    resolved_email = email or os.environ.get("JOBINJA_EMAIL", "")
    resolved_password = password or os.environ.get("JOBINJA_PASSWORD", "")
    if resolved_email and resolved_password:
        login(session, resolved_email, resolved_password)


def main() -> None:
    load_dotenv()
    parser = argparse.ArgumentParser(description="Crawl Jobinja jobs and export to Excel")
    parser.add_argument("--url", default=DEFAULT_URL, help="Jobinja search URL")
    parser.add_argument(
        "--output",
        default="",
        help="Output Excel file path (default: output/jobinja_<timestamp>.xlsx)",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Limit number of pages to crawl (useful for testing)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=REQUEST_DELAY_SECONDS,
        help="Delay between page requests in seconds",
    )
    parser.add_argument(
        "--skip-details",
        action="store_true",
        help="Skip fetching PDP details from individual job pages",
    )
    parser.add_argument("--email", default="", help="Jobinja login email")
    parser.add_argument("--password", default="", help="Jobinja login password")
    parser.add_argument(
        "--cookies-file",
        default="",
        help="Netscape cookies file exported from browser",
    )
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = (
        Path(args.output)
        if args.output
        else Path("output") / f"jobinja_product_manager_{timestamp}.xlsx"
    )

    session = create_session()
    resolve_auth(session, args.email, args.password, args.cookies_file)

    print("Starting crawl...")
    jobs = crawl_jobs(
        session,
        args.url,
        max_pages=args.max_pages,
        delay=args.delay,
        fetch_details=not args.skip_details,
    )

    if not jobs:
        print("No jobs found. Check the URL or site structure.")
        return

    export_to_excel(jobs, output_path)
    print(f"Done! {len(jobs)} jobs saved to: {output_path.resolve()}")


if __name__ == "__main__":
    main()
