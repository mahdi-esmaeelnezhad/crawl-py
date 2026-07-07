"""Crawl Jobvision jobs (list + PDP details) and export to Excel."""

from __future__ import annotations

import argparse
import html
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

DEFAULT_URL = "https://jobvision.ir/jobs/keyword/Product%20Manager?page=1&sort=0"
LIST_API = "https://candidateapi.jobvision.ir/api/v1/JobPost/List"
DETAIL_API = "https://candidateapi.jobvision.ir/api/v1/JobPost/Detail"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)
REQUEST_DELAY_SECONDS = 0.5


@dataclass
class JobvisionRow:
    job_id: int
    title: str
    company: str
    industry: str
    company_size: str
    company_website: str
    company_link: str
    location: str
    work_type: str
    seniority_level: str
    salary: str
    gender: str
    is_remote: str
    is_internship: str
    suitable_for_disabled: str
    required_experience_years: str
    age_range: str
    military_service_required: str
    work_days: str
    benefits: str
    skills: str
    job_categories: str
    labels: str
    description: str
    activation_time: str
    expire_date: str
    source_url: str


def text_fa(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("titleFa", "title", "titleEn", "beautifyFa", "date"):
            if value.get(key):
                return str(value[key]).strip()
        return ""
    return str(value).strip()


def clean_html_text(raw_html: str | None) -> str:
    if not raw_html:
        return ""
    soup = BeautifulSoup(raw_html, "lxml")
    text = soup.get_text(" ", strip=True)
    text = html.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept": "application/json"})
    return session


def parse_keyword_from_url(url: str) -> str:
    parsed = urlparse(url)
    parts = [p for p in parsed.path.split("/") if p]
    if "keyword" in parts:
        idx = parts.index("keyword")
        if idx + 1 < len(parts):
            return unquote(parts[idx + 1]).strip()
    return ""


def parse_sort_from_url(url: str) -> int:
    parsed = urlparse(url)
    for part in parsed.query.split("&"):
        if part.startswith("sort="):
            try:
                return int(part.split("=", 1)[1])
            except ValueError:
                return 0
    return 0


def list_jobs(
    session: requests.Session,
    keyword: str,
    sort: int,
    page: int,
    page_size: int,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "page": page,
        "pageSize": page_size,
        "sort": sort,
        "filters": {"keyword": keyword},
    }
    response = session.post(LIST_API, json=payload, timeout=30)
    response.raise_for_status()
    body = response.json()
    if not body.get("isSuccess"):
        raise RuntimeError(f"Jobvision list failed: {body.get('message')}")
    return body.get("data", {})


def get_job_detail(session: requests.Session, job_post_id: int) -> dict[str, Any]:
    response = session.get(DETAIL_API, params={"jobPostId": job_post_id}, timeout=30)
    response.raise_for_status()
    body = response.json()
    if not body.get("isSuccess"):
        raise RuntimeError(f"Jobvision detail failed for {job_post_id}: {body.get('message')}")
    return body.get("data", {})


def to_bool_fa(value: Any) -> str:
    return "بله" if bool(value) else "خیر"


def compose_location(location: dict[str, Any] | None) -> str:
    if not location:
        return ""
    province = text_fa(location.get("province"))
    city = text_fa(location.get("city"))
    region = text_fa(location.get("region"))
    parts = [p for p in (province, city, region) if p]
    return "، ".join(parts)


def map_detail_to_row(detail: dict[str, Any]) -> JobvisionRow:
    company = detail.get("company") or {}
    company_name = text_fa(company.get("name"))
    company_link = company.get("companyLink") or ""
    if company_link and company_link.startswith("/"):
        company_link = f"https://jobvision.ir{company_link}"

    industries = company.get("industries") or []
    company_industry = "، ".join(text_fa(item) for item in industries if text_fa(item))
    if not company_industry:
        company_industry = text_fa(detail.get("industry"))

    skills = detail.get("skills") or []
    benefits = detail.get("benefits") or []
    categories = detail.get("jobCategories") or []
    labels = detail.get("labels") or []

    age_min = detail.get("requiredAgeMin")
    age_max = detail.get("requiredAgeMax")
    age_range = ""
    if age_min is not None or age_max is not None:
        age_range = f"{age_min or '-'} تا {age_max or '-'}"

    salary = detail.get("salary")
    salary_text = text_fa(salary) if salary else ""

    return JobvisionRow(
        job_id=int(detail.get("id", 0)),
        title=text_fa(detail.get("title")),
        company=company_name,
        industry=company_industry,
        company_size=text_fa((company.get("size") or {}).get("titleFa")),
        company_website=str(company.get("website") or ""),
        company_link=company_link,
        location=compose_location(detail.get("location")),
        work_type=text_fa(detail.get("workType")),
        seniority_level=text_fa(detail.get("seniorityLevel")),
        salary=salary_text,
        gender=text_fa(detail.get("gender")),
        is_remote=to_bool_fa(detail.get("isRemote")),
        is_internship=to_bool_fa(detail.get("isInternship")),
        suitable_for_disabled=to_bool_fa(detail.get("suitableForDisabled")),
        required_experience_years=str(detail.get("requiredRelatedExperienceYears") or ""),
        age_range=age_range,
        military_service_required=to_bool_fa(detail.get("shouldDoneMilitaryService")),
        work_days=str(detail.get("workDays") or ""),
        benefits="، ".join(text_fa(item) for item in benefits if text_fa(item)),
        skills="، ".join(text_fa(item) for item in skills if text_fa(item)),
        job_categories="، ".join(text_fa(item) for item in categories if text_fa(item)),
        labels="، ".join(str(item) for item in labels if str(item).strip()),
        description=clean_html_text(detail.get("description")),
        activation_time=text_fa(detail.get("activationTime")),
        expire_date=text_fa((detail.get("expireTime") or {}).get("date")),
        source_url=f"https://jobvision.ir/jobs/id/{detail.get('id')}",
    )


def crawl_jobvision(
    base_url: str,
    max_pages: int | None,
    page_size: int,
    delay: float,
) -> list[JobvisionRow]:
    keyword = parse_keyword_from_url(base_url)
    if not keyword:
        raise ValueError("Keyword not found in URL. Expected /jobs/keyword/<keyword>")

    sort = parse_sort_from_url(base_url)
    session = build_session()

    first = list_jobs(session, keyword=keyword, sort=sort, page=1, page_size=page_size)
    first_jobs = first.get("jobPosts") or []
    total_count = int(first.get("jobPostCount") or 0)
    total_pages = max(1, (total_count + page_size - 1) // page_size)
    if max_pages:
        total_pages = min(total_pages, max_pages)

    all_items = list(first_jobs)
    print(f"List page 1/{total_pages}: {len(first_jobs)} jobs")

    for page in range(2, total_pages + 1):
        time.sleep(delay)
        data = list_jobs(session, keyword=keyword, sort=sort, page=page, page_size=page_size)
        page_jobs = data.get("jobPosts") or []
        all_items.extend(page_jobs)
        print(f"List page {page}/{total_pages}: {len(page_jobs)} jobs")

    seen_ids: set[int] = set()
    rows: list[JobvisionRow] = []
    total = len(all_items)

    for idx, item in enumerate(all_items, start=1):
        job_id = int(item.get("id", 0))
        if not job_id or job_id in seen_ids:
            continue
        seen_ids.add(job_id)

        time.sleep(delay)
        detail = get_job_detail(session, job_id)
        rows.append(map_detail_to_row(detail))
        print(f"PDP {idx}/{total}: id={job_id}")

    return rows


def export_excel(rows: list[JobvisionRow], output_path: Path) -> None:
    records = [
        {
            "شناسه آگهی": r.job_id,
            "عنوان": r.title,
            "شرکت": r.company,
            "صنعت": r.industry,
            "اندازه شرکت": r.company_size,
            "وب‌سایت شرکت": r.company_website,
            "لینک شرکت": r.company_link,
            "موقعیت": r.location,
            "نوع همکاری": r.work_type,
            "سطح ارشدیت": r.seniority_level,
            "حقوق": r.salary,
            "جنسیت": r.gender,
            "امکان دورکاری": r.is_remote,
            "کارآموزی": r.is_internship,
            "مناسب معلولین": r.suitable_for_disabled,
            "سابقه کار (سال)": r.required_experience_years,
            "بازه سنی": r.age_range,
            "نیاز به خدمت سربازی": r.military_service_required,
            "روزها/ساعات کاری": r.work_days,
            "مزایا": r.benefits,
            "مهارت‌ها": r.skills,
            "دسته‌بندی شغلی": r.job_categories,
            "برچسب‌ها": r.labels,
            "شرح شغل": r.description,
            "تاریخ انتشار": r.activation_time,
            "تاریخ انقضا": r.expire_date,
            "لینک آگهی": r.source_url,
        }
        for r in rows
    ]
    df = pd.DataFrame(records)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")


def main() -> None:
    parser = argparse.ArgumentParser(description="Crawl Jobvision with PDP details")
    parser.add_argument("--url", default=DEFAULT_URL, help="Jobvision search URL")
    parser.add_argument("--max-pages", type=int, default=None, help="Limit list pages")
    parser.add_argument("--page-size", type=int, default=10, help="API page size")
    parser.add_argument("--delay", type=float, default=REQUEST_DELAY_SECONDS, help="Delay between requests")
    parser.add_argument("--output", default="", help="Output excel path")
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = Path(args.output) if args.output else Path("output") / f"jobvision_product_manager_{timestamp}.xlsx"

    print("Starting Jobvision crawl...")
    rows = crawl_jobvision(args.url, args.max_pages, args.page_size, args.delay)
    if not rows:
        print("No jobs found for this keyword/url.")
        return

    export_excel(rows, output)
    print(f"Done! {len(rows)} jobs saved to: {output.resolve()}")


if __name__ == "__main__":
    main()
