"""Crawl Snappfood careers (list + PDP details) and export to Excel."""

from __future__ import annotations

import argparse
import html
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from bs4 import BeautifulSoup

DEFAULT_URL = "https://careers.snappfood.ir/"
LIST_API = "https://careerapi.hrcando.ir/api/v1/CareerPageCompany/GetCompanyCareerPageJobList"
DETAIL_BY_GUID_API = (
    "https://careerapi.hrcando.ir/api/v1/CareerPage/GetCareerPageJobPageInfoByJobGuid"
)
DETAIL_BY_ID_API = "https://careerapi.hrcando.ir/api/v1/CareerPage/GetCareerPageJobPageInfo"
COMPANY_API_KEY = "4bdfd70c-49c5-42e0-a6ec-7622eec64670"
COMPANY_ADDRESS = "https://career.hrcando.ir/co/snappfood"
REQUEST_DELAY_SECONDS = 0.5


@dataclass
class SnappfoodRow:
    job_id: int
    job_guid: str
    title: str
    company: str
    company_website: str
    company_address: str
    branch_title: str
    department: str
    job_category: str
    city: str
    region: str
    work_type: str
    seniority_level: str
    preferred_gender: str
    salary: str
    is_remote: str
    is_internship: str
    required_experience_years: str
    age_range: str
    working_days: str
    description: str
    conditions: str
    requirements: str
    published_at: str
    source_url: str


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def clean_html_text(raw_html: str | None) -> str:
    if not raw_html:
        return ""
    soup = BeautifulSoup(raw_html, "lxml")
    text = soup.get_text(" ", strip=True)
    text = html.unescape(text)
    return clean_text(text)


def text_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("name", "title", "titleFa", "departmentTitle", "cityTitleFa", "regionTitleFa"):
            if value.get(key):
                return clean_text(str(value[key]))
        return ""
    return clean_text(str(value))


def to_bool_fa(value: Any) -> str:
    return "بله" if bool(value) else "خیر"


def format_salary(detail: dict[str, Any]) -> str:
    if not detail.get("salaryCanBeShow"):
        min_salary = detail.get("minSalary")
        max_salary = detail.get("maxSalary")
        if min_salary is None and max_salary is None:
            return ""
    min_salary = detail.get("minSalary")
    max_salary = detail.get("maxSalary")
    if min_salary is not None and max_salary is not None:
        return f"{min_salary:,} - {max_salary:,} تومان"
    if min_salary is not None:
        return f"از {min_salary:,} تومان"
    if max_salary is not None:
        return f"تا {max_salary:,} تومان"
    return ""


def build_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "companyApiKey": COMPANY_API_KEY,
            "careerauthkey": COMPANY_API_KEY,
            "address": COMPANY_ADDRESS,
            "Content-Type": "application/json",
        }
    )
    return session


def list_jobs(
    session: requests.Session,
    page_number: int,
    page_size: int,
    keyword: str,
    city_id: str = "",
    department_id: str = "",
) -> dict[str, Any]:
    payload = {
        "Take": page_size,
        "PageNumber": page_number,
        "CityId": city_id,
        "DepartmentId": department_id,
        "Title": keyword,
    }
    response = session.post(LIST_API, json=payload, timeout=30)
    response.raise_for_status()
    body = response.json()
    if not body.get("succeed", body.get("isSuccess")):
        raise RuntimeError(f"Snappfood list failed: {body.get('message')}")
    return body.get("data", {})


def get_job_detail(session: requests.Session, job_guid: str, job_id: int) -> dict[str, Any]:
    if job_guid:
        url = f"{DETAIL_BY_GUID_API}/{job_guid}"
    else:
        url = f"{DETAIL_BY_ID_API}/{job_id}"
    response = session.get(url, timeout=30)
    response.raise_for_status()
    body = response.json()
    if not body.get("succeed", body.get("isSuccess")):
        raise RuntimeError(f"Snappfood detail failed for {job_id}: {body.get('message')}")
    return body.get("data", {})


def map_list_item(item: dict[str, Any]) -> dict[str, str]:
    region = item.get("region") or {}
    return {
        "job_id": str(item.get("id", "")),
        "job_guid": str(item.get("jobGuid", "")),
        "title": clean_text(str(item.get("title", ""))),
        "department": text_value(item.get("department")),
        "job_category": text_value(item.get("jobCategory")),
        "city": text_value(item.get("city")),
        "region": text_value(region),
        "work_type": clean_text(str(item.get("workType", ""))),
        "is_remote": to_bool_fa(item.get("isRemote")),
        "published_at": clean_text(str(item.get("submitTime", ""))),
        "source_url": str(item.get("redirectUrl") or ""),
    }


def map_detail_to_row(list_item: dict[str, Any], detail: dict[str, Any]) -> SnappfoodRow:
    company = detail.get("companyInfo") or {}
    list_meta = map_list_item(list_item)

    min_age = detail.get("minAge")
    max_age = detail.get("maxAge")
    age_range = ""
    if min_age is not None or max_age is not None:
        age_range = f"{min_age or '-'} تا {max_age or '-'}"

    source_url = list_meta["source_url"]
    if not source_url and detail.get("guid"):
        source_url = f"https://career.hrcando.ir/co/snappfood/job-detail/{detail['guid']}"

    return SnappfoodRow(
        job_id=int(list_item.get("id", detail.get("id", 0)) or 0),
        job_guid=str(detail.get("guid") or list_item.get("jobGuid") or ""),
        title=clean_text(str(detail.get("title") or list_meta["title"])),
        company=clean_text(str(company.get("companyName") or "اسنپ فود")),
        company_website=clean_text(str(company.get("companyWebsite") or "")),
        company_address=clean_text(str(company.get("companyAddress") or "")),
        branch_title=clean_text(str(detail.get("branchTitle") or list_item.get("branchTitle") or "")),
        department=text_value(detail.get("department")) or list_meta["department"],
        job_category=text_value(detail.get("jobCategory")) or list_meta["job_category"],
        city=text_value(detail.get("city")) or list_meta["city"],
        region=list_meta["region"],
        work_type=text_value(detail.get("workType")) or list_meta["work_type"],
        seniority_level=text_value(detail.get("seniorityLevel")),
        preferred_gender=text_value(detail.get("preferredGender")),
        salary=format_salary(detail),
        is_remote=to_bool_fa(detail.get("isRemote")),
        is_internship=to_bool_fa(detail.get("isInternship")),
        required_experience_years=str(detail.get("requiredExperienceYear") or ""),
        age_range=age_range,
        working_days=clean_text(str(detail.get("workingDays") or "")),
        description=clean_html_text(detail.get("description")),
        conditions=clean_html_text(detail.get("conditions")),
        requirements=clean_html_text(detail.get("requirements")),
        published_at=clean_text(str(detail.get("createTime") or list_meta["published_at"])),
        source_url=source_url,
    )


def crawl_snappfood(
    keyword: str = "",
    max_pages: int | None = None,
    page_size: int = 50,
    delay: float = REQUEST_DELAY_SECONDS,
) -> list[SnappfoodRow]:
    session = build_session()

    first_page = list_jobs(session, page_number=1, page_size=page_size, keyword=keyword)
    total_pages = int(first_page.get("totalPages") or 1)
    if max_pages is not None:
        total_pages = min(total_pages, max_pages)

    all_items: list[dict[str, Any]] = list(first_page.get("jobs") or [])
    print(f"List page 1/{total_pages}: {len(first_page.get('jobs') or [])} jobs")

    for page in range(2, total_pages + 1):
        time.sleep(delay)
        data = list_jobs(session, page_number=page, page_size=page_size, keyword=keyword)
        page_jobs = data.get("jobs") or []
        all_items.extend(page_jobs)
        print(f"List page {page}/{total_pages}: {len(page_jobs)} jobs")

    rows: list[SnappfoodRow] = []
    total = len(all_items)
    seen: set[str] = set()

    for index, item in enumerate(all_items, start=1):
        job_guid = str(item.get("jobGuid") or "")
        job_id = int(item.get("id", 0) or 0)
        unique_key = job_guid or str(job_id)
        if not unique_key or unique_key in seen:
            continue
        seen.add(unique_key)

        time.sleep(delay)
        detail = get_job_detail(session, job_guid=job_guid, job_id=job_id)
        rows.append(map_detail_to_row(item, detail))
        print(f"PDP {index}/{total}: id={job_id}")

    return rows


def export_excel(rows: list[SnappfoodRow], output_path: Path) -> None:
    records = [
        {
            "شناسه آگهی": row.job_id,
            "شناسه GUID": row.job_guid,
            "عنوان": row.title,
            "شرکت": row.company,
            "وب‌سایت شرکت": row.company_website,
            "آدرس شرکت": row.company_address,
            "شعبه": row.branch_title,
            "دپارتمان": row.department,
            "دسته‌بندی شغلی": row.job_category,
            "شهر": row.city,
            "منطقه": row.region,
            "نوع همکاری": row.work_type,
            "سطح ارشدیت": row.seniority_level,
            "جنسیت": row.preferred_gender,
            "حقوق": row.salary,
            "امکان دورکاری": row.is_remote,
            "کارآموزی": row.is_internship,
            "سابقه کار (سال)": row.required_experience_years,
            "بازه سنی": row.age_range,
            "روزها/ساعات کاری": row.working_days,
            "شرایط": row.conditions,
            "نیازمندی‌ها": row.requirements,
            "شرح شغل": row.description,
            "تاریخ انتشار": row.published_at,
            "لینک آگهی": row.source_url,
        }
        for row in rows
    ]
    df = pd.DataFrame(records)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False, engine="openpyxl")


def main() -> None:
    parser = argparse.ArgumentParser(description="Crawl Snappfood careers with PDP details")
    parser.add_argument("--keyword", default="", help="Filter jobs by title keyword")
    parser.add_argument("--max-pages", type=int, default=None, help="Limit list pages")
    parser.add_argument("--page-size", type=int, default=50, help="Jobs per list page")
    parser.add_argument("--delay", type=float, default=REQUEST_DELAY_SECONDS, help="Delay between requests")
    parser.add_argument("--output", default="", help="Output Excel path")
    args = parser.parse_args()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = (
        Path(args.output)
        if args.output
        else Path("output") / f"snappfood_careers_{timestamp}.xlsx"
    )

    print("Starting Snappfood crawl...")
    rows = crawl_snappfood(
        keyword=args.keyword,
        max_pages=args.max_pages,
        page_size=args.page_size,
        delay=args.delay,
    )
    if not rows:
        print("No jobs found.")
        return

    export_excel(rows, output_path)
    print(f"Done! {len(rows)} jobs saved to: {output_path.resolve()}")


if __name__ == "__main__":
    main()
