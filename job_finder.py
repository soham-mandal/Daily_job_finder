# job_finder.py
import os
import re
import time
import json
import requests
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import dateparser
from urllib.parse import urljoin, quote_plus
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText

# ---------- CONFIG ----------
OUT_XLSX = "daily_jobs.xlsx"
PLATFORMS = ["indeed", "glassdoor", "naukri"]   # 'linkedin' via SerpAPI optionally
QUERY = ""  # leave blank to fetch all roles; or "strategy" to limit by keyword
COUNTRY = "India"
EXPERIENCE_MIN = 1
EXPERIENCE_MAX = 4
ONLY_LAST_N_DAYS = 1   # 24 hours -> 1 day
SERPAPI_KEY = os.getenv("SERPAPI_KEY")  # optional: for LinkedIn via SerpAPI
# Email config (optional)
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT") or 587)
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")
RECIPIENT = os.getenv("RECIPIENT_EMAIL")
# --------------------------------

def parse_relative_date(text):
    """Try to turn '2 days ago' / 'Today' / 'Posted 1 day ago' into a datetime.date"""
    if not text:
        return None
    text = text.strip().lower()
    # Quick heuristics
    if "just" in text or "today" in text or "posted today" in text:
        return datetime.utcnow().date()
    # Try dateparser
    dt = dateparser.parse(text, settings={"RELATIVE_BASE": datetime.utcnow(), "RETURN_AS_TIMEZONE_AWARE": False})
    if dt:
        return dt.date()
    # Fallback: extract integers like '1 day' or '2 hrs'
    m = re.search(r"(\d+)\s*(day|days|hr|hrs|hour|hours)", text)
    if m:
        n = int(m.group(1))
        unit = m.group(2)
        if "day" in unit:
            return (datetime.utcnow() - timedelta(days=n)).date()
        else:
            return (datetime.utcnow() - timedelta(hours=n)).date()
    return None

def experience_in_range(exp_text):
    """Return True if experience text indicates overlap with 1-4 years."""
    if not exp_text:
        return False
    exp_text = exp_text.lower()
    # common patterns: "1-3 yrs", "2 years", "1 to 4 years", "fresher"
    if "fresher" in exp_text or "intern" in exp_text:
        return False
    nums = re.findall(r"(\d+)", exp_text)
    nums = [int(n) for n in nums]
    if not nums:
        # some postings don't show experience; allow them (optional)
        return True
    # If single number e.g., "3 years", check if it's inside range
    if len(nums) == 1:
        val = nums[0]
        return EXPERIENCE_MIN <= val <= EXPERIENCE_MAX
    # If range given, check overlap
    low, high = min(nums), max(nums)
    return not (high < EXPERIENCE_MIN or low > EXPERIENCE_MAX)

def dedupe_rows(rows):
    seen = set()
    out = []
    for r in rows:
        key = (r.get("Job Link") or "").split("?")[0] or (r.get("Job Title","")+"|"+r.get("Company",""))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out

# ---------- Indeed parser ----------
def fetch_indeed(query="", days=1):
    rows = []
    q = quote_plus(query) if query else ""
    url = f"https://in.indeed.com/jobs?q={q}&l=India&fromage={days}"
    print("Indeed URL:", url)
    r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    for card in soup.select("a.tapItem"):
        try:
            title = card.select_one("h2 span")
            title = title.get_text(strip=True) if title else ""
            comp = card.select_one(".companyName")
            company = comp.get_text(strip=True) if comp else ""
            loc = card.select_one(".companyLocation")
            location = loc.get_text(strip=True) if loc else "India"
            link = card.get("href")
            if link and not link.startswith("http"):
                link = urljoin("https://in.indeed.com", link)
            # Some Indeed cards include date/experience text
            date_el = card.select_one(".date")
            date_text = date_el.get_text(strip=True) if date_el else "Last 24 hours"
            posted_date = parse_relative_date(date_text)
            # Experience: sometimes in snippet
            exp_text = ""
            snippet = card.select_one(".job-snippet")
            if snippet:
                exp_m = re.search(r"(\d+\+?\s*-\s*\d+\s*years|\d+\s*years|\d+\s*yrs)", snippet.get_text(" ",strip=True).lower())
                exp_text = exp_m.group(0) if exp_m else ""
            # Filter
            if posted_date and (datetime.utcnow().date() - posted_date).days > days:
                continue
            if exp_text and not experience_in_range(exp_text):
                continue
            rows.append({
                "Job Title": title,
                "Company": company,
                "Location": location,
                "Experience": exp_text or "1-4 yrs (assumed)",
                "Date Posted": date_text,
                "Job Link": link,
                "Source": "Indeed"
            })
        except Exception as e:
            print("indeed parse error", e)
    return rows

# ---------- Glassdoor parser ----------
def fetch_glassdoor(query="", days=1):
    rows = []
    # glassdoor often uses JS and requires more care — this is a best-effort approach
    q = quote_plus(query) if query else ""
    url = f"https://www.glassdoor.co.in/Job/india-{q}-jobs-SRCH_IL.0,5_IN115.htm?fromAge={days}"
    print("Glassdoor URL:", url)
    r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    # Glassdoor lists are complex; try to find job card anchors
    for li in soup.select("li.react-job-listing") or soup.select("div.jobContainer") or soup.select("li.jl"):
        try:
            a = li.select_one("a[data-test='job-link']") or li.select_one("a.jobLink")
            title = a.get_text(strip=True) if a else ""
            link = "https://www.glassdoor.co.in" + a.get("href") if a and a.get("href") else None
            company = li.select_one(".jobInfoItem .jobEmpolyerName") or li.select_one(".jobEmpolyerName")
            company = company.get_text(strip=True) if company else ""
            date_text = li.get_text(" ", strip=True)
            posted_date = parse_relative_date(date_text)
            # Glassdoor doesn't always expose experience in card; allow if no exp text
            rows.append({
                "Job Title": title,
                "Company": company,
                "Location": "India",
                "Experience": "1-4 yrs (assumed)",
                "Date Posted": date_text[:80],
                "Job Link": link,
                "Source": "Glassdoor"
            })
        except Exception as e:
            # ignore parse errors per card
            pass
    return rows

# ---------- Naukri parser (template) ----------
def fetch_naukri(query="", days=1):
    # NOTE: Naukri page structure changes frequently. Inspect the job-card HTML and update selectors.
    rows = []
    q = quote_plus(query) if query else ""
    url = f"https://www.naukri.com/{q}-jobs-in-india"
    print("Naukri URL:", url)
    r = requests.get(url, headers={"User-Agent":"Mozilla/5.0"}, timeout=15)
    soup = BeautifulSoup(r.text, "html.parser")
    # Example: search for job cards (you will probably need to inspect and tune these selectors)
    for card in soup.select("article")[:80]:
        try:
            a = card.select_one("a.title")
            title = a.get_text(strip=True) if a else ""
            link = a.get("href") if a else None
            comp = card.select_one(".subTitle a") or card.select_one(".subTitle")
            company = comp.get_text(strip=True) if comp else ""
            # Naukri shows experience; try to parse
            exp_el = card.select_one(".experience") or card.select_one(".exp")
            exp_text = exp_el.get_text(strip=True) if exp_el else ""
            date_text = card.get_text(" ", strip=True)[:120]
            # Basic filters:
            if exp_text and not experience_in_range(exp_text):
                continue
            # last 24 hours filter – try to parse using date_text heuristics
            rows.append({
                "Job Title": title,
                "Company": company,
                "Location": "India",
                "Experience": exp_text or "1-4 yrs (assumed)",
                "Date Posted": date_text,
                "Job Link": link,
                "Source": "Naukri"
            })
        except Exception:
            pass
    return rows

# ---------- LinkedIn via SerpAPI (recommended) ----------
def fetch_linkedin_serpapi(query="", days=1):
    """
    SerpAPI supports Job search results (Google Jobs results) which include LinkedIn postings.
    Use SERPAPI_KEY in environment to enable this. SerpAPI handles captchas/proxies.
    See SerpApi docs for engine options.
    """
    if not SERPAPI_KEY:
        print("No SERPAPI_KEY set — skipping LinkedIn results")
        return []
    # Example: use SerpAPI Google Jobs engine with 'q' query; tailor if using serpapi-specific LinkedIn endpoint.
    url = "https://serpapi.com/search.json"
    params = {
        "engine":"google_jobs",   # serpapi engine that returns Google Jobs (which aggregates LinkedIn/others)
        "q": query or "",
        "google_domain":"google.co.in",
        "hl":"en",
        "gl":"in",
        "api_key": SERPAPI_KEY
    }
    resp = requests.get(url, params=params, timeout=20)
    data = resp.json()
    rows = []
    for job in data.get("jobs_results", []):
        title = job.get("title")
        company = job.get("company_name")
        link = job.get("link")
        date_text = job.get("posted_at") or job.get("date")
        posted_date = parse_relative_date(date_text)
        # Some job entries have experience in snippet
        snippet = job.get("snippet") or ""
        exp_m = re.search(r"(\d+\s*-\s*\d+|\d+)\s*(year|yr|yrs|years)", snippet.lower())
        exp_text = exp_m.group(0) if exp_m else ""
        if exp_text and not experience_in_range(exp_text):
            continue
        if posted_date and (datetime.utcnow().date() - posted_date).days > days:
            continue
        rows.append({
            "Job Title": title,
            "Company": company,
            "Location": "India",
            "Experience": exp_text or "1-4 yrs (assumed)",
            "Date Posted": date_text,
            "Job Link": link,
            "Source": "LinkedIn/GoogleJobs(SerpAPI)"
        })
    return rows

# ---------- Save and email ----------
def save_to_excel(all_rows, path=OUT_XLSX):
    # all_rows: dict of platform->list[dict]
    writer = pd.ExcelWriter(path, engine="openpyxl")
    master = []
    for src, rows in all_rows.items():
        df = pd.DataFrame(rows)
        if df.empty:
            df = pd.DataFrame(columns=["Job Title","Company","Location","Experience","Date Posted","Job Link","Source"])
        df.to_excel(writer, sheet_name=src[:30], index=False)
        master.extend(rows)
    pd.DataFrame(master).to_excel(writer, sheet_name="Master", index=False)
    writer.close()
    return path

def send_email_with_attachment(subject, body, filepath):
    if not SMTP_USER or not SMTP_PASS or not RECIPIENT:
        print("Email environment variables not set, skipping email")
        return
    msg = MIMEMultipart()
    msg["From"] = SMTP_USER
    msg["To"] = RECIPIENT
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    with open(filepath, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(filepath))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(filepath)}"'
        msg.attach(part)
    s = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    s.starttls()
    s.login(SMTP_USER, SMTP_PASS)
    s.send_message(msg)
    s.quit()
    print("Email sent to", RECIPIENT)

def main():
    # orchestrate
    all_rows = {}
    q = QUERY
    print("Starting fetch:", datetime.utcnow().isoformat())
    if "indeed" in PLATFORMS:
        all_rows["Indeed"] = fetch_indeed(q, ONLY_LAST_N_DAYS)
        time.sleep(1)
    if "glassdoor" in PLATFORMS:
        all_rows["Glassdoor"] = fetch_glassdoor(q, ONLY_LAST_N_DAYS)
        time.sleep(1)
    if "naukri" in PLATFORMS:
        all_rows["Naukri"] = fetch_naukri(q, ONLY_LAST_N_DAYS)
        time.sleep(1)
    if SERPAPI_KEY:
        all_rows["LinkedIn"] = fetch_linkedin_serpapi(q, ONLY_LAST_N_DAYS)
        time.sleep(1)
    # dedupe per source and combine
    for k in list(all_rows.keys()):
        all_rows[k] = dedupe_rows(all_rows[k])
    path = save_to_excel(all_rows)
    print("Saved to", path)
    subject = f"Daily Jobs - {datetime.utcnow().strftime('%Y-%m-%d')}"
    body = "Attached are the job results (last 24 hours, experience 1-4 yrs)."
    send_email_with_attachment(subject, body, path)

if __name__ == "__main__":
    main()
