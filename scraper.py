from pathlib import Path
from datetime import datetime
import time
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait

DIRECTORY_URL = "https://www.indiainx.com/markets/DirectoryMembers.aspx"
OUTPUT_DIR = Path(__file__).resolve().parent / "output"

MEMBER_COLUMNS = [
    "Member Name","Trade Name","INX Clearing Number","Registered Office",
    "Correspondence Office","Website","SEBI Registration No.","SEBI Registration Date",
    "Date Of Incorporation","Year of INX Membership","Membership Status",
    "Profile Of Member","Belongs To Business Group","Member Entity whether listed at"
]
DIRECTOR_COLUMNS = ["INX Clearing Number","Member Name","Director Name","Designation"]
DESIGNATED_COLUMNS = ["INX Clearing Number","Member Name","Name","Email ID","Telephone No.","Mobile No."]
OTHER_COLUMNS = [
    "INX Clearing Number","Member Name","Category / Source","Company Name","Activity",
    "SEBI Registration No.","SEBI Registration Start Date","Name Of Exchange",
    "Exchange Member No.","SEBI Registration Date","Branch Office Details"
]

SECTIONS_KEYWORDS = [
    "board of directors",
    "designated officer details",
    "other capital market activities under same entity",
    "membership of other indian stock exchanges of same entity",
    "other group companies in the capital market area",
    "profile of member",
    "belongs to business group",
    "member entity whether listed at",
    "branch offices details"
]

def clean(v):
    return "" if v is None else " ".join(str(v).replace("\xa0"," ").split()).strip()

def clean_phone(v):
    if not v:
        return ""
    digits = [c for c in str(v) if c.isdigit()]
    if not digits:
        return ""
    return clean(v)

def driver_factory():
    o = Options()

    o.binary_location = "/usr/bin/chromium"

    o.add_argument("--headless=new")
    o.add_argument("--no-sandbox")
    o.add_argument("--disable-dev-shm-usage")
    o.add_argument("--disable-gpu")
    o.add_argument("--window-size=1920,1080")
    o.add_argument("--log-level=3")
    o.add_argument("--disable-notifications")

    o.add_experimental_option("excludeSwitches", ["enable-logging"])

    return webdriver.Chrome(options=o)

def collect_urls(driver, page_from=None, page_to=None):
    wait=WebDriverWait(driver,20)
    driver.get(DIRECTORY_URL)
    wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR,'a[href*="MemberData.aspx?MemberNo="]'))>0)
    result=[]; seen=set(); page=1

    while True:
        if page_from is None or page_from <= page <= page_to:
            links=driver.find_elements(By.CSS_SELECTOR,'a[href*="MemberData.aspx?MemberNo="]')
            for a in links:
                try:
                    name=clean(a.text); href=a.get_attribute("href")
                    if href and href not in seen:
                        result.append({"name":name,"url":href}); seen.add(href)
                except StaleElementReferenceException:
                    pass

        if page_to is not None and page >= page_to: break

        links=driver.find_elements(By.CSS_SELECTOR,'a[href*="MemberData.aspx?MemberNo="]')
        old=links[0].get_attribute("href") if links else None
        try: nxt=driver.find_element(By.NAME,"ctl00$ContentPlaceHolder1$btnNext")
        except NoSuchElementException: break
        if nxt.get_attribute("disabled") is not None: break

        driver.execute_script("arguments[0].scrollIntoView({block:'center'});",nxt)
        time.sleep(.5)
        nxt=driver.find_element(By.NAME,"ctl00$ContentPlaceHolder1$btnNext")
        driver.execute_script("arguments[0].click();",nxt)

        try:
            wait.until(lambda d: (
                len(d.find_elements(By.CSS_SELECTOR,'a[href*="MemberData.aspx?MemberNo="]'))>0
                and d.find_elements(By.CSS_SELECTOR,'a[href*="MemberData.aspx?MemberNo="]')[0].get_attribute("href") != old
            ))
        except TimeoutException:
            break
        page += 1
        time.sleep(.4)
    return result

def parse_member(driver, url):
    wait = WebDriverWait(driver, 20)
    driver.get(url)
    wait.until(lambda d: len(d.find_elements(By.TAG_NAME, "table")) >= 2)
    
    t = driver.find_elements(By.TAG_NAME, "table")[1]
    raw_rows = []
    for tr in t.find_elements(By.TAG_NAME, "tr"):
        cells = [clean(c.text) for c in tr.find_elements(By.XPATH, "./th|./td")]
        if any(cells):
            raw_rows.append(cells)

    member = {c: "" for c in MEMBER_COLUMNS}
    directors = []
    designated = []
    others = []

    # 1. Parse top member details key-values
    for r in raw_rows:
        if len(r) >= 2:
            k0 = clean(r[0])
            if k0 in member: member[k0] = clean(r[1])
        if len(r) >= 4:
            k2 = clean(r[2])
            if k2 in member: member[k2] = clean(r[3])

    name = member["Member Name"]
    no = member["INX Clearing Number"]

    def find_section_row(keyword):
        for idx, r in enumerate(raw_rows):
            if r and clean(r[0]).lower() == keyword:
                return idx
        return None

    def get_section_end(start_idx):
        for idx in range(start_idx + 1, len(raw_rows)):
            if raw_rows[idx]:
                k0 = clean(raw_rows[idx][0]).lower()
                if any(k0 == kw for kw in SECTIONS_KEYWORDS):
                    return idx
        return len(raw_rows)

    # 2. Parse Profile, Belongs to, Listed at
    for title, col in [
        ("profile of member", "Profile Of Member"),
        ("belongs to business group", "Belongs To Business Group"),
        ("member entity whether listed at", "Member Entity whether listed at")
    ]:
        s_idx = find_section_row(title)
        if s_idx is not None:
            vals = [x for x in raw_rows[s_idx][1:] if x]
            if not vals and s_idx + 1 < len(raw_rows):
                vals = [x for x in raw_rows[s_idx + 1] if x]
            member[col] = " | ".join(vals)

    # 3. Board of Directors
    s_idx = find_section_row("board of directors")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            if len(r) >= 2:
                d_name = r[1] if len(r) > 1 and r[1] else r[0]
                d_desig = r[2] if len(r) > 2 else ""
                if d_name and d_name.lower() not in {"name", "board of directors"}:
                    directors.append({
                        "INX Clearing Number": no,
                        "Member Name": name,
                        "Director Name": d_name,
                        "Designation": d_desig
                    })

    # 4. Designated Officer Details
    s_idx = find_section_row("designated officer details")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            if len(r) >= 2:
                off_name = r[1] if len(r) > 1 and r[1] else (r[0] if r[0] else "")
                if off_name and off_name.lower() not in {"name", "designated officer details"}:
                    designated.append({
                        "INX Clearing Number": no,
                        "Member Name": name,
                        "Name": off_name,
                        "Email ID": r[2] if len(r) > 2 else "",
                        "Telephone No.": clean_phone(r[3]) if len(r) > 3 else "",
                        "Mobile No.": clean_phone(r[4]) if len(r) > 4 else ""
                    })

    # 5. Other Capital Market Activities Under Same Entity
    s_idx = find_section_row("other capital market activities under same entity")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            if len(r) >= 2:
                sebi_no = r[1] if len(r) > 1 else ""
                sebi_start = r[2] if len(r) > 2 else ""
                if (sebi_no or sebi_start) and sebi_no.lower() != "sebi registration no.":
                    others.append({
                        "INX Clearing Number": no,
                        "Member Name": name,
                        "Category / Source": "Other Capital Market Activities",
                        "Company Name": "",
                        "Activity": "",
                        "SEBI Registration No.": sebi_no,
                        "SEBI Registration Start Date": sebi_start,
                        "Name Of Exchange": "",
                        "Exchange Member No.": "",
                        "SEBI Registration Date": "",
                        "Branch Office Details": ""
                    })

    # 6. Membership Of Other Indian Stock Exchanges of Same Entity
    s_idx = find_section_row("membership of other indian stock exchanges of same entity")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            if len(r) >= 2:
                ex_name = r[1] if len(r) > 1 else ""
                ex_no = r[2] if len(r) > 2 else ""
                sebi_no = r[3] if len(r) > 3 else ""
                sebi_date = r[4] if len(r) > 4 else ""
                if ex_name and ex_name.lower() != "name of exchange":
                    others.append({
                        "INX Clearing Number": no,
                        "Member Name": name,
                        "Category / Source": "Membership Of Other Stock Exchanges",
                        "Company Name": "",
                        "Activity": "",
                        "SEBI Registration No.": sebi_no,
                        "SEBI Registration Start Date": "",
                        "Name Of Exchange": ex_name,
                        "Exchange Member No.": ex_no,
                        "SEBI Registration Date": sebi_date,
                        "Branch Office Details": ""
                    })

    # 7. Other Group Companies In The Capital Market Area
    s_idx = find_section_row("other group companies in the capital market area")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            if len(r) >= 2:
                comp_name = r[1] if len(r) > 1 else ""
                activity = r[2] if len(r) > 2 else ""
                if comp_name and comp_name.lower() != "company name":
                    others.append({
                        "INX Clearing Number": no,
                        "Member Name": name,
                        "Category / Source": "Other Group Companies",
                        "Company Name": comp_name,
                        "Activity": activity,
                        "SEBI Registration No.": "",
                        "SEBI Registration Start Date": "",
                        "Name Of Exchange": "",
                        "Exchange Member No.": "",
                        "SEBI Registration Date": "",
                        "Branch Office Details": ""
                    })

    # 8. Branch offices details
    s_idx = find_section_row("branch offices details")
    if s_idx is not None:
        end_idx = get_section_end(s_idx)
        for r in raw_rows[s_idx + 1 : end_idx]:
            vals = [x for x in r if x]
            if vals:
                others.append({
                    "INX Clearing Number": no,
                    "Member Name": name,
                    "Category / Source": "Branch Office",
                    "Company Name": "",
                    "Activity": "",
                    "SEBI Registration No.": "",
                    "SEBI Registration Start Date": "",
                    "Name Of Exchange": "",
                    "Exchange Member No.": "",
                    "SEBI Registration Date": "",
                    "Branch Office Details": " | ".join(vals)
                })

    return member, directors, designated, others

def run_scraper(page_from=None, page_to=None, progress_callback=None):
    OUTPUT_DIR.mkdir(exist_ok=True)
    d = driver_factory()
    try:
        urls = collect_urls(d, page_from, page_to)
        if not urls: raise RuntimeError("No members found.")
        members = []; directors = []; designated = []; others = []; total = len(urls)

        for n, item in enumerate(urls, 1):
            if progress_callback: progress_callback(n - 1, total, f"Processing {n}/{total}: {item['name']}")
            try:
                m, ds, do, ot = parse_member(d, item["url"])
                members.append(m); directors.extend(ds); designated.extend(do); others.extend(ot)
            except Exception as e:
                print(f"Failed: {item['name']} | {type(e).__name__}")

        if progress_callback: progress_callback(total, total, "Creating Excel workbook...")
        filename = f"India_INX_Members_{datetime.now().strftime('%d-%m-%y')}.xlsx"
        path = OUTPUT_DIR / filename

        with pd.ExcelWriter(path, engine="openpyxl") as w:
            pd.DataFrame(members, columns=MEMBER_COLUMNS).to_excel(w, sheet_name="Members", index=False)
            pd.DataFrame(directors, columns=DIRECTOR_COLUMNS).to_excel(w, sheet_name="Directors", index=False)
            pd.DataFrame(designated, columns=DESIGNATED_COLUMNS).to_excel(w, sheet_name="Designated", index=False)
            pd.DataFrame(others, columns=OTHER_COLUMNS).to_excel(w, sheet_name="Others", index=False)

        wb = load_workbook(path)
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                letter = col[0].column_letter
                longest = max((len(str(c.value)) if c.value is not None else 0 for c in col), default=10)
                ws.column_dimensions[letter].width = min(max(longest + 2, 12), 55)
        wb.save(path)

        return {"file": str(path), "filename": filename, "members": len(members),
                "directors": len(directors), "designated": len(designated), "others": len(others)}
    finally:
        d.quit()
