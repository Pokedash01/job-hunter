import os
import re
import json
import requests
import pandas as pd
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from datetime import datetime
from jobspy import scrape_jobs
from google import genai
from google.genai import types
from google.genai.errors import APIError

# ReportLab imports for dense, professional PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY") or os.environ.get("SEARCH_API_KEYSEARCH_API_KEY")

client = genai.Client(api_key=GEMINI_KEY) if GEMINI_KEY else None

# ----------------- KARTIK BHATT MASTER PROFILE -----------------
KARTIK_PROFILE = """
Candidate: Kartik Bhatt
Contact: kb270102@gmail.com | +91-7428062532 | LinkedIn | Portfolio: https://kartikb.vercel.app/
Education: Maharaja Surajmal Institute | Bachelor of Computer Applications | Majors: Computer Science | GPA: 9.3/10 (Top 1%) | Jul'19 – Aug'22
Total Work Experience: ~3.5+ years (3 years 2 months recorded)

Work Experience:
1. KPMG | Analyst | Knowledge Management | Gurugram (May'24 – Present)
Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management & played key role in business development.
Key Projects:
- Built complete Power Platform solution (Power Automate, SharePoint lists, Power Apps, Power BI dashboards) facilitating 20,000 reach outs annually across 30+ member firms in 13 sectors, saving 1,200 hrs annually.
- Built end-to-end solution to facilitate migration of old Excel-based data collection for 45+ pillars to automated SPO list integration, including 3 Power Automate flows for alerts, change management, data migration, and permission governance.
- Built and managed multiple VBA macros solutions refreshing a repository of 30,000+ assets globally, including change management, saving 485 hrs annually.
- Architected a multi-modal Copilot agent to assist with messy data, draft fields, and apply metadata tags based on source and guidelines, saving 325 hrs annually.

Business Development & Operations:
- Saved 2,000+ hrs annually leveraging Power Platform, Copilot Studio, and VBA Macros.
- Handled 100+ RFP and RFI requests and built/maintained 100+ internal site pages per brand values and standards.
- Undertook complete contact management system administration for 10,000+ KGS members.
- Uploaded 5,000+ content assets across 3 content types and 15 libraries.
- Performed Audit Market Share (AMS) analysis for 6+ sectors.
- Catered to 50+ SharePoint governance and administration requests (term store management, change management, metadata management, permission level governance).

Awards & Recognitions:
- Awarded 'KUDOS' for displaying exceptional efficiency and applying lean six sigma methodology saving 2,000+ hrs annually.
- Awarded 'KUDOS' for migrating legacy practices using VBA and Excel to more GenAI focused using agents and Power Platform.
- Earned 'Super Team' award for hosting/organizing employee council events.
- Received 'Ally of Inclusion' accolade and 'Gurus@Work' for contributions to firm learning culture.

2. GlobalLogic Technologies Private Limited | Associate Analyst | Content Engineering | Gurugram (Sep'22 – Oct'23)
Participated in content generation and manipulation projects for clients including Google.
Key Projects:
- Created best practices, process docs, and QA processes for Google GenAI training datasets for Android screen search.
- Piloted project to extract relevant answers from multi-level docs to build training datasets for AI.
Business Development & Operations:
- Designed and implemented QA processes for data entry, reducing errors by 25% and improving data accuracy.
- Managed process documentation for 10+ projects, ensuring compliance and accessibility for stakeholders.
- Connected with onshore stakeholders and improved delivery quality from 74% to 95%.
- QA'd 500+ pieces on a weekly basis and led 3 pilot projects to completion against major MNC competition.

Skills: MS Excel, Copilot GenAI (Agents), Power Automate, Power BI, SharePoint Online, Power Apps, Copilot Studio, SQL, Python, Process Automation, Stakeholder Management, RFP/RFI Bidding.
Certifications:
- Microsoft Certified: Azure AI Fundamentals (AI-901)
- Microsoft Certified: AI Transformation Leader (AB-731)
- Microsoft Certified: AI Business Professional (AB-730)
- Lean Six Sigma: Yellow Belt
- Oracle: Agentic AI Certified Foundations Associate
"""

# ----------------- STRICT FILTERING CONFIG -----------------
SENIORITY_BLACKLIST = [
    r"\bsenior manager\b", r"\bprincipal\b", r"\bdirector\b", r"\bvp\b",
    r"\bhead of\b", r"\bgroup product manager\b", r"\btech lead\b",
    r"\bengineering manager\b", r"\bgeneral manager\b", r"\blead architect\b",
    r"\bassociate director\b", r"\bavp\b", r"\boperations director\b", r"\bstaff\b"
]

ROLES_WHITELIST = [
    r"\bdata analyst\b", r"\bbusiness analyst\b", r"\bproduct analyst\b",
    r"\bassociate product manager\b", r"\bapm\b", r"\bcopilot studio\b",
    r"\bpower automate\b", r"\bpower platform\b", r"\bbi developer\b",
    r"\bprocess automation\b", r"\banalytics engineer\b", r"\boperations analyst\b",
    r"\bproduct operations\b", r"\bai analyst\b", r"\bautomation analyst\b"
]

CORE_SKILLS_KEYWORDS = [
    r"\bpower automate\b", r"\bpower platform\b", r"\bpower apps\b", r"\bpower bi\b",
    r"\bcopilot\b", r"\bgenai\b", r"\bllm\b", r"\bsharepoint\b", r"\bvba\b",
    r"\bprocess automation\b", r"\bsql\b", r"\bpython\b", r"\brpa\b", r"\bworkflows?\b",
    r"\bdata analytics\b", r"\bbusiness intelligence\b", r"\bdecision support\b"
]

EXCLUDED_DOMAINS = [
    "hr", "human resources", "talent acquisition", "recruiter", "recruitment",
    "sales", "business development executive", "bde", "marketing", "digital marketing",
    "telecaller", "content writer", "seo", "graphic designer", "accountant",
    "surveyor", "field data", "manufacturing operating", "master data management"
]

LOCATIONS_TIER_1 = [
    r"\bdelhi\b", r"\bnew delhi\b", r"\bncr\b", r"\bgurgaon\b", r"\bgurugram\b",
    r"\bnoida\b", r"\bfaridabad\b", r"\bghaziabad\b",
    r"\bremote\b", r"\bwfh\b", r"\bwork from home\b"
]
LOCATIONS_TIER_2 = [
    r"\bbangalore\b", r"\bbengaluru\b", r"\bhyderabad\b", r"\bpune\b", r"\bjaipur\b"
]

LOCATION_CANONICAL = [
    (r"\bremote\b|\bwfh\b|\bwork from home\b", "Remote"),
    (r"\bgurugram\b|\bgurgaon\b", "Gurugram"),
    (r"\bnoida\b", "Noida"),
    (r"\bfaridabad\b", "Faridabad"),
    (r"\bghaziabad\b", "Ghaziabad"),
    (r"\bnew delhi\b|\bdelhi\b|\bncr\b", "Delhi"),
    (r"\bbengaluru\b|\bbangalore\b", "Bengaluru"),
    (r"\bhyderabad\b", "Hyderabad"),
    (r"\bpune\b", "Pune"),
    (r"\bmumbai\b", "Mumbai"),
]

MAX_EXPERIENCE_CAP = 5.5

EXCEL_FILE = "job_applications.xlsx"
DOCS_DIR = "generated_docs"
os.makedirs(DOCS_DIR, exist_ok=True)

GITHUB_REPO = "Pokedash01/job-hunter"
GITHUB_BRANCH = "main"

EXCEL_COLUMNS = [
    "Title", "Company", "Salary Range", "Fit Score",
    "Location", "Resume Link", "Cover Letter Link", "Job Link"
]
LOG_FILE = "job_log.csv"


def github_raw_link(local_path: str) -> str:
    clean_path = local_path.replace(os.sep, "/").lstrip("./")
    return f"https://raw.githubusercontent.com/{GITHUB_REPO}/{GITHUB_BRANCH}/{clean_path}"


# ----------------- TELEGRAM SYSTEM ALERTS -----------------
def send_telegram_system_alert(service_name: str, error_detail: str):
    if not TELEGRAM_TOKEN or not CHAT_ID:
        return
    text = (
        f"🚨 <b>SYSTEM NOTIFICATION: Token / Quota Alert</b>\n\n"
        f"<b>Target Service:</b> {service_name}\n"
        f"<b>Details:</b> <code>{str(error_detail)[:450]}</code>\n\n"
        f"⚠️ <i>Please renew tokens or update API secrets to resume uninterrupted execution.</i>"
    )
    endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        requests.post(endpoint, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "HTML"}, timeout=10)
    except Exception as e:
        print(f"Failed to deliver system alert to Telegram: {e}")


# ----------------- STORAGE ENGINE -----------------
def init_tracker():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=EXCEL_COLUMNS)
        df.to_excel(EXCEL_FILE, index=False)


def _load_tracker() -> pd.DataFrame:
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        df = pd.read_excel(EXCEL_FILE)
        for col in EXCEL_COLUMNS:
            if col not in df.columns:
                df[col] = None
        return df
    except Exception as e:
        print(f"Excel read error: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS)


def is_seen(url: str) -> bool:
    df = _load_tracker()
    if df.empty:
        return False
    return df["Job Link"].astype(str).eq(str(url)).any()


def log_job(title, company, salary_range, fit_score, location, resume_link, cover_letter_link, job_link):
    try:
        df = _load_tracker()
        new_row = {
            "Title": title,
            "Company": company,
            "Salary Range": salary_range,
            "Fit Score": fit_score,
            "Location": location,
            "Resume Link": resume_link,
            "Cover Letter Link": cover_letter_link,
            "Job Link": job_link
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)[EXCEL_COLUMNS]
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        print(f"Excel logging error: {e}")


def append_to_log(title, company, salary_range, fit_score, location, resume_link, cover_letter_link, job_link):
    import csv
    file_exists = os.path.exists(LOG_FILE)
    try:
        with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(EXCEL_COLUMNS)
            writer.writerow([
                title, company, salary_range, fit_score,
                location, resume_link, cover_letter_link, job_link
            ])
    except Exception as e:
        print(f"job_log.csv logging error: {e}")


# ----------------- SCREENERS & FILTERS -----------------
def is_seniority_excluded(title: str) -> bool:
    t = title.lower()
    return any(re.search(pat, t) for pat in SENIORITY_BLACKLIST)


def is_role_and_skill_relevant(title: str, description: str) -> bool:
    t = title.lower()
    d = description.lower()
    full_text = f"{t} {d}"

    for bad in EXCLUDED_DOMAINS:
        if re.search(r'\b' + re.escape(bad) + r'\b', t):
            return False

    if is_seniority_excluded(title):
        return False

    if not any(re.search(good, t) for good in ROLES_WHITELIST):
        return False

    matched_skills = sum(1 for kw in CORE_SKILLS_KEYWORDS if re.search(kw, full_text))
    return matched_skills >= 2


def extract_clean_location(text_to_search: str) -> str:
    loc = str(text_to_search).lower()
    matched = []
    for pattern, name in LOCATION_CANONICAL:
        if re.search(pattern, loc) and name not in matched:
            matched.append(name)
    if matched:
        return " / ".join(matched)
    if re.search(r"\bindia\b", loc):
        return "India"
    return "Unspecified"


def classify_location(location_str: str, fallback_text: str = ""):
    clean = extract_clean_location(location_str)
    if clean == "Unspecified" and fallback_text:
        clean = extract_clean_location(fallback_text)

    loc = clean.lower()
    if any(re.search(pat, loc) for pat in LOCATIONS_TIER_1):
        return True, "Tier 1", clean
    if any(re.search(pat, loc) for pat in LOCATIONS_TIER_2):
        return True, "Tier 2", clean
    if "india" in loc:
        return True, "Tier 1", clean

    return False, "Excluded", clean


# ----------------- COMPENSATION ANALYZER -----------------
def get_salary_range_and_check(title: str, company: str, location: str, description: str, tier: str, min_sal=None, max_sal=None):
    min_floor = 1000000 if "Tier 1" in tier else 1400000

    if max_sal and float(max_sal) > 0:
        actual_min = float(min_sal) if min_sal and float(min_sal) > 0 else float(max_sal) * 0.8
        actual_max = float(max_sal)
        if actual_max < min_floor:
            return False, ""
        return True, f"₹{actual_min/100000:.1f} - ₹{actual_max/100000:.1f} LPA"

    prompt = f"""
You are an expert Indian tech industry compensation analyst.
Analyze the expected total annual CTC (in INR / LPA) for this role given Kartik Bhatt's profile (~3.5 years of experience at KPMG & GlobalLogic, BCA CS 9.3 GPA, Power Platform, Copilot Studio, Analytics, SQL, Python):

Job Title: {title}
Company: {company}
Location: {location}
Job Description: {description[:2200]}

Determine:
1. Realistic min and max CTC range in INR.
2. `passes_floor`: true if max_inr >= {min_floor}, false otherwise.

Return ONLY a valid JSON object matching:
{{
  "min_inr": 1300000,
  "max_inr": 1850000,
  "display_range": "₹13 - ₹18.5 LPA (Est.)",
  "passes_floor": true
}}
"""
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        data = json.loads(response.text)
        if not data.get("passes_floor", True):
            return False, ""
        return True, data.get("display_range", "₹12 - ₹16 LPA (Est.)")
    except APIError as e:
        if "RESOURCE_EXHAUSTED" in str(e) or e.code == 429:
            send_telegram_system_alert("Gemini API", f"Compensation Check Rate Limited / Exhausted: {e}")
        return True, "₹12 - ₹16 LPA (Est.)"
    except Exception as e:
        print(f"Salary check notice for {title} @ {company}: {e}")
        return True, "₹12 - ₹16 LPA (Est.)"


# ----------------- PORTAL SEARCH & SCRAPING -----------------
def extract_company_from_url(url: str) -> str:
    try:
        domain = urlparse(url).netloc.lower()
        path = urlparse(url).path.lower()
        parts = domain.split(".")
        if "greenhouse.io" in domain or "lever.co" in domain or "ashbyhq.com" in domain or "smartrecruiters.com" in domain:
            if parts[0] not in ["boards", "job-boards", "jobs", "www"]:
                return parts[0].replace("-", " ").capitalize()
            elif len(path.split("/")) > 1:
                return path.split("/")[1].replace("-", " ").capitalize()
        elif "myworkdayjobs.com" in domain:
            return parts[0].split("-")[0].replace("-", " ").capitalize()
    except Exception:
        pass
    return "Enterprise Portal"


def fetch_portal_description(url: str) -> str:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        res = requests.get(url, headers=headers, timeout=6)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            for tag in soup(["script", "style", "nav", "footer", "header"]):
                tag.decompose()
            return ' '.join(soup.stripped_strings)[:3500]
    except Exception:
        pass
    return ""


def search_enterprise_ats_jobs():
    discovered = []
    if not SEARCH_KEY:
        return discovered

    negatives = '-"Senior Manager" -Director -VP -Intern -Lead -HR -Talent'
    locations = '("Gurgaon" OR "Gurugram" OR "Noida" OR "Delhi" OR "Bangalore" OR "Remote")'
    ats_sites = '(site:myworkdayjobs.com OR site:boards.greenhouse.io OR site:jobs.lever.co OR site:jobs.ashbyhq.com OR site:smartrecruiters.com)'

    ats_queries = [
        f'{ats_sites} intitle:("Power Platform" OR "Power Automate" OR "Process Automation" OR "Copilot Studio" OR "Automation Analyst") {locations} {negatives}',
        f'{ats_sites} intitle:("Data Analyst" OR "Business Analyst" OR "Product Analyst" OR "APM") ("Power BI" OR "SQL" OR "Python" OR "Automate") {locations} {negatives}'
    ]

    for query in ats_queries:
        try:
            url = "https://www.searchapi.io/api/v1/search"
            params = {
                "engine": "google",
                "q": query,
                "api_key": SEARCH_KEY,
                "gl": "in",
                "hl": "en",
                "num": 25,
                "time_period": "last_week"
            }
            res = requests.get(url, params=params, timeout=12)
            if res.status_code in [401, 402, 429]:
                send_telegram_system_alert("SearchAPI.io", f"HTTP {res.status_code}: Exhausted credits or unauthorized.")
                break

            data = res.json()
            for item in data.get("organic_results", []):
                link = item.get("link", "")
                raw_title = item.get("title", "")
                snippet = item.get("snippet", "")
                title = re.sub(
                    r"\s*[-|–]\s*(Greenhouse|Lever|Workday|Ashby|SmartRecruiters|Jobs|Careers|Myworkdayjobs\.com).*",
                    "",
                    raw_title,
                    flags=re.IGNORECASE
                ).strip()
                company = extract_company_from_url(link)
                inferred_loc = extract_clean_location(snippet)
                if inferred_loc == "Unspecified":
                    inferred_loc = "India"

                if link:
                    discovered.append({
                        "title": title,
                        "company": company,
                        "job_url": link,
                        "location": inferred_loc,
                        "description": snippet,
                        "min_amount": 0,
                        "max_amount": 0,
                        "is_direct_ats": True
                    })
        except Exception as e:
            print(f"ATS search notice: {e}")

    return discovered


# ----------------- ACCURATE MATCHING & RESUME ALIGNMENT -----------------
def generate_application_kit(title, company, description):
    prompt = f"""
You are an expert executive tech recruiter evaluating a candidate for a technical role.

CANDIDATE GROUND TRUTH:
{KARTIK_PROFILE}

TARGET JOB DETAILS:
Role: {title}
Company: {company}
JD: {description[:2800]}

STRICT EVALUATION GUIDELINES:
1. EXPERIENCE DETECTION:
   - Identify the exact minimum years of experience required from the JD. If open/unspecified, state 'Not specified (Estimated 2-4 Years)'.
   - Determine `exceeds_cap`: true if required experience > 5.5 years, false otherwise.
2. ACCURATE SKILL GAPS:
   - Carefully identify distinct tools/skills demanded by the JD that Kartik does NOT have in his ground truth (e.g. AWS, Snowflake, Tableau, BigQuery, Kafka, Jira).
   - If Kartik covers all core requirements, state 'None detected'.
3. CALIBRATED FIT SCORE:
   - Score objectively between 0% and 100% comparing his genuine 3.5 YOE stack (Power Platform, Copilot Studio, SQL, Python, Excel/VBA) against the JD.
4. INTELLIGENT RESUME BULLET ALIGNMENT (DO NOT LIE):
   - Adapt the terminology in Kartik's bullets to match the JD's stack without altering facts or numbers.
   - Example 1: If JD focuses on ServiceNow, Jira, or SQL tables instead of SharePoint, rephrase SPO list migration to 'structured data repository migration and list integration (SPO/Enterprise tables)'.
   - Example 2: If JD asks for Automation Scripting over VBA, phrase VBA macros as 'automated VBA/macro scripting and tabular data engines'.
   - Keep ALL numbers exact: 20,000 reach outs, 30+ member firms, 1,200 hrs saved, 45+ pillars, 30,000+ assets, 485 hrs saved, 325 hrs saved, 2,000+ hrs saved, 25% QA improvement, 74% to 95% quality.

Return ONLY a valid JSON object matching this schema:
{{
  "detected_experience": "e.g., 3+ Years",
  "min_years_numeric": 3.0,
  "exceeds_cap": false,
  "match_score": "e.g., 92%",
  "skills_gap": "e.g., Snowflake, Tableau",
  "reason": "1-2 sharp sentences explaining fit and overlap.",
  "kpmg_project_bullets": [
    "Adapted bullet on Power Platform: 20,000 reach outs, 30+ member firms, 1,200 hrs saved",
    "Adapted bullet on data collection migration: 45+ pillars, 3 Power Automate flows, permission governance",
    "Adapted bullet on macro/scripting automation: 30,000+ assets globally, 485 hrs saved",
    "Adapted bullet on multi-modal Copilot agent: messy data cleanup, field drafting, metadata tagging, 325 hrs saved"
  ],
  "kpmg_bd_bullets": [
    "Saved 2,000+ hrs annually using Power Platform, Copilot Studio, and VBA Macros.",
    "Catered to 100+ RFP and RFI requests and 100+ internal site pages according to brand standards.",
    "Undertook contact management system administration for 10,000+ members and uploaded 5,000+ content assets.",
    "Performed Audit Market Share (AMS) for 6+ sectors and handled 50+ governance/administration requests."
  ],
  "globallogic_bullets": [
    "Created best practices, process docs, and QA workflows for Google GenAI Android screen search datasets.",
    "Piloted project to extract relevant answers from multi-level docs to build training datasets for AI.",
    "Designed and implemented QA processes for data entry, reducing errors by 25% and improving data accuracy.",
    "Managed documentation for 10+ projects, lifting project delivery quality from 74% to 95% with 500+ weekly QA reviews."
  ],
  "cover_letter_subject": "Subject: Application for {title} - Kartik Bhatt",
  "cover_letter_paragraphs": [
    "Targeted opening demonstrating alignment with {company} and role {title}.",
    "KPMG achievement paragraph mapped to specific requirements in the JD.",
    "GlobalLogic QA and AI training datasets achievement paragraph.",
    "Closing paragraph reaffirming enthusiasm and availability."
  ]
}}
"""
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except APIError as e:
        if "RESOURCE_EXHAUSTED" in str(e) or e.code == 429:
            send_telegram_system_alert("Gemini API", f"Gemini generation quota exhausted: {e}")
        return None
    except Exception as e:
        print(f"Generation error: {e}")
        return None


# ----------------- EXACT TEMPLATE RESUME BUILDER -----------------
def create_dense_resume(filepath, kit):
    doc = SimpleDocTemplate(
        filepath,
        pagesize=letter,
        leftMargin=28,
        rightMargin=28,
        topMargin=22,
        bottomMargin=22
    )
    styles = getSampleStyleSheet()
    story = []

    DARK_NAVY = colors.HexColor("#0A2540")
    TEXT_CHARCOAL = colors.HexColor("#1A202C")
    MUTED_GRAY = colors.HexColor("#4A5568")
    BORDER_GRAY = colors.HexColor("#CBD5E1")

    name_style = ParagraphStyle('HeaderName', fontSize=13, leading=15, fontName='Helvetica-Bold', textColor=DARK_NAVY)
    contact_style = ParagraphStyle('HeaderContact', fontSize=7.5, leading=10.5, fontName='Helvetica', textColor=MUTED_GRAY, alignment=2)
    section_head_style = ParagraphStyle('SectionHead', fontSize=9, leading=11, fontName='Helvetica-Bold', textColor=DARK_NAVY, spaceBefore=4, spaceAfter=1)
    left_bold_style = ParagraphStyle('LeftBold', fontSize=8, leading=10.5, fontName='Helvetica-Bold', textColor=TEXT_CHARCOAL)
    right_date_style = ParagraphStyle('RightDate', fontSize=8, leading=10.5, fontName='Helvetica-Bold', textColor=MUTED_GRAY, alignment=2)
    body_style = ParagraphStyle('BodyText', fontSize=7.8, leading=10, fontName='Helvetica', textColor=TEXT_CHARCOAL)
    role_desc_style = ParagraphStyle('RoleDesc', fontSize=7.8, leading=9.8, fontName='Helvetica-Oblique', textColor=TEXT_CHARCOAL, spaceAfter=1)
    subhead_style = ParagraphStyle('SubCategoryHead', fontSize=8, leading=10, fontName='Helvetica-Bold', textColor=DARK_NAVY, spaceBefore=2, spaceAfter=1)
    bullet_style = ParagraphStyle('CompactBullet', fontSize=7.5, leading=9.5, fontName='Helvetica', textColor=TEXT_CHARCOAL, leftIndent=8, spaceAfter=1)
    grid_cell_style = ParagraphStyle('GridCell', fontSize=7.8, leading=10, fontName='Helvetica', textColor=TEXT_CHARCOAL)

    # 1. TOP HEADER
    header_table = Table([
        [
            Paragraph("KARTIK BHATT", name_style),
            Paragraph("kb270102@gmail.com | +91-7428062532 | <a href='https://kartikb.vercel.app/'><u>Portfolio</u></a> | Delhi NCR, India", contact_style)
        ]
    ], colWidths=[200, 356])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2)
    ]))
    story.append(header_table)
    story.append(HRFlowable(width="100%", thickness=1, color=DARK_NAVY, spaceBefore=2, spaceAfter=4))

    # 2. EDUCATION
    story.append(Paragraph("EDUCATION", section_head_style))
    edu_table = Table([
        [
            Paragraph("<b>Maharaja Surajmal Institute</b> | Bachelor of Computer Applications | Majors: Computer Science | <b>GPA: 9.3/10 (Top 1%)</b>", body_style),
            Paragraph("Jul’19 – Aug’22", right_date_style)
        ]
    ], colWidths=[460, 96])
    edu_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0)
    ]))
    story.append(edu_table)
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER_GRAY, spaceBefore=1, spaceAfter=3))

    # 3. WORK EXPERIENCE
    exp_header_table = Table([
        [
            Paragraph("WORK EXPERIENCE", section_head_style),
            Paragraph("3 years 2 months", right_date_style)
        ]
    ], colWidths=[430, 126])
    exp_header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0)
    ]))
    story.append(exp_header_table)

    # KPMG Block
    kpmg_title_table = Table([
        [
            Paragraph("<b>KPMG</b> | Analyst | Knowledge Management | Gurugram", left_bold_style),
            Paragraph("May’24 – Present", right_date_style)
        ]
    ], colWidths=[440, 116])
    kpmg_title_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0)
    ]))
    story.append(kpmg_title_table)
    story.append(Paragraph("Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management & played key role in business development.", role_desc_style))

    story.append(Paragraph("Key Projects", subhead_style))
    for b in kit.get("kpmg_project_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))

    story.append(Paragraph("Business Development & Operations", subhead_style))
    for b in kit.get("kpmg_bd_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))

    story.append(Paragraph("Key Achievements & Recognitions", subhead_style))
    story.append(Paragraph("• Awarded 'KUDOS' twice: for Lean Six Sigma efficiency (2,000+ hrs saved) and migrating legacy Excel/VBA to GenAI agents + Power Platform.", bullet_style))
    story.append(Paragraph("• Honored with 'Super Team' (council leadership), 'Ally of Inclusion' (advocate), and 'Gurus@Work' (firmwide learning).", bullet_style))
    story.append(Spacer(1, 2))

    # GlobalLogic Block
    gl_title_table = Table([
        [
            Paragraph("<b>GlobalLogic Technologies Private Limited</b> | Associate Analyst | Content Engineering | Gurugram", left_bold_style),
            Paragraph("Sep’22 – Oct’23", right_date_style)
        ]
    ], colWidths=[440, 116])
    gl_title_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0)
    ]))
    story.append(gl_title_table)
    story.append(Paragraph("Participated in content generation and manipulation workflows for Tier-1 clients including Google.", role_desc_style))

    story.append(Paragraph("Key Deliverables & QA Engineering", subhead_style))
    for b in kit.get("globallogic_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))

    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER_GRAY, spaceBefore=1, spaceAfter=3))

    # 4. SKILLS SECTION (4-Column Balanced Grid matching template)
    story.append(Paragraph("SKILLS", section_head_style))
    skills_data = [
        [
            Paragraph("MS Excel", grid_cell_style),
            Paragraph("| Copilot GenAI (Agents)", grid_cell_style),
            Paragraph("| Power Automate", grid_cell_style),
            Paragraph("| Power BI", grid_cell_style)
        ],
        [
            Paragraph("SharePoint Online", grid_cell_style),
            Paragraph("| Power Apps", grid_cell_style),
            Paragraph("| Copilot Studio", grid_cell_style),
            Paragraph("| SQL & Python", grid_cell_style)
        ]
    ]
    skills_table = Table(skills_data, colWidths=[135, 145, 135, 141])
    skills_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
    ]))
    story.append(skills_table)
    story.append(Spacer(1, 2))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER_GRAY, spaceBefore=1, spaceAfter=3))

    # 5. CERTIFICATIONS SECTION (2-Column Grid matching template)
    story.append(Paragraph("CERTIFICATIONS", section_head_style))
    certs_data = [
        [
            Paragraph("Microsoft Certified: Azure AI Fundamentals (AI-901)", grid_cell_style),
            Paragraph("| Microsoft Certified: AI Transformation Leader (AB-731)", grid_cell_style)
        ],
        [
            Paragraph("Lean Six Sigma: Yellow Belt", grid_cell_style),
            Paragraph("| Microsoft Certified: AI Business Professional (AB-730)", grid_cell_style)
        ],
        [
            Paragraph("Oracle: Agentic AI Certified Foundations Associate", grid_cell_style),
            Paragraph("", grid_cell_style)
        ]
    ]
    certs_table = Table(certs_data, colWidths=[280, 276])
    certs_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1)
    ]))
    story.append(certs_table)

    doc.build(story)


# ----------------- COVER LETTER BUILDER -----------------
def create_dense_cover_letter(filepath, title, company, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=40, rightMargin=40, topMargin=35, bottomMargin=35)
    styles = getSampleStyleSheet()
    story = []

    NAVY = colors.HexColor("#0B2540")
    body_style = ParagraphStyle('CLBody', fontSize=9, leading=13.5, fontName='Helvetica', textColor=colors.HexColor("#1E293B"), spaceBefore=5)
    header_style = ParagraphStyle('CLHead', fontSize=14, leading=16, fontName='Helvetica-Bold', textColor=NAVY)
    sub_style = ParagraphStyle('CLSub', fontSize=8.5, leading=11, fontName='Helvetica', textColor=colors.HexColor("#475569"))
    subj_style = ParagraphStyle('CLSubj', fontSize=9.5, leading=12, fontName='Helvetica-Bold', textColor=NAVY, spaceBefore=4, spaceAfter=4)

    story.append(Paragraph("KARTIK BHATT", header_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | Delhi NCR | kartikb.vercel.app", sub_style))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=1, color=NAVY, spaceAfter=6))

    story.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%B %d, %Y')}", body_style))
    story.append(Paragraph(f"<b>Target Role:</b> {title} | <b>Company:</b> {company}", body_style))
    story.append(Spacer(1, 4))
    story.append(Paragraph(kit.get("cover_letter_subject", f"Subject: Application for {title}"), subj_style))
    story.append(Paragraph("Dear Hiring Team,", body_style))

    for p in kit.get("cover_letter_paragraphs", []):
        story.append(Paragraph(p, body_style))

    story.append(Spacer(1, 6))
    story.append(Paragraph("Warm regards,<br/><b>Kartik Bhatt</b>", body_style))
    doc.build(story)


# ----------------- TELEGRAM CARD DISPATCHER -----------------
def send_telegram_alert(title, company, location, exp_detected, salary_range, kit, resume_link, cl_link, job_url):
    fit = kit.get("match_score", "N/A")
    gaps = kit.get("skills_gap", "None")

    msg = (
        f"🎯 <b>NEW HIGH-FIT OPPORTUNITY</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"💼 <b>Role:</b> {title}\n"
        f"🏢 <b>Company:</b> {company}\n"
        f"📍 <b>Location:</b> {location}\n"
        f"⏳ <b>Exp. Demanded:</b> {exp_detected}\n"
        f"💰 <b>Est. CTC:</b> {salary_range}\n"
        f"📊 <b>ATS Fit Score:</b> <code>{fit}</code>\n"
        f"⚠️ <b>Skill Gaps:</b> {gaps}\n\n"
        f"💡 <b>Match Alignment:</b>\n<i>{kit.get('reason', '')}</i>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"📄 <a href='{resume_link}'><b>Download Tailored Resume (PDF)</b></a>\n"
        f"📝 <a href='{cl_link}'><b>Download Cover Letter (PDF)</b></a>\n"
        f"🔗 <a href='{job_url}'><b>Apply Directly on Portal</b></a>\n\n"
        f"<i>Note: PDF links go live shortly once pushed to GitHub repo.</i>"
    )

    endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        res = requests.post(endpoint, data={
            "chat_id": CHAT_ID,
            "text": msg,
            "parse_mode": "HTML",
            "disable_web_page_preview": True
        }, timeout=10)
        if res.status_code != 200:
            print(f"Telegram dispatch returned {res.status_code}: {res.text}")
    except Exception as e:
        print(f"Telegram dispatch error: {e}")


# ----------------- MAIN PIPELINE -----------------
def run():
    init_tracker()
    all_jobs = []

    # 1. Scrape standard job boards
    try:
        board_jobs = scrape_jobs(
            site_name=["linkedin"],
            search_term='"Power Platform" OR "Copilot" OR "Power Automate" OR "Business Analyst" OR "Data Analyst"',
            location="India",
            results_wanted=35,
            hours_old=48,
            country_indeed='india'
        )
        for _, row in board_jobs.iterrows():
            job_dict = row.to_dict()
            job_dict["is_direct_ats"] = False
            all_jobs.append(job_dict)
        print(f"[DEBUG] Job boards returned {len(board_jobs)} listings.")
    except Exception as e:
        print(f"[DEBUG] Board scraper warning: {e}")

    # 2. Scrape direct enterprise ATS portals
    ats_jobs = search_enterprise_ats_jobs()
    print(f"[DEBUG] Enterprise ATS returned {len(ats_jobs)} listings.")
    all_jobs.extend(ats_jobs)

    print(f"[DEBUG] Total raw jobs to process: {len(all_jobs)}")

    dispatched = 0

    for job in all_jobs:
        if dispatched >= 8:
            break

        url = str(job.get('job_url') or '')
        title = str(job.get('title') or '')
        company = str(job.get('company') or '')
        raw_location = str(job.get('location') or '')
        desc = str(job.get('description') or '')

        if not url or url == 'nan' or is_seen(url):
            continue

        if job.get("is_direct_ats") and len(desc) < 300:
            full_desc = fetch_portal_description(url)
            if full_desc:
                desc = full_desc

        # Filter keywords and role fit
        if not is_role_and_skill_relevant(title, desc):
            continue

        # Location check
        loc_valid, loc_tier, clean_location = classify_location(raw_location, fallback_text=desc)
        if not loc_valid:
            continue

        # Salary check
        passes_sal, salary_range = get_salary_range_and_check(
            title=title,
            company=company,
            location=clean_location,
            description=desc,
            tier=loc_tier,
            min_sal=job.get('min_amount'),
            max_sal=job.get('max_amount')
        )
        if not passes_sal:
            continue

        # Generate custom kit with exact scoring & experience detection
        kit = generate_application_kit(title, company, desc)
        if not kit:
            continue

        # Check experience ceiling
        if kit.get("exceeds_cap", False):
            print(f"[DEBUG] Skipped '{title}' @ {company} (Experience required > {MAX_EXPERIENCE_CAP} years)")
            continue

        safe_company = re.sub(r'[^a-zA-Z0-9]', '_', company)[:15]
        safe_title = re.sub(r'[^a-zA-Z0-9]', '_', title)[:15]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        resume_filename = f"Resume_{safe_company}_{safe_title}_{timestamp}.pdf"
        cl_filename = f"CoverLetter_{safe_company}_{safe_title}_{timestamp}.pdf"
        resume_path = os.path.join(DOCS_DIR, resume_filename)
        cl_path = os.path.join(DOCS_DIR, cl_filename)

        create_dense_resume(resume_path, kit)
        create_dense_cover_letter(cl_path, title, company, kit)

        resume_link = github_raw_link(resume_path)
        cl_link = github_raw_link(cl_path)
        exp_text = kit.get("detected_experience", "2–5 Years")

        send_telegram_alert(
            title=title,
            company=company,
            location=clean_location,
            exp_detected=exp_text,
            salary_range=salary_range,
            kit=kit,
            resume_link=resume_link,
            cl_link=cl_link,
            job_url=url
        )

        log_job(title, company, salary_range, kit.get("match_score"), clean_location, resume_link, cl_link, url)
        append_to_log(title, company, salary_range, kit.get("match_score"), clean_location, resume_link, cl_link, url)

        dispatched += 1
        print(f"Successfully processed & notified: {title} at {company}")

    print(f"\n[DEBUG] Finished run. Dispatched {dispatched} matched applications.")


if __name__ == "__main__":
    run()
