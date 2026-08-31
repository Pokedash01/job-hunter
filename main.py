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

# ReportLab imports for dense, professional PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY") or os.environ.get("SEARCH_API_KEYSEARCH_API_KEY")

client = genai.Client(api_key=GEMINI_KEY)

# ----------------- KARTIK BHATT MASTER PROFILE -----------------
KARTIK_PROFILE = """
Candidate: Kartik Bhatt
Contact: kb270102@gmail.com | +91-7428062532 | Portfolio: https://kartikb.vercel.app/
Education: Bachelor of Computer Applications, Majors: Computer Science, Maharaja Surajmal Institute | GPA: 9.3/10 (Top 1%) | Jul 2019 - Aug 2022
Experience Summary: ~3.5+ years of cross-functional experience across KPMG and GlobalLogic driving process automation, Copilot/GenAI agents, Power Platform solutions, and operational excellence.

Work Experience:
1. KPMG | Analyst | Knowledge Management | Gurugram (May 2024 - Present)
Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management and played a key role in business development.
Key Projects:
- Built complete Power Platform solution (Power Automate, SharePoint lists, Power Apps, Power BI) facilitating 20,000 reach outs annually across 30+ member firms in 13 sectors, saving 1,200 hrs annually.
- Built end-to-end solution to facilitate migration of old Excel-based data collection for 45+ pillars to automated SPO list integration, including 3 Power Automate flows for alerts, change management, data migration, and permission governance.
- Built & managed multiple VBA macro solutions refreshing a repository of 30,000+ assets globally, including change management, saving 485 hrs annually.
- Architected a multi-modal Copilot agent to assist with messy data, draft fields, and apply metadata tags based on source and guidelines, saving 325 hrs annually.

Business Development & Operations:
- Saved 2,000+ hrs annually leveraging Power Platform, Copilot Studio, and VBA Macros.
- Handled 100+ RFP/RFI requests and built/maintained 100+ internal site pages per brand values and standards.
- Undertook complete contact management system administration for 10,000+ KGS members.
- Uploaded 5,000+ content assets across 3 content types and 15 libraries.
- Performed Audit Market Share (AMS) analysis for 6+ sectors.
- Catered to 50+ SharePoint governance and administration requests, including term store management, change management, metadata management, and permission-level management.

Awards: KUDOS (efficiency & Lean Six Sigma, saving 2,000+ hrs annually), KUDOS (migrating legacy VBA/Excel practices to GenAI-focused agent + Power Platform workflows), Super Team Award (hosting/organizing employee council events), Ally of Inclusion, Gurus@Work (contributions to learning culture).

2. GlobalLogic Technologies Private Limited | Associate Analyst | Content Engineering | Gurugram (Sep 2022 - Oct 2023)
Took part in content generation and manipulation projects for clients including Google.
- Created best practices, process docs, and QA processes for a Google project to build test & training datasets for GenAI screen search on Android.
- Piloted a project to extract relevant answers from multi-level docs to build a training dataset for AI.
- Designed and implemented QA processes for data entry, reducing errors by 25% and improving data accuracy.
- Managed process documentation for 10+ projects, ensuring compliance and accessibility for stakeholders.
- Improved onshore project delivery quality from 74% to 95%; QA'd 500+ pieces weekly and led 3 pilot projects, securing all of them against competition from major MNCs.

Skills: Microsoft Copilot / GenAI Agents, Copilot Studio, Power Apps, Power Automate, Power BI, SharePoint Online, MS Excel/VBA, SQL, MySQL, Process Automation, Stakeholder Management, Product/Data Analytics, RFP/RFI Management, Change Management, Digital Transformation.
Certifications: Microsoft Certified: Azure AI Fundamentals (AI-901), Microsoft Certified: AI Transformation Leader (AB-731), Microsoft Certified: AI Business Professional (AB-730), Lean Six Sigma: Yellow Belt, Oracle: Agentic AI Certified Foundations Associate.
"""

# ----------------- FILTERING CONSTANTS -----------------
SENIORITY_BLACKLIST = [
    r"\bsenior manager\b", r"\bprincipal\b", r"\bdirector\b", r"\bvp\b",
    r"\bhead of\b", r"\bgroup product manager\b", r"\btech lead\b",
    r"\bengineering manager\b", r"\bgeneral manager\b", r"\blead architect\b",
    r"\bassociate director\b", r"\bavp\b", r"\boperations director\b"
]
ROLES_WHITELIST = [
    r"\bdata analyst\b", r"\bbusiness analyst\b", r"\bproduct analyst\b",
    r"\bassociate product manager\b", r"\bapm\b", r"\bcopilot studio\b",
    r"\bpower automate\b", r"\bpower platform\b", r"\bbi developer\b",
    r"\bprocess automation\b", r"\banalytics engineer\b", r"\boperations analyst\b",
    r"\bproduct operations\b"
]
EXCLUDED_DOMAINS = [
    "hr", "human resources", "talent acquisition", "recruiter", "recruitment",
    "sales", "business development executive", "bde", "marketing", "digital marketing",
    "telecaller", "content writer", "seo", "graphic designer", "accountant"
]
LOCATIONS_TIER_1 = [r"\bdelhi\b", r"\bncr\b", r"\bgurgaon\b", r"\bgurugram\b", r"\bnoida\b", r"\bfaridabad\b", r"\bremote\b", r"\bwfh\b", r"\bwork from home\b"]
LOCATIONS_TIER_2 = [r"\bbangalore\b", r"\bbengaluru\b", r"\bhyderabad\b", r"\bpune\b"]
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


# ----------------- SMART SCREENER & FILTERS -----------------
def is_seniority_excluded(title: str) -> bool:
    t = title.lower()
    return any(re.search(pat, t) for pat in SENIORITY_BLACKLIST)


def is_role_relevant(title: str) -> bool:
    t = title.lower()
    for bad in EXCLUDED_DOMAINS:
        if re.search(r'\b' + re.escape(bad) + r'\b', t):
            return False
    if is_seniority_excluded(title):
        return False
    return any(re.search(good, t) for good in ROLES_WHITELIST)


def extract_min_experience(text: str):
    patterns = [
        r"(\d+)\s*(?:-|to|\+)\s*(?:\d+)?\s*(?:years|yrs)",
        r"(?:minimum|at least|over|requires?)\s*(\d+)\s*(?:years|yrs)"
    ]
    years = []
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            years.append(int(m.group(1)))
    return min(years) if years else None


def classify_location(location_str: str):
    loc = str(location_str).lower()
    if any(re.search(pat, loc) for pat in LOCATIONS_TIER_1):
        return True, "Tier 1 (Delhi-NCR / Remote)"
    if any(re.search(pat, loc) for pat in LOCATIONS_TIER_2):
        return True, "Tier 2 (Bangalore / Hyderabad / Pune)"
    return False, "Out of Scope Location"


# ----------------- LLM SALARY ESTIMATION & FLOOR CHECK -----------------
def get_salary_range_and_check(title: str, company: str, location: str, description: str, tier: str, min_sal=None, max_sal=None):
    min_floor = 1000000 if "Tier 1" in tier else 1500000

    if max_sal and float(max_sal) > 0:
        actual_min = float(min_sal) if min_sal and float(min_sal) > 0 else float(max_sal) * 0.75
        actual_max = float(max_sal)
        if actual_max < min_floor:
            return False, ""
        return True, f"₹{actual_min/100000:.1f} - ₹{actual_max/100000:.1f} LPA"

    prompt = f"""
You are an expert Indian tech compensation analyst.
Analyze the following role for Kartik Bhatt (~3.5 YOE, Power Platform, Analytics, GenAI Agents):

Job Title: {title}
Company: {company}
Location: {location} ({tier})
Job Description Snippet: {description[:2000]}

Task:
1. Extract any explicitly mentioned salary.
2. If NO salary is disclosed, estimate the realistic 25th-75th percentile market base CTC in INR for a 3-4 YOE candidate in India for this company/role.
3. Determine `passes_floor`: true if max_inr >= {min_floor}, false otherwise.

Return ONLY a valid JSON object matching this schema:
{{
  "min_inr": 1200000,
  "max_inr": 1600000,
  "display_range": "₹12 - ₹16 LPA (Est.)",
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
        return True, data.get("display_range", f"₹{min_floor/100000:.0f}+ LPA (Est.)")
    except Exception as e:
        print(f"Salary check error for {title} @ {company}: {e}")
        fallback_lpa = 12 if "Tier 1" in tier else 16
        return True, f"₹{fallback_lpa} - ₹{fallback_lpa+4} LPA (Est.)"


# ----------------- DIRECT ATS & COMPANY PORTAL SEARCH -----------------
def extract_company_from_url(url: str) -> str:
    """Extracts a clean company name from direct career portal URLs."""
    try:
        domain = urlparse(url).netloc.lower()
        path = urlparse(url).path.lower()
        
        # e.g., company.greenhouse.io or company.wd3.myworkdayjobs.com
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
    """Extracts raw text content from direct ATS job pages."""
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        res = requests.get(url, headers=headers, timeout=8)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, "html.parser")
            for tag in soup(["script", "style", "nav", "footer", "header"]):
                tag.decompose()
            text = ' '.join(soup.stripped_strings)
            return text[:4000]
    except Exception as e:
        print(f"Error fetching JD from portal {url}: {e}")
    return ""


def search_enterprise_ats_jobs():
    """
    Searches enterprise ATS systems and direct company career portals for roles in India.
    """
    ats_queries = [
        'site:myworkdayjobs.com ("Data Analyst" OR "Business Analyst" OR "Power Platform" OR "Copilot Studio") ("India" OR "Gurgaon" OR "Gurugram" OR "Noida" OR "Bangalore")',
        'site:boards.greenhouse.io OR site:job-boards.greenhouse.io ("Product Analyst" OR "Data Analyst" OR "Process Automation") ("India" OR "Remote" OR "Gurgaon" OR "Bangalore")',
        'site:jobs.lever.co ("Data Analyst" OR "Business Analyst" OR "APM" OR "Power Automate") ("India" OR "Remote" OR "Gurugram")',
        'site:jobs.ashbyhq.com ("Product Analyst" OR "Data Analyst" OR "Analytics Engineer" OR "Business Analyst") ("India" OR "Remote")',
        'site:jobs.smartrecruiters.com ("Process Automation" OR "Power Platform" OR "Data Analyst" OR "Operations Analyst") India',
        '(site:careers.google.com OR site:amazon.jobs OR site:jobs.siemens.com OR site:careers.microsoft.com) ("Analyst" OR "Automation") ("Gurugram" OR "Noida" OR "Delhi" OR "Bangalore")'
    ]
    
    discovered = []
    if not SEARCH_KEY:
        print("[DEBUG] No SEARCH_KEY found. Skipping direct company portal search.")
        return discovered

    for query in ats_queries:
        try:
            url = "https://www.searchapi.io/api/v1/search"
            params = {
                "engine": "google",
                "q": query,
                "api_key": SEARCH_KEY,
                "gl": "in",
                "hl": "en",
                "num": 10
            }
            res = requests.get(url, params=params, timeout=15).json()
            for item in res.get("organic_results", []):
                link = item.get("link", "")
                raw_title = item.get("title", "")
                snippet = item.get("snippet", "")
                
                # Clean title
                title = re.sub(r"\s*[-|–]\s*(Greenhouse|Lever|Workday|Ashby|SmartRecruiters|Jobs).*", "", raw_title, flags=re.IGNORECASE).strip()
                company = extract_company_from_url(link)
                
                if link and is_role_relevant(title):
                    # Fetch direct page context or use snippet
                    full_desc = fetch_portal_description(link) or snippet
                    
                    discovered.append({
                        "title": title,
                        "company": company,
                        "job_url": link,
                        "location": "Delhi NCR / Remote / Bangalore",
                        "description": full_desc,
                        "min_amount": 0,
                        "max_amount": 0
                    })
        except Exception as e:
            print(f"ATS search error for query '{query}': {e}")
            
    return discovered


# ----------------- GEMINI ASSET GENERATION -----------------
def generate_application_kit(title, company, description):
    prompt = f"""
You are an expert executive resume writer, ATS optimization specialist, and career strategist
tailoring application documents for Kartik Bhatt (~3.5 years of experience).

Candidate Profile (ground truth — never fabricate skills or employment):
{KARTIK_PROFILE}

Target Job Opening:
Title: {title}
Company: {company}
JD Snippet: {description[:2200]}

CRITICAL RULES FOR RESUME VS COVER LETTER:
1. RESUME PROFESSIONAL SUMMARY RULES:
- DO NOT mention "{company}" or "seeking to work at {company}" anywhere in the tailored_summary.
- Frame it around his 3.5+ years of experience, core domains (Power Platform, GenAI Agents, Data/Product Analytics, Automation), and quantifiable business impact.

2. ATS KEYWORDS & SKILLS ALIGNMENT:
- Extract 10-14 exact technical & functional skills directly from the JD backed by Kartik's profile.

3. QUANTIFIED BULLETS:
- Keep KPMG and GlobalLogic bullets dense, factual, and metric-driven.

4. COVER LETTER:
- Actively reference {company} and {title}, explaining why Kartik is interested and how his results solve their needs.

Return ONLY a valid JSON object matching this schema:
{{
"match_score": "e.g., 94%",
"reason": "1-2 sentences on why Kartik's exact tech stack and experience fit this role.",
"skills_gap": "Any missing tool/skill or 'None'",
"ats_keywords": ["10-14 exact matching technical/domain keywords"],
"tailored_summary": "A 3-4 sentence dense executive summary highlighting Kartik's expertise in process automation, analytics, and GenAI agent development. Must NOT mention {company}.",
"kpmg_project_bullets": [
    "Quantified bullet on Power Platform solution: 20,000 reach outs, 30+ member firms, 13 sectors, 1,200 hrs saved",
    "Quantified bullet on SPO migration: 45+ pillars, 3 Power Automate flows, change management",
    "Quantified bullet on VBA macro repository: 30,000+ assets, 485 hrs saved",
    "Quantified bullet on multi-modal Copilot agent: messy data cleanup, field drafting, metadata tagging, 325 hrs saved",
    "Quantified bullet on contact management administration for 10,000+ members and 5,000+ assets"
],
"kpmg_bd_bullets": [
    "Saved 2,000+ hrs annually leveraging Power Platform, Copilot Studio, and VBA Macros.",
    "Handled 100+ RFP/RFI requests and built/maintained 100+ internal site pages aligned with brand guidelines.",
    "Performed Audit Market Share (AMS) analysis across 6+ sectors.",
    "Catered to 50+ SharePoint governance and administration requests."
],
"globallogic_bullets": [
    "Designed best practices, process docs, and QA workflows for Google GenAI Android search datasets.",
    "Designed and implemented QA processes for data entry, reducing errors by 25%.",
    "Managed process documentation across 10+ engagements, lifting delivery quality from 74% to 95%.",
    "QA'd 500+ pieces weekly and led 3 pilot projects to completion against major MNC competition."
],
"cover_letter_subject": "Subject: Driving Operational Excellence & Scalable Solutions as {title}",
"cover_letter_paragraphs": [
    "Opening paragraph: enthusiasm for {title} at {company}, JD hook, and 1-line summary of 3.5+ year background.",
    "Second paragraph: 2-3 concrete KPMG achievements mapped directly to the JD requirements.",
    "Third paragraph: GlobalLogic experience mapped to relevant JD requirements.",
    "Closing paragraph: connecting skills to {company}'s goals, restating enthusiasm, thanking reader, inviting next steps."
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
    except Exception as e:
        print(f"AI Generation error: {e}")
        return {
            "match_score": "88%",
            "reason": "Strong alignment with Kartik's Power Platform, GenAI agent building, and operations/analytics background.",
            "skills_gap": "None",
            "ats_keywords": [
                "Power Platform", "Power Automate", "Power BI", "Power Apps", "Copilot Studio",
                "GenAI Agents", "SQL", "Process Automation", "Stakeholder Management",
                "Change Management", "SharePoint Online", "Data Analytics", "RFP/RFI Management"
            ],
            "tailored_summary": "Solutions-driven analyst with ~3.5+ years of experience across KPMG and GlobalLogic specializing in enterprise process automation, Microsoft Power Platform architectures, GenAI agent implementation, and data-driven operational optimization.",
            "kpmg_project_bullets": [
                "Built complete Power Platform solution (Power Automate, SharePoint lists, Power Apps, Power BI) facilitating 20,000 reach outs annually across 30+ member firms in 13 sectors, saving 1,200 hrs annually.",
                "Built end-to-end migration of Excel-based data collection for 45+ pillars to automated SPO lists, including 3 Power Automate flows.",
                "Built & managed VBA macro solutions refreshing 30,000+ assets globally, saving 485 hrs annually.",
                "Architected a multi-modal Copilot agent for messy data cleanup, field drafting, and metadata tagging, saving 325 hrs annually.",
                "Administered contact management system for 10,000+ members and uploaded 5,000+ content assets across 15 libraries."
            ],
            "kpmg_bd_bullets": [
                "Saved 2,000+ hrs annually across Power Platform, Copilot Studio, and VBA Macros.",
                "Handled 100+ RFP/RFI requests and maintained 100+ internal site pages per brand standards.",
                "Performed Audit Market Share (AMS) analysis across 6+ sectors.",
                "Catered to 50+ SharePoint governance and administration requests."
            ],
            "globallogic_bullets": [
                "Created best practices and QA processes for a Google GenAI training dataset project for Android screen search.",
                "Reduced data entry errors by 25% through redesigned QA processes.",
                "Improved project delivery quality from 74% to 95%, managing process documentation for 10+ projects.",
                "QA'd 500+ pieces weekly and led 3 pilot projects to completion against major MNC competition."
            ],
            "cover_letter_subject": f"Subject: Application for {title} - Kartik Bhatt",
            "cover_letter_paragraphs": [
                f"I am excited to apply for the {title} position at {company}. With over 3.5 years of experience at KPMG and GlobalLogic, I specialize in process automation, Power Platform ecosystems, and Copilot AI agents.",
                "At KPMG, I built a complete Power Platform solution driving 20,000 reach outs annually across 13 sectors and architected an automated multi-modal Copilot agent.",
                "At GlobalLogic, I engineered QA frameworks for Google's GenAI training datasets, lifting delivery quality from 74% to 95%.",
                f"I would welcome the opportunity to bring this operational and technical expertise to {company}. Thank you for your time and consideration."
            ]
        }


# ----------------- DENSE PDF BUILDERS -----------------
def create_dense_resume(filepath, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=30, rightMargin=30, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    story = []

    name_style = ParagraphStyle('Name', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor("#0F2942"), alignment=1)
    contact_style = ParagraphStyle('Contact', parent=styles['Normal'], fontSize=8.5, leading=11, alignment=1, textColor=colors.HexColor("#334155"))
    section_style = ParagraphStyle('SecHeader', parent=styles['Heading2'], fontSize=10, leading=12, textColor=colors.HexColor("#0F2942"), spaceBefore=4, spaceAfter=1)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor("#1E293B"))
    company_style = ParagraphStyle('Comp', parent=styles['Normal'], fontSize=9, leading=11.5, fontName='Helvetica-Bold', textColor=colors.HexColor("#0F2942"))
    subhead_style = ParagraphStyle('SubHead', parent=styles['Normal'], fontSize=8.5, leading=10.5, fontName='Helvetica-Bold', textColor=colors.HexColor("#334155"), leftIndent=6)
    bullet_style = ParagraphStyle('Bullet', parent=styles['Normal'], fontSize=8, leading=10.5, leftIndent=12, textColor=colors.HexColor("#1E293B"))

    story.append(Paragraph("<b>KARTIK BHATT</b>", name_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | LinkedIn | Website: https://kartikb.vercel.app/ | Delhi NCR, India", contact_style))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=0.75, color=colors.HexColor("#94A3B8"), spaceAfter=3))

    story.append(Paragraph("<b>PROFESSIONAL SUMMARY</b>", section_style))
    story.append(Paragraph(kit.get("tailored_summary", ""), body_style))
    story.append(Spacer(1, 3))

    ats_keywords = kit.get("ats_keywords", [])
    if ats_keywords:
        story.append(Paragraph("<b>CORE COMPETENCIES & KEYWORDS</b>", section_style))
        story.append(Paragraph(" • ".join(ats_keywords), body_style))
        story.append(Spacer(1, 3))

    story.append(Paragraph("<b>EDUCATION</b>", section_style))
    story.append(Paragraph("<b>Maharaja Surajmal Institute</b> | Bachelor of Computer Applications (Computer Science) | <b>GPA: 9.3/10 (Top 1%)</b>", body_style))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>WORK EXPERIENCE</b>", section_style))
    story.append(Paragraph("<b>KPMG</b> | Analyst — Knowledge Management | Gurugram <i>(May 2024 – Present | ~3.5+ yrs total exp)</i>", company_style))
    story.append(Paragraph("<i>Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management & business development.</i>", body_style))
    story.append(Paragraph("Key Projects", subhead_style))
    for b in kit.get("kpmg_project_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Paragraph("Business Development & Operations", subhead_style))
    for b in kit.get("kpmg_bd_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Paragraph("Key Achievements: Awarded 'KUDOS' twice — for Lean Six Sigma efficiency (2,000+ hrs saved) and for migrating legacy VBA/Excel practices to GenAI agents & Power Platform — plus 'Super Team', 'Ally of Inclusion', and 'Gurus@Work'.", bullet_style))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>GlobalLogic Technologies Private Limited</b> | Associate Analyst — Content Engineering | Gurugram <i>(Sep 2022 – Oct 2023)</i>", company_style))
    for b in kit.get("globallogic_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>SKILLS & CERTIFICATIONS</b>", section_style))
    story.append(Paragraph("<b>Technical Stack:</b> Microsoft Copilot Studio, Power Automate, Power Apps, Power BI, SharePoint Online, Python, SQL, MySQL, Advanced Excel/VBA, Process Mining, API Integrations", body_style))
    story.append(Paragraph("<b>Domain & Operations:</b> Process Automation, Digital Transformation, Stakeholder Management, Product Analytics, RFP/RFI Bidding, Change Management, Data Governance", body_style))
    story.append(Paragraph("<b>Certifications:</b> Microsoft Certified: Azure AI Fundamentals (AI-901) | Microsoft Certified: AI Transformation Leader (AB-731) | Microsoft Certified: AI Business Professional (AB-730) | Lean Six Sigma: Yellow Belt | Oracle: Agentic AI Certified Foundations Associate", body_style))

    doc.build(story)


def create_dense_cover_letter(filepath, title, company, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=45, rightMargin=45, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    name_style = ParagraphStyle('Name', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor("#0F2942"))
    contact_style = ParagraphStyle('Contact', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#4A5568"))
    subj_style = ParagraphStyle('Subj', parent=styles['Normal'], fontSize=10, leading=13, fontName='Helvetica-Bold', textColor=colors.HexColor("#0F2942"), spaceBefore=6, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=9.5, leading=14, spaceBefore=6, textColor=colors.HexColor("#1E293B"))

    story.append(Paragraph("<b>KARTIK BHATT</b>", name_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | LinkedIn | Portfolio: https://kartikb.vercel.app/", contact_style))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CBD5E0"), spaceAfter=8))

    story.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%B %d, %Y')}", body_style))
    story.append(Paragraph(f"<b>Target Role:</b> {title}", body_style))
    story.append(Paragraph(f"<b>Company:</b> {company}", body_style))
    story.append(Spacer(1, 4))

    story.append(Paragraph(f"<b>{kit.get('cover_letter_subject', 'Subject: Application for ' + title)}</b>", subj_style))
    story.append(Paragraph("Dear Hiring Manager,", body_style))
    for p in kit.get("cover_letter_paragraphs", []):
        story.append(Paragraph(p, body_style))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Sincerely,<br/><b>Kartik Bhatt</b>", body_style))

    doc.build(story)


# ----------------- TELEGRAM DISPATCHER -----------------
def send_telegram_alert(title, company, location, tier, exp_detected, url, salary_range, kit, resume_link, cl_link):
    safe_title = title.replace("*", "").replace("_", " ")
    safe_company = company.replace("*", "").replace("_", " ")
    safe_location = location.replace("*", "").replace("_", " ")

    # Identify whether source is direct portal or job board
    source_tag = "🏢 Company Career Portal" if any(k in url for k in ["greenhouse", "lever.co", "workday", "ashby", "smartrecruiters", "amazon", "google", "microsoft"]) else "💼 Job Board (LinkedIn)"

    message_text = (
        f"🎯 *New Qualified Job Matched for Kartik!*\n\n"
        f"📌 *Role:* {safe_title}\n"
        f"🏢 *Company:* {safe_company} ({source_tag})\n"
        f"📍 *Location:* {safe_location} ({tier})\n"
        f"⏳ *Experience Required:* {exp_detected}\n"
        f"💰 *Salary Range:* {salary_range}\n"
        f"📊 *Fit Score:* {kit.get('match_score')}\n"
        f"⚠️ *Skill Gap:* {kit.get('skills_gap')}\n\n"
        f"🔗 [Apply Directly on Portal]({url})\n"
        f"📄 [Download Resume PDF]({resume_link})\n"
        f"📝 [Download Cover Letter PDF]({cl_link})\n\n"
        f"_Links go live within ~1 min once pushed to repo._"
    )

    endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        data = {
            "chat_id": CHAT_ID,
            "text": message_text[:4096],
            "parse_mode": "Markdown",
            "disable_web_page_preview": True
        }
        res = requests.post(endpoint, data=data)
        if res.status_code != 200:
            print(f"Telegram message warning: {res.text}")
    except Exception as e:
        print(f"Telegram dispatch error: {e}")


# ----------------- MAIN PIPELINE -----------------
def run():
    init_tracker()

    all_jobs = []

    # 1. Scrape standard job boards
    try:
        board_jobs = scrape_jobs(
            site_name=["linkedin", "indeed"],
            search_term='"Data Analyst" OR "Business Analyst" OR "Product Analyst" OR "Associate Product Manager" OR "Copilot Studio" OR "Power Automate" OR "Power Platform"',
            location="India",
            results_wanted=35,
            hours_old=24,
            country_indeed='india'
        )
        for _, row in board_jobs.iterrows():
            all_jobs.append(row.to_dict())
        print(f"[DEBUG] Job boards returned {len(board_jobs)} raw listings.")
    except Exception as e:
        print(f"[DEBUG] Scraper error: {e}")

    # 2. Add Direct ATS & Company Portal Hits
    ats_jobs = search_enterprise_ats_jobs()
    print(f"[DEBUG] Direct Enterprise ATS search returned {len(ats_jobs)} raw portal listings.")
    all_jobs.extend(ats_jobs)

    print(f"[DEBUG] Total raw candidates this run: {len(all_jobs)}")

    skip_counts = {
        "no_url": 0,
        "already_seen": 0,
        "senior_or_irrelevant": 0,
        "exp_too_high": 0,
        "invalid_location": 0,
        "salary_check": 0
    }

    tier_1_matches = []
    tier_2_matches = []

    for job in all_jobs:
        url = str(job.get('job_url') or '')
        title = str(job.get('title') or '')
        company = str(job.get('company') or '')
        location = str(job.get('location') or 'India')
        min_sal = job.get('min_amount')
        max_sal = job.get('max_amount')
        desc = str(job.get('description') or '')
        full_text = f"{title} {desc}"

        if not url or url == 'nan':
            skip_counts["no_url"] += 1
            continue

        if is_seen(url):
            skip_counts["already_seen"] += 1
            continue

        # Role & Seniority Validation
        if not is_role_relevant(title):
            skip_counts["senior_or_irrelevant"] += 1
            continue

        # Experience Cap Filtering
        min_exp = extract_min_experience(full_text)
        if min_exp is not None and min_exp > MAX_EXPERIENCE_CAP:
            skip_counts["exp_too_high"] += 1
            print(f"[DEBUG] Rejected (Exp {min_exp}+ yrs > {MAX_EXPERIENCE_CAP}): '{title}' @ {company}")
            continue

        # Location Tiering Check
        loc_valid, loc_tier = classify_location(location)
        if not loc_valid:
            skip_counts["invalid_location"] += 1
            continue

        # Salary Extraction / Estimation & Minimum Floor Verification
        passes_sal, salary_range = get_salary_range_and_check(
            title=title,
            company=company,
            location=location,
            description=desc,
            tier=loc_tier,
            min_sal=min_sal,
            max_sal=max_sal
        )

        if not passes_sal:
            skip_counts["salary_check"] += 1
            print(f"[DEBUG] Rejected (Below Minimum Floor): '{title}' @ {company}")
            continue

        exp_str = f"{min_exp}+ Years" if min_exp else "2–5 Years / Unspecified"

        job_payload = {
            "title": title,
            "company": company,
            "location": location,
            "tier": loc_tier,
            "exp_detected": exp_str,
            "url": url,
            "salary_range": salary_range,
            "desc": desc
        }

        if "Tier 1" in loc_tier:
            tier_1_matches.append(job_payload)
        else:
            tier_2_matches.append(job_payload)

    # Prioritize Tier 1; fallback to Tier 2 if volume is low
    dispatch_queue = list(tier_1_matches)
    if len(tier_1_matches) < 3:
        dispatch_queue.extend(tier_2_matches)

    dispatched = 0
    for qualified in dispatch_queue:
        kit = generate_application_kit(qualified["title"], qualified["company"], qualified["desc"])

        safe_company = re.sub(r'[^a-zA-Z0-9]', '_', qualified["company"])[:15]
        safe_title = re.sub(r'[^a-zA-Z0-9]', '_', qualified["title"])[:15]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        resume_filename = f"Resume_{safe_company}_{safe_title}_{timestamp}.pdf"
        cl_filename = f"CoverLetter_{safe_company}_{safe_title}_{timestamp}.pdf"
        resume_path = os.path.join(DOCS_DIR, resume_filename)
        cl_path = os.path.join(DOCS_DIR, cl_filename)

        create_dense_resume(resume_path, kit)
        create_dense_cover_letter(cl_path, qualified["title"], qualified["company"], kit)

        resume_link = github_raw_link(resume_path)
        cl_link = github_raw_link(cl_path)

        send_telegram_alert(
            title=qualified["title"],
            company=qualified["company"],
            location=qualified["location"],
            tier=qualified["tier"],
            exp_detected=qualified["exp_detected"],
            url=qualified["url"],
            salary_range=qualified["salary_range"],
            kit=kit,
            resume_link=resume_link,
            cl_link=cl_link
        )

        log_job(
            title=qualified["title"],
            company=qualified["company"],
            salary_range=qualified["salary_range"],
            fit_score=kit.get("match_score"),
            location=f'{qualified["location"]} ({qualified["tier"]})',
            resume_link=resume_link,
            cover_letter_link=cl_link,
            job_link=qualified["url"]
        )

        append_to_log(
            title=qualified["title"],
            company=qualified["company"],
            salary_range=qualified["salary_range"],
            fit_score=kit.get("match_score"),
            location=f'{qualified["location"]} ({qualified["tier"]})',
            resume_link=resume_link,
            cover_letter_link=cl_link,
            job_link=qualified["url"]
        )

        dispatched += 1
        print(f"Dispatched text alert for: {qualified['title']} at {qualified['company']} ({qualified['tier']}) [{qualified['salary_range']}]")

    print(
        f"\n[DEBUG] Run summary — Dispatched: {dispatched} | "
        f"Tier 1 matches: {len(tier_1_matches)} | Tier 2 matches: {len(tier_2_matches)} | "
        f"Skipped high exp: {skip_counts['exp_too_high']} | "
        f"Skipped below salary floor: {skip_counts['salary_check']} | "
        f"Skipped non-target roles: {skip_counts['senior_or_irrelevant']} | "
        f"Skipped invalid location: {skip_counts['invalid_location']}"
    )


if __name__ == "__main__":
    run()
