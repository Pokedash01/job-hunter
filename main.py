import os
import re
import json
import sqlite3
import requests
import pandas as pd
from datetime import datetime
from jobspy import scrape_jobs
from google import genai
from google.genai import types

# ReportLab for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY")
GITHUB_REPOSITORY = os.environ.get("GITHUB_REPOSITORY", "")  # Auto-injected by GitHub Actions (e.g. user/repo)

client = genai.Client(api_key=GEMINI_KEY)

# ----------------- KARTIK BHATT MASTER PROFILE -----------------
KARTIK_PROFILE = """
Candidate: Kartik Bhatt
Contact: kb270102@gmail.com | +91-7428062532 | Portfolio: https://kartikb.vercel.app/
Education: Bachelor of Computer Applications (Computer Science), Maharaja Surajmal Institute (GPA: 9.3/10, Top 1%)
Total Experience: ~3.5+ years across KPMG and GlobalLogic.

Core Skills & Tools:
- Microsoft Power Platform (Power Apps, Power Automate, Power BI)
- GenAI & Agentic AI (Copilot Studio, Multi-modal Copilot Agents, GenAI dataset training, Azure AI)
- SharePoint Online (Governance, Term Store, Permission Management, List Migrations)
- Process Optimization & Operations (Lean Six Sigma Yellow Belt, VBA Macros, SQL, QA Management)
- Stakeholder Management & Business Development (100+ RFP/RFI responses, 360-degree stakeholder management across 13 sectors)

Key Career Highlights:
1. KPMG (Analyst - Knowledge Management):
   - Built Power Platform solutions facilitating 20,000 reach outs annually across 13 sectors, saving 1,200 hrs/yr.
   - Built multi-modal Copilot agents to assist with messy data and metadata tagging, saving 325 hrs/yr.
   - Migrated legacy Excel processes for 45+ pillars to automated SharePoint Online lists with permission governance.
   - Overall saved 2,000+ hours annually using Power Platform, Copilot Studio, and VBA macros. Awarded KUDOS & Gurus@Work.
2. GlobalLogic (Associate Analyst - Content Engineering):
   - Created QA frameworks for Google GenAI training datasets used for Android on-screen search.
   - Improved onshore project quality from 74% to 95%, managing process documentation across 10+ projects.

Target Roles:
- Technical Product Manager / Product Operations Lead (Internal Tools / AI / Automation)
- GenAI Solutions Specialist / AI Transformation Lead
- Operations Optimization Lead / Business Process Automation Consultant
"""

EXCEL_FILE = "job_applications.xlsx"
DOCS_DIR = "generated_docs"

os.makedirs(DOCS_DIR, exist_ok=True)

# ----------------- STORAGE & EXCEL ENGINE -----------------
def init_tracker():
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS seen_jobs (
            url TEXT PRIMARY KEY,
            title TEXT,
            company TEXT,
            found_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=[
            "Date Found", "Role Title", "Company", "Location", 
            "Estimated CTC", "Match Score", "Application Link", 
            "Status", "Resume PDF Path", "Cover Letter PDF Path", "GitHub Folder Link"
        ])
        df.to_excel(EXCEL_FILE, index=False)

def is_seen(url):
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("SELECT url FROM seen_jobs WHERE url = ?", (url,))
    seen = c.fetchone() is not None
    conn.close()
    return seen

def log_job(url, title, company, location, ctc, score, resume_path, cl_path, github_link):
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO seen_jobs (url, title, company) VALUES (?, ?, ?)", (url, title, company))
    conn.commit()
    conn.close()

    try:
        df = pd.read_excel(EXCEL_FILE)
        new_row = {
            "Date Found": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Role Title": title,
            "Company": company,
            "Location": location,
            "Estimated CTC": ctc,
            "Match Score": score,
            "Application Link": url,
            "Status": "PDFs Generated / Ready to Apply",
            "Resume PDF Path": resume_path,
            "Cover Letter PDF Path": cl_path,
            "GitHub Folder Link": github_link
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        print(f"Excel logging error: {e}")

# ----------------- MATCHING RULES -----------------
def passes_salary_and_location(location_str, min_sal, max_sal):
    loc = str(location_str).lower()
    if any(k in loc for k in ["delhi", "ncr", "gurgaon", "gurugram", "noida", "remote"]):
        min_required = 1000000
    else:
        min_required = 1500000

    if max_sal and max_sal > 0:
        return max_sal >= min_required
    return True

# ----------------- ENTERPRISE ATS DISCOVERY -----------------
def search_enterprise_ats_jobs():
    ats_queries = [
        'site:myworkdayjobs.com ("Product Manager" OR "Operations Lead" OR "Power Platform" OR "AI") India',
        'site:greenhouse.io ("Product Operations" OR "Technical PM" OR "Solutions Consultant") India',
        'site:jobs.lever.co ("Product Manager" OR "Operations" OR "Automation Lead") India',
        'site:jobs.ashbyhq.com ("Product Manager" OR "Operations Lead" OR "AI") India'
    ]
    discovered = []
    if not SEARCH_KEY:
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
                "num": 5
            }
            res = requests.get(url, params=params, timeout=15).json()
            for item in res.get("organic_results", []):
                link = item.get("link", "")
                if link:
                    discovered.append({
                        "title": item.get("title", "Role Opening"),
                        "company": item.get("displayed_link", "Enterprise Portal").split(".")[0],
                        "job_url": link,
                        "location": "Delhi NCR / Hybrid India",
                        "description": item.get("snippet", ""),
                        "min_amount": 0,
                        "max_amount": 0
                    })
        except Exception as e:
            print(f"ATS search error for query {query}: {e}")
            
    return discovered

# ----------------- GEMINI AI ASSET GENERATION -----------------
def generate_application_kit(title, company, description):
    prompt = f"""
    You are an executive career advisor tailoring high-impact job materials for Kartik Bhatt.

    Candidate Profile:
    {KARTIK_PROFILE}

    Target Job:
    Title: {title}
    Company: {company}
    JD Snippet: {description[:1500]}

    Generate tailored content for a professional PDF resume and cover letter.
    Return ONLY a JSON object:
    {{
        "match_score": "e.g., 92%",
        "reason": "1-2 sentences on why Kartik's exact tech stack and KPMG/GlobalLogic experience fit this role.",
        "skills_gap": "Any missing tool/skill or 'None'",
        "tailored_summary": "A 3-line tailored Professional Summary specifically aligned to this JD.",
        "kpmg_bullets": [
            "Tailored KPMG bullet 1 emphasizing Power Platform / Copilot Studio impact",
            "Tailored KPMG bullet 2 emphasizing automation, scale, and stakeholder leadership",
            "Tailored KPMG bullet 3 emphasizing process governance or migration"
        ],
        "globallogic_bullets": [
            "Tailored GlobalLogic bullet 1 emphasizing Google GenAI dataset QA",
            "Tailored GlobalLogic bullet 2 emphasizing process documentation and quality metrics"
        ],
        "skills_list": "Copilot Studio, GenAI Agents, Power Apps, Power Automate, Power BI, SharePoint Online, SQL, Lean Six Sigma, VBA",
        "cover_letter_paragraphs": [
            "Opening paragraph stating interest in the role at {company} and summarizing core value proposition.",
            "Body paragraph highlighting specific quantitative impact from KPMG and GlobalLogic aligned with this role.",
            "Closing paragraph requesting an interview and summarizing readiness to contribute."
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
            "match_score": "High Match",
            "reason": "Strong alignment with Power Platform, Copilot GenAI, and operations background.",
            "skills_gap": "None",
            "tailored_summary": "Results-driven Analyst with 3.5+ years of experience across KPMG and GlobalLogic specializing in Power Platform automation, Copilot AI agents, and cross-functional operations.",
            "kpmg_bullets": [
                "Built Power Platform solutions facilitating 20,000 reach outs annually across 13 sectors, saving 1,200 hrs/yr.",
                "Architected multi-modal Copilot agents for metadata tagging and unstructured data processing, saving 325 hrs/yr.",
                "Migrated legacy Excel systems for 45+ pillars to automated SharePoint Online lists with strict permission governance."
            ],
            "globallogic_bullets": [
                "Engineered QA frameworks for Google GenAI datasets for Android search, boosting delivery quality from 74% to 95%.",
                "Managed process documentation across 10+ high-visibility client engagements."
            ],
            "skills_list": "Copilot Studio, GenAI Agents, Power Apps, Power Automate, Power BI, SharePoint Online, SQL, Lean Six Sigma, VBA",
            "cover_letter_paragraphs": [
                f"I am writing to express my strong interest in the {title} position at {company}.",
                "With over 3.5 years of experience across KPMG and GlobalLogic driving enterprise automation, Power Platform ecosystems, and Copilot AI agents, I bring a proven track record of optimizing cross-functional operations.",
                "I look forward to discussing how my technical background and execution rigor can add immediate value to your team."
            ]
        }

# ----------------- PDF BUILDER ENGINE -----------------
def create_pdf_resume(filepath, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = []

    name_style = ParagraphStyle('Name', parent=styles['Heading1'], fontSize=18, leading=20, textColor=colors.HexColor("#1A365D"), alignment=1)
    contact_style = ParagraphStyle('Contact', parent=styles['Normal'], fontSize=9, leading=12, alignment=1, textColor=colors.HexColor("#4A5568"))
    section_style = ParagraphStyle('SecHeader', parent=styles['Heading2'], fontSize=11, leading=13, textColor=colors.HexColor("#1A365D"), spaceBefore=6, spaceAfter=2)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#2D3748"))
    job_header_style = ParagraphStyle('JobHeader', parent=styles['Normal'], fontSize=9.5, leading=12, fontName='Helvetica-Bold', textColor=colors.HexColor("#1A202C"))
    bullet_style = ParagraphStyle('Bullet', parent=styles['Normal'], fontSize=8.5, leading=11.5, leftIndent=12, textColor=colors.HexColor("#2D3748"))

    # Header
    story.append(Paragraph("<b>KARTIK BHATT</b>", name_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | Portfolio: https://kartikb.vercel.app/ | Delhi NCR, India", contact_style))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CBD5E0"), spaceAfter=6))

    # Professional Summary
    story.append(Paragraph("PROFESSIONAL SUMMARY", section_style))
    story.append(Paragraph(kit.get("tailored_summary", ""), body_style))
    story.append(Spacer(1, 4))

    # Education
    story.append(Paragraph("EDUCATION", section_style))
    story.append(Paragraph("<b>Maharaja Surajmal Institute</b> | Bachelor of Computer Applications (Computer Science) — <b>GPA: 9.3/10 (Top 1%)</b>", body_style))
    story.append(Spacer(1, 4))

    # Experience
    story.append(Paragraph("WORK EXPERIENCE", section_style))
    story.append(Paragraph("<b>KPMG</b> | Analyst — Knowledge Management | Gurugram <i>(May 2024 – Present)</i>", job_header_style))
    for b in kit.get("kpmg_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Spacer(1, 4))

    story.append(Paragraph("<b>GlobalLogic Technologies</b> | Associate Analyst — Content Engineering | Gurugram <i>(Sep 2022 – Oct 2023)</i>", job_header_style))
    for b in kit.get("globallogic_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Spacer(1, 4))

    # Skills & Certifications
    story.append(Paragraph("SKILLS & CERTIFICATIONS", section_style))
    story.append(Paragraph(f"<b>Core Stack:</b> {kit.get('skills_list', '')}", body_style))
    story.append(Paragraph("<b>Certifications:</b> Microsoft Certified: Azure AI Fundamentals (AI-901) | Lean Six Sigma: Yellow Belt | Oracle: Agentic AI Foundations | AI Transformation Leader (AB-731)", body_style))

    doc.build(story)

def create_pdf_cover_letter(filepath, title, company, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=45, rightMargin=45, topMargin=45, bottomMargin=45)
    styles = getSampleStyleSheet()
    story = []

    name_style = ParagraphStyle('Name', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor("#1A365D"))
    contact_style = ParagraphStyle('Contact', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#4A5568"))
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=10, leading=14, spaceBefore=8, textColor=colors.HexColor("#2D3748"))

    # Header
    story.append(Paragraph("<b>KARTIK BHATT</b>", name_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | Portfolio: https://kartikb.vercel.app/", contact_style))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CBD5E0"), spaceAfter=12))

    story.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%B %d, %Y')}", body_style))
    story.append(Paragraph(f"<b>Application For:</b> {title}", body_style))
    story.append(Paragraph(f"<b>Company:</b> {company}", body_style))
    story.append(Spacer(1, 8))

    story.append(Paragraph("Dear Hiring Team,", body_style))

    for p in kit.get("cover_letter_paragraphs", []):
        story.append(Paragraph(p, body_style))

    story.append(Spacer(1, 10))
    story.append(Paragraph("Sincerely,<br/><b>Kartik Bhatt</b>", body_style))

    doc.build(story)

# ----------------- TELEGRAM DISPATCHER -----------------
def send_telegram_alert(title, company, location, url, ctc_label, kit, raw_resume_url, raw_cl_url):
    msg = (
        f"🎯 *New Matched Role for Kartik!*\n\n"
        f"📌 *Role:* {title}\n"
        f"🏢 *Company:* {company}\n"
        f"📍 *Location:* {location}\n"
        f"💰 *CTC Check:* {ctc_label}\n"
        f"📊 *Fit Score:* {kit.get('match_score')}\n"
        f"💡 *Why You Match:* {kit.get('reason')}\n"
        f"⚠️ *Skill Gap:* {kit.get('skills_gap')}\n\n"
        f"🔗 [Apply Directly Here]({url})\n\n"
        f"📂 *Generated Application PDFs (1-Click View & Download):*\n"
        f"📄 [Download Tailored Resume PDF]({raw_resume_url})\n"
        f"📝 [Download Tailored Cover Letter PDF]({raw_cl_url})"
    )

    endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": msg,
        "parse_mode": "Markdown",
        "disable_web_page_preview": False
    }
    requests.post(endpoint, json=payload)

# ----------------- MAIN PIPELINE -----------------
def run():
    init_tracker()
    all_jobs = []

    # 1. Scrape standard boards
    try:
        board_jobs = scrape_jobs(
            site_name=["linkedin", "indeed", "glassdoor"],
            search_term='"Product Manager" OR "Power Platform" OR "Operations Lead" OR "Copilot" OR "AI Specialist"',
            location="India",
            results_wanted=15,
            hours_old=24,
            country_indeed='india'
        )
        for _, row in board_jobs.iterrows():
            all_jobs.append(row.to_dict())
    except Exception as e:
        print(f"Scraper error: {e}")

    # 2. Search enterprise ATS
    ats_jobs = search_enterprise_ats_jobs()
    all_jobs.extend(ats_jobs)

    # 3. Process jobs
    for job in all_jobs:
        url = str(job.get('job_url') or '')
        title = str(job.get('title') or '')
        company = str(job.get('company') or '')
        location = str(job.get('location') or 'India')
        min_sal = job.get('min_amount')
        max_sal = job.get('max_amount')
        desc = str(job.get('description') or '')

        if not url or url == 'nan' or is_seen(url):
            continue

        if not passes_salary_and_location(location, min_sal, max_sal):
            continue

        ctc_label = f"₹{int(max_sal):,}" if (max_sal and max_sal > 0) else "Meets/Exceeds Location Threshold"

        # Generate Tailored Assets via Gemini
        kit = generate_application_kit(title, company, desc)

        # Generate Safe File Names
        safe_company = re.sub(r'[^a-zA-Z0-9]', '_', company)[:15]
        safe_title = re.sub(r'[^a-zA-Z0-9]', '_', title)[:15]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        resume_filename = f"Resume_{safe_company}_{safe_title}_{timestamp}.pdf"
        cl_filename = f"CoverLetter_{safe_company}_{safe_title}_{timestamp}.pdf"

        resume_path = os.path.join(DOCS_DIR, resume_filename)
        cl_path = os.path.join(DOCS_DIR, cl_filename)

        # Build PDFs
        create_pdf_resume(resume_path, kit)
        create_pdf_cover_letter(cl_path, title, company, kit)

        # Build GitHub URLs for direct download
        repo = GITHUB_REPOSITORY if GITHUB_REPOSITORY else "owner/repo"
        raw_resume_url = f"https://raw.githubusercontent.com/{repo}/main/{DOCS_DIR}/{resume_filename}"
        raw_cl_url = f"https://raw.githubusercontent.com/{repo}/main/{DOCS_DIR}/{cl_filename}"
        github_folder_link = f"https://github.com/{repo}/tree/main/{DOCS_DIR}"

        # Send alert to Telegram
        send_telegram_alert(title, company, location, url, ctc_label, kit, raw_resume_url, raw_cl_url)

        # Log into SQLite & Excel
        log_job(
            url=url,
            title=title,
            company=company,
            location=location,
            ctc=ctc_label,
            score=kit.get("match_score"),
            resume_path=resume_path,
            cl_path=cl_path,
            github_link=github_folder_link
        )
        print(f"Generated PDFs & dispatched alert for: {title} at {company}")

if __name__ == "__main__":
    run()
