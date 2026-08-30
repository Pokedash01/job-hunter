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

# ReportLab imports for dense, professional PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY")
GITHUB_REPOSITORY = os.environ.get("GITHUB_REPOSITORY", "")

client = genai.Client(api_key=GEMINI_KEY)

# ----------------- KARTIK BHATT MASTER PROFILE -----------------
KARTIK_PROFILE = """
Candidate: Kartik Bhatt
Contact: kb270102@gmail.com | +91-7428062532 | Portfolio: https://kartikb.vercel.app/
Education: Bachelor of Computer Applications, Maharaja Surajmal Institute (Computer Science) | GPA: 9.3/10 (Top 1%)

Experience Summary: ~3.5+ years of cross-functional experience across KPMG and GlobalLogic driving process automation, Copilot/GenAI agents, Power Platform solutions, and operational excellence.

Work Experience:
1. KPMG | Analyst | Knowledge Management | Gurugram (May 2024 - Present)
   - Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management.
   - Built complete Power Platform solution (Power Automate, SharePoint lists, Power Apps, Power BI) facilitating 20,000 reach outs annually across 30 member firms, saving 1,200 hrs annually.
   - Automated migration of old Excel-based data collection for 45+ pillars to automated SPO lists with permission governance & alerts.
   - Built & managed VBA macro solutions refreshing 30,000+ assets globally, saving 485 hrs annually.
   - Architected multi-modal Copilot agent for messy data, field drafting, and metadata tagging, saving 325 hrs annually.
   - Overall saved 2,000+ hrs annually across Power Platform, Copilot Studio, and VBA. Handled 100+ RFP/RFI requests.
   - Awards: KUDOS (efficiency & Lean Six Sigma), Super Team Award, Ally of Inclusion, Gurus@Work.

2. GlobalLogic Technologies | Associate Analyst | Content Engineering | Gurugram (Sep 2022 - Oct 2023)
   - Created best practices, process docs, and QA processes for Google project to build test & training datasets for GenAI screen search on Android.
   - Piloted project to extract answers from multi-level docs for AI training datasets.
   - Improved onshore project delivery quality from 74% to 95%; QA'd 500+ pieces weekly and led 3 pilot projects.

Skills: Copilot Studio, GenAI Agents, Power Apps, Power Automate, Power BI, SharePoint Online, MS Excel/VBA, SQL.
Certifications: Microsoft Certified: Azure AI Fundamentals (AI-901), Lean Six Sigma: Yellow Belt, Oracle: Agentic AI Certified Foundations Associate, AI Transformation Leader (AB-731), AI Business Professional (AB-730).
"""

# Strict Domain Rules: Only Product, Operations, Technology
EXCLUDED_KEYWORDS = [
    "hr", "human resources", "talent acquisition", "recruiter", "recruitment", 
    "sales", "business development executive", "bde", "marketing", "digital marketing", 
    "telecaller", "content writer", "seo", "graphic designer", "accountant"
]

ALLOWED_DOMAINS = [
    "product", "operations", "tech", "technology", "program", "project", 
    "automation", "analyst", "process", "copilot", "power platform", "ai", "consultant"
]

EXCEL_FILE = "job_applications.xlsx"
DOCS_DIR = "generated_docs"
os.makedirs(DOCS_DIR, exist_ok=True)

# ----------------- STORAGE ENGINE -----------------
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
            "Status": "PDFs Dispatched / Ready to Apply",
            "Resume PDF Path": resume_path,
            "Cover Letter PDF Path": cl_path,
            "GitHub Folder Link": github_link
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        print(f"Excel logging error: {e}")

# ----------------- STRICT DOMAIN & SALARY FILTER -----------------
def is_role_relevant(title: str) -> bool:
    t = title.lower()
    for bad in EXCLUDED_KEYWORDS:
        if re.search(r'\b' + re.escape(bad) + r'\b', t):
            return False
    return any(good in t for good in ALLOWED_DOMAINS)

def passes_salary_and_location(location_str, min_sal, max_sal):
    loc = str(location_str).lower()
    if any(k in loc for k in ["delhi", "ncr", "gurgaon", "gurugram", "noida", "remote"]):
        min_required = 1000000
    else:
        min_required = 1500000

    if max_sal and max_sal > 0:
        return max_sal >= min_required
    return True

# ----------------- ENTERPRISE ATS SEARCH -----------------
def search_enterprise_ats_jobs():
    ats_queries = [
        'site:myworkdayjobs.com ("Product Manager" OR "Operations Lead" OR "Process Automation" OR "Power Platform") India',
        'site:greenhouse.io ("Product Operations" OR "Technical PM" OR "Solutions Consultant" OR "Business Analyst") India',
        'site:jobs.lever.co ("Product Manager" OR "Operations Specialist" OR "AI Specialist") India',
        'site:jobs.ashbyhq.com ("Product Manager" OR "Operations Lead" OR "AI Lead") India'
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
                title = item.get("title", "Role Opening")
                if link and is_role_relevant(title):
                    discovered.append({
                        "title": title,
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

# ----------------- GEMINI TAILORED ASSET GENERATOR -----------------
def generate_application_kit(title, company, description):
    prompt = f"""
    You are an expert executive resume writer and career strategist tailoring application documents for Kartik Bhatt.

    Candidate Profile:
    {KARTIK_PROFILE}

    Target Job Opening:
    Title: {title}
    Company: {company}
    JD Snippet: {description[:1600]}

    Generate tailored, rich, and dense content for both a 1-page full Resume and a comprehensive Cover Letter.

    Return ONLY a valid JSON object matching this exact schema:
    {{
        "match_score": "e.g., 94%",
        "reason": "1-2 sentences on why Kartik's exact tech stack and KPMG/GlobalLogic experience fit this role.",
        "skills_gap": "Any missing tool/skill or 'None'",
        "tailored_summary": "A comprehensive 3-4 line Professional Summary showcasing 3.5+ years of experience across KPMG and GlobalLogic, tailored directly to {title}.",
        "kpmg_project_bullets": [
            "Tailored bullet 1 detailing Power Platform solution for 20,000 reach outs and 1,200 hrs saved, framed for {title}",
            "Tailored bullet 2 detailing end-to-end SPO migration for 45+ pillars and governance",
            "Tailored bullet 3 detailing global VBA macro repository automation for 30,000+ assets saving 485 hrs",
            "Tailored bullet 4 detailing multi-modal Copilot Agent for metadata and messy data saving 325 hrs"
        ],
        "kpmg_bd_bullets": [
            "Saved 2,000+ hrs annually leveraging Power Platform, Copilot Studio, and VBA Macros.",
            "Catered to 100+ RFP/RFI requests and built 100+ internal site pages aligned with enterprise brand guidelines.",
            "Administered contact management system for 10,000+ members and uploaded 5,000+ content assets."
        ],
        "globallogic_bullets": [
            "Designed best practices, process docs, and QA workflows for Google GenAI datasets for Android screen search.",
            "Managed process documentation across 10+ engagements, improving project delivery quality from 74% to 95%."
        ],
        "cover_letter_subject": "Subject: Driving Operational Excellence & Scalable Solutions as {title}",
        "cover_letter_paragraphs": [
            "I am writing to express my strong enthusiasm for the {title} position at {company}. With over 3.5 years of experience across KPMG and GlobalLogic, I have specialized in building scalable automation frameworks, leveraging Microsoft Copilot and Power Platform ecosystems, and driving data-backed operational efficiencies. My background aligns directly with {company}'s focus on innovation and execution.",
            "In my current role at KPMG, I lead cross-functional initiatives across 13 sectors, managing 360-degree stakeholder relationships and translating complex business requirements into high-impact digital solutions. I architected end-to-end Power Platform systems that streamlined 20,000 annual interactions and built multi-modal Copilot agents that saved over 2,000 hours annually. Additionally, during my tenure at GlobalLogic, I established quality assurance frameworks for Google's GenAI training datasets, elevating project benchmark delivery from 74% to 95%.",
            "I am excited about the opportunity to bring my hands-on technical skills in Copilot Studio, Power Platform, and SQL, coupled with Lean Six Sigma methodologies, to accelerate outcomes for {company}. Thank you for considering my application. I look forward to discussing how my experience can deliver measurable value to your team."
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
            "reason": "Strong alignment with Kartik's Power Platform, GenAI agent building, and operations leadership background.",
            "skills_gap": "None",
            "tailored_summary": f"Analytical and solutions-driven professional with 3.5+ years of experience across KPMG and GlobalLogic specializing in Power Platform automation, Copilot AI agents, and cross-functional operations tailored for the {title} role.",
            "kpmg_project_bullets": [
                "Built complete Power Platform solution facilitating 20,000 reach outs annually across 13 sectors, saving 1,200 hrs annually.",
                "Automated migration of data collection for 45+ pillars to automated SPO lists with change management and permissions.",
                "Developed VBA macro solutions for 30,000+ assets globally, saving 485 hrs annually.",
                "Built multi-modal Copilot agent to clean messy data and apply metadata tags, saving 325 hrs annually."
            ],
            "kpmg_bd_bullets": [
                "Saved 2,000+ hrs annually across Power Platform, Copilot Studio, and VBA Macros.",
                "Catered to 100+ RFP and RFI requests and 100+ internal site pages according to brand standards.",
                "Administered contact management system for 10,000+ members and managed 50+ SharePoint governance requests."
            ],
            "globallogic_bullets": [
                "Created best practices and QA processes for Google GenAI training datasets for Android screen search.",
                "Improved project delivery quality from 74% to 95%, managing process documentation for 10+ projects."
            ],
            "cover_letter_subject": f"Subject: Application for {title} - Kartik Bhatt",
            "cover_letter_paragraphs": [
                f"I am excited to apply for the {title} position at {company}. With over 3.5 years of experience at KPMG and GlobalLogic, I specialize in process automation, Power Platform ecosystems, and Copilot AI agents.",
                "At KPMG, I architected automation systems that saved over 2,000 hours annually and led migrations across 45 pillars. At GlobalLogic, I engineered QA frameworks for Google's GenAI datasets, lifting project quality from 74% to 95%.",
                f"I look forward to bringing this operational and technical expertise to {company} to drive scalable outcomes."
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

    # Header
    story.append(Paragraph("<b>KARTIK BHATT</b>", name_style))
    story.append(Paragraph("kb270102@gmail.com | +91-7428062532 | LinkedIn | Website: https://kartikb.vercel.app/ | Delhi NCR, India", contact_style))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=0.75, color=colors.HexColor("#94A3B8"), spaceAfter=3))

    # Professional Summary
    story.append(Paragraph("<b>PROFESSIONAL SUMMARY</b>", section_style))
    story.append(Paragraph(kit.get("tailored_summary", ""), body_style))
    story.append(Spacer(1, 3))

    # Education
    story.append(Paragraph("<b>EDUCATION</b>", section_style))
    story.append(Paragraph("<b>Maharaja Surajmal Institute</b> | Bachelor of Computer Applications (Computer Science) | <b>GPA: 9.3/10 (Top 1%)</b>", body_style))
    story.append(Spacer(1, 3))

    # Experience
    story.append(Paragraph("<b>WORK EXPERIENCE</b>", section_style))
    story.append(Paragraph("<b>KPMG</b> | Analyst — Knowledge Management | Gurugram <i>(May 2024 – Present | 3 yrs 2 mos total exp)</i>", company_style))
    story.append(Paragraph("<i>Led cross-functional projects across 13 sectors demanding 360-degree stakeholder management & business development.</i>", body_style))
    
    story.append(Paragraph("Key Projects", subhead_style))
    for b in kit.get("kpmg_project_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    
    story.append(Paragraph("Business Development & Operations", subhead_style))
    for b in kit.get("kpmg_bd_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    
    story.append(Paragraph("Key Achievements: Awarded 'KUDOS' (Lean Six Sigma saving 2,000+ hrs), 'Super Team', 'Ally of Inclusion', and 'Gurus@Work'.", bullet_style))
    story.append(Spacer(1, 3))

    story.append(Paragraph("<b>GlobalLogic Technologies Private Limited</b> | Associate Analyst — Content Engineering | Gurugram <i>(Sep 2022 – Oct 2023)</i>", company_style))
    for b in kit.get("globallogic_bullets", []):
        story.append(Paragraph(f"• {b}", bullet_style))
    story.append(Spacer(1, 3))

    # Skills & Certifications
    story.append(Paragraph("<b>SKILLS & CERTIFICATIONS</b>", section_style))
    story.append(Paragraph("<b>Skills:</b> Copilot GenAI (Agents), Copilot Studio, Power Apps, Power Automate, Power BI, SharePoint Online, MS Excel/VBA, SQL", body_style))
    story.append(Paragraph("<b>Certifications:</b> Microsoft Certified: Azure AI Fundamentals (AI-901) | Lean Six Sigma: Yellow Belt | Oracle: Agentic AI Certified Foundations Associate | AI Transformation Leader (AB-731) | AI Business Professional (AB-730)", body_style))

    doc.build(story)

def create_dense_cover_letter(filepath, title, company, kit):
    doc = SimpleDocTemplate(filepath, pagesize=letter, leftMargin=45, rightMargin=45, topMargin=40, bottomMargin=40)
    styles = getSampleStyleSheet()
    story = []

    name_style = ParagraphStyle('Name', parent=styles['Heading1'], fontSize=16, leading=18, textColor=colors.HexColor("#0F2942"))
    contact_style = ParagraphStyle('Contact', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor("#4A5568"))
    subj_style = ParagraphStyle('Subj', parent=styles['Normal'], fontSize=10, leading=13, fontName='Helvetica-Bold', textColor=colors.HexColor("#0F2942"), spaceBefore=6, spaceAfter=6)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=9.5, leading=14, spaceBefore=6, textColor=colors.HexColor("#1E293B"))

    # Header
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

# ----------------- TELEGRAM DISPATCHER (DIRECT FILE ATTACHMENTS) -----------------
def send_telegram_alert(title, company, location, url, ctc_label, kit, resume_path, cl_path):
    msg = (
        f"🎯 *New Job Matched for Kartik!*\n\n"
        f"📌 *Role:* {title}\n"
        f"🏢 *Company:* {company}\n"
        f"📍 *Location:* {location}\n"
        f"💰 *CTC Check:* {ctc_label}\n"
        f"📊 *Fit Score:* {kit.get('match_score')}\n"
        f"💡 *Why You Match:* {kit.get('reason')}\n"
        f"⚠️ *Skill Gap:* {kit.get('skills_gap')}\n\n"
        f"🔗 [Apply Directly Here]({url})\n\n"
        f"📎 *Attached below:* Tailored 1-page Resume & Cover Letter PDFs ready to upload."
    )

    msg_endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    doc_endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"

    try:
        # 1. Send text overview
        requests.post(msg_endpoint, json={
            "chat_id": CHAT_ID,
            "text": msg,
            "parse_mode": "Markdown",
            "disable_web_page_preview": False
        })

        # 2. Directly send Tailored Resume PDF file
        if os.path.exists(resume_path):
            with open(resume_path, 'rb') as f:
                requests.post(doc_endpoint, data={
                    "chat_id": CHAT_ID,
                    "caption": f"📄 Tailored Resume — {company} ({title})"
                }, files={"document": f})

        # 3. Directly send Tailored Cover Letter PDF file
        if os.path.exists(cl_path):
            with open(cl_path, 'rb') as f:
                requests.post(doc_endpoint, data={
                    "chat_id": CHAT_ID,
                    "caption": f"📝 Tailored Cover Letter — {company}"
                }, files={"document": f})

    except Exception as e:
        print(f"Telegram dispatch error: {e}")

# ----------------- MAIN PIPELINE -----------------
def run():
    init_tracker()
    all_jobs = []

    # 1. Scrape standard boards with targeted queries
    try:
        board_jobs = scrape_jobs(
            site_name=["linkedin", "indeed", "glassdoor"],
            search_term='"Product Manager" OR "Product Operations" OR "Operations Manager" OR "Operations Lead" OR "Power Platform" OR "Copilot" OR "Solutions Specialist"',
            location="India",
            results_wanted=20,
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

        # Enforce strict domain relevance (No HR/Sales/Marketing)
        if not is_role_relevant(title):
            print(f"Skipping irrelevant domain: {title}")
            continue

        # Enforce CTC and Location rules
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

        # Build Full Dense PDFs
        create_dense_resume(resume_path, kit)
        create_dense_cover_letter(cl_path, title, company, kit)

        github_folder_link = f"https://github.com/{GITHUB_REPOSITORY}/tree/main/{DOCS_DIR}"

        # Send Telegram alert with PDF files attached
        send_telegram_alert(title, company, location, url, ctc_label, kit, resume_path, cl_path)

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
        print(f"Generated comprehensive PDFs & dispatched files for: {title} at {company}")

if __name__ == "__main__":
    run()
