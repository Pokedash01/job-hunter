import os
import json
import sqlite3
import requests
import pandas as pd
from datetime import datetime
from jobspy import scrape_jobs
from google import genai
from google.genai import types

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY")

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

# ----------------- STORAGE & EXCEL ENGINE -----------------
EXCEL_FILE = "job_applications.xlsx"

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
            "Status", "Tailored Pitch", "Cover Letter"
        ])
        df.to_excel(EXCEL_FILE, index=False)

def is_seen(url):
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("SELECT url FROM seen_jobs WHERE url = ?", (url,))
    seen = c.fetchone() is not None
    conn.close()
    return seen

def log_job(url, title, company, location, ctc, score, pitch, cover_letter):
    # Log to SQLite
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO seen_jobs (url, title, company) VALUES (?, ?, ?)", (url, title, company))
    conn.commit()
    conn.close()

    # Log to Excel
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
            "Status": "Discovered / Ready to Apply",
            "Tailored Pitch": pitch,
            "Cover Letter": cover_letter
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
    except Exception as e:
        print(f"Excel logging error: {e}")

# ----------------- MATCHING & SALARY RULES -----------------
def passes_salary_and_location(location_str, min_sal, max_sal):
    loc = str(location_str).lower()
    
    # Delhi NCR / Remote: Min 10 LPA (1,000,000 INR)
    if any(k in loc for k in ["delhi", "ncr", "gurgaon", "gurugram", "noida", "remote"]):
        min_required = 1000000
    # Tier-1 (Bangalore, Pune, Mumbai, Hyderabad, etc.): Min 15 LPA (1,500,000 INR)
    else:
        min_required = 1500000

    if max_sal and max_sal > 0:
        return max_sal >= min_required
    return True

# ----------------- ENTERPRISE ATS SEARCH (Workday, Ashby, Lever, Greenhouse) -----------------
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
            print(f"ATS search error: {e}")
            
    return discovered

# ----------------- AI TAILORING ENGINE (Gemini) -----------------
def generate_application_kit(title, company, description):
    prompt = f"""
    You are an executive career advisor tailoring application materials for Kartik Bhatt.

    Candidate Profile:
    {KARTIK_PROFILE}

    Target Job:
    Title: {title}
    Company: {company}
    JD Snippet: {description[:1500]}

    Generate tailored application assets strictly matching his actual accomplishments (KPMG Power Platform/Copilot work, GlobalLogic GenAI QA, 9.3 BCA GPA).

    Return ONLY a valid JSON object with the following schema:
    {{
        "match_score": "e.g., 90%",
        "reason": "1-2 sentences on why Kartik's exact tech stack and KPMG/GlobalLogic experience fit this role.",
        "skills_gap": "Any missing skill or 'None'",
        "tailored_bullets": [
            "Tailored bullet point 1 quantifying impact for this JD",
            "Tailored bullet point 2 emphasizing Power Platform, Copilot, or Process Ops"
        ],
        "cover_letter": "A polished 3-paragraph (under 200 words) tailored cover letter ready to submit."
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
            "tailored_bullets": [
                "Led end-to-end automation and Copilot Agent workflows at KPMG, saving 2,000+ hours annually across 13 sectors.",
                "Engineered robust QA frameworks for Google GenAI datasets at GlobalLogic, elevating delivery quality to 95%."
            ],
            "cover_letter": "Dear Hiring Team,\n\nI am writing to express my strong interest in this role. With over 3.5 years of experience across KPMG and GlobalLogic driving enterprise automation, Power Platform ecosystems, and Copilot AI agents, I bring a proven track record of optimizing cross-functional operations.\n\nAt KPMG, I architected solutions that automated 20,000 annual interactions and saved over 2,000 hours annually, earning multiple KUDOS accolades. I look forward to bringing this operational and technical rigor to your team.\n\nBest regards,\nKartik Bhatt"
        }

# ----------------- TELEGRAM DISPATCHER -----------------
def send_telegram_alert(title, company, location, url, ctc_label, kit):
    bullets_text = "\n".join([f"• {b}" for b in kit.get("tailored_bullets", [])])
    
    msg = (
        f"🎯 *New Matched Role for Kartik!*\n\n"
        f"📌 *Role:* {title}\n"
        f"🏢 *Company:* {company}\n"
        f"📍 *Location:* {location}\n"
        f"💰 *CTC Rule Check:* {ctc_label}\n"
        f"📊 *Fit Score:* {kit.get('match_score')}\n"
        f"💡 *Why You Match:* {kit.get('reason')}\n"
        f"⚠️ *Skill Gap:* {kit.get('skills_gap')}\n\n"
        f"⚡ *Tailored Resume Bullets:*\n{bullets_text}\n\n"
        f"📝 *Tailored Cover Letter:*\n```\n{kit.get('cover_letter')}\n```\n\n"
        f"🔗 [Apply Directly Here]({url})"
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
            results_wanted=20,
            hours_old=24,
            country_indeed='india'
        )
        for _, row in board_jobs.iterrows():
            all_jobs.append(row.to_dict())
    except Exception as e:
        print(f"Aggregator scraper warning: {e}")

    # 2. Search direct enterprise ATS portals
    ats_jobs = search_enterprise_ats_jobs()
    all_jobs.extend(ats_jobs)

    # 3. Process postings
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
        
        # Generate custom resume bullets + cover letter
        app_kit = generate_application_kit(title, company, desc)

        # Send Telegram alert
        send_telegram_alert(title, company, location, url, ctc_label, app_kit)

        # Log to Excel & SQLite
        log_job(
            url=url,
            title=title,
            company=company,
            location=location,
            ctc=ctc_label,
            score=app_kit.get("match_score"),
            pitch=app_kit.get("reason"),
            cover_letter=app_kit.get("cover_letter")
        )
        print(f"Processed and logged: {title} at {company}")

if __name__ == "__main__":
    run()
