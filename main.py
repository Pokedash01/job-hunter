import os
import json
import sqlite3
import requests
from jobspy import scrape_jobs
from google import genai
from google.genai import types

# ----------------- CREDENTIALS -----------------
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
SEARCH_KEY = os.environ.get("SEARCH_API_KEY")

client = genai.Client(api_key=GEMINI_KEY)

# ----------------- CANDIDATE PROFILE -----------------
MY_PROFILE = """
Target Roles: Product Manager, Operations Manager, Technical PM, Operations Lead
Experience: 3-5 years in cross-functional execution, product/operations roadmaps, SQL, process automation, agile delivery.
Core Skills: Product Strategy, Operations Optimization, Data Analytics, Cross-functional Leadership.
"""

def init_db():
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

def is_seen(url):
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("SELECT url FROM seen_jobs WHERE url = ?", (url,))
    seen = c.fetchone() is not None
    conn.close()
    return seen

def mark_seen(url, title, company):
    conn = sqlite3.connect("job_tracker.db")
    c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO seen_jobs (url, title, company) VALUES (?, ?, ?)", (url, title, company))
    conn.commit()
    conn.close()

def passes_salary_and_location(location_str, min_sal, max_sal):
    loc = str(location_str).lower()
    
    # Rule 1: Delhi NCR or Remote -> Min 10 LPA (1,000,000 INR)
    if any(k in loc for k in ["delhi", "ncr", "gurgaon", "gurugram", "noida", "remote"]):
        min_required = 1000000
    # Rule 2: Other Tier-1 Cities -> Min 15 LPA (1,500,000 INR)
    else:
        min_required = 1500000

    if max_sal and max_sal > 0:
        return max_sal >= min_required
    return True

def search_enterprise_ats_jobs():
    """Finds roles hosted on direct Workday, Greenhouse, Lever, and Ashby portals in India."""
    ats_queries = [
        'site:myworkdayjobs.com ("Product Manager" OR "Operations Manager") India',
        'site:greenhouse.io ("Product Manager" OR "Operations Manager") India',
        'site:jobs.lever.co ("Product Manager" OR "Operations Manager") India',
        'site:jobs.ashbyhq.com ("Product Manager" OR "Operations Manager") India'
    ]
    discovered = []
    if not SEARCH_KEY:
        return discovered

    for query in ats_queries:
        try:
            # SearchApi.io endpoint integration
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
                        "location": "India",
                        "description": item.get("snippet", ""),
                        "min_amount": 0,
                        "max_amount": 0
                    })
        except Exception as e:
            print(f"ATS search error for query {query}: {e}")
            
    return discovered

def generate_pitch(title, company, description):
    prompt = f"""
    Compare this job against my profile:
    Profile: {MY_PROFILE}
    Role: {title} at {company}
    JD Snippet: {description[:1200]}

    Return ONLY a valid JSON object:
    {{
        "match_score": "85%",
        "reason": "One sentence explaining key skill alignment.",
        "skills_gap": "Any missing skill or 'None'",
        "tailored_pitch": "A 2-sentence tailored cover note highlighting relevant achievements."
    }}
    """
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
            config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        return json.loads(response.text)
    except Exception:
        return {
            "match_score": "Match Found",
            "reason": "Matches your target title and domain criteria.",
            "skills_gap": "N/A",
            "tailored_pitch": "Experienced professional with background directly matching requirements."
        }

def send_alert(title, company, location, url, pitch):
    msg = (
        f"🎯 *New Job Opening Found!*\n\n"
        f"📌 *Role:* {title}\n"
        f"🏢 *Company:* {company}\n"
        f"📍 *Location:* {location}\n"
        f"📊 *Fit:* {pitch.get('match_score')}\n"
        f"💡 *Rationale:* {pitch.get('reason')}\n"
        f"⚠️ *Skill Gap:* {pitch.get('skills_gap')}\n\n"
        f"📝 *Custom Pitch:*\n_{pitch.get('tailored_pitch')}_\n\n"
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

def run():
    init_db()
    all_jobs = []

    # 1. Scrape standard job portals (LinkedIn, Indeed, Glassdoor)
    try:
        board_jobs = scrape_jobs(
            site_name=["linkedin", "indeed", "glassdoor"],
            search_term="Product Manager OR Operations Manager",
            location="India",
            results_wanted=15,
            hours_old=24,
            country_indeed='india'
        )
        for _, row in board_jobs.iterrows():
            all_jobs.append(row.to_dict())
    except Exception as e:
        print(f"Board scraper issue: {e}")

    # 2. Search enterprise career portals (Workday, Lever, Greenhouse, Ashby)
    ats_jobs = search_enterprise_ats_jobs()
    all_jobs.extend(ats_jobs)

    # 3. Filter and alert
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

        pitch_data = generate_pitch(title, company, desc)
        send_alert(title, company, location, url, pitch_data)
        mark_seen(url, title, company)
        print(f"Alert dispatched: {title} at {company}")

if __name__ == "__main__":
    run()
