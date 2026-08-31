# 🤖 Autonomous Job Hunter & Application Suite

An automated, intelligent job discovery, qualification, and application generation pipeline designed for **Kartik Bhatt** (~3.5+ YOE).

The bot continuously scrapes listings across major job platforms and enterprise ATS portals (Workday, Greenhouse, Lever, Ashby), runs hard heuristic/experience filters, generates tailored ATS-optimized resumes and cover letters via Google Gemini, archives them to GitHub, and dispatches instant action alerts via Telegram.

---

## 🚀 Key Features

* **Smart Qualification Engine**:
* **Experience Gate**: Hard filter rejecting roles demanding $\ge$ 6 years of experience (focused strictly on the 2–5.5 YOE window).
* **Seniority Blacklist**: Instantly eliminates `Director`, `VP`, `Head of`, `Engineering Manager`, and `Lead` roles.
* **Domain Targeting**: Scans for Data Analytics, Business Analytics, Copilot Studio / Power Platform Development, Product Analysis, and APM openings.


* **Geographic Tiering**:
* **Tier 1 (Priority)**: Delhi-NCR (Gurgaon, Noida, Delhi, Faridabad) and Remote/WFH.
* **Tier 2 (Fallback)**: Bengaluru, Hyderabad, and Pune.


* **AI-Powered Tailoring**:
* Leverages **Gemini** to extract 10–14 exact ATS keywords matching the candidate's profile.
* Dynamically weaves matching keywords into executive summaries and quantified project bullets.
* Generates clean ATS-compliant PDFs using ReportLab with strict company/objective separation.


* **Automated Telegram Alerts & Archival**:
* Sends compact job cards with direct application links and direct raw PDF download links.
* Auto-tracks seen jobs in `job_tracker.db` and logs full records into `job_applications.xlsx`.



---

## 🛠 Tech Stack

* **Language**: Python 3.11+
* **Scraping**: `python-jobspy`, `requests`, `SearchAPI` (Google Search Engine API)
* **LLM Engine**: Google GenAI SDK (`gemini-2.5-flash` / `gemini-3.6-flash`)
* **Document Generation**: ReportLab
* **Storage & Persistence**: SQLite3, Pandas, OpenPyXL
* **CI/CD**: GitHub Actions

---

## 📁 Repository Structure

```text
job-hunter/
├── .github/
│   └── workflows/
│       └── run_engine.yml    # GitHub Actions cron scheduler & auto-commit
├── generated_docs/           # Auto-generated resumes and cover letters (PDFs)
├── main.py                   # Master orchestration, filtering, and dispatch pipeline
├── job_tracker.db            # SQLite database tracking seen job URLs
├── job_applications.xlsx     # Application logs with match scores and metadata
├── requirements.txt          # Python dependencies
└── README.md

```

---

## ⚙️ Environment Variables & Secrets

Add the following secrets to your GitHub repository (**Settings > Secrets and variables > Actions**):

| Secret Name | Description |
| --- | --- |
| `GEMINI_API_KEY` | Google AI Studio API Key for application tailoring |
| `TELEGRAM_BOT_TOKEN` | Bot token from `@BotFather` |
| `TELEGRAM_CHAT_ID` | Target Telegram Chat / Channel ID |
| `SEARCH_API_KEY` | *(Optional)* SearchAPI.io key for enterprise ATS querying |

---

## 📦 Local Installation & Setup

1. **Clone the repository:**
```bash
git clone https://github.com/Pokedash01/job-hunter.git
cd job-hunter

```


2. **Set up a virtual environment:**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

```


3. **Install dependencies:**
```bash
pip install -r requirements.txt

```


*(Or install directly: `pip install python-jobspy google-genai requests pandas openpyxl reportlab`)*
4. **Export environment variables:**
```bash
export GEMINI_API_KEY="your-gemini-key"
export TELEGRAM_BOT_TOKEN="your-telegram-token"
export TELEGRAM_CHAT_ID="your-chat-id"
export SEARCH_API_KEY="your-searchapi-key"

```


5. **Run the pipeline:**
```bash
python main.py

```



---

## 🔄 GitHub Actions Automation

The engine is configured to run automatically on schedule via `.github/workflows/run_engine.yml`. It restores tracked cache databases, executes `main.py`, commits freshly generated PDFs to `generated_docs/`, and pushes them back to `main` with `[skip ci]`.
