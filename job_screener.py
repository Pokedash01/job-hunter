# job_screener.py

import re
from typing import Dict, Any, Optional, Tuple
from config import TARGET_CONFIG

def parse_minimum_experience(text: str) -> Optional[int]:
    """Extracts explicit lower-bound years of experience mentioned in text."""
    patterns = [
        r"(\d+)\s*(?:-|to|\+)\s*(?:\d+)?\s*(?:years|yrs)",
        r"(?:minimum|at least|over|requires?)\s*(\d+)\s*(?:years|yrs)"
    ]
    years_found = []
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            val = int(m.group(1))
            years_found.append(val)
            
    return min(years_found) if years_found else None

def classify_location(location_str: str) -> Tuple[bool, str]:
    """Determines location eligibility and tier rank."""
    loc_lower = location_str.lower()
    
    if any(re.search(pat, loc_lower) for pat in TARGET_CONFIG["LOCATIONS_TIER_1"]):
        return True, "Tier 1 (Delhi-NCR / Remote)"
    
    if any(re.search(pat, loc_lower) for pat in TARGET_CONFIG["LOCATIONS_TIER_2"]):
        return True, "Tier 2 (Bangalore / Hyderabad / Pune)"
        
    return False, "Out of Scope Location"

def extract_matched_skills(full_text: str) -> Dict[str, list]:
    """Identifies skills present in the posting across domains."""
    matched = {}
    for category, patterns in TARGET_CONFIG["SKILL_TAXONOMY"].items():
        cat_matches = []
        for pat in patterns:
            if re.search(pat, full_text, flags=re.IGNORECASE):
                clean_name = pat.replace(r"\b", "")
                cat_matches.append(clean_name)
        if cat_matches:
            matched[category] = cat_matches
    return matched

def screen_job(job: Dict[str, str]) -> Dict[str, Any]:
    """
    Evaluates a raw job dictionary:
    job = {'title': str, 'location': str, 'description': str, 'link': str, 'company': str}
    """
    title = job.get("title", "")
    description = job.get("description", "")
    location = job.get("location", "")
    full_text = f"{title} {description}"

    # 1. Seniority Blacklist Gate
    for blacklist_pat in TARGET_CONFIG["SENIORITY_BLACKLIST"]:
        if re.search(blacklist_pat, title, flags=re.IGNORECASE):
            return {"passed": False, "reason": f"Excluded senior title pattern: '{blacklist_pat}'"}

    # 2. Target Role Whitelist Gate
    if not any(re.search(whitelist_pat, title, flags=re.IGNORECASE) for whitelist_pat in TARGET_CONFIG["ROLES_WHITELIST"]):
        return {"passed": False, "reason": "Title does not match target roles"}

    # 3. Location Evaluation
    loc_valid, loc_tier = classify_location(location)
    if not loc_valid:
        return {"passed": False, "reason": f"Location '{location}' not in Tier 1 or Tier 2"}

    # 4. Experience Limit Gate
    min_exp = parse_minimum_experience(full_text)
    if min_exp is not None and min_exp > TARGET_CONFIG["EXP_MAX_CAP"]:
        return {"passed": False, "reason": f"Requires {min_exp}+ yrs (exceeds cap of {TARGET_CONFIG['EXP_MAX_CAP']} yrs)"}

    # 5. Extract Matching Skills
    matched_skills = extract_matched_skills(full_text)

    return {
        "passed": True,
        "title": title,
        "company": job.get("company", "Unknown"),
        "location_tier": loc_tier,
        "location_raw": location,
        "experience_detected": f"{min_exp}+ years" if min_exp else "Unspecified (Passes limit check)",
        "skills": matched_skills,
        "url": job.get("link", "")
    }

def format_alert_message(evaluated: Dict[str, Any]) -> str:
    """Formats verified jobs into an actionable alert card."""
    skill_lines = []
    for cat, skills in evaluated["skills"].items():
        skill_lines.append(f"• *{cat}*: {', '.join(skills)}")
    skills_text = "\n".join(skill_lines) if skill_lines else "• General domain alignment"

    return (
        f"🎯 *{evaluated['title']}* — *{evaluated['company']}*\n"
        f"📍 *Location*: {evaluated['location_raw']} ({evaluated['location_tier']})\n"
        f"⏳ *Experience*: {evaluated['experience_detected']}\n"
        f"🛠 *Skills Found*:\n{skills_text}\n"
        f"🔗 [Apply Here]({evaluated['url']})"
    )
