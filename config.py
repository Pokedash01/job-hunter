# config.py

TARGET_CONFIG = {
    "EXP_MIN_CAP": 2,       # Minimum years accepted
    "EXP_MAX_CAP": 5.5,     # Hard limit ceiling (rejects 6+, 8+, 10+ YOE)
    
    "ROLES_WHITELIST": [
        r"\bdata analyst\b", r"\bbusiness analyst\b", r"\bproduct analyst\b",
        r"\bassociate product manager\b", r"\bapm\b", r"\bcopilot studio\b",
        r"\bpower automate\b", r"\bpower platform\b", r"\bbi developer\b",
        r"\banalytics engineer\b", r"\bproduct operations\b"
    ],
    
    "SENIORITY_BLACKLIST": [
        r"\bsenior manager\b", r"\bprincipal\b", r"\bdirector\b", r"\bvp\b",
        r"\bhead of\b", r"\bgroup product manager\b", r"\btech lead\b",
        r"\bengineering manager\b", r"\bgeneral manager\b", r"\blead architect\b",
        r"\bassociate director\b", r"\bavp\b"
    ],
    
    "LOCATIONS_TIER_1": [
        r"\bdelhi\b", r"\bnoida\b", r"\bgurgaon\b", r"\bgurugram\b", 
        r"\bfaridabad\b", r"\bncr\b", r"\bremote\b", r"\bwfh\b", r"\bwork from home\b"
    ],
    
    "LOCATIONS_TIER_2": [
        r"\bbangalore\b", r"\bbengaluru\b", r"\bhyderabad\b", r"\bpune\b"
    ],
    
    "SKILL_TAXONOMY": {
        "Analytics & BI": [
            r"\bsql\b", r"\bpower bi\b", r"\btableau\b", r"\bpython\b", 
            r"\bexcel\b", r"\betl\b", r"\bdbt\b", r"\blooker\b"
        ],
        "Automation & Low-Code AI": [
            r"\bpower automate\b", r"\bcopilot studio\b", r"\bpower platform\b", 
            r"\bpower apps\b", r"\bcustom connectors\b", r"\bprompt engineering\b"
        ],
        "Product Operations": [
            r"\ba/b testing\b", r"\buser stories\b", r"\bproduct metrics\b", 
            r"\bmixpanel\b", r"\bamplitude\b", r"\broadmap\b", r"\bjira\b"
        ]
    }
}
