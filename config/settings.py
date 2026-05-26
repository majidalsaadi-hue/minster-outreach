
# Ministry of Investment — CRM Tool
# Branding, constants, and app-wide configuration

APP_TITLE_EN = "Minister of Investment — Investor Relations CRM"
APP_TITLE_AR = "وزارة الاستثمار — إدارة علاقات المستثمرين"

# MISA brand colours extracted from official documents
MISA_GREEN       = "#1B5C3F"
MISA_GREEN_DARK  = "#0F3D2A"
MISA_GREEN_LIGHT = "#2D7A54"
MISA_GOLD        = "#C9974A"
MISA_GOLD_LIGHT  = "#E4B96A"
MISA_WHITE       = "#FFFFFF"
MISA_OFF_WHITE   = "#F7F7F2"
MISA_LIGHT_GRAY  = "#F0F0EA"
MISA_TEXT        = "#1A1A1A"
MISA_TEXT_MUTED  = "#6B6B6B"
MISA_RED         = "#C0392B"
MISA_AMBER       = "#D4A017"

# Status colour mapping
STATUS_COLORS = {
    "Completed":          "#1B5C3F",
    "In Progress":        "#C9974A",
    "Not Started":        "#9B9B9B",
    "Blocked":            "#C0392B",
    "Cancelled":          "#7F8C8D",
    "Active":             "#1B5C3F",
    "Pending":            "#C9974A",
    "On Hold":            "#9B9B9B",
    "Closed":             "#4A4A4A",
    "Escalated":          "#C0392B",
    "Follow-Up Required": "#C9974A",
}

PRIORITY_COLORS = {
    "High":   "#C0392B",
    "Medium": "#C9974A",
    "Low":    "#1B5C3F",
}

TIER_COLORS = {
    "Tier 1 — Strategic":       "#C9974A",
    "Tier 2 — High Potential":  "#1B5C3F",
    "Tier 3 — General":         "#6B6B6B",
}

# Journey stages in order
JOURNEY_STAGES = [
    "Qualification",
    "Opportunity Matching",
    "Technical Engagement",
    "Structuring & Gov. Coordination",
    "Closure / Commitment",
    "Aftercare & Expansion",
]

JOURNEY_STAGE_DURATIONS = {
    "Qualification":                  "5–10 Days",
    "Opportunity Matching":           "10–20 Days",
    "Technical Engagement":           "20–45 Days",
    "Structuring & Gov. Coordination":"30–60 Days",
    "Closure / Commitment":           "60–120 Days",
    "Aftercare & Expansion":          "Ongoing",
}

# Number of days without update before flagging as stalled
STALE_THRESHOLD_DAYS = 14

# Alert thresholds
OVERDUE_WARNING_DAYS = 0    # 0 = already past due date
DUE_SOON_DAYS        = 7    # flag if due within 7 days

# Sectors aligned with MISA classification
SECTORS = [
    "Finance & Investment",
    "Energy & Industrial",
    "Healthcare & Retail",
    "Insurance & Risk Management",
    "Tech & AI",
    "Logistics",
    "Real Estate",
    "Aerospace & Defence",
    "Tourism & Hospitality",
    "Other",
]

# Countries (top investor origins + GCC)
COUNTRIES = [
    "United States", "United Kingdom", "France", "Germany",
    "Switzerland", "China", "Japan", "South Korea",
    "Canada", "Australia", "Singapore", "UAE",
    "Kuwait", "Bahrain", "Qatar", "Oman",
    "Saudi Arabia", "Portugal", "Italy", "Spain",
    "Netherlands", "Sweden", "Norway", "India", "Other",
]

ENGAGEMENT_TYPES = [
    "Opportunity",
    "Challenge",
    "Support",
    "Escalation",
    "Administrative",
]

OPPORTUNITY_TYPES = [
    "HE Recent Meeting",
    "Partner",
    "Mature Opportunity",
    "L4 Deal",
    "Expansion Plan",
]

OPPORTUNITY_SOURCES = [
    "Sector Opportunity",
    "Invest Saudi Platform",
    "Strategic Partners",
    "Company Expansion",
    "HE Meeting",
]

OPPORTUNITY_STAGES = ["Early", "Development", "Ready", "Execution"]

CONFIDENCE_LEVELS = ["High", "Medium", "Low", "Speculative"]

MEETING_TYPES = [
    "In-Person", "Virtual", "Phone", "Event", "Ministerial",
]

MEETING_OBJECTIVES = [
    "Introduction",
    "Opportunity Presentation",
    "Technical Discussion",
    "Deal Negotiation",
    "Follow-Up",
    "Escalation",
    "Aftercare",
]

MEETING_STATUSES = [
    "Completed",
    "Pending",
    "Follow-Up Required",
    "Escalated",
]

INVESTOR_TIERS = [
    "Tier 1 — Strategic",
    "Tier 2 — High Potential",
    "Tier 3 — General",
]

INVESTOR_STATUSES = ["Active", "Pending", "On Hold", "Closed"]

DEPARTMENTS = [
    "Business Development",
    "RHQ Department",
    "Investor Services",
    "International Offices",
    "HE Outreach Office",
    "Ministry Programs",
]

ACTION_STATUSES = [
    "Not Started",
    "In Progress",
    "Completed",
    "Blocked",
    "Cancelled",
]

PROGRESS_OPTIONS = ["0%", "25%", "50%", "75%", "100%"]

ESCALATION_FLAGS = ["None", "Flag for RM", "Flag for Leadership"]

TASK_PRIORITIES = ["High", "Medium", "Low"]

# ── Minister decision-support fields ────────────────────────────────────────

# How the deal classification maps to KSA economic impact categories
DEAL_CLASSIFICATIONS = [
    "Greenfield",
    "Brownfield / Expansion",
    "Joint Venture",
    "Acquisition",
    "Strategic Partnership",
    "Fund / FDI",
]

# Vision 2030 pillar alignment
VISION_2030_PILLARS = [
    "Vibrant Society",
    "Thriving Economy",
    "Ambitious Nation",
    "Giga Projects",
    "National Champions",
    "Privatisation",
    "SME Development",
    "Tourism & Culture",
    "Renewables & Sustainability",
]

# Blocker severity — determines escalation path
BLOCKER_LEVELS = [
    "None",
    "Operational — AM can resolve",
    "Ministerial — requires HE intervention",
    "Cabinet — inter-ministerial coordination required",
]

# Who initiated the meeting/engagement
MEETING_INITIATED_BY = [
    "Investor",
    "Account Manager",
    "Relationship Manager",
    "Minister's Office",
    "Third Party / Event",
]

# Strategic priority score descriptors
STRATEGIC_PRIORITY_LABELS = {
    5: "Critical — Immediate ministerial attention",
    4: "High — HE briefing required",
    3: "Medium — Regular RM oversight",
    2: "Standard — AM-managed",
    1: "Watch — Monitor only",
}

# International IR benchmark KPIs (used in dashboard & reports)
IR_BENCHMARKS = {
    "avg_days_to_close":          120,   # world-class IPA benchmark
    "meeting_to_opp_conversion":  0.40,  # 40% industry target
    "opp_to_commitment":          0.25,  # 25% industry target
    "pipeline_coverage_ratio":    3.0,   # pipeline should be 3x target
    "escalation_resolution_days": 7,     # resolve blockers within 7 days
}

SAUDI_CONTENT_OPTIONS = [
    "< 20%", "20–40%", "40–60%", "60–80%", "> 80%", "TBD",
]

TECH_TRANSFER_OPTIONS = ["Yes", "No", "Partial", "TBD"]

MINISTER_ACTION_TYPES = [
    "None Required",
    "Decision Required",
    "Approval Required",
    "Meeting Requested",
    "Escalation — Unblock Deal",
    "Letter / Communication",
    "Site Visit",
]
