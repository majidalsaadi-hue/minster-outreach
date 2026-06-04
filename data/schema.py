
# Expected Excel schema — used for upload validation and template generation.
# Each sheet definition lists required columns, optional columns, and dtypes.

SCHEMA = {
    "Investor Master": {
        "required": [
            "Investor ID", "Company Name", "Country", "Sector",
            "Investor Tier", "Relationship Manager", "Journey Stage",
            "Relationship Status",
        ],
        "optional": [
            "Account Manager", "Outreach Manager",
            "Key Contact Name", "Key Contact Title",
            "Est. Investment Value (SAR)", "Actual Commitment (SAR)",
            "Est. Jobs Created", "Saudi Content %", "Technology Transfer",
            "Deal Classification", "Vision 2030 Pillar",
            "Strategic Priority Score",
            "Minister Action Required", "Decision Required By",
            "Blocker Level",
            "Last Meeting Date", "Next Meeting Date", "Last Updated",
            "Escalation Flag", "Notes",
        ],
        "dtypes": {
            "Last Meeting Date":           "date",
            "Next Meeting Date":           "date",
            "Last Updated":                "date",
            "Decision Required By":        "date",
            "Est. Investment Value (SAR)": "float",
            "Actual Commitment (SAR)":     "float",
            "Est. Jobs Created":           "float",
            "Strategic Priority Score":    "float",
        },
    },

    "Meeting Log": {
        "required": [
            "Meeting ID", "Investor ID", "Company Name",
            "Meeting Date", "Meeting Type", "Meeting Status",
        ],
        "optional": [
            "Location", "MISA Attendees", "Investor Attendees",
            "Meeting Objective", "Key Discussion Points", "Decisions Made",
            "Blockers Identified", "Next Steps", "Follow-Up Owner",
            "Follow-Up Due Date", "RM Reviewed", "Logged By",
        ],
        "dtypes": {
            "Meeting Date":      "date",
            "Follow-Up Due Date":"date",
        },
    },

    "Opportunity Pipeline": {
        "required": [
            "Opportunity ID", "Investor ID", "Company Name",
            "Opportunity Name", "Sector", "Opportunity Stage",
            "Opportunity Status",
        ],
        "optional": [
            "Opportunity Type", "Opportunity Source",
            "Est. Value (SAR)", "Confidence Level",
            "Assigned AM", "Start Date", "Target Closure Date",
            "Blockers", "Escalation Required", "Last Updated", "Notes",
        ],
        "dtypes": {
            "Start Date":          "date",
            "Target Closure Date": "date",
            "Last Updated":        "date",
            "Est. Value (SAR)":    "float",
        },
    },

    "Action Items": {
        "required": [
            "Action ID", "Investor ID", "Company Name",
            "Action Description", "Assigned To",
            "Priority", "Status",
        ],
        "optional": [
            "Meeting ID", "Opportunity ID", "Department",
            "Sector", "Type of Engagement",
            "Start Date", "Due Date", "Progress",
            "Escalation Flag", "Remarks",
            "Outcome", "Next Action", "Next Action Date",
            "Last Updated", "Updated By",
        ],
        "dtypes": {
            "Start Date":      "date",
            "Due Date":        "date",
            "Next Action Date":"date",
            "Last Updated":    "date",
        },
    },

    "RM Tasks": {
        "required": [
            "Task ID", "Task Title", "Priority", "Status",
        ],
        "optional": [
            "Linked Investor", "Due Date", "Notes", "Created Date",
        ],
        "dtypes": {
            "Due Date":     "date",
            "Created Date": "date",
        },
    },

    "Deal Progress": {
        "required": [
            "Deal ID", "Investor ID", "Company Name",
            "Deal Name", "Deal Stage", "Deal Status",
            "Challenge Severity", "Escalation Required",
        ],
        "optional": [
            "Linked Opportunity ID",
            "Est. Value (SAR)",
            "Challenge Description", "Challenge Classification",
            "Proposed Solution",
            "Escalation Level", "Escalation Status",
            "Assigned Owner", "Target Resolution Date",
            "Last Updated", "Notes",
        ],
        "dtypes": {
            "Est. Value (SAR)":       "float",
            "Target Resolution Date": "date",
            "Last Updated":           "date",
        },
    },
}

# Legacy sheet names from the original tracker — used to detect old-format uploads
LEGACY_SHEET_PREFIXES = ["Action Items"]

# Columns used to detect which row is the header in legacy files
LEGACY_HEADER_ROW_MARKERS = ["ID", "Action Item", "Status"]
