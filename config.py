# config.py

# ==========================================
# 1. WEEKLY REPORT (OSS) SETTINGS & METRIC NAMES
# ==========================================
WEEKLY_CONFIG = {
    "sheet_name": "2G-3G Data",
    "date_col": ["DATE_ID", "Date", "Timestamp"],
    "network_col": ["Network", "Tech", "Technology"],
    "networks": ["2G", "3G"],

    # Custom Metric Display Names / Chart Titles
    "metric_titles": {
        "Avail_Active": "Average Availability",
        "CFR": "CFR%",
        "CDR_CELL": "CDR%",
        "Calls": "Total Calls",
        "DATA_MB": "Data Volume (MB)",
        "SPEECH_TRAFFIC": "Speech Traffic"
    },

    # Mapping: Placeholder Tag -> (Column Alias Key, Aggregation, Is_Percentage, Network_Filter)
    "charts": {
        "{{CHART_2G_AVAIL}}":         ("Avail_Active",   "mean", False, "2G"),
        "{{CHART_2G_CFR}}":           ("CFR",            "mean", True,  "2G"),
        "{{CHART_2G_CDR}}":           ("CDR_CELL",       "mean", True,  "2G"),
        "{{CHART_2G_CALLS}}":         ("Calls",          "sum",  False, "2G"),
        "{{CHART_2G_DATAMB}}":        ("DATA_MB",        "sum",  False, "2G"),
        "{{CHART_2G_SPEECHTRAFFIC}}": ("SPEECH_TRAFFIC", "sum",  False, "2G"),

        "{{CHART_3G_AVAIL}}":         ("Avail_Active",   "mean", False, "3G"),
        "{{CHART_3G_CFR}}":           ("CFR",            "mean", True,  "3G"),
        "{{CHART_3G_CDR}}":           ("CDR_CELL",       "mean", True,  "3G"),
        "{{CHART_3G_CALLS}}":         ("Calls",          "sum",  False, "3G"),
        "{{CHART_3G_DATAMB}}":        ("DATA_MB",        "sum",  False, "3G"),
        "{{CHART_3G_SPEECHTRAFFIC}}": ("SPEECH_TRAFFIC", "sum",  False, "3G"),
    },

    # Optional alias lists for Weekly OSS columns
    "column_aliases": {
        "Avail_Active":   ["Avail_Active", "Availability"],
        "CFR":            ["CFR", "CFR%"],
        "CDR_CELL":       ["CDR_CELL", "CDR%"],
        "Calls":          ["Calls", "Total Calls"],
        "DATA_MB":        ["DATA_MB", "Data Volume"],
        "SPEECH_TRAFFIC": ["SPEECH_TRAFFIC", "Speech Traffic"]
    }
}

# ==========================================
# 2. CHART VISUAL STYLING
# ==========================================
CHART_STYLE = {
    "fig_width": 8.5,
    "fig_height": 4.4,                # <-- Reduced height from 4.5 to 4.0 (~10% shorter)
    "dpi": 300,
    "line_color": "#0F5C7E",          # Deep teal line
    "line_width": 2.8,
    "title_color": "#FF0000",         # Red chart title
    "title_size": 14,
    "border_color": "#BFBFBF",        # Grey outer border around chart
    "border_width": 1.5,
    "border_pad_inches": 0.25,
    "x_tick_interval": 1,             # 1 = Show every weekly date point
    "x_label_rotation": 45,           # Tilted date labels
    "y_axis_bins": 12,                # Denser Y-axis ticks & gridlines
    "word_image_width_inches": 6.1
}

# ==========================================
# 3. COLUMN MATCHING RULES
# ==========================================
COLUMN_MATCHING = {
    "enable_fuzzy": True,
    "similarity_cutoff": 0.65,
    "warn_on_fallback": True,
    "prompt_on_missing": True         # Opens interactive column-picker dialog if all tiers fail
}

# ==========================================
# 4. DRIVE TEST (DT) REPORT SETTINGS & PLACEHOLDERS
# ==========================================
DT_CONFIG = {
    "sheets": {
        "cdr_cbr":    "CDR & CBR",
        "rx_2g":      "2G Idle - received level",
        "rx_3g":      "3G Idle- received level",
        "polqa":      "Polqa",
        "data_exp":   "Data Export",
        "setup_time": "session setup time "
    },
    # Each key can be a single string OR a list of known column aliases
    "columns": {
        "attempt":     ["Call Attempt", "Call Attempt CSFB"],
        "setup":       ["Call Setup"],
        "established": ["Call Established"],
        "end":         ["Call End"],
        "rx_2g_val":   ["RxLevSub (dBm) - .Server", "RxLevSub"],
        "rx_3g_val":   ["Categorized RSCP:A1", "RSCP"],
        "polqa_val":   ["POLQA SWB Score DL", "POLQA"],
        "setup_val":   ["session setup time ", "Setup_Time_sec"],
        "http_start":  ["HTTP Start"],
        "http_end":    ["HTTP End"],
        "post_start":  ["HTTP Post Start"],
        "post_end":    ["HTTP Post End"]
    },
    "thresholds": {
        "rx_min_dbm": -90.0,
        "polqa_min_mos": 2.8
    },
    "placeholders": {
        "{{DT_ATTEMPTS}}": "attempts",
        "{{DT_BLOCKS}}": "blocks",
        "{{DT_BLOCKS_%}}": "blocks_pct",
        "{{DT_DROPS}}": "drops",
        "{{DT_DROPS_%}}": "drops_pct",
        "{{DT_SIGNAL_AVG_2G}}": "avg_2g",
        "{{DT_TOTAL_SAMPLES_2G}}": "total_2g",
        "{{DT_BAD_SAMPLES_2G}}": "bad_2g",
        "{{DT_GOOD_SAMPLES_2G}}": "good_2g",
        "{{DT_SIGNAL_AVG_3G}}": "avg_3g",
        "{{DT_TOTAL_SAMPLES_3G}}": "total_3g",
        "{{DT_BAD_SAMPLES_3G}}": "bad_3g",
        "{{DT_GOOD_SAMPLES_3G}}": "good_3g",
        "{{DT_SESSION_BLOCK_%}}": "session_block_pct",
        "{{DT_SESSION_SETUP_TIME_AVG}}": "avg_setup",
        "{{DT_VOICE_SAMPLES_BELOW_2.8_%}}": "bad_polqa_pct"
    }
}