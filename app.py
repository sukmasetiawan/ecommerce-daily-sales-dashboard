
from pathlib import Path
import base64
import calendar
from datetime import date, datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Daily Sales Monitoring E-Commerce",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# CONSTANTS / THEME
# =========================================================
DB_FILE = "eCommerce - Daily Sales Monitoring Dashboard.xlsx"
LOGO_FILE = "Final Logo Group_KSP.png"
SALES_SHEET = "DB Penjualan Produk 2026"
TARGET_SHEET = "Monthly Sales Target"

NAVY = "#111B3A"
TEXT = "#26324C"
MUTED = "#7A849C"
PURPLE = "#6538E6"
PURPLE_2 = "#9A69F2"
PURPLE_LIGHT = "#C6B6FF"
PINK = "#F32F79"
TEAL = "#24AFC8"
ORANGE = "#FF8A34"
GREEN = "#15B968"
RED = "#FF3B30"
GRID = "#E7EAF1"

PLATFORM_ORDER = ["TikTok-Tokped", "Shopee", "Others"]
PLATFORM_COLORS = {
    "TikTok-Tokped": TEAL,
    "Shopee": ORANGE,
    "Others": PURPLE_2,
}

# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at 14% 0%, rgba(147, 113, 255, 0.06), transparent 28%),
            linear-gradient(135deg, #F8F9FC 0%, #EEF1F6 100%);
        color: #26324C;
    }

    #MainMenu, footer {visibility: hidden;}
    header[data-testid="stHeader"] {visibility: hidden;}
    [data-testid="stSidebar"] {display: none;}

    .block-container {
        max-width: 1540px;
        padding-top: 0.35rem;
        padding-bottom: 1.8rem;
        padding-left: 2.2rem;
        padding-right: 2.2rem;
    }

    .dashboard-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 24px;
        margin: 0 0 8px 0;
    }

    .dashboard-title {
        margin: 0;
        line-height: 1.04;
        letter-spacing: -1.4px;
        color: #111B3A;
        font-size: clamp(28px, 2.45vw, 43px);
        font-weight: 800;
        white-space: nowrap;
    }

    .dashboard-title .accent {
        background: linear-gradient(90deg, #6E36E8 0%, #C43DD6 52%, #F22D84 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    .asof {
        text-align: right;
        color: #7A849C;
        font-size: 12px;
        line-height: 1.45;
        padding-top: 3px;
        white-space: nowrap;
    }

    .asof strong {
        display: block;
        color: #26324C;
        font-size: 13px;
        font-weight: 650;
    }

    .dashboard-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 22px;
        margin: 0 0 12px 0;
        width: 100%;
    }

    .header-left {
        display: flex;
        align-items: center;
        gap: 14px;
        min-width: 0;
        flex: 1 1 auto;
    }

    .ksp-logo-img {
        width: 92px;
        height: 52px;
        object-fit: contain;
        object-position: left center;
        flex: 0 0 auto;
        display: block;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }

    .ksp-logo-fallback {
        width: 92px;
        height: 52px;
        display: flex;
        align-items: center;
        justify-content: flex-start;
        flex: 0 0 auto;
        color: #E73579;
        font-size: 18px;
        font-weight: 800;
        background: transparent;
        border: none;
        box-shadow: none;
    }

    .dashboard-title {
        margin: 0;
        line-height: 1.04;
        letter-spacing: -1.4px;
        color: #111B3A;
        font-size: clamp(28px, 2.45vw, 43px);
        font-weight: 800;
        white-space: nowrap;
    }

    .dashboard-title .accent {
        background: linear-gradient(90deg, #6E36E8 0%, #C43DD6 52%, #F22D84 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }

    .asof-badge {
        flex: 0 0 auto;
        min-width: 132px;
        padding: 9px 13px 10px 13px;
        border-radius: 13px;
        background: linear-gradient(135deg, #F0ECFF 0%, #F8F4FF 100%);
        border: 1px solid #DDD4F5;
        box-shadow: 0 5px 16px rgba(83, 60, 155, .06);
        text-align: left;
    }

    .asof-label {
        color: #7A849C;
        font-size: 10px;
        font-weight: 600;
        line-height: 1.1;
        margin-bottom: 3px;
    }

    .asof-date {
        color: #33245F;
        font-size: 13px;
        font-weight: 800;
        line-height: 1.1;
        white-space: nowrap;
    }

    .filter-caption {
        font-size: 11px;
        color: #8D96A9;
        font-weight: 600;
        margin-bottom: -5px;
    }

    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stDateInput"] > div > div {
        background: rgba(255,255,255,.88);
        border: 1px solid #E2E6EF;
        border-radius: 13px;
        min-height: 44px;
        box-shadow: 0 3px 15px rgba(34, 47, 72, .035);
    }

    div[data-testid="stSelectbox"] label,
    div[data-testid="stDateInput"] label {
        color: #6F7890 !important;
        font-size: 11px !important;
        font-weight: 600 !important;
    }

    div.stButton > button {
        border-radius: 12px;
        border: 1px solid #E1E5EE;
        background: rgba(255,255,255,.92);
        color: #4F5A73;
        height: 44px;
        font-weight: 600;
        box-shadow: 0 3px 15px rgba(34, 47, 72, .035);
    }

    div.stButton > button:hover {
        border-color: #BEB4F6;
        color: #6538E6;
    }

    .card {
        background: rgba(255,255,255,.94);
        border: 1px solid rgba(222, 226, 236, .96);
        border-radius: 19px;
        box-shadow: 0 8px 28px rgba(30, 45, 75, .055);
    }

    .kpi-card {
        min-height: 224px;
        height: 224px;
        padding: 20px 22px;
        position: relative;
        overflow: hidden;
        box-sizing: border-box;
    }

    .kpi-grid {
        display: grid;
        grid-template-columns: 112px minmax(0, 1fr);
        gap: 20px;
        align-items: start;
    }

    .kpi-icon {
        width: 96px;
        height: 96px;
        border-radius: 21px;
        display: flex;
        justify-content: center;
        align-items: center;
        color: white;
        font-size: 40px;
        box-shadow: 0 12px 24px rgba(101, 56, 230, .17);
    }

    .kpi-icon.sales {
        background: linear-gradient(135deg, #EE2879 0%, #FF8264 100%);
    }

    .kpi-icon.target {
        background: linear-gradient(135deg, #6636E3 0%, #B27BF1 100%);
    }

    .kpi-title {
        color: #17213D;
        font-size: 15px;
        font-weight: 760;
        letter-spacing: -.15px;
        margin: 1px 0 8px 0;
    }

    .kpi-value {
        color: #121C3A;
        font-size: clamp(34px, 2.75vw, 47px);
        line-height: 1.02;
        font-weight: 800;
        letter-spacing: -1.5px;
        margin: 2px 0 8px 0;
    }

    .kpi-value.purple {
        color: #6338D7;
    }

    .subtle {
        color: #7D879E;
        font-size: 12px;
    }

    .growth {
        color: #16B968;
        font-size: 16px;
        font-weight: 750;
        margin-left: 8px;
        white-space: nowrap;
    }

    .growth.negative {
        color: #F04444;
    }

    .sales-meta {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
    }

    .divider {
        height: 1px;
        background: #E5E8EF;
        margin: 11px 0 10px 0;
    }

    .daily-avg-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 16px;
    }

    .daily-avg-label {
        color: #18213C;
        font-size: 13px;
        font-weight: 700;
    }

    .daily-avg-value {
        color: #17213D;
        font-size: 20px;
        font-weight: 800;
        text-align: right;
    }

    .daily-avg-value span {
        color: #63708B;
        font-size: 13px;
        font-weight: 500;
    }

    .progress-track {
        height: 7px;
        width: 100%;
        background: #E7E6F0;
        border-radius: 999px;
        overflow: hidden;
        margin: 10px 0 12px 0;
    }

    .progress-fill {
        height: 100%;
        border-radius: 999px;
        background: linear-gradient(90deg, #7F55E9 0%, #6635DE 100%);
    }

    .target-stats {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: 0;
        margin-top: 6px;
        align-items: start;
    }

    .target-stat {
        padding: 0 14px;
        border-right: 1px solid #E4E7EE;
        min-width: 0;
    }

    .target-stat:first-child {
        padding-left: 0;
    }

    .target-stat:last-child {
        border-right: none;
        padding-right: 0;
    }

    .target-stat-label {
        color: #7E879C;
        font-size: 10px;
        margin-bottom: 4px;
    }

    .target-stat-value {
        color: #1C2744;
        font-size: 12px;
        font-weight: 720;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    .target-stat-value.red { color: #F04444; }
    .target-stat-value.green { color: #15A967; }

    .panel-title {
        color: #17213C;
        font-size: 17px;
        font-weight: 760;
        margin: 0 0 1px 0;
    }

    .panel-subtitle {
        color: #929AAD;
        font-size: 11px;
        margin-bottom: 4px;
    }

    .chart-card {
        padding: 18px 20px 6px 20px;
    }

    .chart-card.compact {
    }

    .top-card {
        padding: 17px 20px 13px 20px;
    }


    .insight-title {
        color: #1867D8;
        font-size: 15px;
        font-weight: 800;
        margin-bottom: 18px;
    }

    .insight-block {
        padding: 0 0 14px 0;
    }

    .insight-block + .insight-block {
        border-top: 1px solid #CFE0EF;
        padding-top: 14px;
    }

    .insight-label {
        color: #2B3855;
        font-size: 13px;
        line-height: 1.45;
    }

    .insight-value {
        color: #1867D8;
        font-size: 18px;
        font-weight: 800;
        margin-top: 2px;
    }

    .insight-note {
        color: #71809A;
        font-size: 11px;
        margin-top: 3px;
    }

    .footer-note {
        color: #8790A4;
        font-size: 11px;
        margin-top: 10px;
    }

    .stPlotlyChart {
        margin-top: -2px;
    }

    .upload-note {
        color: #7A849C;
        font-size: 11px;
    }

    /* Hide extra top whitespace Streamlit occasionally creates */
    div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stAlert"]) {
        margin-top: 0;
    }


    .section-gap {
        height: 14px;
    }

    @media (max-width: 900px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
            padding-top: .8rem;
        }
        .dashboard-header {
            display: block;
            margin-bottom: 8px;
        }
        .dashboard-title {
            white-space: normal;
            font-size: 28px;
            line-height: 1.1;
        }
        .asof {
            text-align: left;
            margin-top: 8px;
        }
        .kpi-grid {
            grid-template-columns: 80px 1fr;
            gap: 15px;
        }
        .kpi-icon {
            width: 72px;
            height: 72px;
            border-radius: 17px;
            font-size: 30px;
        }
        .kpi-card {
            min-height: unset;
            padding: 18px;
        }
        .kpi-value {
            font-size: 34px;
            letter-spacing: -.8px;
        }
        .daily-avg-value {
            font-size: 19px;
        }
        .target-stats {
            grid-template-columns: repeat(3, minmax(0, 1fr));
        }
        .target-stat {
            border-right: 1px solid #E4E7EE;
            border-bottom: none;
            padding: 0 8px;
        }
        .target-stat:first-child {
            padding-left: 0;
        }
        .target-stat:last-child {
            border-right: none;
        }
    }

    /* V3: stable Streamlit layout */
    .dashboard-header {
        width: 100%;
        align-items: flex-end;
        margin: 0 0 12px 0;
    }

    .dashboard-title {
        font-size: clamp(30px, 3.0vw, 43px);
        white-space: nowrap;
        flex: 1 1 auto;
    }

    .asof {
        flex: 0 0 auto;
        text-align: right;
        padding: 0 2px 3px 16px;
        font-size: 11px;
    }

    .asof strong {
        display: inline;
        margin-left: 5px;
        font-size: 12px;
    }

    .filter-button-spacer {
        height: 27px;
    }

    /* Prevent utility controls from breaking into stacked letters */
    div[data-testid="stPopover"] button,
    div.stButton > button {
        white-space: nowrap !important;
        min-width: 0 !important;
        padding-left: .55rem !important;
        padding-right: .55rem !important;
    }

    @media (max-width: 1100px) {
        .dashboard-title {
            font-size: 34px;
            letter-spacing: -1px;
        }
    }

    @media (max-width: 760px) {
        .dashboard-header {
            display: block;
        }
        .dashboard-title {
            white-space: normal;
            font-size: 28px;
        }
        .asof {
            text-align: left;
            padding: 6px 0 0 0;
        }
    }


    /* V5 chart cards: explicit white surfaces using Streamlit keyed containers */
    .st-key-daily_trend_card,
    .st-key-platform_card,
    .st-key-top_products_card {
        background: #FFFFFF !important;
        border: 1px solid #E1E5ED !important;
        border-radius: 18px !important;
        box-shadow: 0 7px 24px rgba(31, 45, 72, 0.055) !important;
        padding: 17px 18px 8px 18px !important;
        overflow: hidden !important;
    }

    .st-key-daily_trend_card > div,
    .st-key-platform_card > div,
    .st-key-top_products_card > div {
        background: transparent !important;
    }

    .st-key-daily_trend_card [data-testid="stPlotlyChart"],
    .st-key-platform_card [data-testid="stPlotlyChart"],
    .st-key-top_products_card [data-testid="stPlotlyChart"] {
        background: transparent !important;
        margin-top: -2px !important;
        margin-bottom: -2px !important;
    }

    .st-key-daily_trend_card iframe,
    .st-key-platform_card iframe,
    .st-key-top_products_card iframe {
        background: transparent !important;
    }


    /* V6: Quick Insight stretches to match Top 5 row height */
    div[data-testid="stHorizontalBlock"]:has(.st-key-top_products_card):has(.st-key-quick_insight_card) {
        align-items: stretch !important;
    }

    div[data-testid="stColumn"]:has(.st-key-quick_insight_card) {
        display: flex !important;
        flex-direction: column !important;
    }

    .st-key-quick_insight_card {
        flex: 1 1 auto !important;
        height: 100% !important;
        min-height: 100% !important;
        background: linear-gradient(145deg, #F7FBFF 0%, #EDF7FF 100%) !important;
        border: 1px solid #D9EAF8 !important;
        border-radius: 19px !important;
        box-shadow: 0 8px 28px rgba(30, 80, 130, .045) !important;
        padding: 18px 20px !important;
        overflow: hidden !important;
    }

    .st-key-quick_insight_card > div {
        height: 100% !important;
    }

    .insight-content {
        min-height: 100%;
    }

    @media (max-width: 900px) {
        .dashboard-header {
            align-items: flex-start;
        }
        .header-left {
            gap: 10px;
        }
        .ksp-logo-img,
        .ksp-logo-fallback {
            width: 78px;
            height: 44px;
        }
        .asof-badge {
            min-width: 118px;
            padding: 8px 10px;
        }
    }

    @media (max-width: 760px) {
        .dashboard-header {
            display: block;
        }
        .header-left {
            align-items: flex-start;
        }
        .dashboard-title {
            white-space: normal;
        }
        .asof-badge {
            display: inline-block;
            margin-top: 10px;
        }
    }


    /* V7: use empty right-side space for growth KPI */
    .sales-top-row {
        display: grid;
        grid-template-columns: minmax(0, 1.45fr) minmax(180px, .8fr);
        gap: 22px;
        align-items: end;
    }

    .sales-main {
        min-width: 0;
    }

    .growth-panel {
        min-width: 0;
        padding-bottom: 7px;
        text-align: left;
    }

    .growth-label {
        color: #7D879E;
        font-size: 12px;
        font-weight: 500;
        margin-bottom: 5px;
        white-space: nowrap;
    }

    .growth-large {
        margin-left: 0 !important;
        font-size: clamp(22px, 1.75vw, 29px) !important;
        line-height: 1.05;
        font-weight: 800 !important;
        letter-spacing: -.4px;
    }

    @media (max-width: 980px) {
        .sales-top-row {
            grid-template-columns: 1fr;
            gap: 4px;
        }
        .growth-panel {
            padding-bottom: 0;
        }
        .growth-label {
            white-space: normal;
        }
    }


    /* V9 mobile filter layout: max 2 rows */
    @media (max-width: 760px) {
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"]) {
            display: flex !important;
            flex-wrap: wrap !important;
            gap: 8px !important;
            align-items: flex-end !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(1) {
            flex: 1 1 34% !important;
            min-width: 0 !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(2) {
            flex: 1 1 28% !important;
            min-width: 0 !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(3) {
            flex: 1 1 34% !important;
            min-width: 0 !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(4) {
            flex: 0 0 72px !important;
            min-width: 72px !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(5) {
            flex: 0 0 132px !important;
            min-width: 132px !important;
        }

        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {
            font-size: 10px !important;
            margin-bottom: 2px !important;
        }

        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            min-height: 40px !important;
            height: 40px !important;
            border-radius: 11px !important;
        }

        .filter-button-spacer {
            height: 21px !important;
        }

        div.stButton > button,
        div[data-testid="stPopover"] button {
            height: 40px !important;
            min-height: 40px !important;
            border-radius: 11px !important;
            font-size: 12px !important;
        }
    }

    @media (max-width: 430px) {
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(1) {
            flex: 1 1 100% !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(2),
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(3) {
            flex: 1 1 calc(50% - 4px) !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(4) {
            flex: 0 0 58px !important;
            min-width: 58px !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(5) {
            flex: 0 0 118px !important;
            min-width: 118px !important;
        }
    }

    </style>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# HELPERS
# =========================================================
def rp_jt(value, decimals=1):
    value = 0 if pd.isna(value) else float(value)
    return f"Rp {value / 1_000_000:,.{decimals}f} jt".replace(",", "X").replace(".", ",").replace("X", ".")

def pct(value, decimals=1):
    if value is None or pd.isna(value):
        return "—"
    return f"{value:.{decimals}f}%"

def platform_group(value):
    s = str(value).strip().lower()
    if "tiktok" in s or "tokped" in s or "tokopedia" in s:
        return "TikTok-Tokped"
    if "shopee" in s:
        return "Shopee"
    return "Others"

def image_data_uri(path):
    path = Path(path)
    if not path.exists():
        return ""
    mime = "image/png" if path.suffix.lower() == ".png" else "image/jpeg"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"

@st.cache_data(show_spinner=False)
def load_workbook(source):
    sales = pd.read_excel(source, sheet_name=SALES_SHEET)
    target = pd.read_excel(source, sheet_name=TARGET_SHEET, header=1)

    sales.columns = [str(c).strip() for c in sales.columns]
    target.columns = [str(c).strip() for c in target.columns]

    required_sales = ["Tanggal", "Platform", "Product", "Sales Quantity", "Sales Value"]
    missing = [c for c in required_sales if c not in sales.columns]
    if missing:
        raise ValueError(f"Kolom sales tidak ditemukan: {', '.join(missing)}")

    required_target = ["Year", "Month No", "Month", "Target Sales Value"]
    missing_target = [c for c in required_target if c not in target.columns]
    if missing_target:
        raise ValueError(f"Kolom target tidak ditemukan: {', '.join(missing_target)}")

    sales["Tanggal"] = pd.to_datetime(sales["Tanggal"], errors="coerce")
    sales["Sales Quantity"] = pd.to_numeric(sales["Sales Quantity"], errors="coerce").fillna(0)
    sales["Sales Value"] = pd.to_numeric(sales["Sales Value"], errors="coerce").fillna(0)
    sales["Platform"] = sales["Platform"].astype(str).str.strip()
    sales["Product"] = sales["Product"].astype(str).str.strip()
    sales = sales.dropna(subset=["Tanggal"])
    sales["Platform Group"] = sales["Platform"].map(platform_group)

    target["Year"] = pd.to_numeric(target["Year"], errors="coerce")
    target["Month No"] = pd.to_numeric(target["Month No"], errors="coerce")
    target["Target Sales Value"] = pd.to_numeric(target["Target Sales Value"], errors="coerce")
    target = target.dropna(subset=["Year", "Month No", "Target Sales Value"])
    target["Year"] = target["Year"].astype(int)
    target["Month No"] = target["Month No"].astype(int)

    return sales, target

def plot_layout(height=330):
    return dict(
        height=height,
        margin=dict(l=8, r=8, t=18, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=TEXT, size=11),
        hoverlabel=dict(bgcolor="#FFFFFF", font_color=NAVY, bordercolor="#E0E3EB"),
        legend=dict(
            orientation="h",
            x=0,
            y=1.12,
            xanchor="left",
            yanchor="bottom",
            font=dict(size=10, color="#64708A"),
        ),
    )

# =========================================================
# DATA SOURCE
# =========================================================
default_path = Path(__file__).parent / DB_FILE

uploaded = None
data_source = default_path

if uploaded is None and not default_path.exists():
    st.error(f"Database default **{DB_FILE}** tidak ditemukan.")
    st.stop()

try:
    sales, targets = load_workbook(data_source)
except Exception as e:
    st.error(f"Gagal membaca database: {e}")
    st.stop()

if sales.empty:
    st.warning("Database sales kosong.")
    st.stop()

data_max_date = sales["Tanggal"].max().date()
data_min_date = sales["Tanggal"].min().date()

# =========================================================
# HEADER PLACEHOLDER
# =========================================================
header_slot = st.empty()

# =========================================================
# FILTERS
# =========================================================
f1, f2, f3, f4, f5 = st.columns([1.0, 1.0, 1.45, .28, .42], gap="small")

with f1:
    selected_date = st.date_input(
        "Date",
        value=data_max_date,
        min_value=data_min_date,
        max_value=data_max_date,
    )

selected_date = min(selected_date, data_max_date)

with f2:
    platform_options = ["All Platform"] + PLATFORM_ORDER
    selected_platform = st.selectbox("Platform", platform_options, index=0)

products_available = (
    sales.loc[sales["Tanggal"].dt.date <= selected_date, "Product"]
    .dropna()
    .sort_values()
    .unique()
    .tolist()
)
with f3:
    selected_product = st.selectbox(
        "Product",
        ["All Product"] + products_available,
        index=0,
    )

with f4:
    st.markdown('<div class="filter-button-spacer"></div>', unsafe_allow_html=True)
    reset = st.button("↻", help="Reset filters", use_container_width=True)
    if reset:
        st.rerun()

with f5:
    st.markdown('<div class="filter-button-spacer"></div>', unsafe_allow_html=True)
    with st.popover("↥ Update", use_container_width=True):
        uploaded = st.file_uploader(
            "Upload Excel terbaru",
            type=["xlsx"],
            label_visibility="collapsed",
            help="File harus mempertahankan sheet DB Penjualan Produk 2026 dan Monthly Sales Target.",
        )
        st.markdown(
            '<div class="upload-note">Upload hanya berlaku untuk sesi aktif.</div>',
            unsafe_allow_html=True,
        )

if uploaded is not None:
    try:
        sales, targets = load_workbook(uploaded)
        data_max_date = sales["Tanggal"].max().date()
        data_min_date = sales["Tanggal"].min().date()
    except Exception as e:
        st.error(f"Gagal membaca database upload: {e}")
        st.stop()

# Render header using the selected dashboard date.
logo_uri = image_data_uri(Path(__file__).parent / LOGO_FILE)
logo_html = (
    f'<img class="ksp-logo-img" src="{logo_uri}" alt="KSP logo">'
    if logo_uri else
    '<div class="ksp-logo-fallback">KSP</div>'
)

header_slot.markdown(
    f"""
    <div class="dashboard-header">
      <div class="header-left">
        {logo_html}
        <h1 class="dashboard-title">
          DAILY SALES MONITORING <span class="accent">E-COMMERCE</span>
        </h1>
      </div>
      <div class="asof-badge">
        <div class="asof-label">Data per</div>
        <div class="asof-date">{selected_date.strftime("%d %b %Y")}</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

selected_ts = pd.Timestamp(selected_date)
year, month, day = selected_ts.year, selected_ts.month, selected_ts.day
month_start = selected_ts.replace(day=1)
prev_month_end = month_start - pd.Timedelta(days=1)
prev_month_start = prev_month_end.replace(day=1)
prev_compare_day = min(day, calendar.monthrange(prev_month_start.year, prev_month_start.month)[1])
prev_compare_end = prev_month_start + pd.Timedelta(days=prev_compare_day - 1)
month_end = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])

# Base dimension filters
filtered = sales.copy()
if selected_platform != "All Platform":
    filtered = filtered[filtered["Platform Group"] == selected_platform]
if selected_product != "All Product":
    filtered = filtered[filtered["Product"] == selected_product]

mtd = filtered[
    (filtered["Tanggal"] >= month_start) &
    (filtered["Tanggal"] <= selected_ts)
].copy()

prev_same = filtered[
    (filtered["Tanggal"] >= prev_month_start) &
    (filtered["Tanggal"] <= prev_compare_end)
].copy()

sales_mtd = float(mtd["Sales Value"].sum())
prev_sales_same = float(prev_same["Sales Value"].sum())
growth = ((sales_mtd / prev_sales_same) - 1) * 100 if prev_sales_same != 0 else np.nan
daily_avg = sales_mtd / day if day > 0 else 0

# Monthly target: targets are total e-commerce; dimension filtering changes achievement meaning.
target_row = targets[(targets["Year"] == year) & (targets["Month No"] == month)]
monthly_target = float(target_row["Target Sales Value"].iloc[0]) if not target_row.empty else np.nan

# When filtered to specific platform/product, achievement is still vs total target only if unfiltered.
is_total_view = selected_platform == "All Platform" and selected_product == "All Product"
achievement = (sales_mtd / monthly_target * 100) if (not pd.isna(monthly_target) and monthly_target != 0 and is_total_view) else np.nan
gap = (monthly_target - sales_mtd) if (not pd.isna(monthly_target) and is_total_view) else np.nan

remaining_days = max((month_end.date() - selected_date).days, 0)
required_daily = (
    max(monthly_target - sales_mtd, 0) / remaining_days
    if (is_total_view and not pd.isna(monthly_target) and remaining_days > 0)
    else np.nan
)

# =========================================================
# KPI ROW
# =========================================================
k1, k2 = st.columns([1, 1], gap="medium")

growth_class = "growth negative" if (not pd.isna(growth) and growth < 0) else "growth"
growth_arrow = "▼" if (not pd.isna(growth) and growth < 0) else "▲"
growth_text = "—" if pd.isna(growth) else f"{growth_arrow} {growth:+.1f}%"

with k1:
    st.markdown(
        f"""
        <div class="card kpi-card">
          <div class="kpi-grid">
            <div class="kpi-icon sales">◉</div>
            <div>
              <div class="sales-top-row">
                <div class="sales-main">
                  <div class="kpi-title">SALES - MONTH TO DATE</div>
                  <div class="kpi-value">{rp_jt(sales_mtd)}</div>
                </div>
                <div class="growth-panel">
                  <div class="growth-label">vs Same Period Last Month</div>
                  <div class="{growth_class} growth-large">{growth_text}</div>
                </div>
              </div>
              <div class="divider"></div>
              <div class="daily-avg-row">
                <div>
                  <div class="daily-avg-label">Daily Sales Avg</div>
                  <div class="subtle">({day} days elapsed)</div>
                </div>
                <div class="daily-avg-value">{rp_jt(daily_avg)} <span>/day</span></div>
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k2:
    if is_total_view and not pd.isna(monthly_target):
        progress = max(0, min(100, achievement))
        gap_class = "green" if gap <= 0 else "red"
        gap_display = (
            f"+{rp_jt(abs(gap))}" if gap < 0 else f"-{rp_jt(abs(gap))}"
        )
        achievement_display = pct(achievement)
        target_note = "Achievement vs Monthly Target"
        actual_display = rp_jt(sales_mtd)
        target_display = rp_jt(monthly_target)
    else:
        progress = 0
        gap_class = ""
        gap_display = "—"
        achievement_display = "—"
        target_note = "Available on All Platform + All Product"
        actual_display = rp_jt(sales_mtd)
        target_display = rp_jt(monthly_target) if not pd.isna(monthly_target) else "—"

    st.markdown(
        f"""
        <div class="card kpi-card">
          <div class="kpi-grid">
            <div class="kpi-icon target">◎</div>
            <div>
              <div class="kpi-title">SALES TARGET ACHIEVEMENT</div>
              <div class="kpi-value purple">{achievement_display}</div>
              <div class="subtle">{target_note}</div>
              <div class="progress-track">
                <div class="progress-fill" style="width:{progress:.1f}%"></div>
              </div>
              <div class="target-stats">
                <div class="target-stat">
                  <div class="target-stat-label">Actual MTD</div>
                  <div class="target-stat-value">{actual_display}</div>
                </div>
                <div class="target-stat">
                  <div class="target-stat-label">Monthly Target</div>
                  <div class="target-stat-value">{target_display}</div>
                </div>
                <div class="target-stat">
                  <div class="target-stat-label">Gap</div>
                  <div class="target-stat-value {gap_class}">{gap_display}</div>
                </div>
              </div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)

# =========================================================
# DAILY SALES TREND
# =========================================================
c1, c2 = st.columns([1.65, 1.0], gap="medium")

with c1:
    with st.container(key="daily_trend_card", border=False):
        st.markdown(
            f'<div class="panel-title">DAILY SALES TREND <span style="font-size:11px;color:#929AAD;font-weight:500">(Rp jt)</span></div>',
            unsafe_allow_html=True,
        )

        curr_daily = (
            filtered[
                (filtered["Tanggal"] >= month_start) &
                (filtered["Tanggal"] <= selected_ts)
            ]
            .assign(Day=lambda x: x["Tanggal"].dt.day)
            .groupby("Day", as_index=False)["Sales Value"].sum()
        )
        prev_full = (
            filtered[
                (filtered["Tanggal"] >= prev_month_start) &
                (filtered["Tanggal"] <= prev_month_end)
            ]
            .assign(Day=lambda x: x["Tanggal"].dt.day)
            .groupby("Day", as_index=False)["Sales Value"].sum()
        )

        current_days = pd.DataFrame({"Day": range(1, calendar.monthrange(year, month)[1] + 1)})
        prev_days = pd.DataFrame({"Day": range(1, calendar.monthrange(prev_month_start.year, prev_month_start.month)[1] + 1)})
        curr_daily = current_days.merge(curr_daily, on="Day", how="left")
        prev_full = prev_days.merge(prev_full, on="Day", how="left")

        # no line after selected date for current month
        curr_daily.loc[curr_daily["Day"] > day, "Sales Value"] = np.nan
        curr_daily["Sales Value"] = curr_daily["Sales Value"] / 1_000_000
        prev_full["Sales Value"] = prev_full["Sales Value"].fillna(0) / 1_000_000

        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=curr_daily["Day"],
                y=curr_daily["Sales Value"],
                mode="lines+markers",
                name=f"This Month ({selected_ts.strftime('%b %Y')})",
                line=dict(color=PURPLE, width=3),
                marker=dict(size=5, color="#FFFFFF", line=dict(color=PURPLE, width=1.8)),
                connectgaps=False,
                hovertemplate="Day %{x}<br>Rp %{y:,.1f} jt<extra></extra>",
            )
        )
        fig.add_trace(
            go.Scatter(
                x=prev_full["Day"],
                y=prev_full["Sales Value"],
                mode="lines",
                name=f"Last Month ({prev_month_start.strftime('%b %Y')})",
                line=dict(color=PURPLE_LIGHT, width=2.2, dash="dash"),
                hovertemplate="Day %{x}<br>Rp %{y:,.1f} jt<extra></extra>",
            )
        )

        fig.update_layout(**plot_layout(292))
        fig.update_xaxes(
            title=None,
            tickmode="linear",
            dtick=1,
            showgrid=False,
            zeroline=False,
            color="#6F7890",
            tickfont=dict(size=9),
        )
        fig.update_yaxes(
            title=None,
            showgrid=True,
            gridcolor=GRID,
            griddash="dot",
            zeroline=False,
            color="#6F7890",
            ticksuffix="",
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# =========================================================
# SALES BY PLATFORM
# =========================================================
with c2:
    with st.container(key="platform_card", border=False):
        st.markdown(
            '<div class="panel-title">SALES BY PLATFORM <span style="font-size:11px;color:#929AAD;font-weight:500">(Rp jt)</span></div>',
            unsafe_allow_html=True,
        )

        by_platform = (
            mtd.groupby("Platform Group", as_index=False)["Sales Value"]
            .sum()
            .set_index("Platform Group")
            .reindex(PLATFORM_ORDER, fill_value=0)
            .reset_index()
        )
        total_platform = by_platform["Sales Value"].sum()
        by_platform["Share"] = np.where(
            total_platform != 0,
            by_platform["Sales Value"] / total_platform * 100,
            0,
        )
        by_platform["ValueJt"] = by_platform["Sales Value"] / 1_000_000

        platform_text = [
            f"{rp_jt(v)}<br><span style='font-size:10px'>{s:.1f}%</span>"
            for v, s in zip(by_platform["Sales Value"], by_platform["Share"])
        ]

        fig2 = go.Figure(
            go.Bar(
                y=by_platform["Platform Group"],
                x=by_platform["ValueJt"],
                orientation="h",
                marker_color=[PLATFORM_COLORS[p] for p in by_platform["Platform Group"]],
                text=platform_text,
                textposition="outside",
                cliponaxis=False,
                hovertemplate="%{y}<br>Rp %{x:,.1f} jt<extra></extra>",
            )
        )
        fig2.update_layout(**plot_layout(292))
        fig2.update_layout(showlegend=False, bargap=.48, margin=dict(l=8, r=78, t=18, b=20))
        fig2.update_xaxes(
            showgrid=True,
            gridcolor=GRID,
            griddash="dot",
            zeroline=False,
            color="#6F7890",
            title=None,
        )
        fig2.update_yaxes(
            showgrid=False,
            autorange="reversed",
            color="#42506B",
            tickfont=dict(size=11),
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)

# =========================================================
# TOP 5 PRODUCTS + QUICK INSIGHT
# =========================================================
b1, b2 = st.columns([2.9, 1.0], gap="medium")

with b1:
    with st.container(key="top_products_card", border=False):
        st.markdown(
            '<div class="panel-title">TOP 5 PRODUCTS <span style="font-size:11px;color:#929AAD;font-weight:500">(by Sales Value · Rp jt)</span></div>',
            unsafe_allow_html=True,
        )

        top5 = (
            mtd.groupby("Product", as_index=False)["Sales Value"]
            .sum()
            .sort_values("Sales Value", ascending=False)
            .head(5)
            .sort_values("Sales Value", ascending=True)
        )
        total_mtd_for_share = float(mtd["Sales Value"].sum())
        top5["Share"] = np.where(total_mtd_for_share != 0, top5["Sales Value"] / total_mtd_for_share * 100, 0)
        top5["ValueJt"] = top5["Sales Value"] / 1_000_000

        top_colors = ["#F4B51F", "#4BC675", "#4E85EB", "#8A5AE2", "#EF347D"][:len(top5)]

        fig3 = go.Figure(
            go.Bar(
                y=top5["Product"],
                x=top5["ValueJt"],
                orientation="h",
                marker_color=top_colors,
                text=[f"{rp_jt(v)}  ·  {s:.1f}%" for v, s in zip(top5["Sales Value"], top5["Share"])],
                textposition="outside",
                cliponaxis=False,
                hovertemplate="%{y}<br>Rp %{x:,.1f} jt<extra></extra>",
            )
        )
        fig3.update_layout(**plot_layout(230))
        fig3.update_layout(
            showlegend=False,
            bargap=.48,
            margin=dict(l=8, r=120, t=16, b=12),
        )
        fig3.update_xaxes(showgrid=True, gridcolor=GRID, griddash="dot", zeroline=False, color="#6F7890")
        fig3.update_yaxes(showgrid=False, color="#34415C", tickfont=dict(size=10))
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

with b2:
    if is_total_view and not pd.isna(required_daily):
        pace_msg = "Current pace is above required pace." if daily_avg >= required_daily else "Current pace is below required pace."
        req_value = rp_jt(required_daily)
        req_note = f"Based on {remaining_days} days remaining"
    elif is_total_view and remaining_days == 0 and not pd.isna(monthly_target):
        pace_msg = "Month has reached its final day."
        req_value = "—"
        req_note = "No remaining selling days"
    else:
        pace_msg = "Target pace is shown on total view."
        req_value = "—"
        req_note = "Select All Platform + All Product"

    with st.container(key="quick_insight_card", border=False):
        st.markdown(
            f"""
            <div class="insight-content">
              <div class="insight-title">💡 QUICK INSIGHT</div>
              <div class="insight-block">
                <div class="insight-label">Sales pace (Daily Avg)</div>
                <div class="insight-value">{rp_jt(daily_avg)}/day</div>
                <div class="insight-note">{pace_msg}</div>
              </div>
              <div class="insight-block">
                <div class="insight-label">Required daily sales to reach monthly target</div>
                <div class="insight-value">{req_value}/day</div>
                <div class="insight-note">{req_note}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown(
    '<div class="footer-note">Note: All sales values are in Rupiah (Rp). Sales Target Achievement uses the total monthly e-commerce target and is therefore shown only on the unfiltered total view.</div>',
    unsafe_allow_html=True,
)
