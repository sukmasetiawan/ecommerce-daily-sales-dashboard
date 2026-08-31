
from pathlib import Path
from io import BytesIO
from urllib.parse import quote
import base64
import requests
import html
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


    /* =====================================================
       V10 MOBILE ONLY — desktop remains unchanged
       ===================================================== */
    .mobile-only-card,
    .st-key-daily_trend_card_mobile,
    .st-key-platform_card_mobile,
    .st-key-top_products_card_mobile {
        display: none;
    }

    @media (max-width: 760px) {
        .block-container {
            padding: 0.55rem 0.72rem 1.15rem 0.72rem !important;
        }

        /* ---------- Header ---------- */
        .dashboard-header {
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) auto !important;
            grid-template-areas:
                "left left"
                "date date" !important;
            gap: 7px !important;
            margin-bottom: 10px !important;
            align-items: start !important;
        }

        .header-left {
            grid-area: left;
            display: grid !important;
            grid-template-columns: 74px minmax(0, 1fr) !important;
            gap: 9px !important;
            align-items: center !important;
            width: 100% !important;
        }

        .ksp-logo-img,
        .ksp-logo-fallback {
            width: 72px !important;
            height: 44px !important;
        }

        .dashboard-title {
            display: block !important;
            font-size: clamp(21px, 6.2vw, 27px) !important;
            line-height: 1.04 !important;
            letter-spacing: -0.65px !important;
            white-space: normal !important;
            margin: 0 !important;
        }

        .dashboard-title .title-main,
        .dashboard-title .title-accent {
            display: block !important;
            white-space: nowrap !important;
        }

        .dashboard-title .title-accent {
            margin-top: 3px !important;
        }

        .asof-badge {
            grid-area: date;
            display: inline-flex !important;
            align-items: center !important;
            gap: 7px !important;
            width: fit-content !important;
            min-width: 0 !important;
            margin: 0 !important;
            padding: 6px 9px !important;
            border-radius: 10px !important;
        }

        .asof-label {
            margin: 0 !important;
            font-size: 9px !important;
        }

        .asof-date {
            font-size: 11px !important;
        }

        /* ---------- Filters: EXACTLY 2 rows ----------
           Row 1: Date | Platform
           Row 2: Product | Reset | Update
        */
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"]) {
            display: grid !important;
            grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) 54px 104px !important;
            grid-template-areas:
                "date platform platform platform"
                "product product reset update" !important;
            gap: 7px !important;
            align-items: end !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"] {
            width: auto !important;
            min-width: 0 !important;
            flex: none !important;
        }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(1) { grid-area: date !important; }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(2) { grid-area: platform !important; }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(3) { grid-area: product !important; }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(4) { grid-area: reset !important; }

        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"])
        > div[data-testid="stColumn"]:nth-child(5) { grid-area: update !important; }

        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {
            font-size: 9px !important;
            margin-bottom: 1px !important;
        }

        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            height: 38px !important;
            min-height: 38px !important;
            border-radius: 10px !important;
        }

        .filter-button-spacer {
            height: 20px !important;
        }

        div.stButton > button,
        div[data-testid="stPopover"] button {
            height: 38px !important;
            min-height: 38px !important;
            border-radius: 10px !important;
            font-size: 11px !important;
        }

        /* ---------- KPI cards ---------- */
        .kpi-card {
            height: auto !important;
            min-height: 0 !important;
            padding: 14px !important;
            border-radius: 16px !important;
        }

        .kpi-grid {
            grid-template-columns: 62px minmax(0, 1fr) !important;
            gap: 11px !important;
        }

        .kpi-icon {
            width: 58px !important;
            height: 58px !important;
            border-radius: 15px !important;
            font-size: 25px !important;
        }

        .kpi-title {
            font-size: 12px !important;
            margin-bottom: 4px !important;
        }

        .kpi-value {
            font-size: clamp(29px, 8.1vw, 36px) !important;
            line-height: 1 !important;
            margin: 1px 0 5px 0 !important;
            letter-spacing: -0.7px !important;
        }

        .sales-top-row {
            grid-template-columns: minmax(0, 1.25fr) minmax(105px, .75fr) !important;
            gap: 8px !important;
            align-items: end !important;
        }

        .growth-panel {
            padding-bottom: 2px !important;
        }

        .growth-label {
            font-size: 9px !important;
            line-height: 1.2 !important;
            white-space: normal !important;
            margin-bottom: 3px !important;
        }

        .growth-large {
            font-size: 20px !important;
            white-space: nowrap !important;
        }

        .divider {
            margin: 8px 0 !important;
        }

        .daily-avg-label {
            font-size: 11px !important;
        }

        .subtle {
            font-size: 9px !important;
        }

        .daily-avg-value {
            font-size: 17px !important;
            white-space: nowrap !important;
        }

        .daily-avg-value span {
            font-size: 10px !important;
        }

        /* Target KPI: Actual + Target row, Gap full-width second row */
        .progress-track {
            margin: 7px 0 8px 0 !important;
            height: 6px !important;
        }

        .target-stats {
            grid-template-columns: 1fr 1fr !important;
            gap: 0 !important;
            margin-top: 4px !important;
        }

        .target-stat {
            padding: 0 7px !important;
            border-right: 1px solid #E4E7EE !important;
            border-bottom: none !important;
        }

        .target-stat:first-child {
            padding-left: 0 !important;
        }

        .target-stat:nth-child(2) {
            border-right: none !important;
        }

        .target-stat:last-child {
            grid-column: 1 / -1 !important;
            display: flex !important;
            justify-content: space-between !important;
            align-items: center !important;
            border-right: none !important;
            border-top: 1px solid #E7EAF0 !important;
            margin-top: 7px !important;
            padding: 7px 0 0 0 !important;
        }

        .target-stat-label {
            font-size: 8px !important;
            margin-bottom: 2px !important;
        }

        .target-stat:last-child .target-stat-label {
            margin-bottom: 0 !important;
        }

        .target-stat-value {
            font-size: 10px !important;
            overflow: visible !important;
            text-overflow: clip !important;
        }

        /* ---------- Section spacing ---------- */
        .section-gap {
            height: 9px !important;
        }

        /* Streamlit columns become stacked cards */
        div[data-testid="stHorizontalBlock"]:has(.st-key-daily_trend_card),
        div[data-testid="stHorizontalBlock"]:has(.st-key-top_products_card) {
            gap: 10px !important;
        }

        /* ---------- Desktop chart cards hidden on mobile ---------- */
        .st-key-daily_trend_card,
        .st-key-platform_card,
        .st-key-top_products_card {
            display: none !important;
        }

        /* ---------- Mobile-specific chart cards ---------- */
        .st-key-daily_trend_card_mobile,
        .st-key-platform_card_mobile,
        .st-key-top_products_card_mobile {
            display: block !important;
            background: #FFFFFF !important;
            border: 1px solid #E1E5ED !important;
            border-radius: 16px !important;
            box-shadow: 0 6px 20px rgba(31,45,72,.045) !important;
            padding: 13px 12px 7px 12px !important;
            overflow: hidden !important;
        }

        .panel-title {
            font-size: 13px !important;
        }

        .st-key-daily_trend_card_mobile [data-testid="stPlotlyChart"],
        .st-key-platform_card_mobile [data-testid="stPlotlyChart"] {
            margin-top: -4px !important;
            margin-bottom: -6px !important;
        }

        /* ---------- Top 5 mobile list ---------- */
        .mobile-top5-list {
            display: block !important;
            margin-top: 9px !important;
        }

        .mobile-top5-item {
            padding: 7px 0 !important;
            border-top: 1px solid #EEF0F5 !important;
        }

        .mobile-top5-item:first-child {
            border-top: none !important;
            padding-top: 2px !important;
        }

        .mobile-top5-head {
            display: grid !important;
            grid-template-columns: 20px minmax(0, 1fr) auto !important;
            gap: 6px !important;
            align-items: center !important;
        }

        .mobile-rank {
            width: 18px !important;
            height: 18px !important;
            border-radius: 999px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            font-size: 9px !important;
            font-weight: 800 !important;
            color: white !important;
        }

        .mobile-product-name {
            font-size: 10px !important;
            line-height: 1.2 !important;
            color: #26324C !important;
            font-weight: 650 !important;
            overflow-wrap: anywhere !important;
        }

        .mobile-product-value {
            font-size: 9px !important;
            color: #59657D !important;
            font-weight: 700 !important;
            white-space: nowrap !important;
        }

        .mobile-product-bar {
            height: 5px !important;
            margin: 5px 0 0 26px !important;
            border-radius: 999px !important;
            background: #EEF0F5 !important;
            overflow: hidden !important;
        }

        .mobile-product-bar > span {
            display: block !important;
            height: 100% !important;
            border-radius: 999px !important;
        }

        /* ---------- Quick Insight ---------- */
        div[data-testid="stColumn"]:has(.st-key-quick_insight_card) {
            display: block !important;
        }

        .st-key-quick_insight_card {
            min-height: 0 !important;
            height: auto !important;
            padding: 14px !important;
            border-radius: 16px !important;
        }

        .insight-title {
            font-size: 13px !important;
            margin-bottom: 11px !important;
        }

        .insight-block {
            padding-bottom: 10px !important;
        }

        .insight-block + .insight-block {
            padding-top: 10px !important;
        }

        .insight-label {
            font-size: 10px !important;
        }

        .insight-value {
            font-size: 16px !important;
        }

        .insight-note {
            font-size: 9px !important;
        }

        .footer-note {
            font-size: 8px !important;
            line-height: 1.35 !important;
            margin-top: 7px !important;
        }
    }

    @media (max-width: 390px) {
        .dashboard-title {
            font-size: 20px !important;
        }
        .header-left {
            grid-template-columns: 66px minmax(0, 1fr) !important;
        }
        .ksp-logo-img,
        .ksp-logo-fallback {
            width: 64px !important;
        }
        .sales-top-row {
            grid-template-columns: minmax(0, 1fr) 96px !important;
        }
        .growth-large {
            font-size: 18px !important;
        }
    }


    @media (max-width: 760px) {
        /* FILTERS: shorter + tighter */
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"]) {
            gap: 4px 6px !important;
            margin-top: -2px !important;
            margin-bottom: 2px !important;
        }
        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {
            font-size: 8px !important;
            line-height: 1 !important;
            margin-bottom: 0 !important;
            min-height: 10px !important;
        }
        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            height: 34px !important;
            min-height: 34px !important;
            border-radius: 9px !important;
        }
        div[data-testid="stDateInput"] input,
        div[data-testid="stSelectbox"] div[role="button"] {
            font-size: 10px !important;
        }
        .filter-button-spacer { height: 15px !important; }

        /* REFRESH + UPDATE: readable, not black */
        div.stButton > button,
        div[data-testid="stPopover"] button {
            height: 34px !important;
            min-height: 34px !important;
            border-radius: 9px !important;
            font-size: 10px !important;
            padding-top: 0 !important;
            padding-bottom: 0 !important;
            background: #FFFFFF !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
            box-shadow: 0 2px 8px rgba(25,40,70,.035) !important;
        }
        div.stButton > button *,
        div[data-testid="stPopover"] button * {
            color: #25304A !important;
        }

        /* CHART TITLES */
        .st-key-daily_trend_card_mobile .panel-title,
        .st-key-platform_card_mobile .panel-title,
        .st-key-top_products_card_mobile .panel-title {
            margin-bottom: 9px !important;
            color: #17213C !important;
        }

        /* PLOTLY TEXT: darker for readability */
        .st-key-daily_trend_card_mobile .js-plotly-plot text,
        .st-key-platform_card_mobile .js-plotly-plot text {
            fill: #4A556D !important;
        }

        /* TOP 5 */
        .st-key-top_products_card_mobile {
            color: #26324C !important;
        }
        .mobile-top5-list {
            width: 100% !important;
            overflow: visible !important;
        }
    }


    /* V12 MOBILE ONLY — desktop unchanged */
    @media (max-width: 760px) {
        /* 1. Header -> Data per spacing */
        .dashboard-header {
            gap: 2px !important;
            margin-bottom: 5px !important;
        }
        .asof-badge {
            margin-top: -2px !important;
            padding-top: 5px !important;
            padding-bottom: 5px !important;
        }

        /* 2. Date + Platform shorter */
        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {
            font-size: 8px !important;
            line-height: 1 !important;
            margin-bottom: 0 !important;
            min-height: 9px !important;
        }
        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            height: 31px !important;
            min-height: 31px !important;
            border-radius: 8px !important;
        }
        div[data-testid="stDateInput"] input,
        div[data-testid="stSelectbox"] div[role="button"] {
            font-size: 9px !important;
            line-height: 1 !important;
        }

        /* 3. Product + buttons row shorter */
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"]) {
            gap: 2px 5px !important;
            margin-top: -2px !important;
            margin-bottom: 0 !important;
        }
        .filter-button-spacer { height: 12px !important; }

        div.stButton > button,
        div[data-testid="stPopover"] button {
            height: 31px !important;
            min-height: 31px !important;
            border-radius: 8px !important;
            font-size: 9px !important;
            line-height: 1 !important;
            padding: 0 8px !important;
        }

        /* 4. Refresh button white/light + dark icon */
        div.stButton > button,
        div[data-testid="stPopover"] button {
            background: #FFFFFF !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
            box-shadow: 0 2px 7px rgba(25,40,70,.04) !important;
        }
        div.stButton > button *,
        div[data-testid="stPopover"] button * {
            color: #25304A !important;
            fill: #25304A !important;
        }

        /* 5. Section titles 15px */
        .panel-title,
        .insight-title {
            font-size: 15px !important;
            line-height: 1.15 !important;
        }

        /* 6. Sales Target Achievement text tuning */
        .kpi-card .kpi-value.purple {
            font-size: calc(clamp(29px, 8.1vw, 36px) - 1px) !important;
        }
        .kpi-card .subtle {
            font-size: 10px !important;
        }
        .target-stat-label {
            font-size: 9px !important;
        }

        /* 7. Sales MTD value -1px */
        .kpi-card .sales-main .kpi-value {
            font-size: calc(clamp(29px, 8.1vw, 36px) - 1px) !important;
        }

        /* 8. Daily Sales Trend legend bold */
        .st-key-daily_trend_card_mobile .legendtext {
            font-weight: 700 !important;
        }

        /* 9. Top 5 auto-height and more bottom padding */
        .st-key-top_products_card_mobile {
            height: auto !important;
            min-height: 0 !important;
            padding-bottom: 16px !important;
            overflow: visible !important;
        }
        .st-key-top_products_card_mobile > div {
            height: auto !important;
            overflow: visible !important;
        }
        .mobile-top5-list {
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
            padding-bottom: 2px !important;
        }
        .mobile-top5-item:last-child {
            padding-bottom: 2px !important;
        }
    }


    /* V13 MOBILE ONLY — stronger Streamlit secondary-button override */
    @media (max-width: 760px) {
        button[data-testid="baseButton-secondary"],
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"],
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] {
            background: #FFFFFF !important;
            background-color: #FFFFFF !important;
            background-image: none !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
            box-shadow: 0 2px 7px rgba(25,40,70,.04) !important;
        }

        button[data-testid="baseButton-secondary"] *,
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"] *,
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] * {
            color: #25304A !important;
            fill: #25304A !important;
            stroke: #25304A !important;
        }

        button[data-testid="baseButton-secondary"]:hover,
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"]:hover,
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"]:hover {
            background: #F8F6FF !important;
            background-color: #F8F6FF !important;
            border-color: #CFC5F5 !important;
            color: #6538E6 !important;
        }

        button[data-testid="baseButton-secondary"]:hover *,
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"]:hover *,
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"]:hover * {
            color: #6538E6 !important;
            fill: #6538E6 !important;
            stroke: #6538E6 !important;
        }
    }


    /* =====================================================
       V14 MOBILE ONLY — desktop remains unchanged
       ===================================================== */
    @media (max-width: 760px) {

        /* 1. Header -> "Data per" tighter */
        .dashboard-header {
            gap: 0 !important;
            margin-bottom: 3px !important;
        }

        .asof-badge {
            margin-top: -4px !important;
            padding: 4px 8px !important;
            border-radius: 9px !important;
        }

        /* 2. Date + Platform controls shorter */
        div[data-testid="stDateInput"] label,
        div[data-testid="stSelectbox"] label {
            font-size: 8px !important;
            line-height: 1 !important;
            margin-bottom: 0 !important;
            min-height: 8px !important;
        }

        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            height: 29px !important;
            min-height: 29px !important;
            border-radius: 8px !important;
        }

        div[data-testid="stDateInput"] input,
        div[data-testid="stSelectbox"] div[role="button"] {
            font-size: 9px !important;
            line-height: 1 !important;
        }

        /* 3. Product + Refresh + Update row shorter */
        div[data-testid="stHorizontalBlock"]:has([data-testid="stDateInput"]):has([data-testid="stSelectbox"]) {
            gap: 1px 5px !important;
            margin-top: -3px !important;
            margin-bottom: 0 !important;
        }

        .filter-button-spacer {
            height: 10px !important;
        }

        div.stButton > button,
        div[data-testid="stPopover"] button,
        button[data-testid="baseButton-secondary"] {
            height: 29px !important;
            min-height: 29px !important;
            border-radius: 8px !important;
            font-size: 9px !important;
            line-height: 1 !important;
            padding: 0 7px !important;
        }

        /* 4. Refresh + Update must stay light/white */
        button[data-testid="baseButton-secondary"],
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"],
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"],
        div.stButton > button,
        div[data-testid="stPopover"] button {
            background: #FFFFFF !important;
            background-color: #FFFFFF !important;
            background-image: none !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
            box-shadow: 0 2px 7px rgba(25,40,70,.04) !important;
        }

        button[data-testid="baseButton-secondary"] *,
        div[data-testid="stButton"] button[data-testid="baseButton-secondary"] *,
        div[data-testid="stPopover"] button[data-testid="baseButton-secondary"] *,
        div.stButton > button *,
        div[data-testid="stPopover"] button * {
            color: #25304A !important;
            fill: #25304A !important;
            stroke: #25304A !important;
        }

        /* 5. ALL mobile panel/section titles = 15px */
        .panel-title,
        .insight-title,
        .kpi-title {
            font-size: 15px !important;
            line-height: 1.15 !important;
            font-weight: 760 !important;
        }

        /* 6. Sales Target Achievement typography */
        .kpi-card .kpi-value.purple {
            font-size: calc(clamp(29px, 8.1vw, 36px) - 1px) !important;
        }

        /* Achievement vs Monthly Target: +1px from previous mobile 10px */
        .kpi-card .subtle {
            font-size: 11px !important;
        }

        /* Actual MTD + Monthly Target labels: +1px */
        .target-stat-label {
            font-size: 10px !important;
        }

        /* Keep values readable */
        .target-stat-value {
            font-size: 10px !important;
        }

        /* 7. Sales MTD main value -1px */
        .kpi-card .sales-main .kpi-value {
            font-size: calc(clamp(29px, 8.1vw, 36px) - 1px) !important;
        }

        /* 8. Daily Sales Trend legend bold */
        .st-key-daily_trend_card_mobile .legendtext {
            font-weight: 700 !important;
        }

        /* 9 & 10. Top 5 spacing + auto height + bottom breathing room */
        .st-key-top_products_card_mobile {
            height: auto !important;
            min-height: 0 !important;
            padding-top: 14px !important;
            padding-bottom: 20px !important;
            overflow: visible !important;
        }

        .st-key-top_products_card_mobile > div {
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
        }

        .st-key-top_products_card_mobile .panel-title {
            margin-bottom: 10px !important;
        }

        .mobile-top5-list {
            height: auto !important;
            min-height: 0 !important;
            overflow: visible !important;
            padding-top: 0 !important;
            padding-bottom: 4px !important;
        }

        .mobile-top5-item:first-child {
            padding-top: 4px !important;
        }

        .mobile-top5-item:last-child {
            padding-bottom: 10px !important;
        }

        .mobile-product-bar {
            margin-bottom: 0 !important;
        }
    }


    /* =====================================================
       V15 MOBILE ONLY — desktop remains unchanged
       ===================================================== */

    /* Mobile title duplicate is hidden on desktop */
    .sales-mobile-title {
        display: none;
    }

    @media (max-width: 760px) {

        /* Refresh button: target the keyed Streamlit wrapper directly.
           This is intentionally more specific than Streamlit's theme CSS. */
        .st-key-refresh_btn,
        .st-key-refresh_btn > div,
        .st-key-refresh_btn [data-testid="stButton"] {
            background: transparent !important;
        }

        .st-key-refresh_btn button,
        .st-key-refresh_btn button[data-testid="baseButton-secondary"],
        .st-key-refresh_btn [data-testid="stButton"] button {
            background: #FFFFFF !important;
            background-color: #FFFFFF !important;
            background-image: none !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
            box-shadow: 0 2px 7px rgba(25,40,70,.04) !important;
            -webkit-appearance: none !important;
            appearance: none !important;
        }

        .st-key-refresh_btn button *,
        .st-key-refresh_btn button svg,
        .st-key-refresh_btn button span,
        .st-key-refresh_btn button p {
            color: #25304A !important;
            fill: #25304A !important;
            stroke: #25304A !important;
        }

        .st-key-refresh_btn button:hover {
            background: #F8F6FF !important;
            background-color: #F8F6FF !important;
            border-color: #CFC5F5 !important;
        }

        /* SALES - MONTH TO DATE title: retain 15px but force one full-width line */
        .sales-mobile-title {
            display: block !important;
            font-size: 15px !important;
            line-height: 1.15 !important;
            white-space: nowrap !important;
            margin: 0 0 7px 0 !important;
            width: 100% !important;
        }

        .sales-desktop-title {
            display: none !important;
        }

        /* Once the title is outside this row, value + growth can use all remaining width */
        .sales-top-row {
            grid-template-columns: minmax(0, 1fr) auto !important;
            gap: 12px !important;
            align-items: end !important;
        }

        .sales-main {
            min-width: 0 !important;
        }

        .growth-panel {
            min-width: 112px !important;
            max-width: 132px !important;
            padding-bottom: 1px !important;
        }
    }


    /* V16 MOBILE ONLY — rebalance Data per spacing */
    @media (max-width: 760px) {
        .dashboard-header {
            margin-bottom: 10px !important;
        }

        .asof-badge {
            margin-top: -8px !important;
            margin-bottom: 10px !important;
        }
    }


    /* V17 MOBILE ONLY — enlarge Actual MTD, Monthly Target, and Gap values by 1px */
    @media (max-width: 760px) {
        .target-stat:nth-child(1) .target-stat-value,
        .target-stat:nth-child(2) .target-stat-value,
        .target-stat:nth-child(3) .target-stat-value {
            font-size: 11px !important;
        }
    }


    /* V19 Admin Mode */
    @media (max-width: 760px) {
        .st-key-admin_login_btn button,
        .st-key-admin_logout_btn button {
            background: #FFFFFF !important;
            color: #25304A !important;
            border: 1px solid #D9DEE9 !important;
        }

        .st-key-admin_confirm_update button {
            background: #6538E6 !important;
            color: #FFFFFF !important;
            border: 1px solid #6538E6 !important;
        }

        .st-key-admin_confirm_update button * {
            color: #FFFFFF !important;
        }
    }


    /* =====================================================
       V22 — SOFT NEON PREMIUM THEME
       Unified white→grey card fill, unique neon outline by panel.
       ===================================================== */

    :root {
        --card-fill: linear-gradient(145deg, #FFFFFF 0%, #F7F8FB 100%);
        --page-bg: linear-gradient(135deg, #F7F9FC 0%, #EDF1F7 48%, #E9EEF6 100%);
        --neon-pink: #FF4FA8;
        --neon-violet: #9B5CFF;
        --neon-blue: #4EA1FF;
        --neon-cyan: #29D3E3;
        --neon-yellow: #FFD44A;
        --neon-orange: #FF9A5B;
        --neon-lilac: #C39BFF;
    }

    .stApp {
        background:
            radial-gradient(circle at 16% 0%, rgba(154, 92, 255, .055), transparent 30%),
            radial-gradient(circle at 92% 10%, rgba(41, 211, 227, .045), transparent 26%),
            var(--page-bg) !important;
    }

    /* Filters and utility boxes: same fill, restrained lavender neon */
    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stDateInput"] > div > div {
        background: var(--card-fill) !important;
        border: 1px solid rgba(195, 155, 255, .72) !important;
        box-shadow:
            0 0 0 1px rgba(195, 155, 255, .10),
            0 0 12px rgba(155, 92, 255, .12),
            0 5px 16px rgba(31,45,72,.045) !important;
    }

    div.stButton > button,
    div[data-testid="stPopover"] > button,
    button[data-testid="baseButton-secondary"] {
        background: var(--card-fill) !important;
        border: 1px solid rgba(195, 155, 255, .72) !important;
        color: #25304A !important;
        box-shadow:
            0 0 0 1px rgba(195, 155, 255, .08),
            0 0 12px rgba(155, 92, 255, .11),
            0 5px 16px rgba(31,45,72,.04) !important;
    }

    .asof-badge {
        background: var(--card-fill) !important;
        border: 1px solid rgba(195, 155, 255, .78) !important;
        box-shadow:
            0 0 0 1px rgba(195, 155, 255, .09),
            0 0 14px rgba(155, 92, 255, .13),
            0 5px 16px rgba(83,60,155,.05) !important;
    }

    /* Base card fill: all major panels use the same white→grey gradient */
    .card,
    .st-key-daily_trend_card,
    .st-key-platform_card,
    .st-key-top_products_card,
    .st-key-quick_insight_card,
    .st-key-daily_trend_card_mobile,
    .st-key-platform_card_mobile,
    .st-key-top_products_card_mobile {
        background: var(--card-fill) !important;
    }

    /* 1) SALES MTD — pink */
    .sales-kpi-card {
        border: 1px solid rgba(255, 79, 168, .82) !important;
        box-shadow:
            0 0 0 1px rgba(255,79,168,.08),
            0 0 15px rgba(255,79,168,.18),
            0 9px 26px rgba(30,45,75,.06) !important;
    }

    /* 2) TARGET ACHIEVEMENT — violet */
    .target-kpi-card {
        border: 1px solid rgba(155, 92, 255, .82) !important;
        box-shadow:
            0 0 0 1px rgba(155,92,255,.08),
            0 0 15px rgba(155,92,255,.18),
            0 9px 26px rgba(30,45,75,.06) !important;
    }

    /* 3) DAILY SALES TREND — blue */
    .st-key-daily_trend_card,
    .st-key-daily_trend_card_mobile {
        position: relative !important;
        border: 1px solid rgba(78, 161, 255, .85) !important;
        box-shadow:
            0 0 0 1px rgba(78,161,255,.07),
            0 0 15px rgba(78,161,255,.17),
            0 9px 26px rgba(30,45,75,.055) !important;
    }

    /* 4) SALES BY PLATFORM — cyan */
    .st-key-platform_card,
    .st-key-platform_card_mobile {
        position: relative !important;
        border: 1px solid rgba(41, 211, 227, .88) !important;
        box-shadow:
            0 0 0 1px rgba(41,211,227,.07),
            0 0 15px rgba(41,211,227,.18),
            0 9px 26px rgba(30,45,75,.055) !important;
    }

    /* 5) TOP 5 PRODUCTS — yellow */
    .st-key-top_products_card,
    .st-key-top_products_card_mobile {
        position: relative !important;
        border: 1px solid rgba(255, 212, 74, .92) !important;
        box-shadow:
            0 0 0 1px rgba(255,212,74,.08),
            0 0 15px rgba(255,212,74,.19),
            0 9px 26px rgba(30,45,75,.055) !important;
    }

    /* 6) QUICK INSIGHT — orange */
    .st-key-quick_insight_card {
        position: relative !important;
        border: 1px solid rgba(255, 154, 91, .90) !important;
        box-shadow:
            0 0 0 1px rgba(255,154,91,.08),
            0 0 15px rgba(255,154,91,.18),
            0 9px 26px rgba(30,45,75,.055) !important;
    }

    /* Small corner badges like the approved preview */
    .st-key-daily_trend_card::before,
    .st-key-platform_card::before,
    .st-key-top_products_card::before,
    .st-key-quick_insight_card::before {
        position: absolute;
        z-index: 3;
        top: 12px;
        right: 14px;
        width: 34px;
        height: 34px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        background: linear-gradient(145deg, #FFFFFF, #F3F5F9);
        font-size: 17px;
        font-weight: 800;
        line-height: 1;
        box-shadow: 0 4px 12px rgba(31,45,72,.08);
    }

    .st-key-daily_trend_card::before {
        content: "↗";
        color: #315FD7;
        border: 1px solid rgba(78,161,255,.32);
    }

    .st-key-platform_card::before {
        content: "◔";
        color: #20AFC2;
        border: 1px solid rgba(41,211,227,.34);
    }

    .st-key-top_products_card::before {
        content: "★";
        color: #D99F00;
        border: 1px solid rgba(255,212,74,.42);
    }

    .st-key-quick_insight_card::before {
        content: "↗";
        color: #315FD7;
        border: 1px solid rgba(255,154,91,.34);
    }

    /* Keep chart canvases transparent so card gradient remains visible */
    .st-key-daily_trend_card [data-testid="stPlotlyChart"],
    .st-key-platform_card [data-testid="stPlotlyChart"],
    .st-key-top_products_card [data-testid="stPlotlyChart"],
    .st-key-daily_trend_card_mobile [data-testid="stPlotlyChart"],
    .st-key-platform_card_mobile [data-testid="stPlotlyChart"] {
        background: transparent !important;
    }

    /* Quick Insight now uses same neutral card fill */
    .st-key-quick_insight_card .insight-block + .insight-block {
        border-top-color: #D8DEE8 !important;
    }

    /* Slightly stronger card separation while staying clean */
    .sales-kpi-card,
    .target-kpi-card,
    .st-key-daily_trend_card,
    .st-key-platform_card,
    .st-key-top_products_card,
    .st-key-quick_insight_card {
        transition: box-shadow .18s ease, transform .18s ease;
    }

    @media (hover: hover) and (min-width: 761px) {
        .sales-kpi-card:hover,
        .target-kpi-card:hover,
        .st-key-daily_trend_card:hover,
        .st-key-platform_card:hover,
        .st-key-top_products_card:hover,
        .st-key-quick_insight_card:hover {
            transform: translateY(-1px);
        }
    }

    /* Mobile: preserve color identity but reduce glow intensity */
    @media (max-width: 760px) {
        .sales-kpi-card {
            box-shadow: 0 0 10px rgba(255,79,168,.12), 0 6px 20px rgba(30,45,75,.045) !important;
        }
        .target-kpi-card {
            box-shadow: 0 0 10px rgba(155,92,255,.12), 0 6px 20px rgba(30,45,75,.045) !important;
        }
        .st-key-daily_trend_card_mobile {
            box-shadow: 0 0 10px rgba(78,161,255,.12), 0 6px 20px rgba(30,45,75,.045) !important;
        }
        .st-key-platform_card_mobile {
            box-shadow: 0 0 10px rgba(41,211,227,.12), 0 6px 20px rgba(30,45,75,.045) !important;
        }
        .st-key-top_products_card_mobile {
            box-shadow: 0 0 10px rgba(255,212,74,.13), 0 6px 20px rgba(30,45,75,.045) !important;
        }
        .st-key-quick_insight_card {
            box-shadow: 0 0 10px rgba(255,154,91,.12), 0 6px 20px rgba(30,45,75,.045) !important;
        }

        /* Hide decorative desktop corner icons on mobile to save space */
        .st-key-daily_trend_card::before,
        .st-key-platform_card::before,
        .st-key-top_products_card::before,
        .st-key-quick_insight_card::before {
            display: none !important;
        }
    }


    /* V23 — COMPACT NEON DASHBOARD */
    .block-container {
        padding-top: 0.15rem !important;
        padding-bottom: 1.0rem !important;
    }
    .dashboard-header { margin-bottom: 6px !important; }
    .section-gap { height: 8px !important; }

    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stDateInput"] > div > div {
        min-height: 36px !important;
        height: 36px !important;
        border-radius: 11px !important;
    }
    div[data-testid="stSelectbox"] label,
    div[data-testid="stDateInput"] label {
        margin-bottom: 1px !important;
        font-size: 10px !important;
    }
    div.stButton > button,
    div[data-testid="stPopover"] > button,
    button[data-testid="baseButton-secondary"] {
        height: 36px !important;
        min-height: 36px !important;
        border-radius: 10px !important;
    }
    .filter-button-spacer { height: 16px !important; }
    .asof-badge {
        padding: 7px 11px 8px 11px !important;
        min-width: 122px !important;
    }

    .kpi-card {
        min-height: 158px !important;
        height: 158px !important;
        padding: 14px 18px !important;
    }
    .kpi-grid {
        grid-template-columns: 90px minmax(0, 1fr) !important;
        gap: 16px !important;
    }
    .kpi-icon {
        width: 74px !important;
        height: 74px !important;
        border-radius: 18px !important;
        font-size: 30px !important;
    }
    .kpi-title { margin: 0 0 5px 0 !important; }
    .kpi-value { margin: 1px 0 5px 0 !important; }
    .growth-panel { padding-bottom: 3px !important; }
    .growth-label { margin-bottom: 3px !important; }
    .divider { margin: 7px 0 7px 0 !important; }
    .daily-avg-row { align-items: center !important; }
    .progress-track {
        margin: 6px 0 7px 0 !important;
        height: 6px !important;
    }
    .target-stats { margin-top: 2px !important; }

    .st-key-daily_trend_card,
    .st-key-platform_card,
    .st-key-top_products_card {
        padding: 12px 16px 5px 16px !important;
        border-radius: 17px !important;
    }
    .st-key-quick_insight_card {
        padding: 13px 16px !important;
        border-radius: 17px !important;
    }
    .panel-title { margin-bottom: 0 !important; }
    .insight-title { margin-bottom: 10px !important; }
    .insight-block { padding-bottom: 9px !important; }
    .insight-block + .insight-block { padding-top: 9px !important; }

    .st-key-daily_trend_card::before,
    .st-key-platform_card::before,
    .st-key-top_products_card::before,
    .st-key-quick_insight_card::before {
        top: 9px !important;
        right: 11px !important;
        width: 29px !important;
        height: 29px !important;
        border-radius: 9px !important;
        font-size: 14px !important;
    }

    .st-key-daily_trend_card [data-testid="stPlotlyChart"],
    .st-key-platform_card [data-testid="stPlotlyChart"],
    .st-key-top_products_card [data-testid="stPlotlyChart"] {
        margin-top: -6px !important;
        margin-bottom: -7px !important;
    }

    @media (max-width: 760px) {
        .block-container {
            padding: 0.35rem 0.65rem 0.8rem 0.65rem !important;
        }
        .dashboard-header { margin-bottom: 4px !important; }
        .asof-badge {
            padding: 4px 8px !important;
            margin-bottom: 7px !important;
        }

        div[data-testid="stDateInput"] > div > div,
        div[data-testid="stSelectbox"] > div > div {
            height: 28px !important;
            min-height: 28px !important;
        }
        div.stButton > button,
        div[data-testid="stPopover"] button,
        button[data-testid="baseButton-secondary"] {
            height: 28px !important;
            min-height: 28px !important;
        }
        .filter-button-spacer { height: 8px !important; }

        .kpi-card {
            min-height: 0 !important;
            height: auto !important;
            padding: 11px 12px !important;
        }
        .kpi-grid {
            grid-template-columns: 54px minmax(0, 1fr) !important;
            gap: 9px !important;
        }
        .kpi-icon {
            width: 50px !important;
            height: 50px !important;
            border-radius: 14px !important;
            font-size: 22px !important;
        }
        .divider { margin: 6px 0 !important; }

        .st-key-daily_trend_card_mobile,
        .st-key-platform_card_mobile,
        .st-key-top_products_card_mobile {
            padding: 10px 10px 5px 10px !important;
            border-radius: 14px !important;
        }
        .st-key-quick_insight_card {
            padding: 11px 12px !important;
            border-radius: 14px !important;
        }
        .insight-title { margin-bottom: 8px !important; }
        .insight-block { padding-bottom: 7px !important; }
        .insight-block + .insight-block { padding-top: 7px !important; }
    }


    /* =====================================================
       V24 — COMPACT LAYOUT CORRECTIONS
       ===================================================== */
    @media (min-width: 761px) {
        /* Fix clipped Target Achievement values while keeping both KPI cards aligned */
        .kpi-card {
            min-height: 176px !important;
            height: 176px !important;
        }

        /* Ensure the bottom stat values remain fully visible */
        .target-stats {
            margin-top: 3px !important;
        }

        .target-stat-label {
            margin-bottom: 2px !important;
            line-height: 1.15 !important;
        }

        .target-stat-value {
            line-height: 1.2 !important;
            overflow: visible !important;
            text-overflow: clip !important;
        }

        /* Reduce the gap between chart row and Top 5 / Quick Insight row */
        .section-gap {
            height: 3px !important;
        }

        /* Keep Daily Trend and Platform cards visually equal in height */
        .st-key-daily_trend_card,
        .st-key-platform_card {
            min-height: 0 !important;
            height: auto !important;
        }
    }


    /* =====================================================
       V25 — REDUCE GAP BETWEEN CHART ROW AND BOTTOM ROW
       Desktop only. Card heights remain unchanged.
       ===================================================== */
    @media (min-width: 761px) {
        /* The row spacer after Daily Trend / Sales by Platform was still too tall.
           Pull the bottom row upward without altering card dimensions. */
        .section-gap {
            height: 0 !important;
            margin: 0 !important;
        }

        div[data-testid="stHorizontalBlock"]:has(.st-key-top_products_card):has(.st-key-quick_insight_card) {
            margin-top: -8px !important;
        }
    }


    /* V26 — direct bottom-row gap fix */
    @media (min-width: 761px) {
        .st-key-top_products_card,
        .st-key-quick_insight_card {
            margin-top: -26px !important;
        }

        .section-gap {
            height: 0 !important;
            min-height: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
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


def validate_database_bytes(file_bytes):
    if not file_bytes:
        return False, "File kosong.", None
    if len(file_bytes) > 50 * 1024 * 1024:
        return False, "Ukuran file terlalu besar (maks. 50 MB).", None

    try:
        sales_check, target_check = load_workbook(BytesIO(file_bytes))
    except Exception as exc:
        return False, f"Struktur database tidak valid: {exc}", None

    if sales_check.empty:
        return False, "Sheet sales tidak boleh kosong.", None
    if target_check.empty:
        return False, "Sheet Monthly Sales Target tidak boleh kosong.", None

    max_date = sales_check["Tanggal"].max()
    info = {
        "rows": int(len(sales_check)),
        "max_date": max_date.strftime("%d %b %Y") if pd.notna(max_date) else "—",
        "max_date_obj": max_date.date() if pd.notna(max_date) else None,
        "sales_value": float(sales_check["Sales Value"].sum()),
    }
    return True, "Database valid.", info


def github_fetch_database():
    """Fetch the latest master Excel directly from GitHub.

    This avoids relying on the stale local repository copy inside a running
    Streamlit instance.
    """
    required = [
        "GITHUB_OWNER",
        "GITHUB_REPO",
        "GITHUB_BRANCH",
        "DATABASE_PATH",
        "GITHUB_TOKEN",
    ]
    if not all(key in st.secrets for key in required):
        return None

    owner = st.secrets["GITHUB_OWNER"]
    repo = st.secrets["GITHUB_REPO"]
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    path = st.secrets["DATABASE_PATH"]
    token = st.secrets["GITHUB_TOKEN"]

    encoded_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/{encoded_path}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.raw+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "Cache-Control": "no-cache",
    }

    response = requests.get(
        api_url,
        headers=headers,
        params={"ref": branch, "_ts": datetime.now().timestamp()},
        timeout=60,
    )

    if response.status_code != 200:
        raise RuntimeError(
            f"Gagal mengambil database terbaru dari GitHub "
            f"(HTTP {response.status_code})."
        )

    # An .xlsx file is a ZIP container and normally starts with PK.
    if not response.content.startswith(b"PK"):
        raise RuntimeError("Respons GitHub bukan file Excel yang valid.")

    return response.content


def github_update_database(file_bytes):
    owner = st.secrets["GITHUB_OWNER"]
    repo = st.secrets["GITHUB_REPO"]
    branch = st.secrets.get("GITHUB_BRANCH", "main")
    path = st.secrets["DATABASE_PATH"]
    token = st.secrets["GITHUB_TOKEN"]

    encoded_path = quote(path, safe="/")
    api_url = f"https://api.github.com/repos/{owner}/{repo}/contents/{encoded_path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }

    current = requests.get(api_url, headers=headers, params={"ref": branch}, timeout=30)
    if current.status_code != 200:
        raise RuntimeError(f"Gagal membaca file master dari GitHub (HTTP {current.status_code}).")

    sha = current.json().get("sha")
    if not sha:
        raise RuntimeError("GitHub tidak mengembalikan SHA file database.")

    payload = {
        "message": "Update e-commerce dashboard database via Admin Mode",
        "content": base64.b64encode(file_bytes).decode("ascii"),
        "sha": sha,
        "branch": branch,
    }

    updated = requests.put(api_url, headers=headers, json=payload, timeout=60)
    if updated.status_code not in (200, 201):
        try:
            detail = updated.json().get("message", "")
        except Exception:
            detail = ""
        raise RuntimeError(
            f"Gagal update database ke GitHub (HTTP {updated.status_code})"
            + (f": {detail}" if detail else ".")
        )

    return updated.json()

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

# Always try the GitHub master first. The local Excel bundled with the
# Streamlit deployment is only a fallback if GitHub is temporarily unavailable.
try:
    github_database_bytes = github_fetch_database()
except Exception as github_exc:
    github_database_bytes = None
    github_fetch_error = str(github_exc)
else:
    github_fetch_error = None

try:
    if github_database_bytes is not None:
        sales, targets = load_workbook(BytesIO(github_database_bytes))
        database_source = "GitHub master"
    else:
        if not default_path.exists():
            raise FileNotFoundError(f"Database default **{DB_FILE}** tidak ditemukan.")
        sales, targets = load_workbook(default_path)
        database_source = "Local fallback"
except Exception as e:
    st.error(f"Gagal membaca database: {e}")
    st.stop()

if sales.empty:
    st.warning("Database sales kosong.")
    st.stop()

data_max_date = sales["Tanggal"].max().date()
data_min_date = sales["Tanggal"].min().date()

if st.session_state.pop("admin_update_success", False):
    st.toast("Database berhasil di-update dan dashboard sudah membaca data terbaru.", icon="✅")

# =========================================================
# HEADER PLACEHOLDER
# =========================================================
header_slot = st.empty()

# =========================================================
# FILTERS
# =========================================================
f1, f2, f3, f4, f5 = st.columns([1.0, 1.0, 1.45, .28, .42], gap="small")

with f1:
    if "force_latest_date" in st.session_state:
        st.session_state["date_filter"] = st.session_state.pop("force_latest_date")
    elif "date_filter" not in st.session_state:
        st.session_state["date_filter"] = data_max_date

    # Keep the stored filter inside the available database range.
    if st.session_state["date_filter"] < data_min_date:
        st.session_state["date_filter"] = data_min_date
    if st.session_state["date_filter"] > data_max_date:
        st.session_state["date_filter"] = data_max_date

    selected_date = st.date_input(
        "Date",
        min_value=data_min_date,
        max_value=data_max_date,
        key="date_filter",
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
    reset = st.button("↻", help="Reset filters", use_container_width=True, key="refresh_btn")
    if reset:
        st.rerun()

with f5:
    st.markdown('<div class="filter-button-spacer"></div>', unsafe_allow_html=True)

    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False

    admin_label = "🔒 Admin" if not st.session_state.admin_authenticated else "✓ Admin"

    with st.popover(admin_label, use_container_width=True):
        if not st.session_state.admin_authenticated:
            st.markdown("**Admin Mode**")
            admin_password_input = st.text_input(
                "Password",
                type="password",
                key="admin_password_input",
                placeholder="Admin password",
            )

            if st.button("Login", key="admin_login_btn", use_container_width=True):
                expected_password = st.secrets.get("ADMIN_PASSWORD", "")
                if expected_password and admin_password_input == expected_password:
                    st.session_state.admin_authenticated = True
                    st.session_state.pop("admin_password_input", None)
                    st.rerun()
                else:
                    st.error("Password salah.")
        else:
            st.markdown("**Update Database**")
            st.caption("File akan divalidasi sebelum mengganti database master di GitHub.")

            admin_upload = st.file_uploader(
                "Upload Excel terbaru",
                type=["xlsx"],
                key="admin_database_upload",
                label_visibility="collapsed",
                help="Harus mempertahankan sheet DB Penjualan Produk 2026 dan Monthly Sales Target.",
            )

            if admin_upload is not None:
                upload_bytes = admin_upload.getvalue()
                valid, validation_message, validation_info = validate_database_bytes(upload_bytes)

                if valid:
                    st.success(validation_message)
                    st.caption(
                        f"Rows: {validation_info['rows']:,} · "
                        f"Data terakhir: {validation_info['max_date']} · "
                        f"Total Sales Value: {rp_jt(validation_info['sales_value'])}"
                    )
                    st.warning("Confirm Update akan mengganti database master di GitHub.")

                    if st.button(
                        "Confirm Update",
                        key="admin_confirm_update",
                        type="primary",
                        use_container_width=True,
                    ):
                        try:
                            with st.spinner("Updating master database..."):
                                github_update_database(upload_bytes)
                            st.cache_data.clear()
                            if validation_info.get("max_date_obj") is not None:
                                st.session_state["force_latest_date"] = validation_info["max_date_obj"]
                            st.session_state["admin_update_success"] = True
                            st.session_state.pop("admin_database_upload", None)
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Update gagal: {exc}")
                else:
                    st.error(validation_message)

            if st.button("Logout Admin", key="admin_logout_btn", use_container_width=True):
                st.session_state.admin_authenticated = False
                st.session_state.pop("admin_database_upload", None)
                st.rerun()

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
          <span class="title-main">DAILY SALES MONITORING</span>
          <span class="accent title-accent">E-COMMERCE</span>
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
        <div class="card kpi-card sales-kpi-card">
          <div class="kpi-grid">
            <div class="kpi-icon sales">◉</div>
            <div>
              <div class="kpi-title sales-mobile-title">SALES - MONTH TO DATE</div>
              <div class="sales-top-row">
                <div class="sales-main">
                  <div class="kpi-title sales-desktop-title">SALES - MONTH TO DATE</div>
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
        <div class="card kpi-card target-kpi-card">
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

        fig.update_layout(**plot_layout(215))
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

    # Mobile-only Daily Sales Trend (desktop card above is hidden by CSS on mobile)
    with st.container(key="daily_trend_card_mobile", border=False):
        st.markdown(
            '<div class="panel-title">DAILY SALES TREND <span style="font-size:9px;color:#929AAD;font-weight:500">(Rp jt)</span></div>',
            unsafe_allow_html=True,
        )
        fig_mobile = go.Figure()
        fig_mobile.add_trace(
            go.Scatter(
                x=curr_daily["Day"],
                y=curr_daily["Sales Value"],
                mode="lines+markers",
                name="<b>This Month</b>",
                line=dict(color=PURPLE, width=2.6),
                marker=dict(size=4, color="#FFFFFF", line=dict(color=PURPLE, width=1.5)),
                connectgaps=False,
                hovertemplate="Day %{x}<br>Rp %{y:,.1f} jt<extra></extra>",
            )
        )
        fig_mobile.add_trace(
            go.Scatter(
                x=prev_full["Day"],
                y=prev_full["Sales Value"],
                mode="lines",
                name="<b>Last Month</b>",
                line=dict(color=PURPLE_LIGHT, width=2.0, dash="dash"),
                hovertemplate="Day %{x}<br>Rp %{y:,.1f} jt<extra></extra>",
            )
        )
        mobile_tickvals = [d for d in [1, 5, 10, 15, 20, 25, 31] if d <= max(current_days["Day"].max(), prev_days["Day"].max())]
        fig_mobile.update_layout(
            height=175,
            margin=dict(l=8, r=6, t=52, b=24),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter, sans-serif", color="#4A556D", size=9),
            hoverlabel=dict(bgcolor="#FFFFFF", font_color=NAVY, bordercolor="#E0E3EB"),
            legend=dict(
                orientation="h",
                x=0,
                y=1.06,
                xanchor="left",
                yanchor="bottom",
                font=dict(size=9, color="#4A556D"),
                bgcolor="rgba(0,0,0,0)",
            ),
        )
        fig_mobile.update_xaxes(
            title=None,
            tickmode="array",
            tickvals=mobile_tickvals,
            ticktext=[str(v) for v in mobile_tickvals],
            range=[0.5, max(current_days["Day"].max(), prev_days["Day"].max()) + 0.5],
            showgrid=False,
            zeroline=False,
            color="#4A556D",
            tickfont=dict(size=8, color="#4A556D"),
        )
        fig_mobile.update_yaxes(
            title=None,
            showgrid=True,
            gridcolor=GRID,
            griddash="dot",
            zeroline=False,
            color="#4A556D",
            tickfont=dict(size=8, color="#4A556D"),
            nticks=5,
        )
        st.plotly_chart(fig_mobile, use_container_width=True, config={"displayModeBar": False})

# =========================================================
# SALES BY PLATFORM
# =========================================================
with c2:
    with st.container(key="platform_card", border=False):
        st.markdown(
            '<div class="panel-title">SALES BY PLATFORM <span style="font-size:11px;color:#929AAD;font-weight:500">(Share of MTD Sales)</span></div>',
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
            go.Pie(
                labels=by_platform["Platform Group"],
                values=by_platform["Sales Value"],
                hole=0.58,
                sort=False,
                direction="clockwise",
                marker=dict(colors=[PLATFORM_COLORS[p] for p in by_platform["Platform Group"]]),
                textinfo="percent",
                textposition="inside",
                textfont=dict(size=12, color="#FFFFFF", family="Inter, sans-serif"),
                hovertemplate="<b>%{label}</b><br>Rp %{value:,.0f}<br>%{percent}<extra></extra>",
            )
        )
        fig2.update_layout(
            height=215,
            margin=dict(l=8, r=8, t=4, b=4),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter, sans-serif", color=TEXT, size=10),
            showlegend=True,
            legend=dict(
                orientation="v",
                x=1.02,
                y=0.5,
                xanchor="left",
                yanchor="middle",
                font=dict(size=10, color="#4A556D"),
            ),
            annotations=[
                dict(
                    text=f"<b>{rp_jt(total_platform)}</b><br><span style='font-size:10px'>MTD Sales</span>",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    align="center",
                    font=dict(size=14, color=NAVY, family="Inter, sans-serif"),
                )
            ],
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    # Mobile-only Sales by Platform
    with st.container(key="platform_card_mobile", border=False):
        st.markdown(
            '<div class="panel-title">SALES BY PLATFORM <span style="font-size:9px;color:#929AAD;font-weight:500">(Share of MTD Sales)</span></div>',
            unsafe_allow_html=True,
        )
        fig2_mobile = go.Figure(
            go.Pie(
                labels=by_platform["Platform Group"],
                values=by_platform["Sales Value"],
                hole=0.60,
                sort=False,
                direction="clockwise",
                marker=dict(colors=[PLATFORM_COLORS[p] for p in by_platform["Platform Group"]]),
                textinfo="percent",
                textposition="inside",
                textfont=dict(size=10, color="#FFFFFF", family="Inter, sans-serif"),
                hovertemplate="<b>%{label}</b><br>Rp %{value:,.0f}<br>%{percent}<extra></extra>",
            )
        )
        fig2_mobile.update_layout(
            height=185,
            margin=dict(l=4, r=4, t=6, b=44),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter, sans-serif", color="#4A556D", size=9),
            showlegend=True,
            legend=dict(
                orientation="h",
                x=0.5,
                y=-0.08,
                xanchor="center",
                yanchor="top",
                font=dict(size=9, color="#4A556D"),
            ),
            annotations=[
                dict(
                    text=f"<b>{rp_jt(total_platform)}</b><br><span style='font-size:9px'>MTD Sales</span>",
                    x=0.5,
                    y=0.53,
                    showarrow=False,
                    align="center",
                    font=dict(size=12, color=NAVY, family="Inter, sans-serif"),
                )
            ],
        )
        st.plotly_chart(fig2_mobile, use_container_width=True, config={"displayModeBar": False})

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
        fig3.update_layout(**plot_layout(165))
        fig3.update_layout(
            showlegend=False,
            bargap=.48,
            margin=dict(l=6, r=110, t=8, b=6),
        )
        fig3.update_xaxes(showgrid=True, gridcolor=GRID, griddash="dot", zeroline=False, color="#6F7890")
        fig3.update_yaxes(showgrid=False, color="#34415C", tickfont=dict(size=10))
        st.plotly_chart(fig3, use_container_width=True, config={"displayModeBar": False})

    # Mobile-only Top 5 list: readable product names, compact bars
    with st.container(key="top_products_card_mobile", border=False):
        st.markdown(
            '<div class="panel-title">TOP 5 PRODUCTS <span style="font-size:9px;color:#929AAD;font-weight:500">(by Sales Value)</span></div>',
            unsafe_allow_html=True,
        )
        top5_mobile = top5.sort_values("Sales Value", ascending=False).reset_index(drop=True)
        max_top5_value = float(top5_mobile["Sales Value"].max()) if not top5_mobile.empty else 1.0
        mobile_colors = ["#EF347D", "#8A5AE2", "#4E85EB", "#4BC675", "#F4B51F"]
        mobile_items = []
        for idx, row in top5_mobile.iterrows():
            width_pct = (float(row["Sales Value"]) / max_top5_value * 100) if max_top5_value else 0
            product_name = html.escape(str(row["Product"]))
            value_text = html.escape(f"{rp_jt(row['Sales Value'])} · {row['Share']:.1f}%")
            color = mobile_colors[idx % len(mobile_colors)]
            item_html = (
                f'<div class="mobile-top5-item">'
                f'<div class="mobile-top5-head">'
                f'<span class="mobile-rank" style="background:{color}">{idx+1}</span>'
                f'<span class="mobile-product-name">{product_name}</span>'
                f'<span class="mobile-product-value">{value_text}</span>'
                f'</div>'
                f'<div class="mobile-product-bar">'
                f'<span style="width:{width_pct:.1f}%;background:{color}"></span>'
                f'</div>'
                f'</div>'
            )
            mobile_items.append(item_html)
        st.markdown('<div class="mobile-top5-list">' + ''.join(mobile_items) + '</div>', unsafe_allow_html=True)

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
