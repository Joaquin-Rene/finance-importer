from __future__ import annotations

from datetime import datetime
from html import escape
from pathlib import Path
from typing import Callable

import altair as alt
import pandas as pd
import streamlit as st

from src.finanzas_importer.analytics import AlertItem, MonthlyKpis
from src.finanzas_importer.mp_parser import FILTER_REASON_RENDIMIENTOS, FILTER_REASON_SELF_TRANSFER, ParseResult
from src.finanzas_importer.workbook_writer import DATE_FILTER_MODE_AFTER_MAX, DATE_FILTER_MODE_SKIP_EXISTING, ImportPlan


def inject_styles() -> None:
    st.markdown(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Space+Grotesk:wght@500;700&display=swap');

            :root {
                --bg-main: #0b1020;
                --surface: rgba(18, 26, 43, 0.9);
                --surface-strong: rgba(26, 37, 64, 0.94);
                --surface-soft: rgba(20, 29, 49, 0.84);
                --line: rgba(36, 50, 74, 0.95);
                --text-main: #f8fafc;
                --text-soft: #94a3b8;
                --brand: #6d5ef5;
                --brand-2: #22d3ee;
                --success: #22c55e;
                --warning: #f59e0b;
                --danger: #ef4444;
                --radius-lg: 28px;
                --radius-md: 20px;
                --radius-sm: 14px;
            }

            html, body, [class*="css"] {
                font-family: "Manrope", "Segoe UI", sans-serif;
            }

            .stApp {
                color: var(--text-main);
                background:
                    radial-gradient(circle at top left, rgba(109, 94, 245, 0.2), rgba(11, 16, 32, 0) 26%),
                    radial-gradient(circle at top right, rgba(34, 211, 238, 0.18), rgba(11, 16, 32, 0) 22%),
                    linear-gradient(180deg, #0b1020 0%, #0f1527 45%, #0d1324 100%);
            }

            .block-container {
                max-width: 1320px;
                padding-top: 2rem;
                padding-bottom: 3rem;
            }

            h1, h2, h3, h4 {
                font-family: "Space Grotesk", "Segoe UI", sans-serif;
                letter-spacing: -0.03em;
                color: var(--text-main);
            }

            [data-testid="stSidebar"][aria-expanded="true"] {
                background: linear-gradient(180deg, rgba(10, 15, 28, 0.98) 0%, rgba(13, 19, 36, 0.98) 100%);
                border-right: 1px solid rgba(36, 50, 74, 0.9);
                min-width: 360px !important;
                max-width: 360px !important;
            }

            [data-testid="stSidebar"][aria-expanded="true"] > div:first-child {
                width: 360px !important;
            }

            [data-testid="stSidebar"][aria-expanded="false"] {
                min-width: 0 !important;
                max-width: 0 !important;
            }

            [data-testid="stSidebar"][aria-expanded="false"] > div:first-child {
                width: 0 !important;
                min-width: 0 !important;
            }

            [data-testid="stSidebar"] .stTextInput > div > div,
            [data-testid="stSidebar"] .stTextArea > div > div,
            [data-testid="stSidebar"] .stSelectbox > div > div,
            [data-testid="stSidebar"] .stSlider,
            [data-testid="stSidebar"] .stToggle,
            [data-testid="stSidebar"] .stRadio {
                background: rgba(18, 26, 43, 0.96);
                border: 1px solid rgba(36, 50, 74, 0.95);
                border-radius: 18px;
                padding: 0.15rem 0.45rem;
            }

            [data-testid="stSidebar"] label,
            [data-testid="stSidebar"] p,
            [data-testid="stSidebar"] div,
            [data-testid="stSidebar"] span {
                color: var(--text-main);
            }

            .sidebar-title {
                font-family: "Space Grotesk", "Segoe UI", sans-serif;
                font-size: 1.08rem;
                font-weight: 700;
                margin-bottom: 0.9rem;
            }

            .sidebar-section {
                margin-top: 1.1rem;
                margin-bottom: 0.4rem;
                font-size: 0.74rem;
                letter-spacing: 0.14em;
                text-transform: uppercase;
                color: var(--text-soft);
                font-weight: 700;
            }

            .hero {
                position: relative;
                overflow: hidden;
                padding: 2.25rem;
                margin-bottom: 1.15rem;
                border-radius: var(--radius-lg);
                border: 1px solid rgba(57, 73, 110, 0.65);
                color: #f4fbf8;
                background:
                    radial-gradient(circle at 80% 20%, rgba(34, 211, 238, 0.22), rgba(34, 211, 238, 0) 28%),
                    radial-gradient(circle at 18% 18%, rgba(109, 94, 245, 0.25), rgba(109, 94, 245, 0) 30%),
                    linear-gradient(125deg, #121a2b 0%, #10172a 48%, #0b1020 100%);
                box-shadow: 0 28px 80px rgba(2, 6, 23, 0.45);
            }

            .hero-grid {
                display: grid;
                grid-template-columns: minmax(0, 1.5fr) minmax(260px, 0.8fr);
                gap: 1.2rem;
                position: relative;
                z-index: 1;
            }

            .hero-kicker {
                display: inline-block;
                margin-bottom: 0.8rem;
                padding: 0.35rem 0.7rem;
                border-radius: 999px;
                background: rgba(109, 94, 245, 0.14);
                border: 1px solid rgba(109, 94, 245, 0.22);
                font-size: 0.73rem;
                font-weight: 700;
                letter-spacing: 0.1em;
                text-transform: uppercase;
            }

            .hero h1 {
                margin: 0;
                font-size: clamp(2rem, 3vw, 3.3rem);
                line-height: 0.98;
                max-width: 12ch;
                color: #f7fffc;
            }

            .hero-copy {
                margin-top: 0.95rem;
                max-width: 48ch;
                color: rgba(226, 232, 240, 0.92);
                font-size: 0.98rem;
                line-height: 1.55;
            }

            .hero-list {
                display: flex;
                flex-wrap: wrap;
                gap: 0.55rem;
                margin-top: 1rem;
            }

            .hero-chip {
                padding: 0.46rem 0.82rem;
                border-radius: 999px;
                background: rgba(255, 255, 255, 0.06);
                border: 1px solid rgba(148, 163, 184, 0.18);
                font-size: 0.82rem;
                color: #e2e8f0;
            }

            .hero-panel {
                align-self: end;
                padding: 1rem;
                border-radius: 22px;
                background: linear-gradient(180deg, rgba(26, 37, 64, 0.82) 0%, rgba(18, 26, 43, 0.78) 100%);
                border: 1px solid rgba(71, 85, 105, 0.28);
                backdrop-filter: blur(8px);
            }

            .hero-panel-label {
                font-size: 0.74rem;
                letter-spacing: 0.15em;
                text-transform: uppercase;
                color: rgba(148, 163, 184, 0.88);
                margin-bottom: 0.6rem;
            }

            .hero-panel-status {
                font-size: 1.05rem;
                font-weight: 700;
                color: #ffffff;
                margin-bottom: 0.45rem;
            }

            .hero-panel-copy {
                color: rgba(226, 232, 240, 0.86);
                line-height: 1.55;
                font-size: 0.92rem;
            }

            .status-pill {
                display: inline-flex;
                align-items: center;
                padding: 0.5rem 0.75rem;
                margin-top: 0.8rem;
                border-radius: 999px;
                font-weight: 700;
                font-size: 0.78rem;
                border: 1px solid transparent;
            }

            .status-ok {
                background: rgba(34, 197, 94, 0.12);
                border-color: rgba(34, 197, 94, 0.35);
                color: #dcfce7;
            }

            .status-bad {
                background: rgba(239, 68, 68, 0.12);
                border-color: rgba(239, 68, 68, 0.35);
                color: #fee2e2;
            }

            .subtle {
                color: var(--text-soft);
                font-size: 0.93rem;
                line-height: 1.55;
            }

            [data-testid="stTextInputRootElement"] input,
            [data-testid="stTextArea"] textarea,
            [data-testid="stNumberInput"] input,
            textarea {
                color: var(--text-main) !important;
            }

            [data-baseweb="input"] > div,
            [data-baseweb="base-input"] > div,
            [data-baseweb="select"] > div {
                background: rgba(18, 26, 43, 0.96) !important;
                border: 1px solid rgba(36, 50, 74, 0.95) !important;
                border-radius: 14px !important;
            }

            [data-baseweb="input"] input,
            [data-baseweb="base-input"] input,
            [data-baseweb="textarea"] textarea,
            [data-baseweb="select"] > div,
            [data-baseweb="select"] span {
                color: var(--text-main) !important;
            }

            [data-testid="stSidebar"] [data-baseweb="select"] > div {
                min-height: 46px !important;
                padding-top: 0 !important;
                padding-bottom: 0 !important;
                display: flex !important;
                align-items: center !important;
            }

            [data-testid="stSidebar"] [data-baseweb="select"] input,
            [data-testid="stSidebar"] [data-baseweb="select"] span,
            [data-testid="stSidebar"] [data-baseweb="select"] div {
                line-height: 1.25 !important;
            }

            [data-testid="stSidebar"] [data-baseweb="textarea"] textarea {
                line-height: 1.35 !important;
            }

            [data-testid="stTextInputRootElement"] input::placeholder,
            textarea::placeholder {
                color: rgba(148, 163, 184, 0.8) !important;
            }

            [data-testid="stFileUploader"] > div {
                background: linear-gradient(180deg, rgba(18, 26, 43, 0.92) 0%, rgba(15, 22, 39, 0.92) 100%);
                border: 1px solid rgba(36, 50, 74, 0.95);
                border-radius: 24px;
                padding: 0.45rem;
                box-shadow: 0 18px 40px rgba(2, 6, 23, 0.35);
            }

            [data-testid="stFileUploader"] section {
                border: 2px dashed rgba(34, 211, 238, 0.38);
                border-radius: 18px;
                background: rgba(26, 37, 64, 0.78);
                padding: 0.8rem;
            }

            [data-testid="stFileUploader"] small {
                color: var(--text-soft);
            }

            [data-testid="stFileUploader"] section [data-testid="stMarkdownContainer"] p,
            [data-testid="stFileUploader"] section span,
            [data-testid="stFileUploader"] section small {
                color: #dbe7f5 !important;
                font-weight: 500 !important;
            }

            [data-testid="stFileUploader"] button {
                background: rgba(255, 255, 255, 0.06) !important;
                border: 1px solid rgba(148, 163, 184, 0.3) !important;
                color: #f8fafc !important;
            }

            [data-testid="stFileUploaderFile"] {
                align-items: flex-start !important;
                gap: 0.45rem;
                padding: 0.45rem 0;
            }

            [data-testid="stFileUploaderFile"] .stFileUploaderFileData {
                display: flex !important;
                flex-direction: column !important;
                align-items: flex-start !important;
                gap: 0.15rem;
                min-width: 0;
                padding-left: 0.3rem !important;
            }

            [data-testid="stFileUploaderFileName"] {
                width: 100%;
                margin-right: 0 !important;
                margin-bottom: 0 !important;
                color: #f8fafc !important;
                font-weight: 700 !important;
                line-height: 1.3 !important;
            }

            [data-testid="stFileUploaderFile"] small,
            [data-testid="stFileUploaderFile"] [data-testid="stFileUploaderFileErrorMessage"] {
                color: #dbe7f5 !important;
                line-height: 1.25 !important;
            }

            [data-testid="stFileUploaderFile"] svg,
            [data-testid="stFileUploaderDeleteBtn"] button,
            [data-testid="stFileUploaderDeleteBtn"] button span {
                color: #f8fafc !important;
            }

            [data-testid="stMetric"] {
                background: linear-gradient(180deg, rgba(18, 26, 43, 0.95) 0%, rgba(15, 22, 39, 0.98) 100%);
                border: 1px solid var(--line);
                border-radius: 20px;
                padding: 1rem 1rem 0.8rem;
                box-shadow: 0 16px 36px rgba(2, 6, 23, 0.32);
                min-height: 140px;
            }

            [data-testid="stMetricLabel"] {
                color: var(--text-soft);
                font-size: 0.78rem;
                letter-spacing: 0.02em;
                text-transform: uppercase;
                font-weight: 700;
            }

            [data-testid="stMetricValue"] {
                font-family: "Space Grotesk", "Segoe UI", sans-serif;
                color: var(--text-main);
                font-size: 2rem;
                line-height: 1.05;
            }

            [data-testid="stMetricDelta"] {
                color: var(--brand-2);
                font-weight: 700;
            }

            [data-testid="stVegaLiteChart"] {
                background: linear-gradient(180deg, rgba(18, 26, 43, 0.95) 0%, rgba(15, 22, 39, 0.98) 100%);
                border: 1px solid var(--line);
                border-radius: 22px;
                padding: 0.5rem 0.7rem;
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.28);
            }

            [data-testid="stVegaLiteChart"] > div,
            .vega-embed,
            .vega-embed details,
            .vega-embed summary {
                background: transparent !important;
            }

            [data-testid="stDataFrame"] {
                border-radius: 22px;
                overflow: hidden;
                border: 1px solid rgba(36, 50, 74, 0.95);
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.35);
            }

            [data-testid="stDataEditor"] {
                border-radius: 22px;
                overflow: hidden;
                border: 1px solid rgba(36, 50, 74, 0.95);
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.35);
                background: linear-gradient(180deg, rgba(18, 26, 43, 0.98) 0%, rgba(15, 22, 39, 0.98) 100%) !important;
            }

            [data-testid="stDataEditor"],
            [data-testid="stDataEditor"] *,
            [data-testid="stDataEditor"] [data-testid="stDataFrameResizable"],
            [data-testid="stDataEditor"] [data-testid="stDataFrameGlideDataEditor"] {
                --gdg-bg-cell: #0f172a;
                --gdg-bg-cell-medium: #162033;
                --gdg-bg-header: #1e293b;
                --gdg-bg-header-has-focus: #243247;
                --gdg-bg-search-result: rgba(34, 211, 238, 0.18);
                --gdg-border-color: rgba(51, 65, 85, 0.78);
                --gdg-header-font-style: 700 12px Manrope;
                --gdg-font-style: 500 13px Manrope;
                --gdg-text-dark: #e2e8f0;
                --gdg-text-medium: #cbd5e1;
                --gdg-text-light: #94a3b8;
                --gdg-accent-color: #22d3ee;
            }

            [data-testid="stDataEditor"] input,
            [data-testid="stDataEditor"] textarea,
            [data-testid="stDataEditor"] div {
                color: #e2e8f0 !important;
            }

            [data-testid="stDataEditor"] > div,
            [data-testid="stDataEditor"] section,
            [data-testid="stDataEditor"] canvas,
            [data-testid="stDataEditor"] [role="grid"],
            [data-testid="stDataEditor"] [data-testid="stDataFrameResizable"],
            [data-testid="stDataEditor"] [data-testid="stDataFrameGlideDataEditor"],
            [data-testid="stDataEditor"] [data-testid="stDataFrameGlideDataEditor"] > div {
                background: #0f172a !important;
            }

            [data-testid="stDataEditor"] canvas {
                border-radius: 0 !important;
            }

            .preview-table-wrap {
                border-radius: 22px;
                overflow: hidden;
                border: 1.5px solid rgba(241, 245, 249, 0.85);
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.35);
                background: linear-gradient(180deg, rgba(18,26,43,0.98) 0%, rgba(15,22,39,0.98) 100%);
                margin-top: 0.55rem;
            }

            .preview-table {
                width: 100%;
                border-collapse: collapse;
            }

            .preview-table thead th {
                padding: 0.82rem 0.9rem;
                text-align: left;
                background: rgba(30, 41, 59, 0.96);
                color: #cbd5e1;
                font-size: 0.78rem;
                letter-spacing: 0.05em;
                text-transform: uppercase;
                border-bottom: 1px solid rgba(51, 65, 85, 0.9);
            }

            .preview-table tbody td {
                padding: 0.78rem 0.9rem;
                color: #e2e8f0;
                border-top: 1px solid rgba(36, 50, 74, 0.68);
                vertical-align: top;
                font-size: 0.92rem;
            }

            .preview-table tbody tr:hover {
                background: rgba(30, 41, 59, 0.52);
            }

            .tipo-chip {
                display: inline-flex;
                align-items: center;
                gap: 0.42rem;
                font-weight: 700;
            }

            .tipo-dot {
                font-size: 0.9rem;
                line-height: 1;
            }

            .tipo-ingreso {
                color: #86efac;
            }

            .tipo-gasto {
                color: #fda4af;
            }

            .compartido-flag {
                font-weight: 700;
                color: #f8fafc;
            }

            .manual-editor-kicker {
                display: inline-flex;
                align-items: center;
                padding: 0.38rem 0.72rem;
                margin-bottom: 0.55rem;
                border-radius: 999px;
                background: rgba(34, 211, 238, 0.12);
                border: 1px solid rgba(34, 211, 238, 0.24);
                color: #8be9f8;
                font-size: 0.75rem;
                font-weight: 800;
                letter-spacing: 0.12em;
                text-transform: uppercase;
            }

            .manual-editor-intro {
                margin-bottom: 0.9rem;
                color: #cbd5e1;
                font-size: 0.98rem;
                line-height: 1.45;
            }

            .manual-editor-head {
                display: flex;
                align-items: center;
                justify-content: center;
                min-height: 48px;
                padding: 0.95rem 1rem;
                color: #f8fafc;
                font-family: "Space Grotesk", "Segoe UI", sans-serif;
                font-size: 1.02rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-align: center;
                border: 1px solid rgba(71, 85, 105, 0.55);
                border-radius: 16px;
                background: linear-gradient(180deg, rgba(30, 41, 59, 0.96) 0%, rgba(22, 32, 51, 0.96) 100%);
                box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.03);
                margin-bottom: 0.45rem;
            }

            .manual-editor-row {
                padding: 0.35rem 0 0.15rem;
                border-top: 1px solid rgba(51, 65, 85, 0.45);
            }

            .manual-type-chip {
                display: inline-flex;
                align-items: center;
                gap: 0.46rem;
                min-height: 52px;
                padding: 0.2rem 0.35rem;
                font-size: 0.96rem;
                font-weight: 800;
            }

            .manual-type-ingreso {
                color: #86efac;
            }

            .manual-type-gasto {
                color: #fda4af;
            }

            .vg-tooltip {
                background: rgba(15, 23, 42, 0.96) !important;
                border: 1px solid rgba(51, 65, 85, 0.9) !important;
                border-radius: 14px !important;
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.35) !important;
                color: #e2e8f0 !important;
                padding: 0.55rem 0.7rem !important;
            }

            .vg-tooltip table {
                border-collapse: collapse !important;
            }

            .vg-tooltip td,
            .vg-tooltip th {
                border: none !important;
                padding: 0.2rem 0.35rem !important;
                text-align: left !important;
            }

            .vg-tooltip td.key {
                color: #94a3b8 !important;
                font-weight: 600 !important;
            }

            .vg-tooltip td.value {
                color: #f8fafc !important;
                font-weight: 800 !important;
            }

            [data-testid="stMarkdownContainer"] code {
                background: rgba(109, 94, 245, 0.14);
                color: #ddd6fe;
                border-radius: 8px;
                padding: 0.15rem 0.4rem;
            }

            [data-testid="stButton"] button,
            .stDownloadButton button {
                border-radius: 14px;
                border: 1px solid rgba(36, 50, 74, 0.95);
                background: rgba(18, 26, 43, 0.95);
                color: var(--text-main);
                font-weight: 700;
                box-shadow: 0 14px 26px rgba(2, 6, 23, 0.28);
            }

            [data-testid="stButton"] button[kind="primary"] {
                background: linear-gradient(135deg, #6d5ef5 0%, #22d3ee 100%);
                color: #f8fffc;
                border-color: transparent;
            }

            [data-testid="stAlert"] {
                border-radius: 18px;
                border: 1px solid rgba(36, 50, 74, 0.95);
                box-shadow: 0 16px 30px rgba(2, 6, 23, 0.28);
            }

            [data-testid="stExpander"] {
                border-radius: 18px;
                border: 1px solid var(--line);
                background: rgba(18, 26, 43, 0.78);
                overflow: hidden;
            }

            [data-testid="stSidebar"] div[role="radiogroup"] > label {
                display: flex;
                align-items: flex-start;
                gap: 0.55rem;
                width: 100%;
                margin-bottom: 0.45rem;
                border-radius: 18px;
                padding: 0.65rem 0.8rem;
            }

            [data-testid="stSegmentedControl"] {
                background: transparent !important;
            }

            [data-testid="stSegmentedControl"] button,
            [data-testid="stSegmentedControl"] [role="tab"] {
                border-radius: 999px !important;
                border: 1px solid rgba(36, 50, 74, 0.95) !important;
                background: rgba(18, 26, 43, 0.9) !important;
                color: #dbe7f5 !important;
                font-weight: 700 !important;
                box-shadow: none !important;
            }

            [data-testid="stSegmentedControl"] button[aria-pressed="true"],
            [data-testid="stSegmentedControl"] [role="tab"][aria-selected="true"] {
                background: linear-gradient(135deg, rgba(34, 211, 238, 0.22) 0%, rgba(20, 184, 166, 0.24) 100%) !important;
                border-color: rgba(34, 211, 238, 0.48) !important;
                color: #f8fafc !important;
                box-shadow: 0 10px 24px rgba(2, 6, 23, 0.24);
            }

            [data-testid="stSegmentedControl"] button[aria-pressed="false"],
            [data-testid="stSegmentedControl"] [role="tab"][aria-selected="false"] {
                background: rgba(15, 23, 42, 0.96) !important;
                color: #cbd5e1 !important;
            }

            [data-testid="stSegmentedControl"] svg,
            [data-testid="stSegmentedControl"] [data-baseweb="icon"] {
                display: none !important;
            }

            div[role="radiogroup"] > label {
                background: rgba(18, 26, 43, 0.9);
                border: 1px solid rgba(36, 50, 74, 0.95);
                border-radius: 999px;
                padding: 0.4rem 0.9rem;
                color: var(--text-main);
            }

            [data-testid="stSidebar"] [data-baseweb="radio"] {
                align-items: flex-start;
            }

            [data-baseweb="radio"] > div:first-child {
                background: rgba(26, 37, 64, 0.98);
                border-color: rgba(71, 85, 105, 0.9);
            }

            [data-baseweb="radio"] input:checked + div {
                border-color: var(--brand-2);
            }

            [data-testid="stCaptionContainer"],
            .stCaptionContainer,
            .stMarkdown p {
                color: var(--text-soft);
            }

            [data-testid="stAlert"] *,
            [data-testid="stExpander"] summary,
            [data-testid="stExpander"] summary * {
                color: var(--text-main) !important;
            }

            [data-testid="stSidebar"] [data-testid="stTextInputRootElement"] label,
            [data-testid="stSidebar"] [data-testid="stTextArea"] label,
            [data-testid="stSidebar"] [data-testid="stSelectbox"] label,
            [data-testid="stSidebar"] [data-testid="stSlider"] label,
            [data-testid="stSidebar"] [data-testid="stToggle"] label,
            [data-testid="stSidebar"] .stMarkdown p,
            [data-testid="stSidebar"] .stCaptionContainer {
                color: var(--text-main) !important;
            }

            [data-testid="stCheckbox"] label,
            [data-testid="stCheckbox"] p,
            [data-testid="stCheckbox"] span,
            [data-testid="stCheckbox"] div {
                color: var(--text-main) !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] {
                background: linear-gradient(180deg, rgba(18,26,43,0.98) 0%, rgba(15,22,39,0.98) 100%) !important;
                border: 1.5px solid rgba(241, 245, 249, 0.85) !important;
                border-radius: 22px !important;
                box-shadow: 0 18px 36px rgba(2, 6, 23, 0.35);
                padding: 0.55rem 0.75rem 0.8rem !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) {
                padding: 0.95rem 1rem 1rem !important;
                border-color: rgba(148, 163, 184, 0.42) !important;
                background:
                    radial-gradient(circle at top right, rgba(34, 211, 238, 0.08), transparent 28%),
                    linear-gradient(180deg, rgba(18,26,43,0.98) 0%, rgba(15,22,39,0.98) 100%) !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-testid="stTextInputRootElement"] > div,
            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-baseweb="input"] > div,
            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-baseweb="base-input"] > div {
                min-height: 52px !important;
                border-radius: 16px !important;
                border: 1px solid rgba(71, 85, 105, 0.52) !important;
                background: linear-gradient(180deg, rgba(15, 23, 42, 0.96) 0%, rgba(17, 24, 39, 0.96) 100%) !important;
                box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.02);
                padding-left: 0.55rem !important;
                padding-right: 0.55rem !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-testid="stTextInputRootElement"] input {
                min-height: 50px !important;
                font-size: 0.98rem !important;
                font-weight: 600 !important;
                letter-spacing: 0.01em;
                padding-left: 0.15rem !important;
                padding-right: 0.15rem !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-testid="stTextInputRootElement"] > div:focus-within,
            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-baseweb="input"] > div:focus-within,
            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-baseweb="base-input"] > div:focus-within {
                border: 1px solid rgba(34, 211, 238, 0.62) !important;
                box-shadow: 0 0 0 3px rgba(34, 211, 238, 0.14) !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-testid="stTextInputRootElement"] input::placeholder {
                color: rgba(148, 163, 184, 0.66) !important;
                font-weight: 500 !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"]:has(.manual-editor-scope) [data-testid="stButton"] button {
                min-height: 52px !important;
                border-radius: 16px !important;
                font-size: 0.92rem !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] > div,
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stSelectbox"] [data-baseweb="select"] > div {
                background: transparent !important;
                border: none !important;
                outline: none !important;
                box-shadow: none !important;
                min-height: 42px !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="input"],
            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="base-input"],
            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="input"] > div,
            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="base-input"] > div,
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"],
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] > div {
                background: transparent !important;
                border: none !important;
                outline: none !important;
                box-shadow: none !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] input,
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stSelectbox"] input,
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stSelectbox"] span,
            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stSelectbox"] div {
                color: #f8fafc !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] input {
                background: transparent !important;
                border: none !important;
                outline: none !important;
                box-shadow: none !important;
                padding-left: 0.15rem !important;
                padding-right: 0.15rem !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] > div:focus-within,
            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="input"] > div:focus-within,
            [data-testid="stVerticalBlockBorderWrapper"] [data-baseweb="base-input"] > div:focus-within {
                border: none !important;
                outline: none !important;
                box-shadow: none !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stButton"] button {
                border-radius: 999px !important;
                min-height: 38px !important;
            }

            [data-testid="stVerticalBlockBorderWrapper"] [data-testid="stTextInputRootElement"] input::placeholder,
            [data-testid="stVerticalBlockBorderWrapper"] textarea::placeholder {
                color: rgba(148, 163, 184, 0.72) !important;
            }

            [data-testid="stNumberInput"] button {
                background: #1e293b !important;
                color: #e2e8f0 !important;
                border-color: rgba(51, 65, 85, 0.9) !important;
            }

            [data-testid="stToolbar"] {
                opacity: 0.92;
            }

            .st-emotion-cache-16txtl3,
            .st-emotion-cache-1r4qj8v {
                color: var(--text-main);
            }

            @media (max-width: 980px) {
                .hero-grid {
                    grid-template-columns: 1fr;
                }

                .hero h1 {
                    max-width: none;
                }

                .block-container {
                    padding-top: 1rem;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def setup_sidebar() -> tuple[str, str, str, bool, float]:
    with st.sidebar:
        st.markdown("<div class='sidebar-title'>Import panel</div>", unsafe_allow_html=True)
        demo_default = (Path.cwd() / "FINANZAS_demo.xlsx").resolve()
        local_default = (Path.cwd() / "FINANZAS.xlsx").resolve()
        if demo_default.exists():
            default_path = str(demo_default)
        elif local_default.exists():
            default_path = str(local_default)
        else:
            default_path = ""

        st.markdown("<div class='sidebar-section'>Destination workbook</div>", unsafe_allow_html=True)
        st.caption("Destination path")
        finanzas_path = st.text_area("Destination path", value=default_path, height=82, label_visibility="collapsed")
        if demo_default.exists():
            st.caption("Demo workbook detected. It is selected by default for safer testing.")

        st.markdown("<div class='sidebar-section'>Date filter</div>", unsafe_allow_html=True)
        st.caption("Cutoff rule")
        date_mode_label = st.selectbox(
            "Date cutoff",
            options=[
                "Only after latest existing date",
                "Exclude existing dates",
            ],
            index=0,
            label_visibility="collapsed",
        )
        date_filter_mode = (
            DATE_FILTER_MODE_AFTER_MAX
            if date_mode_label.startswith("Only after")
            else DATE_FILTER_MODE_SKIP_EXISTING
        )

        st.markdown("<div class='sidebar-section'>Safety</div>", unsafe_allow_html=True)
        st.caption("Duplicate checks run by reference and compound key.")

        st.markdown("<div class='sidebar-section'>Fallback category</div>", unsafe_allow_html=True)
        st.caption("Used when no category rule matches")
        uncategorized_label = st.selectbox(
            "Fallback category",
            options=["Otros", "(vacio)"],
            index=0,
            label_visibility="collapsed",
        )
        default_category = "" if uncategorized_label == "(vacio)" else "Otros"
        ahorro_meta_pct = st.slider("Savings target (%)", min_value=0, max_value=40, value=10, step=1)

        st.markdown("<div class='sidebar-section'>Technical details</div>", unsafe_allow_html=True)
        show_technical = st.toggle("Show technical details", value=False)
    return finanzas_path, date_filter_mode, default_category, show_technical, float(ahorro_meta_pct)


def file_size_mb(size_bytes: int) -> str:
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def is_path_writable(path_value: str) -> tuple[bool, str]:
    if not path_value.strip():
        return False, "Set the destination workbook path."
    path = Path(path_value)
    if not path.exists():
        return False, "The file does not exist on disk."
    try:
        with path.open("ab"):
            pass
        return True, "Workbook is writable."
    except PermissionError:
        return False, "Workbook is locked. Close Excel or let sync finish."
    except OSError:
        return False, "Could not validate write permissions."


def render_hero(finanzas_ok: bool, finanzas_msg: str) -> None:
    status_class = "status-ok" if finanzas_ok else "status-bad"
    status_text = "Workbook ready" if finanzas_ok else "Review workbook"
    st.markdown(
        f"""
        <section class="hero">
            <div class="hero-grid">
                <div>
                    <div class="hero-kicker">Transaction intake</div>
                    <h1>Financial Operations Importer</h1>
                    <div class="hero-copy">
                        Load transactions from Excel exports or bank captures, review duplicates, and confirm the update
                        only when the outcome is already clear.
                    </div>
                    <div class="hero-list">
                        <span class="hero-chip">Pre-import review</span>
                        <span class="hero-chip">Automatic backup</span>
                        <span class="hero-chip">Monthly insights</span>
                    </div>
                </div>
                <div class="hero-panel">
                    <div class="hero-panel-label">Workbook status</div>
                    <div class="hero-panel-status">{status_text}</div>
                    <div class="hero-panel-copy">{finanzas_msg}</div>
                    <span class="status-pill {status_class}">{status_text}</span>
                </div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_mode_panel(mode: str) -> None:
    if mode == "Excel file":
        title = "Excel file"
        copy = "Import transactions from an Excel export. The current parser is optimized for Mercado Pago exports."
        tags = ["Transactions", "Review step", "Safe import"]
    else:
        title = "Bank capture"
        copy = "OCR-assisted flow with manual editing to register movements from a bank screenshot."
        tags = ["OCR assisted", "Manual editing", "Batch intake"]

    tags_html = "".join(f"<span class='hero-chip'>{tag}</span>" for tag in tags)
    st.markdown(
        f"""
        <div style="margin:0.45rem 0 1rem; padding:1rem 1.1rem; border-radius:20px; background:linear-gradient(135deg, rgba(18,26,43,0.94) 0%, rgba(26,37,64,0.96) 100%); border:1px solid rgba(36,50,74,0.95); box-shadow:0 18px 40px rgba(2,6,23,0.32);">
            <div style="font-family:'Space Grotesk','Segoe UI',sans-serif; font-size:1.08rem; font-weight:700; color:#f8fafc; margin-bottom:0.3rem;">{title}</div>
            <div style="color:#94a3b8; line-height:1.55; margin-bottom:0.75rem;">{copy}</div>
            <div style="display:flex; flex-wrap:wrap; gap:0.45rem;">{tags_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def format_preview_df(df: pd.DataFrame) -> pd.DataFrame:
    preview = df.copy()
    if "tipo" in preview.columns:
        preview["tipo_visual"] = preview["tipo"].map(_preview_tipo).fillna("Expense")
    if "compartido" in preview.columns:
        preview["compartido"] = (
            preview["compartido"]
            .map({True: "Yes", False: "No"})
            .fillna(preview["compartido"])
            .replace({"S?": "Yes", "SI": "Yes", "NO": "No", "": "No"})
        )
    return preview


def render_step_title(step_number: int, title: str) -> None:
    st.markdown(
        f"""
        <div style="display:flex; align-items:center; gap:0.95rem; margin:1.7rem 0 0.75rem; padding:0.95rem 1.1rem; border-radius:22px; border:1px solid rgba(36,50,74,0.95); background:linear-gradient(180deg, rgba(18,26,43,0.96) 0%, rgba(15,22,39,0.98) 100%); box-shadow:0 18px 42px rgba(2,6,23,0.3);">
            <div style="width:42px; height:42px; border-radius:14px; display:flex; align-items:center; justify-content:center; background:linear-gradient(145deg, #6d5ef5 0%, #22d3ee 100%); color:#f8fffc; font-family:'Space Grotesk','Segoe UI',sans-serif; font-size:1rem; font-weight:700; flex:0 0 auto;">{step_number:02d}</div>
            <div>
                <div style="font-family:'Space Grotesk','Segoe UI',sans-serif; font-size:1.22rem; font-weight:700; color:#f8fafc;">{title}</div>
                <div style="color:#94a3b8; font-size:0.9rem; margin-top:0.12rem;">Revision simple y ordenada antes de avanzar.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def init_upload_state() -> None:
    if "uploader_key" not in st.session_state:
        st.session_state.uploader_key = "mp_uploader"
    if "upload_at" not in st.session_state:
        st.session_state.upload_at = None


def _render_meta_card(label: str, value: str) -> None:
    st.markdown(
        f"""
        <div style="border-radius:18px; padding:0.95rem 1rem; background:linear-gradient(180deg, rgba(18,26,43,0.95) 0%, rgba(15,22,39,0.98) 100%); border:1px solid rgba(36,50,74,0.95); box-shadow:0 14px 28px rgba(2,6,23,0.28);">
            <div style="color:#94a3b8; text-transform:uppercase; letter-spacing:0.12em; font-size:0.72rem; font-weight:700; margin-bottom:0.3rem;">{label}</div>
            <div style="color:#f8fafc; font-weight:700; font-size:0.95rem; word-break:break-word;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_upload_meta(uploaded_file) -> None:
    meta_cols = st.columns([1.8, 1, 1, 0.9])
    with meta_cols[0]:
        _render_meta_card("Archivo", str(uploaded_file.name))
    with meta_cols[1]:
        _render_meta_card("Tamano", file_size_mb(uploaded_file.size))
    with meta_cols[2]:
        loaded_at = st.session_state.upload_at.strftime("%d/%m/%Y %H:%M:%S") if st.session_state.upload_at else "-"
        _render_meta_card("Cargado", loaded_at)
    with meta_cols[3]:
        st.write("")
        st.write("")
        if st.button("Quitar archivo"):
            st.session_state.uploader_key = f"mp_uploader_{datetime.now().timestamp()}"
            st.session_state.upload_at = None
            st.rerun()


def render_upload_dropzone_hint() -> None:
    st.markdown(
        """
        <div style="border:1px dashed rgba(34,211,238,0.3); background:linear-gradient(180deg, rgba(18,26,43,0.92) 0%, rgba(26,37,64,0.86) 100%); color:#dbeafe; border-radius:18px; padding:0.95rem 1rem; margin-bottom:0.8rem;">
            Arrastra el archivo al area de carga o usa <strong>Browse files</strong>. Primero se revisa y despues, si corresponde, se importa.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_footer() -> None:
    st.markdown("---")
    st.markdown(
        f"<div style='color:#94a3b8; font-size:0.85rem; margin-top:1rem;'>finanzas importer | ejecucion: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</div>",
        unsafe_allow_html=True,
    )


def _money(value: float) -> str:
    return f"ARS {value:,.0f}"


def _metric_delta(value: float, improve_when_down: bool = False) -> tuple[str, str]:
    sign = "+" if value > 0 else ""
    delta = f"{sign}ARS {value:,.0f}"
    if value == 0:
        return delta, "off"
    if improve_when_down:
        return delta, "inverse"
    return delta, "normal"


def _preview_tipo(value: object) -> str:
    return "Income" if str(value).strip().lower() == "ingreso" else "Expense"


def render_preview_table(df: pd.DataFrame, columns: list[str], labels: dict[str, str]) -> None:
    if df.empty:
        st.info("There are no rows to display.")
        return

    def format_cell(column: str, value: object) -> str:
        if pd.isna(value):
            return ""
        if column == "date":
            date_value = pd.to_datetime(value, errors="coerce")
            return "" if pd.isna(date_value) else date_value.strftime("%d/%m/%Y")
        if column == "monto":
            return f"ARS {float(value):,.2f}"
        if column == "tipo_visual":
            tipo = str(value).strip().lower()
            dot_class = "tipo-ingreso" if tipo == "ingreso" else "tipo-gasto"
            return (
                f"<span class='tipo-chip {dot_class}'>"
                f"<span class='tipo-dot'>&#9679;</span>{escape(str(value))}</span>"
            )
        if column == "compartido":
            return f"<span class='compartido-flag'>{escape(str(value))}</span>"
        return escape(str(value))

    headers = "".join(f"<th>{escape(labels.get(column, column.title()))}</th>" for column in columns)
    rows = []
    for _, row in df[columns].iterrows():
        cells = "".join(f"<td>{format_cell(column, row[column])}</td>" for column in columns)
        rows.append(f"<tr>{cells}</tr>")

    st.markdown(
        f"""
        <div class="preview-table-wrap">
            <table class="preview-table">
                <thead><tr>{headers}</tr></thead>
                <tbody>{''.join(rows)}</tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _render_alert_card(level: str, title: str, detail: str, action: str) -> None:
    palette = {
        "critical": {
            "border": "rgba(239,68,68,0.32)",
            "bg": "linear-gradient(180deg, rgba(127,29,29,0.28) 0%, rgba(30,41,59,0.96) 100%)",
            "title": "#fecaca",
            "text": "#fee2e2",
        },
        "warning": {
            "border": "rgba(245,158,11,0.30)",
            "bg": "linear-gradient(180deg, rgba(120,53,15,0.24) 0%, rgba(30,41,59,0.96) 100%)",
            "title": "#fde68a",
            "text": "#fef3c7",
        },
        "info": {
            "border": "rgba(34,211,238,0.28)",
            "bg": "linear-gradient(180deg, rgba(8,47,73,0.24) 0%, rgba(30,41,59,0.96) 100%)",
            "title": "#bae6fd",
            "text": "#e0f2fe",
        },
    }[level]
    st.markdown(
        f"""
        <div style="margin-bottom:0.8rem; padding:1rem 1.1rem; border-radius:18px; border:1px solid {palette['border']}; background:{palette['bg']};">
            <div style="font-family:'Space Grotesk','Segoe UI',sans-serif; font-size:1rem; margin-bottom:0.25rem; color:{palette['title']};">{title}</div>
            <div style="color:{palette['text']};">{detail}</div>
            <div style="margin-top:0.6rem; color:{palette['text']};"><strong>Accion sugerida:</strong> {action}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_review_step(
    parsed_df: pd.DataFrame,
    parse_result: ParseResult,
    plan: ImportPlan | None,
    date_filter_mode: str,
    show_technical: bool,
    plan_error: str | None,
) -> int:
    total_rows = len(parsed_df)
    filtered_rendimientos = parse_result.filtered_reasons.get(FILTER_REASON_RENDIMIENTOS, 0)
    filtered_self_transfer = parse_result.filtered_reasons.get(FILTER_REASON_SELF_TRANSFER, 0)
    marked_shared_transfers = parse_result.shared_transfer_count
    skipped_by_date = int(plan.filtered_by_date_df.shape[0]) if plan else 0
    dupes_ref = int(plan.duplicate_mp_ref_df.shape[0]) if plan else 0
    dupes_key = int(plan.duplicate_compound_df.shape[0]) if plan else 0
    to_import = int(plan.to_import_df.shape[0]) if plan else total_rows

    render_step_title(2, "Review before import")
    st.markdown(
        """
        <div style="margin:0.2rem 0 1rem; padding:1rem 1.15rem; border-radius:20px; border:1px solid rgba(36,50,74,0.95); background:linear-gradient(180deg, rgba(18,26,43,0.94) 0%, rgba(15,22,39,0.98) 100%); box-shadow:0 16px 36px rgba(2,6,23,0.3);">
            <div style="font-size:0.8rem; letter-spacing:0.12em; text-transform:uppercase; color:#22d3ee; font-weight:800; margin-bottom:0.3rem;">Control de calidad</div>
            <div style="color:#94a3b8; line-height:1.55; font-size:0.94rem;">
                This block summarizes detections, filters, and duplicates so the import decision is obvious at a glance.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric_cols = st.columns(5)
    metric_cols[0].metric("Detected", f"{total_rows}")
    metric_cols[1].metric("Skipped by date", f"{skipped_by_date}")
    metric_cols[2].metric("Duplicates", f"{dupes_ref + dupes_key}")
    metric_cols[3].metric("Flagged shared", f"{marked_shared_transfers}")
    metric_cols[4].metric("To import", f"{to_import}")

    hidden_filters = filtered_rendimientos + filtered_self_transfer + dupes_ref + dupes_key
    st.caption(
        "The main view keeps only the most useful indicators for a fast decision. "
        f"The audit trail still holds {hidden_filters} technical records across filters and duplicates."
    )

    if to_import == 0:
        st.markdown(
            "<div style='padding:1rem 1.1rem; margin:0.8rem 0 1rem; border-radius:18px; border:1px solid rgba(34,197,94,0.22); background:linear-gradient(180deg, rgba(6,78,59,0.34) 0%, rgba(18,26,43,0.96) 100%); color:#dcfce7; box-shadow:0 14px 30px rgba(2,6,23,0.24);'><strong>Up to date.</strong> There are no new rows to import with the current rules.</div>",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"<div style='padding:1rem 1.1rem; margin:0.8rem 0 1rem; border-radius:18px; border:1px solid rgba(34,211,238,0.24); background:linear-gradient(180deg, rgba(109,94,245,0.16) 0%, rgba(18,26,43,0.96) 100%); color:#e0f2fe; box-shadow:0 14px 30px rgba(2,6,23,0.24);'><strong>Ready to import.</strong> {to_import} new rows will be added after the safety filters pass.</div>",
            unsafe_allow_html=True,
        )

    preview_df = plan.to_import_df if plan is not None and not plan.to_import_df.empty else parsed_df
    preview_df = format_preview_df(preview_df)
    if preview_df.shape[0] > 50:
        st.caption("Showing the first 50 rows to keep the app responsive.")
    st.markdown("#### Preview")
    render_preview_table(
        preview_df.head(50),
        ["date", "tipo_visual", "categoria", "descripcion", "monto", "compartido"],
        {
            "date": "Date",
            "tipo_visual": "Type",
            "categoria": "Category",
            "descripcion": "Description",
            "monto": "Amount",
            "compartido": "Shared",
        },
    )

    show_audit = st.toggle("Show audit trail", value=False)
    if show_audit:
        audit_cols = st.columns(4)
        audit_cols[0].metric("Filtered earnings", f"{filtered_rendimientos}")
        audit_cols[1].metric("Own transfers", f"{filtered_self_transfer}")
        audit_cols[2].metric("Duplicate refs", f"{dupes_ref}")
        audit_cols[3].metric("Duplicate keys", f"{dupes_key}")

        with st.expander("Detail: rows to import", expanded=False):
            if plan is not None and not plan.to_import_df.empty:
                render_preview_table(
                    format_preview_df(plan.to_import_df).head(50),
                    ["date", "tipo_visual", "categoria", "descripcion", "monto", "compartido", "mp_ref"],
                    {
                        "date": "Date",
                        "tipo_visual": "Type",
                        "categoria": "Category",
                        "descripcion": "Description",
                        "monto": "Amount",
                        "compartido": "Shared",
                        "mp_ref": "Reference",
                    },
                )
            else:
                st.info("There are no rows ready to import.")

        with st.expander("Detail: skipped by date", expanded=False):
            if plan is not None and not plan.filtered_by_date_df.empty:
                render_preview_table(
                    format_preview_df(plan.filtered_by_date_df).head(50),
                    ["date", "tipo_visual", "descripcion", "monto", "mp_ref"],
                    {
                        "date": "Date",
                        "tipo_visual": "Type",
                        "descripcion": "Description",
                        "monto": "Amount",
                        "mp_ref": "Reference",
                    },
                )
            else:
                st.info("There were no rows skipped by date.")

        with st.expander("Detail: excluded self-transfers", expanded=False):
            if parse_result.self_transfer_examples is not None and not parse_result.self_transfer_examples.empty:
                render_preview_table(
                    format_preview_df(parse_result.self_transfer_examples).head(50),
                    ["date", "tipo_visual", "descripcion", "monto", "mp_ref"],
                    {
                        "date": "Date",
                        "tipo_visual": "Type",
                        "descripcion": "Description",
                        "monto": "Amount",
                        "mp_ref": "Reference",
                    },
                )
            else:
                st.info("There were no excluded internal transfers.")

    if show_technical:
        st.markdown("#### Technical details")
        mode_caption = (
            "Only import > latest date" if date_filter_mode == DATE_FILTER_MODE_AFTER_MAX else "Exclude existing dates"
        )
        st.caption(f"Date mode: `{mode_caption}`")
        st.caption(f"Parser date column: `{parse_result.date_source_column}`")
        if plan is not None:
            st.caption(f"Detected workbook table: `{plan.table_name}`")
        if plan_error:
            st.error(f"Could not validate duplicates against the workbook: {plan_error}")

        with st.expander("Technical: category rules", expanded=False):
            rules_summary = (
                parsed_df["regla_categoria"].value_counts(dropna=False).rename_axis("regla_categoria").reset_index(name="count")
            )
            st.dataframe(rules_summary, hide_index=True, width="stretch")
            rules_examples = parsed_df.sort_values(["regla_categoria", "date"]).groupby("regla_categoria", as_index=False).head(5)
            st.dataframe(
                rules_examples[["regla_categoria", "date", "categoria", "descripcion", "monto", "mp_ref"]],
                hide_index=True,
                width="stretch",
            )

    return to_import


def render_insights_step(
    analytics_df: pd.DataFrame,
    kpis: MonthlyKpis | None,
    projection: dict[str, float] | None,
    alerts: list[AlertItem],
    ref_date: pd.Timestamp | None = None,
    executive_summary: str = "",
    ahorro_meta_pct: float = 10.0,
) -> None:
    render_step_title(3, "Monthly snapshot")
    if analytics_df.empty or kpis is None:
        st.info("There is not enough historical data to show monthly insights.")
        return

    ref = ref_date if ref_date is not None else analytics_df["date"].max()
    st.markdown(
        f"""
        <div style="margin:0.2rem 0 1rem; padding:1rem 1.15rem; border-radius:20px; border:1px solid rgba(36,50,74,0.95); background:linear-gradient(180deg, rgba(18,26,43,0.94) 0%, rgba(15,22,39,0.98) 100%); box-shadow:0 16px 36px rgba(2,6,23,0.3);">
            <div style="font-size:0.8rem; letter-spacing:0.12em; text-transform:uppercase; color:#22d3ee; font-weight:800; margin-bottom:0.3rem;">Analysis period</div>
            <div style="color:#94a3b8; line-height:1.55; font-size:0.94rem;">{ref.strftime('%m/%Y')} | Executive view of the current month against recent history.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    if executive_summary:
        st.info(executive_summary)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Monthly income", _money(kpis.ingresos))
    gasto_delta, gasto_delta_color = _metric_delta(kpis.delta_gasto_prev, improve_when_down=True)
    balance_delta, balance_delta_color = _metric_delta(kpis.delta_balance_prev)
    balance_avg_delta, balance_avg_delta_color = _metric_delta(kpis.delta_balance_avg3)
    c2.metric("Monthly expense", _money(kpis.gastos), delta=gasto_delta, delta_color=gasto_delta_color)
    c3.metric("Monthly balance", _money(kpis.balance), delta=balance_delta, delta_color=balance_delta_color)
    c4.metric("Savings %", f"{kpis.ahorro_pct:.1f}%")

    c5, c6, c7 = st.columns(3)
    c5.metric("Estimated fixed spend", _money(kpis.gasto_fijo_aprox))
    c6.metric("Estimated variable spend", _money(kpis.gasto_variable_aprox))
    c7.metric("Balance vs 3m average", _money(kpis.balance), delta=balance_avg_delta, delta_color=balance_avg_delta_color)

    month_df = analytics_df[
        (analytics_df["date"].dt.year == ref.year)
        & (analytics_df["date"].dt.month == ref.month)
    ].copy()
    month_df["day"] = month_df["date"].dt.day
    day_agg = month_df.groupby(["day"], as_index=False)["signed_monto"].sum().sort_values("day")
    day_agg["balance_acum"] = day_agg["signed_monto"].cumsum()

    chart_theme = {
        "background": "#0f172a",
        "view": {"stroke": None, "fill": "#0f172a"},
        "axis": {
            "labelColor": "#d7e4f3",
            "labelFontSize": 12,
            "titleColor": "#f8fafc",
            "titleFontSize": 13,
            "gridColor": "rgba(100, 116, 139, 0.26)",
            "domainColor": "rgba(148, 163, 184, 0.35)",
            "tickColor": "rgba(148, 163, 184, 0.35)",
        },
        "legend": {
            "labelColor": "#d7e4f3",
            "titleColor": "#f8fafc",
            "orient": "top",
        },
        "title": {"color": "#f8fafc", "font": "Space Grotesk", "fontSize": 18, "anchor": "start"},
    }

    line_chart = (
        alt.Chart(day_agg)
        .mark_line(point=alt.OverlayMarkDef(size=72, filled=True), color="#2dd4bf", strokeWidth=3)
        .encode(
            x=alt.X("day:Q", title="Day of month", axis=alt.Axis(format="d")),
            y=alt.Y("balance_acum:Q", title="Cumulative balance", axis=alt.Axis(format=",.0f")),
            tooltip=[
                alt.Tooltip("day:Q", title="Day"),
                alt.Tooltip("balance_acum:Q", title="Cumulative balance", format=",.2f"),
            ],
        )
        .properties(height=290, title="Cumulative daily trend")
        .configure(**chart_theme)
    )

    prev_ref = ref.replace(day=1) - pd.Timedelta(days=1)
    prev_df = analytics_df[
        (analytics_df["date"].dt.year == prev_ref.year)
        & (analytics_df["date"].dt.month == prev_ref.month)
    ].copy()
    curr_cat = (
        month_df[month_df["signed_monto"] < 0]
        .groupby("categoria", as_index=False)["signed_monto"]
        .sum()
        .assign(gasto=lambda d: d["signed_monto"].abs())
        [["categoria", "gasto"]]
    )
    prev_cat = (
        prev_df[prev_df["signed_monto"] < 0]
        .groupby("categoria", as_index=False)["signed_monto"]
        .sum()
        .assign(gasto=lambda d: d["signed_monto"].abs())
        [["categoria", "gasto"]]
    )
    top_curr = curr_cat.sort_values("gasto", ascending=False).head(6)["categoria"].tolist()
    cmp_df = pd.concat(
        [
            curr_cat[curr_cat["categoria"].isin(top_curr)].assign(periodo=f"{ref.strftime('%m/%Y')}"),
            prev_cat[prev_cat["categoria"].isin(top_curr)].assign(periodo=f"{prev_ref.strftime('%m/%Y')}"),
        ],
        ignore_index=True,
    )
    bar_chart = (
        alt.Chart(cmp_df)
        .mark_bar(cornerRadiusTopRight=8, cornerRadiusBottomRight=8)
        .encode(
            x=alt.X("gasto:Q", title="ARS", axis=alt.Axis(format=",.0f")),
            y=alt.Y("categoria:N", sort="-x", title="Category"),
            color=alt.Color("periodo:N", scale=alt.Scale(range=["#2dd4bf", "#38bdf8"])),
            opacity=alt.Opacity("periodo:N", scale=alt.Scale(range=[1.0, 0.55])),
            xOffset="periodo:N",
            tooltip=[
                alt.Tooltip("categoria:N", title="Category"),
                alt.Tooltip("periodo:N", title="Period"),
                alt.Tooltip("gasto:Q", title="Spend", format=",.2f"),
            ],
        )
        .properties(height=290, title="Categories: current month vs previous month")
        .configure(**chart_theme)
    )

    weekday_df = month_df[month_df["signed_monto"] < 0].copy()
    day_map = {
        "Monday": "Monday",
        "Tuesday": "Tuesday",
        "Wednesday": "Wednesday",
        "Thursday": "Thursday",
        "Friday": "Friday",
        "Saturday": "Saturday",
        "Sunday": "Sunday",
    }
    weekday_df["weekday"] = weekday_df["date"].dt.day_name().map(day_map)
    order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday_df["weekday"] = pd.Categorical(weekday_df["weekday"], categories=order, ordered=True)
    weekday_df = weekday_df.groupby("weekday", as_index=False)["signed_monto"].sum()
    weekday_df["gasto"] = weekday_df["signed_monto"].abs()
    weekday_base = alt.Chart(weekday_df).encode(
        x=alt.X(
            "weekday:N",
            sort=order,
            title="Weekday",
            axis=alt.Axis(labelAngle=0, labelPadding=10),
        ),
        y=alt.Y("gasto:Q", title="Total spend", axis=alt.Axis(format=",.0f")),
        tooltip=[
            alt.Tooltip("weekday:N", title="Weekday"),
            alt.Tooltip("gasto:Q", title="Total spend", format=",.2f"),
        ],
    )
    weekday_area = weekday_base.mark_area(color="#0ea5a4", opacity=0.18, interpolate="monotone")
    weekday_line = weekday_base.mark_line(color="#67e8f9", strokeWidth=3, interpolate="monotone")
    weekday_points = weekday_base.mark_circle(color="#67e8f9", size=88)
    weekday_chart = (weekday_area + weekday_line + weekday_points).properties(height=240, title="Weekly spending rhythm").configure(
        **chart_theme
    )

    left, right = st.columns(2)
    left.altair_chart(line_chart, width="stretch")
    right.altair_chart(bar_chart, width="stretch")
    st.altair_chart(weekday_chart, width="stretch")

    if projection is not None:
        st.markdown("#### Monthly signal")
        if projection["balance_proyectado"] < 0:
            st.error("Critical: at the current pace, the projected closing balance is negative.")
        elif kpis.ahorro_pct < ahorro_meta_pct:
            st.warning(f"Alert: monthly savings are running below {ahorro_meta_pct:.0f}%.")
        else:
            st.success("Healthy projection: savings are within target.")

    if projection is not None:
        st.markdown("#### Closing projection")
        p1, p2, p3 = st.columns(3)
        p1.metric("Projected income", _money(projection["ingreso_proyectado"]))
        p2.metric("Projected spend", _money(projection["gasto_proyectado"]))
        p3.metric("Projected balance", _money(projection["balance_proyectado"]))

    st.markdown("#### Actionable alerts")
    for alert in alerts:
        _render_alert_card(alert.level, alert.title, alert.detail, alert.action)


def render_import_step(
    parsed_df: pd.DataFrame,
    parse_result: ParseResult,
    to_import_final: int,
    finanzas_path: str,
    date_filter_mode: str,
    default_category: str,
    importer_fn: Callable[..., object],
) -> None:
    render_step_title(4, "Update workbook")
    st.markdown(
        """
        <div style="margin:0.2rem 0 1rem; padding:1rem 1.15rem; border-radius:20px; border:1px solid rgba(36,50,74,0.95); background:linear-gradient(180deg, rgba(18,26,43,0.94) 0%, rgba(15,22,39,0.98) 100%); box-shadow:0 16px 36px rgba(2,6,23,0.3);">
            <div style="font-size:0.8rem; letter-spacing:0.12em; text-transform:uppercase; color:#22d3ee; font-weight:800; margin-bottom:0.3rem;">Ultimo control</div>
            <div style="color:#94a3b8; line-height:1.55; font-size:0.94rem;">
                Writing stays locked until you confirm the impact. If there are no new rows, the default protection remains in place.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    importar_igual = False
    if parse_result.suspicious_dates:
        st.warning("The file dates look unusual. Review them before continuing to avoid incorrect imports.")
        importar_igual = st.checkbox("I confirm the dates are correct", value=False)

    force_write = False
    if to_import_final == 0:
        force_write = st.checkbox("Force write even if there are no new rows", value=False)
    confirm_import = st.checkbox("I understand the destination workbook will be modified", value=False)

    disable_import = (
        parsed_df is None
        or (parse_result.suspicious_dates and not importar_igual)
        or (to_import_final == 0 and not force_write)
        or not confirm_import
    )
    if disable_import:
        st.caption("The button becomes available after confirming the write and passing the safety checks.")

    if st.button("Importar", type="primary", disabled=disable_import):
        if parsed_df is None:
            st.warning("Upload a file before importing.")
        elif parse_result.suspicious_dates and not importar_igual:
            st.warning("Confirm the dates to continue.")
        elif not finanzas_path.strip():
            st.warning("Complete the destination workbook path.")
        else:
            try:
                result = importer_fn(
                    parsed_df,
                    Path(finanzas_path),
                    date_filter_mode=date_filter_mode,
                    default_category=default_category,
                )
                if result.note == "No rows to import":
                    st.info("There are no new rows to import. The workbook was left unchanged.")
                    st.stop()
                st.success(
                    "Import completed: "
                    f"{result.added_rows} added, "
                    f"{result.skipped_rows_by_date} skipped by date, "
                    f"{result.duplicate_rows_mp_ref} duplicate references, "
                    f"{result.duplicate_rows_compound_key} duplicate compound keys."
                )
                if result.backup_path is not None:
                    st.info("Backup created:")
                    st.code(str(result.backup_path))
            except PermissionError:
                st.error(
                    "The workbook could not be saved. Close it in Excel and, if it is cloud-synced, mark it as available offline."
                )
            except Exception as exc:
                st.error(f"An error occurred while importing into the destination workbook: {exc}")
