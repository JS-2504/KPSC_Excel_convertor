"""
Kerala PSC Rank List Converter — Streamlit web UI
==================================================

A single-page web app that lets anyone drop a Kerala PSC ranked-list PDF
in their browser and download the structured Excel file. Run locally with

    streamlit run app.py

or deploy free on Streamlit Community Cloud (see README.md).
"""
from __future__ import annotations

import tempfile
import time
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st

from psc_pdf_to_xlsx import extract_rank_list, write_xlsx


# ---------------------------------------------------------------------------
# Page configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="PSC Rank List Converter",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed",
    menu_items={
        "About": (
            "Convert Kerala Public Service Commission ranked-list PDFs into "
            "structured Excel files. Free and open source."
        ),
    },
)


# ---------------------------------------------------------------------------
# Theme tokens — exposed as CSS variables so styles stay consistent
# ---------------------------------------------------------------------------
CSS = """
<style>
:root {
    --brand: #2da77d;
    --brand-hover: #248a66;
    --brand-soft: #e7f6f0;
    --brand-softer: #f3faf7;
    --ink-900: #0b1f17;
    --ink-700: #1f2a26;
    --ink-500: #4b5b54;
    --ink-300: #8a9690;
    --ink-100: #d6dcd9;
    --line: #e8ecea;
    --line-soft: #f1f4f3;
    --bg: #ffffff;
}

/* Reset Streamlit chrome */
#MainMenu, header, footer {visibility: hidden;}
.stDeployButton {display: none !important;}
[data-testid="stToolbar"] {display: none !important;}
[data-testid="stHeader"] {display: none !important;}
[data-testid="stStatusWidget"] {display: none !important;}

/* Clean system font stack */
html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, "Inter", "Segoe UI",
                 "Helvetica Neue", Arial, sans-serif !important;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
}

.stApp {
    background: var(--bg);
    color: var(--ink-700);
}

.block-container {
    padding-top: 3.5rem !important;
    padding-bottom: 4rem !important;
    max-width: 720px !important;
}

/* ----- Brand mark + hero ------------------------------------------------ */
.brand {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 10px;
    margin-bottom: 2.25rem;
    color: var(--ink-700);
    font-weight: 600;
    font-size: 0.95rem;
    letter-spacing: -0.01em;
}
.brand-mark {
    width: 28px;
    height: 28px;
    border-radius: 7px;
    background: var(--brand);
    display: flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-weight: 700;
    font-size: 0.85rem;
    letter-spacing: -0.02em;
}

.hero {
    text-align: center;
    margin-bottom: 2rem;
}
.hero h1 {
    font-size: 2.5rem !important;
    font-weight: 700 !important;
    color: var(--ink-900) !important;
    letter-spacing: -0.035em !important;
    line-height: 1.1 !important;
    margin: 0 0 0.85rem 0 !important;
}
.hero p.subtitle {
    font-size: 1.06rem;
    color: var(--ink-500);
    max-width: 520px;
    margin: 0 auto;
    line-height: 1.55;
    font-weight: 400;
}

/* ----- File uploader ---------------------------------------------------- */
section[data-testid="stFileUploaderDropzone"],
section[data-testid="stFileUploadDropzone"] {
    background: var(--brand-softer) !important;
    border: 1.5px dashed #b5d8c9 !important;
    border-radius: 14px !important;
    padding: 2.5rem 1.5rem !important;
    transition: all 0.18s ease;
}
section[data-testid="stFileUploaderDropzone"]:hover,
section[data-testid="stFileUploadDropzone"]:hover {
    border-color: var(--brand) !important;
    background: var(--brand-soft) !important;
}
section[data-testid="stFileUploaderDropzone"] button,
section[data-testid="stFileUploadDropzone"] button {
    background: var(--brand) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.55rem 1.25rem !important;
    transition: background 0.15s ease !important;
}
section[data-testid="stFileUploaderDropzone"] button:hover,
section[data-testid="stFileUploadDropzone"] button:hover {
    background: var(--brand-hover) !important;
}
/* Dropzone instruction text */
[data-testid="stFileUploaderDropzoneInstructions"],
section[data-testid="stFileUploaderDropzone"] small,
section[data-testid="stFileUploadDropzone"] small {
    color: var(--ink-500) !important;
}

/* Uploaded-file pill */
[data-testid="stFileUploaderFile"] {
    background: white !important;
    border: 1px solid var(--line) !important;
    border-radius: 10px !important;
}

/* ----- Buttons ---------------------------------------------------------- */
.stDownloadButton button, .stButton button {
    background: var(--brand) !important;
    color: white !important;
    border: none !important;
    padding: 0.85rem 2rem !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
    letter-spacing: -0.005em !important;
    box-shadow: 0 1px 2px rgba(45, 167, 125, 0.10) !important;
    transition: all 0.15s ease !important;
}
.stDownloadButton button:hover, .stButton button:hover {
    background: var(--brand-hover) !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(45, 167, 125, 0.22) !important;
}
.stDownloadButton button:focus,
.stButton button:focus,
.stDownloadButton button:active,
.stButton button:active {
    background: var(--brand-hover) !important;
    box-shadow: 0 0 0 3px rgba(45, 167, 125, 0.18) !important;
    color: white !important;
}

/* ----- Progress bar ----------------------------------------------------- */
.stProgress > div > div > div > div {
    background: var(--brand) !important;
}
.stProgress > div > div > div {
    background: var(--line-soft) !important;
}

/* ----- Metric cards ----------------------------------------------------- */
[data-testid="stMetric"] {
    background: white;
    padding: 1.1rem 1.25rem;
    border-radius: 12px;
    border: 1px solid var(--line);
    box-shadow: none;
    transition: border-color 0.15s ease;
}
[data-testid="stMetric"]:hover {
    border-color: #c8d5d0;
}
[data-testid="stMetricLabel"] p,
[data-testid="stMetricLabel"] {
    color: var(--ink-300) !important;
    font-size: 0.78rem !important;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    font-weight: 600 !important;
}
[data-testid="stMetricValue"] {
    color: var(--ink-900) !important;
    font-size: 1.85rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em !important;
}

/* ----- Banners ---------------------------------------------------------- */
.banner {
    padding: 0.95rem 1.15rem;
    border-radius: 10px;
    font-weight: 500;
    margin: 1.4rem 0 1.6rem;
    display: flex;
    align-items: center;
    gap: 0.7rem;
    font-size: 0.96rem;
}
.banner-success {
    background: var(--brand-soft);
    border: 1px solid #c0e2d3;
    color: #0d5a3e;
}
.banner-success .dot {
    width: 20px; height: 20px;
    background: var(--brand);
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: white;
    font-size: 12px;
    font-weight: 700;
    flex-shrink: 0;
}
.banner-error {
    background: #fef2f2;
    border: 1px solid #fecaca;
    color: #7f1d1d;
}

/* ----- Section labels --------------------------------------------------- */
.section-label {
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 0.74rem;
    color: var(--ink-300);
    font-weight: 600;
    margin: 1.8rem 0 0.65rem;
}

/* ----- Dataframes ------------------------------------------------------- */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid var(--line);
}

/* ----- Expander --------------------------------------------------------- */
[data-testid="stExpander"] {
    border: 1px solid var(--line) !important;
    border-radius: 10px !important;
    box-shadow: none !important;
}
[data-testid="stExpander"] summary {
    font-weight: 500;
    color: var(--ink-700);
}

/* ----- Footer ----------------------------------------------------------- */
.app-footer {
    text-align: center;
    color: var(--ink-300);
    font-size: 0.83rem;
    padding-top: 3rem;
    border-top: 1px solid var(--line-soft);
    margin-top: 3rem;
}
.app-footer a {
    color: var(--ink-500);
    text-decoration: none;
}
.app-footer a:hover {
    color: var(--brand);
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Header — brand mark + hero
# ---------------------------------------------------------------------------
st.markdown(
    """
    # <div class="brand">
    #     <div class="brand-mark">R</div>
    #     <span>RankSheet</span>
    # </div>
    <div class="hero">
        <h1>Convert PSC rank lists to Excel</h1>
        <p class="subtitle">
            Drop a Kerala Public Service Commission ranked-list PDF and get a
            clean, structured Excel file — every candidate, every category,
            every count, sorted out for you.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def convert_pdf(pdf_bytes: bytes):
    """Run the converter; return (rows, meta, xlsx_bytes)."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        pdf_path = Path(tmp.name)

    progress = st.progress(0.0, text="Preparing…")

    def on_progress(current: int, total: int, msg: str) -> None:
        # leave the last 10% for the XLSX-build phase
        pct = 0.05 + 0.85 * (current / total if total else 0)
        progress.progress(pct, text=msg)

    rows, meta = extract_rank_list(pdf_path, progress_callback=on_progress)

    progress.progress(0.92, text=f"Building Excel file with {len(rows):,} rows…")
    xlsx_path = pdf_path.with_suffix(".xlsx")
    write_xlsx(rows, meta, xlsx_path)
    xlsx_bytes = xlsx_path.read_bytes()

    progress.progress(1.0, text="Done")
    time.sleep(0.25)
    progress.empty()

    # cleanup
    try:
        pdf_path.unlink()
        xlsx_path.unlink()
    except OSError:
        pass

    return rows, meta, xlsx_bytes


# ---------------------------------------------------------------------------
# File uploader
# ---------------------------------------------------------------------------
uploaded = st.file_uploader(
    "Upload PDF",
    type=["pdf"],
    label_visibility="collapsed",
    help="PDF files up to 200 MB",
)

# Reset cached result when a new file is uploaded
if uploaded is not None and st.session_state.get("last_file") != uploaded.name:
    st.session_state["last_file"] = uploaded.name
    st.session_state.pop("result", None)


# ---------------------------------------------------------------------------
# Convert + show results
# ---------------------------------------------------------------------------
if uploaded is not None:
    pdf_bytes = uploaded.getvalue()
    size_mb = len(pdf_bytes) / (1024 * 1024)

    if "result" not in st.session_state:
        try:
            rows, meta, xlsx_bytes = convert_pdf(pdf_bytes)
            st.session_state["result"] = {
                "rows": rows,
                "meta": meta,
                "xlsx_bytes": xlsx_bytes,
                "size_mb": size_mb,
                "filename": uploaded.name,
            }
        except Exception as exc:                # noqa: BLE001
            st.markdown(
                f'<div class="banner banner-error">⚠ Conversion failed: '
                f"{exc}</div>",
                unsafe_allow_html=True,
            )
            with st.expander("Error details"):
                st.exception(exc)
            st.stop()

    result = st.session_state["result"]
    rows = result["rows"]
    meta = result["meta"]
    xlsx_bytes = result["xlsx_bytes"]

    if not rows:
        st.markdown(
            '<div class="banner banner-error">⚠ No candidates found in this '
            "PDF. Is it a Kerala PSC ranked list with the standard "
            "<em>Rank · Reg.No · Name · DOB · Commy · Remarks</em> layout?"
            "</div>",
            unsafe_allow_html=True,
        )
        st.stop()

    # ---- Success banner ----------------------------------------------------
    n_sections = len({r["Category"] for r in rows})
    st.markdown(
        f"""
        <div class="banner banner-success">
            <span class="dot">✓</span>
            <span>Converted <strong>{len(rows):,}</strong> candidates across
            <strong>{n_sections}</strong> sections</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ---- Stats -------------------------------------------------------------
    c1, c2, c3 = st.columns(3)
    c1.metric("Candidates", f"{len(rows):,}")
    c2.metric("Sections", n_sections)
    c3.metric("PDF size", f"{result['size_mb']:.2f} MB")

    # ---- Download ----------------------------------------------------------
    st.markdown('<div class="section-label">Download</div>',
                unsafe_allow_html=True)
    download_name = Path(result["filename"]).stem + "_converted.xlsx"
    st.download_button(
        label="Download Excel file",
        data=xlsx_bytes,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ---- Preview -----------------------------------------------------------
    st.markdown('<div class="section-label">Preview · first 25 rows</div>',
                unsafe_allow_html=True)
    df = pd.DataFrame(rows)
    if "Rank" in df.columns:
        df["Rank"] = df["Rank"].astype(int)
    st.dataframe(df.head(25), use_container_width=True, hide_index=True,
                 height=320)

    # ---- Section breakdown -------------------------------------------------
    st.markdown('<div class="section-label">Section breakdown</div>',
                unsafe_allow_html=True)
    counts = Counter(r["Category"] for r in rows)
    breakdown = pd.DataFrame(
        [(cat, cnt) for cat, cnt in counts.items()],
        columns=["Category", "Count"],
    )
    st.dataframe(breakdown, use_container_width=True, hide_index=True,
                 height=min(36 * len(breakdown) + 38, 460))

    # ---- Metadata ----------------------------------------------------------
    if meta:
        with st.expander("Document metadata"):
            pretty_keys = {
                "ranked_list_no": "Ranked List No.",
                "category_no":    "Category No.",
                "omr_test_date":  "OMR Test Date",
                "in_force_from":  "In Force From",
            }
            for k, v in meta.items():
                st.write(f"**{pretty_keys.get(k, k)}:** {v}")


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="app-footer">
        Built with Streamlit &nbsp;·&nbsp; Free &amp; open source
    </div>
    """,
    unsafe_allow_html=True,
)
