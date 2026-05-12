"""
Kerala PSC Rank List Converter — Streamlit web UI
==================================================

A single-page web app that lets anyone drop a Kerala PSC ranked-list PDF
in their browser and download the structured Excel file. Run locally with

    streamlit run app.py

or deploy free on Streamlit Community Cloud (see README.md).
"""
from __future__ import annotations

import io
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
    page_title="Kerala PSC Rank List Converter",
    page_icon="📋",
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
# Custom CSS — gives the page a clean, premium look
# ---------------------------------------------------------------------------
CSS = """
<style>
/* Hide the default Streamlit chrome */
#MainMenu, header, footer {visibility: hidden;}
.stDeployButton {display: none;}

/* Page background */
.stApp {
    background: linear-gradient(180deg, #f8fafc 0%, #eef2f7 100%);
    min-height: 100vh;
}

/* Tighten top padding */
.block-container {
    padding-top: 2rem;
    padding-bottom: 4rem;
    max-width: 780px;
}

/* Hero */
.hero {
    text-align: center;
    margin-bottom: 1.5rem;
}
.hero-badge {
    display: inline-block;
    background: #e0e7ff;
    color: #4338ca;
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.04em;
    margin-bottom: 1rem;
}
.hero h1 {
    font-size: 2.3rem !important;
    font-weight: 700;
    color: #0f172a;
    margin: 0 0 0.6rem 0 !important;
    letter-spacing: -0.025em;
    line-height: 1.15;
}
.hero p.subtitle {
    font-size: 1.05rem;
    color: #475569;
    max-width: 540px;
    margin: 0 auto;
    line-height: 1.55;
}

/* Style Streamlit's file uploader to look like a premium drop zone */
section[data-testid="stFileUploadDropzone"] {
    background: white !important;
    border: 2px dashed #cbd5e1 !important;
    border-radius: 14px !important;
    padding: 2.5rem 1.5rem !important;
    transition: all 0.2s ease;
}
section[data-testid="stFileUploadDropzone"]:hover {
    border-color: #1F4E79 !important;
    background: #f0f5fa !important;
}
section[data-testid="stFileUploadDropzone"] button {
    background: #1F4E79 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.5rem 1.25rem !important;
}

/* Download button */
.stDownloadButton button {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%) !important;
    color: white !important;
    border: none !important;
    padding: 0.85rem 2rem !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 1.02rem !important;
    box-shadow: 0 2px 6px rgba(31, 78, 121, 0.25) !important;
    transition: all 0.18s ease !important;
    width: 100%;
}
.stDownloadButton button:hover {
    transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(31, 78, 121, 0.35) !important;
}

/* Progress bar */
.stProgress > div > div > div > div {
    background: linear-gradient(90deg, #2E75B6 0%, #1F4E79 100%) !important;
}

/* Metric cards */
[data-testid="stMetric"] {
    background: white;
    padding: 1rem 1.25rem;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
[data-testid="stMetricLabel"] {
    color: #64748b !important;
    font-size: 0.82rem !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    font-weight: 600 !important;
}
[data-testid="stMetricValue"] {
    color: #0f172a !important;
    font-size: 1.7rem !important;
    font-weight: 700 !important;
}

/* Success / error banners */
.banner {
    padding: 1rem 1.25rem;
    border-radius: 10px;
    font-weight: 500;
    margin: 1.2rem 0;
    display: flex;
    align-items: center;
    gap: 0.6rem;
}
.banner-success {
    background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%);
    border-left: 4px solid #16a34a;
    color: #14532d;
}
.banner-error {
    background: #fee2e2;
    border-left: 4px solid #dc2626;
    color: #7f1d1d;
}

/* Section labels */
.section-label {
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-size: 0.78rem;
    color: #64748b;
    font-weight: 600;
    margin: 1.6rem 0 0.6rem;
}

/* Tables (st.dataframe) */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #e2e8f0;
}

/* Footer */
.app-footer {
    text-align: center;
    color: #94a3b8;
    font-size: 0.85rem;
    padding-top: 2rem;
}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Hero header
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
       
        <h1>Kerala PSC Rank List Converter</h1>
        <p class="subtitle">
            Drop a Kerala Public Service Commission ranked-list PDF below and
            get a clean, structured Excel file in seconds — every candidate,
            every category, every count, all sorted out for you.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def convert_pdf(pdf_bytes: bytes, original_name: str):
    """Run the converter and return (rows, meta, xlsx_bytes, n_pages)."""
    # Save PDF to a temp file because pdfplumber wants a path
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        pdf_path = Path(tmp.name)

    progress = st.progress(0.0, text="Initialising…")

    def on_progress(current: int, total: int, msg: str) -> None:
        # reserve the last 10% for the XLSX build step
        pct = 0.05 + 0.85 * (current / total if total else 0)
        progress.progress(pct, text=msg)

    rows, meta = extract_rank_list(pdf_path, progress_callback=on_progress)

    progress.progress(0.92, text=f"Building Excel file with {len(rows)} rows…")
    xlsx_path = pdf_path.with_suffix(".xlsx")
    write_xlsx(rows, meta, xlsx_path)
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()

    progress.progress(1.0, text="Done!")
    time.sleep(0.3)
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
    " ",
    type=["pdf"],
    label_visibility="collapsed",
    help="Up to 200 MB. Your file is processed in memory and never stored.",
)

# Reset state if a new file is uploaded
if uploaded is not None:
    if st.session_state.get("last_file") != uploaded.name:
        st.session_state["last_file"] = uploaded.name
        st.session_state.pop("result", None)


# ---------------------------------------------------------------------------
# Convert + show results
# ---------------------------------------------------------------------------
if uploaded is not None:
    pdf_bytes = uploaded.getvalue()
    size_mb = len(pdf_bytes) / (1024 * 1024)

    # Convert only once per file (cache result in session state)
    if "result" not in st.session_state:
        try:
            rows, meta, xlsx_bytes = convert_pdf(pdf_bytes, uploaded.name)
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
                f'{exc}</div>',
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
        f'<div class="banner banner-success">✓ Converted '
        f"<strong>{len(rows):,}</strong> candidates across "
        f"<strong>{n_sections}</strong> sections</div>",
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
        label="⬇  Download Excel file",
        data=xlsx_bytes,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ---- Preview -----------------------------------------------------------
    st.markdown('<div class="section-label">Preview (first 25 rows)</div>',
                unsafe_allow_html=True)
    df = pd.DataFrame(rows)
    # Make Rank an int so it sorts naturally
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

else:
    # Empty state — show a small "how it works" panel below the uploader
    st.markdown(
        """
        <div style="text-align:center; color:#64748b; font-size:0.95rem;
                    margin-top:1.5rem;">
            Drop a PDF above to begin. Files are processed in memory and never
            stored on the server.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
st.markdown(
    """
    <div class="app-footer">
        Built with Streamlit · Free &amp; open source
    </div>
    """,
    unsafe_allow_html=True,
)
