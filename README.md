# Kerala PSC Rank List Converter

A free web app that converts Kerala Public Service Commission ranked-list PDFs into structured Excel files. Drop a PDF, get an XLSX — no sign-up, no watermarks.

![status: working](https://img.shields.io/badge/status-working-green) ![license: MIT](https://img.shields.io/badge/license-MIT-blue)

## Features

- Drag-and-drop PDF upload
- Real progress bar (page-by-page, not faked)
- Every candidate routed to the correct category (Main List, each Supplementary community, each Differently Abled list)
- Live preview of the first 25 rows + per-section breakdown
- One-click Excel download with summary sheet, freeze-pane and auto-filter
- Files are processed in memory, never stored on the server

## Run locally

```bash
git clone <your-repo-url> psc-rank-converter
cd psc-rank-converter
pip install -r requirements.txt
streamlit run app.py
```

Open <http://localhost:8501>.

## Deploy free — Streamlit Community Cloud (recommended)

This is the easiest path. The app stays up 24/7 with no cold starts, runs on a `.streamlit.app` subdomain, and is genuinely free for public repos.

**Step 1.** Create a free account at <https://github.com> if you don't already have one.

**Step 2.** Create a new public repository (call it anything, e.g. `psc-rank-converter`). Upload these four files into the repo root:

```
app.py
psc_pdf_to_xlsx.py
requirements.txt
.streamlit/config.toml
```

You can do this by clicking *Add file → Upload files* on GitHub's web interface — no git command line needed.

**Step 3.** Go to <https://share.streamlit.io>, sign in with your GitHub account, click **New app**, and pick:

- **Repository:** `your-username/psc-rank-converter`
- **Branch:** `main`
- **Main file path:** `app.py`

**Step 4.** Click **Deploy**. The first build takes 1–3 minutes; afterwards your app is live at `https://your-username-psc-rank-converter.streamlit.app` (or a custom URL you choose). Every push to `main` redeploys automatically.

## Deploy free — Hugging Face Spaces (alternative)

If you'd rather not use GitHub, Hugging Face Spaces accepts direct file uploads.

1. Create a free account at <https://huggingface.co>.
2. Click your profile → **New Space**.
3. Name it, choose **Streamlit** as the SDK, **Free CPU** as the hardware, and **Public** as visibility.
4. In the new Space's **Files** tab, upload the same four files listed above.
5. The Space builds automatically and is live at `https://huggingface.co/spaces/your-username/your-space-name`.

## How it works

The converter has been built and tested against the [Assistant Salesman / Ernakulam ranked list (382/2026/DOE)](https://www.keralapsc.gov.in/) and reproduces all 600 candidates exactly. The technique is summarized in the docstring at the top of `psc_pdf_to_xlsx.py`; the short version:

1. **Detect the header on every page** by finding the line containing `Rank Reg.No Name DOB Commy Remarks` — this adapts automatically if the PSC adjusts margins or paper size.
2. **Compute column boundaries** using a weighted formula `(2·prev_x1 + next_x0) / 3` that biases each boundary toward the next column, keeping dates tight while giving the Community column room to absorb wrap-continuation text.
3. **Sanity-check the DOB column** — any non-date tokens are misclassified neighbours, so pre-date words go back to Name and post-date words forward to Community. This rescues edge cases like long names with trailing initials.
4. **Split glued tokens** like `"16/07/2001OBC-VILAKKITHALA"` which pdfplumber occasionally produces.
5. **Smart wrap-joining** — fragments of a wrapped cell are joined with a space, except when the previous fragment ends with `-` (e.g. `OBC-` + `VEERASAIVAS`) or sits inside an unclosed `[` (e.g. `VEERASAIVAS[YOG` + `ISYOGEESWARA]`).
6. **Recognise every section header** (Main List, the 13 supplementary community lists, the 4 Differently Abled categories) and attach the right Category label to each row.
7. **Harvest metadata** (Ranked List No., Cat. No., OMR test date, in-force date) from page 1 for the Summary sheet.

## File layout

```
psc-rank-converter/
├── app.py                    # Streamlit UI
├── psc_pdf_to_xlsx.py        # Converter library (also a CLI)
├── requirements.txt          # Python dependencies
├── .streamlit/
│   └── config.toml           # Theme + server config
└── README.md
```

## Use as a CLI

The converter also works on its own from the command line:

```bash
python psc_pdf_to_xlsx.py input.pdf [output.xlsx]
```

If you omit `output.xlsx` it writes `<input>_converted.xlsx` next to the PDF.

## License

MIT.
