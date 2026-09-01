# Skyscanner SEO AI-LQA Pipeline (DOCX)

## Project Overview
This repository runs a 3-agent LQA pipeline for localized Skyscanner content.

**Primary input (current):** `input/*.docx` (MosAIQ H1/H2).
**Also supported:** CSV and legacy HTML by changing `parsing.input_format`.

It produces per-language **scorecard** (Excel + JSON) and a **client-ready edited file** (CSV or DOCX matching the input) under `output/<LANG>/`.

The main entrypoint is `lqa_pipeline.py`.

**Current ticket languages:** SA, CZ, GR, HU, PL, PT, TH, TR

## 3-Agent Architecture
1. Agent 1 (Draft Auditor) — proposes objective LQA findings per H2 batch.
2. Agent 2 (Validator) — finalizes findings and returns full-section `corrected_content`.
3. Agent 3 (Language Consolidator) — language-level `overall_feedback` for the scorecard.

## End-to-End Flow
1. Load `config.yaml`.
2. Resolve Gemini API key (`GEMINI_API_KEY` first, `config.yaml` fallback).
3. Load prompts, locale packs (`prompts/<LANG>_instructions.txt`), and glossary v2.
4. Discover `input/*.docx` and infer language from filename (`-sa-`, `-cz-`, `-gr-`, …).
5. Parse DOCX into H2 batches.
6. Run Agent 1 and Agent 2 in windows; retry failed Agent 2 batches once.
7. Run Agent 3 once per language.
8. Write scorecard XLSX + JSON and edited DOCX into `output/<LANG>/` (e.g. `output/BR/`).

## Prerequisites
- Python `3.10+` (validated on `3.12` / `3.14`)
- Internet access (Gemini API calls)
- Gemini API key
- Pip package installation ability

## Repository Layout
- `lqa_pipeline.py`: Main pipeline script.
- `config.yaml`: Runtime configuration.
- `requirements.txt`: Python dependencies.
- `input/`: Active input folder. Only `*.docx` here are processed (when `parsing.input_format: docx`).
- `prompts/`: System/user prompt templates and `<LANG>_instructions.txt` (Jan 2026 style distillations).
- `reference/Skyscanner_Glossary_v2.xlsx`: Multilingual glossary.
- `Style Guide Jan 2026-*/`: Source style-guide DOCXs (authoritative reference for locale packs).
- `training/`: Client unedited/edited sample pair (used to gap-fill prompts; not injected as few-shot).
- `output/`: Run logs at root; per-locale deliverables under `output/<LANG>/`.

## Installation (Windows PowerShell)
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks activation:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

## Configuration
### API key
1. `GEMINI_API_KEY` environment variable (recommended)
2. `config.yaml` → `api_key` (non-placeholder)

```powershell
$env:GEMINI_API_KEY = "<your-real-gemini-api-key>"
```

### Key config sections
- `lang_map`: ticket/runtime codes → locale codes (for example, SA → ar-SA and GR → el-GR)
- `file_names.glossary`: `Skyscanner_Glossary_v2.xlsx`
- `file_names.edited_output_prefix`: `Edited`
- `parsing.input_format`: `docx` (or `html` / `auto`)
- `parsing.heading_level_preference`: `2` (H2 batches under H1 URL)
- `batching` / `limits` / `scoring`: throughput and MQM threshold

## Input Preparation
Place processable files in `input/`.

### CSV (optional)
Required columns:
- `page_url` — full page URL (also used as Candidate in the scorecard)
- `content_html` — HTML body; split into batches on `<h3>` (configurable via `parsing.csv_heading_tag`)

Example filename language token: `NEW_2026-2027-pt-car-hire-airport.csv` → **PT** (`pt-PT`).

### DOCX / HTML
Still supported when `parsing.input_format` is `docx`, `html`, or `auto`.

Supported language inference:
1. Short-code prefix: `FR_1.csv`
2. Delimited short code in stem: `...-pt-car-hire...`, `...-us-flights...`
3. ISO from `lang_map` locale (`-de-`, `-fr-`, …)

## Run
```powershell
.\.venv\Scripts\python.exe lqa_pipeline.py
```

### Expected log progression
- `File logging enabled -> ...`
- `===== Language XX: parsing input files -> sections/batches =====`
- `[XX] Total batches to process: N`
- `[XX] Agent 1 / Agent 2 windows ...`
- `[XX] Agent 3 (consolidator) - starting`
- `[XX] Wrote report -> ...xlsx`
- `Wrote edited DOCX -> ...`
- `DOCX LQA pipeline finished successfully (3-agent).`

## Outputs
| Artifact | Location |
|----------|----------|
| Scorecard Excel | `output/<LANG>/LQA_Report_<LANG>_<timestamp>.xlsx` |
| Final JSON | `output/<LANG>/LQA_Report_<LANG>_<timestamp>_final.json` |
| Edited file | `output/<LANG>/Edited_<LANG>_<input-stem>_<timestamp>.csv` (or `.docx`) |
| Run log | `output/lqa_run_<timestamp>.log` (shared, all locales) |

Example: European Portuguese CSV → `output/PT/`.

Agent 2 JSON includes `corrected_content` (full section body) used to rebuild the edited DOCX.

## Troubleshooting
1. **Missing API key** — set `GEMINI_API_KEY` or a real `api_key` in config.
2. **No files discovered** — put `*.docx` in `input/` with inferable language tokens.
3. **Glossary empty for a language** — confirm `reference/Skyscanner_Glossary_v2.xlsx` sheet `Multilingual Glossary` has the locale column (`pt-BR`, `de-DE`, …).
4. **Transient Gemini 503** — retry; lower `max_batches_per_window`; increase `window_pause_seconds`.

## Security
Do not commit real API keys. Prefer environment variables.
