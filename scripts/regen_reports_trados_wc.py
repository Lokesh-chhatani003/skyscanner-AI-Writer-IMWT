# -*- coding: utf-8 -*-
"""Regenerate LQA Excel scorecards using Trados-style WC from input/ DOCX bodies."""
from __future__ import annotations

import json
import logging
import shutil
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parents[1]
import sys

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lqa_pipeline import (  # noqa: E402
    compute_ewt,
    parse_docx_sections_to_df,
    write_language_report,
)

INPUT = ROOT / "input"
OUTPUT = ROOT / "output"
SCORECARDS = OUTPUT / "scorecards"

LANG_GLOBS = {
    "JP": "*-jp-*.docx",
    "TW": "*-tw-*.docx",
    "TH": "*-th-*.docx",
}


def _input_ewt(lang: str) -> tuple[int, int, str]:
    files = sorted(INPUT.glob(LANG_GLOBS[lang]))
    if not files:
        raise FileNotFoundError(f"No input files for {lang} matching {LANG_GLOBS[lang]}")
    parts: list[str] = []
    sections = 0
    for path in files:
        rows = parse_docx_sections_to_df(path)
        sections += len(rows)
        parts.extend(r.get("content") or "" for r in rows)
    text = "\n\n".join(parts)
    # Word-based Asian tokenization (Janome / jieba / pythainlp)
    return compute_ewt(text, lang_code=lang), sections, text


def _latest_final_json(lang: str) -> Path:
    matches = sorted((OUTPUT / lang).glob("LQA_Report_*_final.json"))
    if not matches:
        raise FileNotFoundError(f"No final.json for {lang}")
    return matches[-1]


def regen(lang: str) -> Path:
    ewt, sections, evaluated_text = _input_ewt(lang)
    final_path = _latest_final_json(lang)
    data = json.loads(final_path.read_text(encoding="utf-8"))
    findings = data.get("review_findings") or []
    ofb = data.get("overall_feedback") or {}
    extras = data.get("extras") if isinstance(data.get("extras"), dict) else {}

    # Best-effort remove prior Excel reports (skip if locked / open in Excel/IDE)
    for old in list((OUTPUT / lang).glob("LQA_Report_*.xlsx")) + list(
        SCORECARDS.glob(f"LQA_Report_{lang}_*.xlsx")
    ):
        try:
            old.unlink()
            logging.info("Removed %s", old)
        except OSError as exc:
            logging.warning("Could not remove %s (%s); writing a new file alongside it.", old, exc)

    scope = (
        f"DOCX heading-based batches ({sections} batches; "
        f"Trados word-based WC from input/; source={final_path.name})"
    )
    write_language_report(
        lang_code=lang,
        evaluated_scope_label=scope,
        evaluated_text=evaluated_text,
        issues_with_url=findings,
        output_dir=OUTPUT / lang,
        prefix="LQA_Report",
        pass_threshold=90,
        save_raw_json=False,
        overall_feedback_override=ofb if ofb else None,
        extra_dump={
            **extras,
            "wc_method": "trados_word_based_asian_tokenization",
            "wc_engines": {"JP": "janome", "TW": "jieba", "TH": "pythainlp_newmm"},
            "wc_source": "input_docx_section_bodies",
            "wc_ewt": ewt,
            "regenerated_at": datetime.now(ZoneInfo("Africa/Cairo")).isoformat(),
        },
        ewt_override=ewt,
    )

    new_xlsx = sorted((OUTPUT / lang).glob("LQA_Report_*.xlsx"))[-1]
    SCORECARDS.mkdir(parents=True, exist_ok=True)
    dest = SCORECARDS / new_xlsx.name
    shutil.copy2(new_xlsx, dest)
    logging.info("[%s] EWT=%s findings=%s -> %s", lang, ewt, len(findings), new_xlsx.name)
    return new_xlsx


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    for lang in ("JP", "TW", "TH"):
        path = regen(lang)
        print(f"{lang}: {path}")


if __name__ == "__main__":
    main()
