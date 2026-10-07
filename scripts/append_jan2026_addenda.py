# -*- coding: utf-8 -*-
"""Append Style Guide Jan 2026 SEO/editorial addenda to locale instruction packs.

Does NOT delete or rewrite existing content. If an addendum marker is already
present, the file is left unchanged.
"""
from __future__ import annotations

from pathlib import Path

PROMPTS = Path(__file__).resolve().parents[1] / "prompts"
MARKER = "## SEO / Editorial LQA Addendum (Jan 2026)"

# Full base packs for locales that had no prior ticket-language file.
NEW_BASE: dict[str, str] = {
    "HR": """# Croatian (hr-HR) Style Guide for Skyscanner

## Critical requirements
- Address the user with formal/friendly **vi** (2nd person plural) when Skyscanner addresses, encourages, or instructs the user.
- Use **ti** (2nd person singular) only when the user is instructing the system.
- Prefer gender-neutral constructions where natural. Helpful > geeky > playful.
- The global project-specific exclusion in the master prompt takes priority over any name or geo-name guidance.

## Language and typography
- Follow Hrvatski pravopis. Use Croatian quotation marks `„…”` and the ellipsis character `…` (not three dots).
- Do not use dashes (`-`/`–`) to mark a pause; rephrase, start a new sentence, or use a colon/semicolon.
- Translate ampersand as `i`. Avoid excessive exclamation marks.
- Use sentence case; do not end headings with a full stop.
- Decimal separator is a comma; thousands separator is a point (`1.000.000,00`). Use a non-breaking space before units, `%`, and currency.

## Currency and units (home market: Croatia)
- Home-market currency: euro (`€`), symbol separated from the number with a non-breaking space. Use it for every price whatever the destination, never the destination's own currency.
- Home-market measurement system: metric (km, m, °C, l, kg). Convert destination measurements to metric and show only the metric value.
- Exception — speed limits and other posted road signage at the destination: keep the destination's unit first, then the metric value in parentheses. A speed limit shown in only one unit, or with km/h first, → **Compliance · Major**.

## Technical content
- Preserve placeholders, variables, tags, numbers, and product tokens exactly.
- Do not duplicate the airport word when it is already part of the source name.
- Keep Skyscanner unchanged.

## Content and compliance
- Use natural SEO language; no keyword stuffing or preference-only rewrites.
- Do not externally fact-check laws, prices, limits, or policies.
- Do not claim Skyscanner provides discounts. Use “best” only with a supplied measurable basis.
- Avoid greenwashing and absolute claims that travel is green, clean, eco-friendly, or sustainable.
""",
    "MX": """# Mexican Spanish (es-MX) Style Guide for Skyscanner

## Critical requirements
- Use informal **tú** for UI, marketing, SEO, and help. Use formal **usted** only for legal/policy text.
- Prefer natural gender-neutral structures when they read well; do not force unnatural neutrality. Helpful > geeky > playful.
- The global project-specific exclusion in the master prompt takes priority over any name or geo-name guidance.

## Language and typography
- Use English-style quotation marks `“…”` (not angular `«»`). Period goes outside closing quotes/brackets when they end a sentence.
- No space before an ellipsis. Use exclamation marks sparingly.
- Prefer sentence case; do not end headings/buttons/menu labels with a full stop.
- Decimal separator is a comma.

## Currency and units (home market: Mexico)
- Home-market currency: Mexican peso (`$` before the amount, no space; ISO `MXN` when needed). Use home-market currency for every price whatever the destination unless the source explicitly requires otherwise.
- Home-market measurement system: metric (km, m, °C, l, kg). Convert destination measurements to metric and show only the metric value.
- Exception — speed limits and other posted road signage at the destination: keep the destination's unit first, then the metric value in parentheses → **Compliance · Major** if violated.

## Technical content
- Preserve placeholders, variables, tags, numbers, and product tokens exactly. Use workarounds when gender agreement depends on an unknown placeholder value.
- Do not duplicate the airport word when it is already part of the source name.
- Keep Skyscanner unchanged.

## Content and compliance
- Use natural SEO language; no keyword stuffing or preference-only rewrites.
- Do not externally fact-check laws, prices, limits, or policies.
- Do not claim Skyscanner provides discounts. Use “best” only with a supplied measurable basis.
- Avoid greenwashing and absolute claims that travel is green, clean, eco-friendly, or sustainable.
""",
    "CH": """# Chinese Simplified (zh-CN) Style Guide for Skyscanner

## Critical requirements
- Write clear, natural Simplified Chinese. Helpful > geeky > playful; avoid stiff literal translation and forced wordplay.
- Prefer short active sentences and direct CTAs.
- The global project-specific exclusion in the master prompt takes priority over any name or geo-name guidance.

## Language and typography
- Use Simplified Chinese characters consistently (never mix Traditional).
- Use Chinese punctuation appropriately. Prefer commas/colons over long em dashes (`——`) when space is tight.
- Put a space between Chinese characters and Latin letters/emojis where needed for readability.
- Do not end headings with a full stop. Follow source formatting tags (including italics tags) even if Chinese has no italic glyph.
- Avoid multiple exclamation marks.

## Currency and units (home market: Mainland China)
- Prefer CNY/RMB presentation already established in the source when present. Do not invent currency conversions without evidence in the text.
- Always convert measure units to metric (e.g. inch → cm/m) unless the unit is inside a variable (e.g. `@@miles@@`). Do not convert airline loyalty “miles”.
- Exception — posted road signage at the destination: keep destination unit first, metric in parentheses when both are required by project rules.

## Technical content
- Preserve placeholders, variables, tags, numbers, and product tokens exactly; do not add spaces inside tags.
- URLs stay lowercase Latin with hyphens.
- Airport names: prefer official/popular Simplified translations (official site → major OTAs → Google Maps); if none, keep Latin source.
- Keep Skyscanner unchanged.

## Content and compliance
- Use natural SEO language; no keyword stuffing or preference-only rewrites.
- Do not externally fact-check laws, prices, limits, or policies.
- Do not claim Skyscanner provides discounts. Use “best” only with a supplied measurable basis.
- Avoid greenwashing and absolute claims that travel is green, clean, eco-friendly, or sustainable.
""",
    "TW": """# Chinese Traditional (zh-TW) Style Guide for Skyscanner

## Critical requirements
- Write clear, natural Traditional Chinese for Taiwan. Helpful > geeky > playful; avoid stiff literal translation and forced wordplay.
- Prefer short active sentences and direct CTAs.
- The global project-specific exclusion in the master prompt takes priority over any name or geo-name guidance.

## Language and typography
- Use Traditional Chinese characters consistently (never mix Simplified).
- Use Chinese quotation marks when referring to UI labels or specific content.
- Use a single horizontal ellipsis `…` in UI when needed. Do not end titles with a full stop unless the punctuation is intentional tone (e.g. `?`/`!`).
- Avoid multiple exclamation marks.

## Currency and units (home market: Taiwan)
- Follow standardized currency terms from the style guide / source. Do not invent conversions without evidence in the text.
- Convert measure units to metric unless the unit is inside a variable. Do not convert airline loyalty “miles”.
- Exception — posted road signage at the destination: keep destination unit first, metric in parentheses when both are required by project rules.

## Technical content
- Preserve placeholders, variables, tags, numbers, and product tokens exactly.
- URLs stay lowercase Latin with hyphens.
- Airport/geo names: prefer Taiwan Traditional Chinese usage (set search/locale to Taiwan); if no established translation, keep Latin source.
- Keep Skyscanner unchanged.

## Content and compliance
- Use natural SEO language; no keyword stuffing or preference-only rewrites.
- Do not externally fact-check laws, prices, limits, or policies.
- Do not claim Skyscanner provides discounts. Use “best” only with a supplied measurable basis.
- Avoid greenwashing and absolute claims that travel is green, clean, eco-friendly, or sustainable.
""",
    "AE": """# Arabic (ar-SA / AE market) Style Guide for Skyscanner

## Critical requirements
- Use clear Modern Standard Arabic: friendly, respectful, concise, and natural. Helpful > geeky > playful; avoid slang, corporate language, forced jokes, and excessive exclamation marks.
- Keep the register and form of address polite and consistent. Prefer active voice and direct imperative CTAs.
- The global project-specific exclusion in the master prompt takes priority over any name or transliteration guidance.

## Language and typography
- Write right-to-left Arabic using Arabic punctuation where appropriate: comma `،` and question mark `؟`.
- Do not put a space before punctuation; use one space after it. Use straight double quotation marks consistently.
- Do not overuse exclamation marks; prefer comma, colon, or period when English overuses `!`.
- Avoid hyphens/dashes to join or break Arabic words the English way.
- Use Arabized (Latin-derived) month/day names for consistency across regions.
- Use sentence case for any Latin-script text.

## Currency and units (home market: UAE for AE content)
- Home-market currency for this AE batch: the UAE dirham (`د.إ` / `AED`). Use it for every price whatever the destination, never the destination's own currency, when a price must be localized.
- Home-market measurement system: metric (كم، م، °C، لتر، كجم). Convert destination measurements to metric and show only the metric value.
- Exception — speed limits and other posted road signage at the destination: keep the destination's unit first, then the metric value in parentheses. Keep the Latin-script destination unit unaltered inside the RTL sentence.

## Technical content
- Preserve placeholders, variables, tags, numbers, and product tokens exactly; they may move only when Arabic grammar requires it.
- Keep URLs in lowercase basic Latin characters, with hyphens and no spaces or special characters.
- Do not translate city names unless a commonly known Arabic form exists; otherwise keep the source spelling for searchability.
- Airport names: transliterate proper nouns; translate descriptive words (e.g. International → دولي).
- Keep Skyscanner recognizable and unaltered.

## Content and compliance
- Use natural SEO language; do not keyword-stuff or rewrite correct text for preference.
- Do not externally fact-check laws, prices, limits, or policies.
- Do not claim Skyscanner provides discounts. Use neutral wording about offers or price reductions.
- Use “best” or other superlatives only when the supplied text includes a clear measurable basis.
- Avoid greenwashing and absolute claims that a flight is green, clean, eco-friendly, or sustainable.
""",
}

# Alias packs: start from an existing locale file, then append addendum.
ALIASES_FROM: dict[str, str] = {
    "JP": "JA",
    "KR": "KO",
}

ADDENDA: dict[str, str] = {
    "PL": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Polish). Additive — do not discard earlier rules.

### Register (mandatory)
- Do **not** use `Pan`/`Pani`. Use capitalized direct-address forms (`Ty`, `Tobie`, `Ciebie`, `Wy`, `Wam`, `Was`, etc.).
- Prefer gender-neutral wording; avoid slash forms when a neutral rewrite is natural.
- Skyscanner may be inflected when Polish grammar requires it (`Skyscannera`, `Skyscannerze`).

### Typography reinforcements
- Polish quotes `„…”`; ellipsis `…`; en dash `–` (never em dash `—`).
- No full stop at the end of headings. Avoid excessive `!`.

### Geo naming
- Official/common local names; consistent spelling; no invented forms.
- Never duplicate “lotnisko/airport” when already in the name.

### SEO body readability (objective Style only)
- If one SEO body paragraph packs multiple distinct ideas and is objectively hard to scan, flag **Style** and split into shorter paragraphs in `updated_sentence` / `corrected_content`.
- No synonym-only or preference rewrites.
""",
    "CZ": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Czech). Additive — do not discard earlier rules.

### Register (mandatory)
- Friendly formal **V-form** consistently. Write `vám`/`váš`/`vás` lowercase except sentence-initial.
- Avoid bureaucratic, bookish, or archaic phrasing while staying respectful.

### Geo naming
- Official/common Czech forms; consistent across the document.
- Never duplicate “letiště/airport” when already in the name.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "TH": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Thai). Additive — do not discard earlier rules.

### Register
- Friendly, respectful “helpful travel pal” tone; keep pronoun/particle formality consistent across the document.
- Do not switch register without a documented legal/policy exception.

### Geo naming
- Official/common Thai usage; consistent transliteration; keep brand names in Latin when expected.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "TR": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Turkish). Additive — do not discard earlier rules.

### Register (mandatory)
- Standard address is friendly formal **Siz** with modern imperatives (`-ın/-in`: yapın, kullanın). Flag informal `sen` and archaic `-ınız/-iniz` as **Compliance**.
- Country name: **Türkiye** (not Turkey).

### Typography reinforcements
- Straight quotes `"…"`; ellipsis `…`; no em dash — rephrase with comma/colon/parentheses or a new sentence.
- Sentence case; no full stop at end of headings. Dates: `DD.MM.YYYY`.

### Geo / cultural
- Official/common local names; consistent spelling.
- Avoid sensitive religious/political references not present in source.

### SEO body readability (objective Style only)
- Prefer splitting objectively overlong multi-idea SEO body paragraphs (**Style**) over comma-heavy run-ons.
- No synonym-only preference edits.
""",
    "AE": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Arabic). Additive — do not discard earlier rules.

### Register / punctuation
- MSA with polite consistent address; Arabic comma `،` and question mark `؟`.
- Do not mirror English exclamation-mark habits.

### Geo naming
- Prefer commonly used Arabic forms; if no established Arabic form, keep the source Latin spelling (searchability).
- Airport: transliterate proper nouns; translate descriptive words (e.g. International → دولي).

### Compliance reminders
- Neutral discount/deal wording; no Skyscanner-offered discounts.
- “Best” only with a measurable basis/disclaimer in the text.
- No absolute green/eco/sustainable claims.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "HR": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Croatian). Additive — do not discard earlier rules.

### Register (mandatory)
- Skyscanner → user: **vi**. User → system: **ti**. Keep consistent.
- Croatian quotes `„…”`; no pause-dashes; ampersand → `i`.

### Currency
- Home currency EUR with non-breaking space before `€`.

### Geo naming
- Official/common local names; consistent spelling; no over-translation of expected originals.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "MX": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Spanish Mexico). Additive — do not discard earlier rules.

### Register (mandatory)
- Informal **tú** for SEO/UI/marketing/help. **Usted** only in legal/policy. Flag wrong register as **Compliance**.
- Quotes: English-style `“…”` (not `«»`).

### Currency
- Mexican peso: `$` before amount, no space (ISO `MXN` when needed).

### Geo naming
- Official/common Mexican Spanish forms; consistent spelling.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "CH": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Chinese Simplified). Additive — do not discard earlier rules.

### Script / readability
- Simplified only. Prefer short active SEO sentences; natural search language, no keyword stuffing.
- Metric conversions required for measures (except variables / airline miles).

### Airport / geo
- Prefer official/popular Simplified names; if none, keep Latin source.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "TW": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Chinese Traditional). Additive — do not discard earlier rules.

### Script / locale
- Traditional Chinese for Taiwan only. Prefer Taiwan-local geo/airport naming; if none, keep Latin source.
- Natural SEO language; no keyword stuffing.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "JP": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Japanese). Additive — do not discard earlier rules.

### Units / currency
- Convert imperial measures to metric with standard rounding; no space before `km` in examples like `5.8km`.
- Do **not** convert airline loyalty miles. Do not localize currency unless the job explicitly requires it.

### SEO
- Natural Japanese search language; no keyword stuffing. Spaces around keywords only if still natural.
- Clear descriptive headings; concise helpful meta-style titles when present.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "KR": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Korean). Additive — do not discard earlier rules.

### Register
- Friendly respectful honorifics; omit explicit “you” when honorific verbs already imply it.
- Sparing `!`; avoid `!!` / `!!!`, especially in notices/errors/support tone.

### SEO
- Natural Korean search language; no keyword stuffing; clear descriptive headings.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    # Also append to source packs that JP/KR alias from, if they lack addenda.
    "JA": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Japanese). Additive — do not discard earlier rules.

### Units / currency
- Convert imperial measures to metric with standard rounding; no space before `km` in examples like `5.8km`.
- Do **not** convert airline loyalty miles. Do not localize currency unless the job explicitly requires it.

### SEO
- Natural Japanese search language; no keyword stuffing. Spaces around keywords only if still natural.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "KO": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Korean). Additive — do not discard earlier rules.

### Register
- Friendly respectful honorifics; omit explicit “you” when honorific verbs already imply it.
- Sparing `!`; avoid `!!` / `!!!`, especially in notices/errors/support tone.

### SEO
- Natural Korean search language; no keyword stuffing; clear descriptive headings.

### Compliance reminders
- Neutral discount/deal wording; “best” only with measurable basis; no greenwashing absolutes.

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
    "SA": """
## SEO / Editorial LQA Addendum (Jan 2026)
Source: Style Guide Jan 2026 (Arabic). Additive — do not discard earlier rules.

### Register / punctuation
- MSA with polite consistent address; Arabic comma `،` and question mark `؟`.
- Do not mirror English exclamation-mark habits.

### Geo naming
- Prefer commonly used Arabic forms; if no established Arabic form, keep the source Latin spelling (searchability).
- Airport: transliterate proper nouns; translate descriptive words (e.g. International → دولي).

### SEO body readability (objective Style only)
- Split objectively overlong multi-idea SEO body paragraphs (**Style**); provide split text in `updated_sentence` / `corrected_content`.
- No synonym-only preference edits.
""",
}


def _ensure_trailing_newline(text: str) -> str:
    return text if text.endswith("\n") else text + "\n"


def append_addendum(code: str, addendum: str) -> str:
    path = PROMPTS / f"{code}_instructions.txt"
    if not path.exists():
        return f"SKIP {code}: missing base file"
    text = path.read_text(encoding="utf-8")
    if MARKER in text:
        return f"KEEP {code}: addendum already present"
    path.write_text(_ensure_trailing_newline(text) + "\n" + addendum.strip() + "\n", encoding="utf-8")
    return f"APPEND {code}"


def main() -> None:
    actions: list[str] = []

    for code, base in NEW_BASE.items():
        path = PROMPTS / f"{code}_instructions.txt"
        if path.exists():
            actions.append(f"KEEP-BASE {code}: already exists")
        else:
            path.write_text(_ensure_trailing_newline(base), encoding="utf-8")
            actions.append(f"CREATE {code}")

    for dest, src in ALIASES_FROM.items():
        dest_path = PROMPTS / f"{dest}_instructions.txt"
        src_path = PROMPTS / f"{src}_instructions.txt"
        if dest_path.exists():
            actions.append(f"KEEP-ALIAS {dest}: already exists")
            continue
        if not src_path.exists():
            actions.append(f"FAIL-ALIAS {dest}: source {src} missing")
            continue
        # Copy source base only (without duplicating if source already has addendum);
        # dest gets its own addendum below.
        src_text = src_path.read_text(encoding="utf-8")
        if MARKER in src_text:
            src_text = src_text.split(MARKER)[0].rstrip() + "\n"
        header = f"# Alias pack for filename code {dest} (locale shared with {src})\n"
        dest_path.write_text(header + _ensure_trailing_newline(src_text), encoding="utf-8")
        actions.append(f"CREATE-ALIAS {dest} from {src}")

    for code, addendum in ADDENDA.items():
        actions.append(append_addendum(code, addendum))

    print("\n".join(actions))


if __name__ == "__main__":
    main()
