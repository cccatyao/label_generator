# Label 4 Implementation Plan

## Goal
Implement `label4` generation in the same SVG->PDF pipeline as labels 2/19, with independent row handling and server-safe embedded fonts.

## Scope and Logic Alignment
- Reuse existing `cairosvg` conversion flow (render SVG, then convert to PDF bytes).
- Keep label2/label19 behavior unchanged.
- Add label4 as an additional output in the same Generate action.
- Handle label4 validation/skips independently from label2/19 row gating.

## Data Mapping
- `Column A` (`default_code`): output filename prefix (`{default_code}-label4.pdf`).
- `Column E` (`Firm`): company selector (`castlery` / `mopio`) for distributor block.
- `Column F` (`Origin`): origin country mapping for `MADE IN ...`.
- `Column G` (`Washing Material`): material composition source for label4 material lines.
- `Column H` (`Washing Guide`): washing instruction source, normalized and mapped to icon snippets.

## Template Updates (`template/label4.svg`)
- Keep fixed title text:
  - `Outer Cover`
  - `(Recouverture exterieure)`
- Replace hardcoded sample content with placeholders:
  - `{{material_lines}}`
  - `{{washing_text_lines}}`
  - `{{washing_icons}}`
  - `{{distributor_lines}}`
  - `{{origin_country}}`
- Add `{{embedded_font_faces}}` placeholder in `<style>` so fonts are embedded at generation time.

## Configuration (`term_config.py`)
- Add label4-specific material French mappings (dictionary + matcher).
- Add washing instruction normalization + icon key mapping scaffolding.
- Add distributor text blocks per firm (`castlery`, `mopio`) in config.
- Keep mapping extensible so new washing terms can be added without generator refactor.

## Generator (`generate_label4.py`)
- Implement `generate_label4_from_dataframe(template_content, df, generate_pdf=True)`.
- Parse/material formatting from column G:
  - split by comma/newline,
  - normalize percentage/material format,
  - map material terms to French,
  - no blank spacer lines.
- Parse washing instructions from column H:
  - normalize each line,
  - map to icon keys,
  - render icon snippets + optional display text lines.
- Embed fonts from local `font/` directory as `@font-face data:` blocks before conversion.
- Generate warnings and skip row (label4 only) when required fields/mappings are missing.

## App Integration (`app.py`)
- Load `template/label4.svg` alongside existing templates.
- Add session counters for label4 outputs.
- Keep existing label2/19 validation pass untouched.
- Run label4 generation on original dataframe (independent skip logic inside generator).
- Merge label4 PDFs/warnings into ZIP and warning panel.
- Update success and footer text to include label4.

## Validation/Failure Rules (Label4 only)
- Skip row with warning when:
  - column G empty,
  - column H empty,
  - firm unknown (no distributor config),
  - no mappable washing instructions.
- Do not fail full run due to label4 row issues.

## Verification
- Run fast syntax check for touched python files.
- Smoke test with `data/import_label_data.xlsx`:
  - ensure rows with valid G/H produce `-label4.pdf` files,
  - ensure missing/invalid rows emit warnings only,
  - ensure output ZIP includes labels 2/19/4 as available.
