---
name: pptx-to-pypptx
description: "Use when the user wants to migrate, edit, inspect, or round-trip a PowerPoint deck through python-pptx code. Triggers on 'migrate pptx to python', 'python-pptx project', 'pptx to python-pptx', 'turn deck into code', 'sync pptx changes to code', 'inspect slide', 'what is on this slide', or any request to recreate, inspect, or maintain a .pptx using python-pptx."
---

# PPTX ⇄ python-pptx Round-Trip Skill

Turn an existing `.pptx` into a maintainable `python-pptx` project, edit the code, and regenerate the deck. The generated code recreates all common PowerPoint features, so the workflow is a single round-trip loop.

Replace `<pptx-to-pypptx-dir>` below with the directory that contains this `SKILL.md`.

## Rules

- **Target & iteration.** The **initial target** is the `.pptx` the user gave you; it is **read-only** — never overwrite, modify, or move it. After the first build the working target becomes **`out/<filename>.pptx`** (unless the user names a different file). Every successful build is archived to `backup/`, so you can edit the `out/` file freely and roll back.
- **Run the scripts; don't hand-craft XML.** The scripts handle extraction, codegen, and rebuilding. Your job is to pick the right one and verify its output. Run them as documented — don't pre-emptively edit, move, rename, or delete files.
- **Slide-number sync.** Always use `sync_slide_numbers.py` to add, remove, or reorder slides. Never manually rename or delete `slides/s*.py`.
- **Skill core is off-limits.** Don't edit anything under `<pptx-to-pypptx-dir>/` (templates, scripts, helpers) — including `lib/shapes.py` — unless the user explicitly asks you to fix or upgrade the skill itself. For deck work, only edit generated project files (`slides/*.py`, `lib/design.py`, etc.).
- **Preserve human files.** `scaffold.py` only overwrites its own generated files (`build_deck.py`, `slides/*.py`, `lib/*.py`, `assets/`, `backup/`). Never delete a `.pptx` a human placed in the project root or `out/` unless the user asks.
- **Backups.** `backup/` keeps the last 10 successful builds (`backup_YYYYMMDD_HHMMSS.pptx`). You may copy one back into `out/` or the project root if the user wants to restore a version.

## What you need

- The **target** `.pptx` to recreate or keep in sync.
- Python with `uv`.
- These packages in the **project-root** environment (add to `pyproject.toml`/`requirements.txt`, then `uv sync` from the project root):
  - `python-pptx>=1.0.0`, `cairosvg>=2.0`, `pillow-heif>=1.0`

## Project layout this skill produces

```
my-deck/
├── build_deck.py              # orchestrator: imports slides/ in filename order
├── backup/                    # last 10 successful builds for rollback
├── assets/                    # media + freeform SVGs extracted from the target
├── lib/
│   ├── design.py              # colors, fonts, layout constants (edit to match deck)
│   └── shapes.py              # add_box, add_shape, add_image, add_connector, add_chart, ...
├── slides/
│   ├── s01_title.py           # one file per slide; deck number = filename order
│   └── s02_outline.py
└── out/
    └── my-deck.pptx
```

- **One file per slide**, imported by `build_deck.py` in sorted filename order (that order assigns deck numbers).
- **Chrome is normal shapes** — title bars, footers, separators, and slide numbers are drawn inside each slide file via `shapes.py`; there is no separate chrome module.

## Commands

`scripts/*.py` are the tools you run (`uv run python <pptx-to-pypptx-dir>/scripts/<name> ...`). `scripts/helpers/` holds their shared internals — never run or edit those.

| Command | What it does |
|---|---|
| `scaffold.py` | Create the project structure, copy assets, and auto-detect the footer into `lib/design.py` (`FOOTER_TEXT`). The built deck is named after `<output-dir>`. Does **not** generate slide code or a `pyproject.toml`. |
| `generate_slides.py` | Fully overwrite selected `slides/sNN_*.py` from the target. `--slides` is required (`4` \| `2-5` \| `3,7,9`); there is no `all`. |
| `sync_slide_numbers.py` | Reserve slots (`--add`) or close gaps (`--delete`) by renaming `slides/s*.py`. Run **before** `generate_slides.py`; only renames/deletes files. Add `--apply` to act (default is a dry run). |
| `extract_slide.py` | Dump a slide's shapes — position, size, text, fill, font, z-order, `[HIDDEN]`. `--verbose` for detail, `--screenshot` for a PNG, `--json` for machine output. Accepts `all`. |
| `extract_notes.py` | Export speaker notes from `slides/*.py` to a Markdown file. |
| `list_layouts.py` | List layout indices in a deck (for a slide's `LAYOUT` constant). |
| `detect_project.py` | List existing projects (current dir or one level down) with each one's slides, backups, and output path. Run before a partial update. Returns a `projects` array (`count` 0 → exit 1). |
| `build_deck.py` | *(inside the project)* Build `slides/` into `out/<name>.pptx`, archiving the prior build to `backup/`. |

Canonical invocations:

```bash
# Scaffold a new project from the target deck (the built deck is named after <output-dir>)
uv run python <pptx-to-pypptx-dir>/scripts/scaffold.py \
  --target "<target.pptx>" --output-dir <output-dir>

# Generate slide code (accepts 14 | 8-12 | 4,5,9 — no "all")
uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
  --target "<target.pptx>" --project-dir <output-dir> --slides 1-5

# Build. Run with --directory so build_deck.py's relative imports resolve;
# uv resolves the environment from the project root.
uv run --directory <output-dir> python build_deck.py --target "<target.pptx>"

# Reserve slots 3 and 6 (or --delete 2,5 to close gaps); default is a dry run
uv run python <pptx-to-pypptx-dir>/scripts/sync_slide_numbers.py \
  --project-dir <output-dir> --add 3,6 --apply

# Inspect a slide without a screenshot
uv run python <pptx-to-pypptx-dir>/scripts/extract_slide.py "<target.pptx>" 7 --verbose
```

## Workflow

### New project (initial migration)

1. **Inspect** the current directory with `detect_project.py`. If a project already exists, ask the user whether to use it or scaffold a new folder — multiple projects can coexist, so only scaffold when the target location is confirmed.
2. **Scaffold** from the target (`scaffold.py`).
3. **Generate** code for **all** slides (`generate_slides.py --slides 1-N`).
4. **Build once** (`build_deck.py --target "<target.pptx>"`). Don't build again until there are edits or a regeneration.
5. **Check for TODOs** in the generated `slides/s*.py`. No `# TODO` → the deck is done. TODOs → implement them, improve the skill, or flag them to the user.

### Partial update (human edited a `.pptx`)

By default, assume the human edited the working deck at `out/<filename>.pptx`. Treat the original target as the source only if the user explicitly says they edited it.

1. **Detect** the project with `detect_project.py` to confirm its structure (`build_deck.py`, `slides/`, `lib/`, `assets/`) and locate `slides/`.
2. **If slides were added or deleted**, run `sync_slide_numbers.py` (`--add` / `--delete`, see Commands) to reserve slots or close gaps first.
3. **Generate** only the affected slides (`generate_slides.py --slides 3,7,12`). Do **not** run `scaffold.py` again.
4. **Build** with `build_deck.py --target "<output-dir>/out/<filename>.pptx"` (or the original target if the user said they edited it).
5. **Check for TODOs** in the regenerated files. (Optionally spot-check the output.)

> **Agent rule:** No need to visually inspect slides unless the user asks. Use `extract_slide.py` to see what is on a slide without screenshots. Success = the requested slides regenerate with no unexpected `# TODO` comments.

### Direct code edits / free inspection

You can edit the generated Python directly without touching the target. See [`references/SLIDE_FORMAT.md`](./references/SLIDE_FORMAT.md) for the required file layout and helper conventions.

Canonical round-trip:

1. **Edit** `slides/s*.py` (or `lib/shapes.py`).
2. **If you add or remove slide files**, run `sync_slide_numbers.py` to keep indices contiguous.
3. **Build** with `build_deck.py --target "<output-dir>/out/my-deck.pptx"`.
4. **Canonicalize** each edited slide by regenerating it from the freshly built output, so the code matches what `python-pptx` actually serialized:
   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
     --target "<output-dir>/out/my-deck.pptx" --project-dir <output-dir> --slides 7
   ```
5. **Build again** to verify.

Use `extract_slide.py` (text, `--screenshot`, or `--json`) whenever you need the facts about a slide — positions, sizes, text, fills. Work in small batches: **3–5 slides at a time**, or **one at a time** for complex diagram slides.

## Supported features

The generator and `lib/shapes.py` round-trip all commonly used PowerPoint constructs:

- Preset shapes (rectangles, rounded rectangles, ovals, triangles, arrows, chevrons, parallelograms, stars, …); non-listed presets are preserved rather than downgraded to rectangles.
- Lines and straight/elbow/curved connectors with color, width, dash, heads, cap, compound style, rotation, and shape connections.
- Solid, theme, gradient, and pattern fills; no-fill; and style references (`lnRef`/`fillRef`/`effectRef`).
- Rich text: multiple paragraphs and runs; font, size, color, bold/italic/underline/strike, super/subscript, highlight, hyperlinks; paragraph alignment, indentation, bullets (char and auto-numbered), and spacing.
- Shape effects — shadow, glow, reflection, soft edge.
- Images with cropping and luminance adjustments. Native: PNG/JPEG/GIF/BMP/TIFF/WMF/EMF. SVG is rasterized; WebP/HEIC/HEIF convert via Pillow (`pillow-heif` for HEIC); WDP/JPEG-XR needs ImageMagick. Freeform/custom shapes are exported as SVG assets and placed via `add_image`, so arbitrary vector artwork survives.
- Tables with per-cell fills, borders, alignment, margins, row/column sizes, and merged cells.
- True group shapes with group-relative child positioning.
- Charts (column, bar, line, pie, area, …).
- Embedded video/movie shapes with poster frames.
- Slide backgrounds (solid and gradient) and speaker notes.
- Hidden slides — re-emitted as `shapes.set_slide_hidden(slide)`; `extract_slide.py` reports them.
- Native Office Math (`m:oMath` / `a14:m`) in text runs, preserved as editable equations.

Anything truly exotic (custom geometry that can't be vectorized, SmartArt, animations, transitions) becomes a `# TODO` comment in the generated slide file.

## Tips

- Always use the actual deck slide number (`slideN.xml`), not the descriptive function names inside `slides/`.
- Implement the hardest slides first (usually diagram slides); title, outline, and divider slides are quick.
- Lean on `extract_slide.py` for diagram slides where exact positions matter.
- Keep shared drawing code in `lib/shapes.py` so slide files stay focused on layout.
