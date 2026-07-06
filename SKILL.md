---
name: pptx-to-pypptx
description: "Use when the user wants to migrate, edit, inspect, or round-trip a PowerPoint deck through python-pptx code. Triggers on 'migrate pptx to python', 'python-pptx project', 'pptx to python-pptx', 'turn deck into code', 'sync pptx changes to code', 'inspect slide', 'what is on this slide', or any request to recreate, inspect, or maintain a .pptx using python-pptx."
---

# PPTX ⇄ python-pptx Round-Trip Skill

Turn an existing `.pptx` into a maintainable `python-pptx` project, edit the code, and regenerate the deck. The generated code is now comprehensive enough to recreate all common PowerPoint features, so the workflow is a single round-trip loop instead of a separate diff script.

> **Target-file rule:** The **initial target `.pptx`** is the file the user provided and asked you to turn into code (or sync from). It is read-only. Never overwrite, modify, or move it. After the first build, the working target becomes `out/<filename>.pptx` unless the user explicitly names a different file. 
>
> **Iteration convention:** The initial migration is a single build. After that, **the working target becomes `out/<filename>.pptx`** — unless the user explicitly says they edited a different `.pptx`. Every build archives the previous output to `backup/`, so you can safely edit the `out/` file and roll back if needed. The original target file stays untouched as the starting reference.
>
> **Agent execution rule:** Do not overthink or hand-craft XML. Run the provided scripts. The scripts handle extraction, codegen, and rebuilding. Your job is to pick the right script and verify the output.
>
> **Follow instructions directly:** Run the scripts as documented. Do not pre-emptively edit, move, rename, or delete files.
>
> **Slide-number sync rule:** Always use `sync_slide_numbers.py` for adding, removing, or reordering slides. Never manually rename or delete `slides/s*.py`.
>
> **Skill-core rule:** Do not edit `lib/shapes.py` or other files under `<pptx-to-pypptx-dir>/` (templates, scripts, helpers) unless the user explicitly asks you to upgrade or fix the skill itself. For normal deck work, only edit generated project files (`slides/*.py`, `lib/design.py`, etc.).
>
> **Preservation rule:** Never delete a `.pptx` file that a human placed in the output directory. `scaffold.py` only overwrites its own generated files (`build_deck.py`, `slides/*.py`, `lib/*.py`, the `assets/` directory, and the `backup/` directory). Any existing `.pptx` files in the project root or `out/` must be left alone unless the user explicitly asks you to remove them.
>
> **Backup rule:** The `backup/` directory stores the last 10 successful builds (`backup_YYYYMMDD_HHMMSS.pptx`) so the user can roll back. The agent may copy one of these backups into `out/` or back to the project root if the user asks to restore a previous version.

## What you need

- The **target** `.pptx` you want to recreate or keep in sync.
- Python with `uv` available.
- These packages in the **project root** environment:
  - `python-pptx>=1.0.0`
  - `cairosvg>=2.0`
  - `pillow-heif>=1.0`
  - `defusedxml` (dev/optional, used by some helper scripts)

Add them to the root `pyproject.toml` or `requirements.txt`, run `uv sync` from the project root, and let `uv` resolve the environment from there.

## Project layout this skill produces

```
my-deck/
├── build_deck.py              # orchestrator: loads slides/ in filename order
├── backup/                    # last 10 successful builds for rollback
│   ├── backup_20260620_070530.pptx
│   ├── backup_20260620_071245.pptx
│   └── ...
├── assets/                    # images/videos + freeform SVGs extracted from the target .pptx (PNG/JPEG/GIF/BMP/TIFF/WMF/EMF/SVG/WebP/HEIC/WDP supported)
├── lib/
│   ├── __init__.py
│   ├── design.py              # colors, fonts, layout constants (edit to match deck)
│   └── shapes.py              # add_box, add_shape, add_image, add_arrow, add_line, add_connector, add_custom_table, add_group, add_chart, add_movie, ...
├── slides/
│   ├── s01_title.py
│   ├── s02_outline.py
│   └── ...
└── out/
    └── my-deck.pptx           
```

- **One file per slide.** Find and edit a slide quickly.
- **Slide numbers come from filename order.** `build_deck.py` imports `slides/s*.py` in sorted order and assigns deck numbers automatically.
- **Chrome is normal shapes.** Title bars, footers, separators, and slide numbers are generated inside each slide file using `shapes.py`; there is no separate chrome module.

## Quick reference

Replace `<pptx-to-pypptx-dir>` with the directory that contains this `SKILL.md` file.

```bash
# 1. Scaffold a new python-pptx project from the target deck
uv run python <pptx-to-pypptx-dir>/scripts/scaffold.py \
  --target "<target.pptx>" \
  --output-dir <output-dir> \
  --project-name <project-name> \
  --output-filename <output-filename>

# 2. Generate baseline slide code from the target deck
#    Accepts: 14 | 8-12 | 4,5,9  (no --all)
uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
  --target "<target.pptx>" \
  --project-dir <output-dir> \
  --slides 1-5

# 3. Build the generated deck
#    Run from the project root; uv will resolve the environment there.
#    --directory sets the cwd so build_deck.py's relative imports work.
#    Post-migration, use out/<filename>.pptx as the target unless the user
#    explicitly says they edited a different file.
uv run --directory <output-dir> python build_deck.py --target "<target.pptx>"

# 4. Check the generated slide files for TODO comments
#    If none, the deck is ready. If there are TODOs, implement or flag them.
```

### Keeping code in sync after manual PPTX edits

If the human edited `out/<filename>.pptx`, regenerate the affected slides from that output file to pull the changes back into code:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
  --target "<output-dir>/out/<filename>.pptx" \
  --project-dir <output-dir> \
  --slides 3,7,12
```

If the user explicitly says they edited the original target `.pptx` instead, use that file as the `--target`.

Then rebuild and check the regenerated slide files for TODO comments. This replaces the old `pptx-sync` compare-diff workflow.

## Workflow

### New project (initial migration)

1. **Inspect** the project root (or current directory) with `detect_project.py`. If an existing pptx-to-pypptx project is found, ask the user whether to use it or scaffold a new folder. Multiple projects can coexist, so only scaffold when the user confirms the target location.
2. **Scaffold** the project from the target `.pptx` (`scaffold.py`).
3. **Generate** baseline slide code for **all** slides (`generate_slides.py --slides 1-N`).
4. **Build once** the generated deck with `build_deck.py --target "<target.pptx>"`. Do not run a second build immediately; the next build only happens after edits or a slide regeneration.
5. **Check for TODOs**. Inspect the generated `slides/s*.py` files. If there are no `# TODO` comments, the deck is considered successfully generated. If TODOs exist, either implement them manually, improve the skill, or flag them to the user.

### Partial update (human edited a `.pptx`)

By default, assume the human edited the working deck in `out/<filename>.pptx`. Only treat the original target `.pptx` as the source if the user explicitly says they edited that file.

1. **Detect** the existing project with `detect_project.py` to confirm it has the structure `generate_slides.py` expects (`build_deck.py`, `slides/`, `lib/`, `assets/`) and to locate `slides/`.
2. **If slides were added or deleted**, run `sync_slide_numbers.py` to reserve slots or close gaps before generating:
   
   ```bash
   # reserve slots 3 and 6
   uv run python <pptx-to-pypptx-dir>/scripts/sync_slide_numbers.py \
     --project-dir <output-dir> --add 3,6 --apply
   
   # or close gaps after deleting slides 2 and 5
   uv run python <pptx-to-pypptx-dir>/scripts/sync_slide_numbers.py \
     --project-dir <output-dir> --delete 2,5 --apply
   ```
3. **Generate** only the affected slides (`generate_slides.py --slides 3,7,12`). Do **not** run `scaffold.py` again.
4. **Build** with `build_deck.py --target "<output-dir>/out/<filename>.pptx"` (or the original target `.pptx` if the user explicitly said they edited it).
5. **Check for TODOs** in the regenerated slide files.
6. (Optional) Spot-check the output `.pptx` if needed.

> **Agent rule:** No need to visually inspect slides unless the user explicitly asks. Use `extract-slide.py` to see what is on a slide without screenshots. The success criterion is that the requested slides are generated with no unexpected `# TODO` comments.

### Direct code edits / free inspection

You can also edit the generated Python code directly without touching the target `.pptx`. See [`references/SLIDE_FORMAT.md`](./references/SLIDE_FORMAT.md) for the required file layout and helper conventions so generated slide files stay consistent.

The recommended canonical round-trip is:

1. **Edit** `slides/s*.py` (or `lib/shapes.py`).
2. **If you add or remove slide files**, run `sync_slide_numbers.py` to keep filename indices contiguous and avoid ordering surprises.
3. **Build** with `build_deck.py --target "<output-dir>/out/my-deck.pptx"`.
4. **Canonicalize** the edited slide(s): if you edited a `slides/s*.py` file, regenerate just that slide from the freshly built output `.pptx` so the code matches what `python-pptx` actually serialized:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
  --target "<output-dir>/out/my-deck.pptx" \
  --project-dir <output-dir> \
  --slides 7
```

5. **Build** again with `build_deck.py --target "<output-dir>/out/my-deck.pptx"` to verify.

When you need to know what a slide originally contained — positions, sizes, text, fills, etc. — use `extract-slide.py` to get the facts:

```bash
# textual dump
uv run python <pptx-to-pypptx-dir>/scripts/extract-slide.py \
  "<target.pptx>" 7 --verbose

# render a PNG screenshot
uv run python <pptx-to-pypptx-dir>/scripts/extract-slide.py \
  "<target.pptx>" 7 --screenshot --screenshot-dir ./shots

# machine-readable JSON with screenshot path
uv run python <pptx-to-pypptx-dir>/scripts/extract-slide.py \
  "<target.pptx>" 7 --json --screenshot --screenshot-dir ./shots
```

Then edit `slides/s07_....py` (or `lib/shapes.py`) and follow the canonical round-trip above.

Work in small batches. A good sweet spot is **3–5 slides at a time**; for complex diagram slides, do **one at a time**.

## Supported features

The generator and `lib/shapes.py` can round-trip all commonly used PowerPoint constructs:

- Rectangles, rounded rectangles, ovals, triangles, arrows, chevrons, parallelograms, stars, and many other preset shapes.
- Lines, arrows, straight/elbow/curved connectors with color, width, dash, head/tail, cap, and compound styles.
- Solid, theme, gradient, and pattern fills; no-fill; and style references (`lnRef`/`fillRef`/`effectRef`).
- Text boxes and shape text with multiple paragraphs, runs, fonts, sizes, colors, bold/italic/underline/strike, baseline (super/sub), highlight, and hyperlinks.
- Paragraph alignment, indentation, bullets (char and auto-numbered), and spacing.
- Shadows, glow, reflection, soft edge, and other shape effects.
- Images with cropping and luminance adjustments. Native formats: PNG, JPEG, GIF, BMP, TIFF, WMF, EMF. SVG is rasterized automatically; WebP/HEIC/HEIF are converted via Pillow (install `pillow-heif` for HEIC); WDP/JPEG-XR requires ImageMagick (`convert`). Freeform/custom shapes are exported as SVG assets and rendered via `add_image` so arbitrary vector artwork (arrows, brush strokes, etc.) survives round-trips.
- Tables with per-cell fills, borders, alignment, margins, row/column sizes, and merged cells.
- True group shapes with group-relative child positioning.
- Charts (column, bar, line, pie, area, etc.).
- Embedded video/movie shapes with poster frames.
- Slide backgrounds (solid and gradient) and speaker notes.
- Hidden slides (`<p:sld show="0">`) — detected on generation and re-emitted as `shapes.set_slide_hidden(slide)`; `extract-slide.py` reports them.
- Native Office Math (`m:oMath` / `a14:m`) equations inside text runs; they are preserved as editable math in PowerPoint.

Anything truly exotic (custom geometry, SmartArt, animations, transitions) becomes a `TODO` comment in the generated slide file.

## Helper scripts

- `scaffold.py` — copies generic templates from `<pptx-to-pypptx-dir>/template/` and creates the project structure (`assets/`, `lib/`, `slides/`, `build_deck.py`, and an empty `backup/` directory). It does **not** generate slide code and does **not** create a `pyproject.toml`.
- `generate_slides.py` — fully overwrites selected `slides/sNN_*.py` files from the target `.pptx`. Requires explicit `--slides` (single, range, or comma list); no `--all` option.
- `detect_project.py` — locates an existing pptx-to-pypptx project (current dir or immediate subdir) and reports its path, slide files, and output filename. Use this before a partial update to confirm you should run `generate_slides.py`, not `scaffold.py`.
- `extract-slide.py` — prints a compact summary of every shape/image/table on a slide: position, size, text, fill, font, z-order, and whether the slide is hidden (`[HIDDEN]`, or `"hidden": true` under `--json`). Use `--verbose` for extra details, `--screenshot` to render the slide to a PNG, and `--json` for machine-readable output that includes the screenshot path.
- `sync_slide_numbers.py` — renames `slides/s*.py` files to reserve slots for added slides or close gaps after deleted slides. Run it **before** `generate_slides.py`; it only renames/deletes files and never modifies slide code.
- `extract_notes.py` — reads all `slides/s*.py` files and writes a Markdown document with slide numbers, titles, and speaker notes.

## Tips

- Always use the actual deck slide number (`slideN.xml`), not the descriptive function names inside `slides/`.
- Implement the hardest/most important slides first (often the diagram slides). Title, outline, and divider slides are usually quick.
- Use `extract-slide.py` heavily for diagram slides where exact positions matter, or whenever you need to know what is on a slide without taking a screenshot.
- Keep shared drawing code in `lib/shapes.py` so slide files stay focused on layout.
