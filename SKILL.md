---
name: pptx-to-pypptx
description: "Use when the user wants to migrate, edit, inspect, round-trip, or upgrade a PowerPoint deck through python-pptx code. Triggers on 'migrate pptx to python', 'python-pptx project', 'pptx to python-pptx', 'turn deck into code', 'sync pptx changes to code', 'inspect slide', 'what is on this slide', 'upgrade the pptx-to-pypptx skill', or any request to recreate, inspect, or maintain a .pptx using python-pptx."
---

# PPTX ⇄ python-pptx Round-Trip Skill

Turn an existing `.pptx` into a maintainable `python-pptx` project, edit the code, and regenerate the deck. The generated code recreates all common PowerPoint features, so the workflow is a single round-trip loop.

Replace `<pptx-to-pypptx-dir>` below with the directory that contains this `SKILL.md`.

## Rules

- **Target & iteration.** The **initial target** is the `.pptx` the user gave you; it is **read-only** — never overwrite, modify, or move it. After the first build the working target becomes **`out/<filename>.pptx`** (unless the user names a different file). Every successful build is archived to `backup/`, so you can edit the `out/` file freely and roll back.
- **Run the scripts; don't hand-craft XML.** The scripts handle extraction, codegen, and rebuilding. Your job is to pick the right one and act on its reported status. Run them as documented — don't pre-emptively edit, move, rename, or delete files.
- **Execute deterministic steps decisively — don't overthink them.** These scripts are deterministic and print an explicit status. Run the right one and act on that line; do **not** re-run "to be sure", re-verify its work, or inspect slides unless the status or the user asks. Trust these success signals and move on:
  - `autosync.py` → `OK` (nothing to do — proceed) · `SYNCED` (code updated — proceed) · `SKIPPED` (do what the message says).
  - `build_deck.py` → `Wrote …` followed by `Validator passed …` and `Recorded round-trip sync state …`.
  - `generate_slides.py` → one `Generated …` line per requested slide.
  - `sync_slide_numbers.py` → the printed rename/delete/reserve plan (add `--apply` to act).
  - `detect_project.py` → the JSON `projects` array (exit 1 only means "none found").
- **Slide-number sync.** Always use `sync_slide_numbers.py` to add, remove, or reorder slides. Never manually rename or delete `slides/s*.py`.
- **Skill core is off-limits.** Don't edit anything under `<pptx-to-pypptx-dir>/` (templates, scripts, helpers) — including `lib/shapes.py` — unless the user explicitly asks you to fix or upgrade the skill itself. For deck work, only edit generated project files (`slides/*.py`, `lib/design.py`, etc.).
- **Preserve human files.** `scaffold.py` only overwrites its own generated files (`build_deck.py`, `slides/*.py`, `lib/*.py`, `lib/base.pptx`, `assets/`, `backup/`). Never delete a `.pptx` a human placed in the project root or `out/` unless the user asks.
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
├── comments/                  # preserved PowerPoint comments (only if the deck has any)
├── lib/
│   ├── design.py              # colors, fonts, layout constants (edit to match deck)
│   ├── shapes.py              # add_box, add_shape, add_image, add_connector, add_chart, ...
│   ├── comments.py            # re-attaches preserved comments after each build
│   ├── roundtrip_state.py     # shared sync-state helper (auto-sync + build_deck)
│   └── base.pptx              # template shell (masters/layouts/theme, no slides); build input
├── slides/
│   ├── s01_title.py           # one file per slide; deck number = filename order
│   └── s02_outline.py
├── out/
│   └── my-deck.pptx
└── .roundtrip_state.json      # marks the deck version the code is in sync with (auto-sync)
```

- **One file per slide**, imported by `build_deck.py` in sorted filename order (that order assigns deck numbers).
- **Chrome is normal shapes** — title bars, footers, separators, and slide numbers are drawn inside each slide file via `shapes.py`; there is no separate chrome module.

## Commands

`scripts/*.py` are the tools you run (`uv run python <pptx-to-pypptx-dir>/scripts/<name> ...`). `scripts/helpers/` holds their shared internals — never run or edit those.

| Command | What it does |
|---|---|
| `scaffold.py` | Create the project structure, copy assets, capture the base deck (`lib/base.pptx`: masters/layouts/theme, no slides), preserve any PowerPoint comments into `comments/`, and auto-detect the footer into `lib/design.py` (`FOOTER_TEXT`). The built deck is named after `<output-dir>`. Does **not** generate slide code or a `pyproject.toml`. |
| `autosync.py` | **Run this first** on any deck task to fold in PowerPoint edits (see **Auto-sync** below). Detects whether `out/<name>.pptx` changed since the last build/sync and, if a human edited it, regenerates the affected `slides/*.py`. Deck→code only (never rebuilds), mechanical (no TODO review), auto-detects changed slides, and never errors out — a cheap no-op when nothing changed. |
| `generate_slides.py` | Fully overwrite selected `slides/sNN_*.py` from the target. `--slides` is required (`4` \| `2-5` \| `3,7,9`); there is no `all`. |
| `sync_slide_numbers.py` | Reserve slots (`--add`) or close gaps (`--delete`) by renaming `slides/s*.py`. Run **before** `generate_slides.py`; only renames/deletes files. Add `--apply` to act (default is a dry run). |
| `extract_slide.py` | Dump a slide's shapes — position, size, text, fill, font, z-order, `[HIDDEN]`. `--verbose` for detail, `--screenshot` for a PNG, `--json` for machine output. Accepts `all`. |
| `extract_notes.py` | Export speaker notes from `slides/*.py` to a Markdown file. |
| `add_comment.py` | Leave a **Claude-authored** comment on a slide (`--project-dir`, `--slide`, `--text`). Writes into the project's `comments/` store so it re-attaches on the next `build_deck.py`. Use it when your edit is substantial, fixes a perceived error, or addresses an existing comment — see **Annotating your own changes** below. |
| `list_layouts.py` | List layout indices in a deck (for a slide's `LAYOUT` constant). |
| `detect_project.py` | List existing projects (current dir or one level down) with each one's slides, backups, and output path. Run before a partial update. Returns a `projects` array (`count` 0 → exit 1). |
| `build_deck.py` | *(inside the project)* Build `slides/` into `out/<name>.pptx`, archiving the prior build to `backup/`. Self-contained and takes no arguments — it uses the bundled `lib/base.pptx` for masters/layouts/theme. On success it stamps `.roundtrip_state.json` so auto-sync never mistakes this build for a human edit. |
| `recapture_base.py` | Refresh a project's `lib/base.pptx` from an edited deck (default source: its `out/<name>.pptx`) after you change masters/layouts/theme in PowerPoint. Non-destructive — only rewrites `lib/base.pptx`, never touches `slides/`. |

Canonical invocations:

```bash
# Scaffold a new project from the target deck (the built deck is named after <output-dir>)
uv run python <pptx-to-pypptx-dir>/scripts/scaffold.py \
  --target "<target.pptx>" --output-dir <output-dir>

# Generate slide code (accepts 14 | 8-12 | 4,5,9 — no "all")
uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
  --target "<target.pptx>" --project-dir <output-dir> --slides 1-5

# Build (self-contained; uses lib/base.pptx). Run with --directory so
# build_deck.py's relative imports resolve; uv resolves the env from the root.
uv run --directory <output-dir> python build_deck.py

# Auto-sync code from a PowerPoint edit — run first on any deck task (cheap
# no-op when nothing changed). Auto-detects and regenerates only changed slides.
uv run --directory <output-dir> python <pptx-to-pypptx-dir>/scripts/autosync.py \
  --project-dir <output-dir>

# Recapture lib/base.pptx after editing masters/layouts/theme in PowerPoint
uv run python <pptx-to-pypptx-dir>/scripts/recapture_base.py --project-dir <output-dir>

# Reserve slots 3 and 6 (or --delete 2,5 to close gaps); default is a dry run
uv run python <pptx-to-pypptx-dir>/scripts/sync_slide_numbers.py \
  --project-dir <output-dir> --add 3,6 --apply

# Inspect a slide without a screenshot
uv run python <pptx-to-pypptx-dir>/scripts/extract_slide.py "<target.pptx>" 7 --verbose

# Leave a concise Claude-authored comment on slide 71 (attaches on next build)
uv run python <pptx-to-pypptx-dir>/scripts/add_comment.py \
  --project-dir <output-dir> --slide 71 \
  --text "Corrected data-size figure per reviewer note: ~20-33 GB, not 40 TB."
```

## Workflow

### New project (initial migration)

1. **Inspect** the current directory with `detect_project.py`. If a project already exists, ask the user whether to use it or scaffold a new folder — multiple projects can coexist, so only scaffold when the target location is confirmed.
2. **Scaffold** from the target (`scaffold.py`).
3. **Generate** code for **all** slides (`generate_slides.py --slides 1-N`).
4. **Build once** (`build_deck.py`). Don't build again until there are edits or a regeneration. This first build stamps `.roundtrip_state.json`, the baseline for auto-sync.
5. **Check for TODOs** in the generated `slides/s*.py`. No `# TODO` → the deck is done. TODOs → implement them, improve the skill, or flag them to the user.
6. **Tell the user** they can now edit `out/<name>.pptx` in PowerPoint; next time they ask you to work on the deck you'll run `autosync.py` and pull those edits back into code automatically (they don't have to say which slides changed).

### Auto-sync a PowerPoint edit (run this first)

The user may edit `out/<name>.pptx` in PowerPoint between turns. **At the start of any deck task — before you read or regenerate slide code — run `autosync.py`** so the code reflects those edits. It is a cheap no-op when nothing changed, so run it every time:

```bash
uv run --directory <output-dir> python <pptx-to-pypptx-dir>/scripts/autosync.py \
  --project-dir <output-dir>
```

It compares the deck against `.roundtrip_state.json` and, if the user edited it, regenerates the affected `slides/*.py` (or all slides on an add/delete/reorder) and re-stamps the marker. It **auto-detects which slides changed**, so the user never has to say. It always prints exactly one status line — act on it and continue, don't re-check:

- `autosync: <name>: OK — …` → nothing changed; the code already matches the deck. **Proceed.**
- `autosync: <name>: SYNCED — N slide(s) [..]; code now matches the deck.` → code updated. **Proceed** (no need to inspect the synced slides).
- `autosync: <name>: SKIPPED — …` → couldn't sync; do what the message says (usually re-scaffold, or the deck is broken).

A status line may end with `(comments mirrored from deck)`: auto-sync also mirrors the deck's **modern comments** back into the `comments/` store (deck → store), so a human reply or deletion made in PowerPoint sticks and the next build no longer resurrects a removed comment. This runs even when no slide changed (a human may edit only comments).

What auto-sync deliberately does **not** do — handle these yourself:

- It never rebuilds the deck, never reviews `# TODO`s, and never runs `recapture_base.py`. After a sync, check the regenerated files for new `# TODO`s if the user cares.
- If the user changed **masters/layouts/theme** in PowerPoint, auto-sync only refreshes slide code — run `recapture_base.py` to bake in the new layouts.
- Pure inspection (`extract_slide.py`) reads the live deck directly, so you don't need auto-sync just to answer "what's on slide N?" — only before editing or building **code**.

> Auto-sync is the fast path for "the human edited the deck." The **Partial update** steps below are the manual equivalent for finer control (choosing slides, TODO review, structural changes).

### Partial update (human edited a `.pptx`)

By default, assume the human edited the working deck at `out/<filename>.pptx`. Treat the original target as the source only if the user explicitly says they edited it.

1. **Detect** the project with `detect_project.py` to confirm its structure (`build_deck.py`, `slides/`, `lib/`, `assets/`) and locate `slides/`.
2. **If slides were added or deleted**, run `sync_slide_numbers.py` (`--add` / `--delete`, see Commands) to reserve slots or close gaps first.
3. **Generate** only the affected slides (`generate_slides.py --slides 3,7,12`). Do **not** run `scaffold.py` again.
4. **Build** with `build_deck.py`. (If the human changed masters/layouts/theme in PowerPoint, first run `recapture_base.py --project-dir <output-dir>` to bake the new layouts into `lib/base.pptx`.)
5. **Check for TODOs** in the regenerated files. (Optionally spot-check the output.)

> **Agent rule:** No need to visually inspect slides unless the user asks. Use `extract_slide.py` to see what is on a slide without screenshots. Success = the requested slides regenerate with no unexpected `# TODO` comments.

### Upgrade the skill

Triggered by "upgrade the pptx-to-pypptx skill." Pull, then migrate projects only if code changed.

1. **Pull and diff:**
   ```bash
   OLD=$(git -C <pptx-to-pypptx-dir> rev-parse HEAD)
   git -C <pptx-to-pypptx-dir> pull
   git -C <pptx-to-pypptx-dir> diff --name-only "$OLD" HEAD
   ```
   "Already up to date." → stop.
2. **Re-read `SKILL.md` if it changed** — your loaded copy is now stale.
3. **Decide.** Migrate only if the diff touched **`scripts/`** or **`template/`**. Docs-only (`*.md`, `LICENSE`, `references/`) → report "no migration needed" and stop.
4. **Re-baseline each project** (`detect_project.py`) from its own `out/<name>.pptx`:
   1. **Build** (`build_deck.py`), then **commit** the project — re-scaffolding resets `lib/design.py`, re-syncs `assets/`, and empties `slides/`.
   2. **Re-scaffold in place** (`scaffold.py --target <output-dir>/out/<name>.pptx --output-dir <output-dir>`).
   3. **Regenerate all slides** (`generate_slides.py --slides 1-N`).
   4. **Build** (`build_deck.py`) and check for new `# TODO`s.

### Direct code edits / free inspection

You can edit the generated Python directly without touching the target. See [`references/SLIDE_FORMAT.md`](./references/SLIDE_FORMAT.md) for the required file layout and helper conventions.

Canonical round-trip:

1. **Edit** `slides/s*.py` (or `lib/shapes.py`).
2. **If you add or remove slide files**, run `sync_slide_numbers.py` to keep indices contiguous.
3. **Build** with `build_deck.py`.
4. **Canonicalize** each edited slide by regenerating it from the freshly built output, so the code matches what `python-pptx` actually serialized:
   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/generate_slides.py \
     --target "<output-dir>/out/my-deck.pptx" --project-dir <output-dir> --slides 7
   ```
5. **Build again** to verify.

Use `extract_slide.py` (text, `--screenshot`, or `--json`) whenever you need the facts about a slide — positions, sizes, text, fills. Work in small batches: **3–5 slides at a time**, or **one at a time** for complex diagram slides.

### Annotating your own changes (leave a comment)

When you change a slide, leave a concise comment on it — authored as **Claude** — so the human reviewers can see what you did and why. Do this **only** when the change is worth a reviewer's attention:

- **Substantial** — you rewrote/restructured content or notes, not just reformatted or fixed a typo.
- **Corrected a perceived error** — say what was wrong and what you changed it to (e.g. a wrong figure, a contradiction with the on-slide text, a broken example). This is your judgment, so flag it for a human to confirm.
- **Addressed an existing comment** — reference the reviewer's point and describe how the edit resolves it.

Do **not** comment on routine passes (spelling out contractions, teleprompter reformatting, adding an end marker, etc.) — that would just add noise.

Keep it to one or two sentences, factual, no preamble. Run it after editing and before the build (the comment attaches on the next `build_deck.py`):

```bash
uv run python <pptx-to-pypptx-dir>/scripts/add_comment.py \
  --project-dir <output-dir> --slide 71 \
  --text "Corrected data-size figure: BooksCorpus + Wikipedia is ~20-33 GB, not 40 TB (addresses reviewer note)."
```

The comment is pinned at slide level and rides the round-trip like any preserved comment. Notes:

- It appears in `out/<name>.pptx` only after the next build.
- It behaves like a **normal** PowerPoint comment: a human reviewer can reply to it or delete it in PowerPoint, and `autosync.py` mirrors that change back into the store on the next deck task (so a deleted comment is not resurrected, and replies are kept).

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
- Modern threaded comments — captured verbatim at scaffold time and re-attached to the rebuilt slides on every build (author, timestamp, text, and thread replies survive). Comment edits made later in PowerPoint (replies, deletions) are mirrored back into the store by `autosync.py` at the start of a deck task, so they round-trip like slide edits; shape-level comment anchoring may relax to a slide-level pin since rebuilt shape ids differ. Claude can also **add its own** comments via `add_comment.py` (author "Claude") to annotate the changes it makes — see **Annotating your own changes**.

Anything truly exotic (custom geometry that can't be vectorized, SmartArt, animations, transitions) becomes a `# TODO` comment in the generated slide file.

## Tips

- Always use the actual deck slide number (`slideN.xml`), not the descriptive function names inside `slides/`.
- Implement the hardest slides first (usually diagram slides); title, outline, and divider slides are quick.
- Lean on `extract_slide.py` for diagram slides where exact positions matter.
- Keep shared drawing code in `lib/shapes.py` so slide files stay focused on layout.
