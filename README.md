# pptx-to-pypptx

Turn an existing PowerPoint deck (`.pptx`) into a maintainable
[`python-pptx`](https://python-pptx.readthedocs.io/) project, edit the generated
code, and regenerate the deck — a clean round-trip loop between binary slides and
readable Python.

This repository is packaged as a [Claude Code](https://claude.com/claude-code)
**skill** ([`SKILL.md`](./SKILL.md)), but the scripts under [`scripts/`](./scripts)
are plain command-line tools and can be run on their own.

## Why

`.pptx` files are opaque zipped XML. Once a deck lives as one small Python file
per slide, you can diff it, review it, script bulk edits, and rebuild the deck
deterministically — while keeping the original untouched as a reference.

## What it does

- **Scaffold** a project skeleton (`build_deck.py`, `lib/`, `slides/`, `assets/`,
  `backup/`) from a target deck.
- **Generate** one `slides/sNN_*.py` file per slide, recreating shapes, text,
  images, tables, charts, groups, connectors, backgrounds, and speaker notes.
- **Build** the deck back into a `.pptx`, archiving the previous output to
  `backup/` on every run.
- **Sync** code from a manually edited deck by regenerating only the affected
  slides.
- **Inspect** any slide's shapes/positions/text (and optionally render a PNG)
  without opening PowerPoint.

## Repository layout

```
pptx-to-pypptx/
├── SKILL.md                 # skill entry point / full workflow docs
├── references/
│   └── SLIDE_FORMAT.md      # conventions for generated slide files
├── scripts/                 # command-line tools (run with uv)
│   ├── scaffold.py          # create a new project from a target deck
│   ├── generate_slides.py   # generate/refresh slides/sNN_*.py from a deck
│   ├── detect_project.py    # locate an existing project
│   ├── extract_slide.py     # dump/screenshot a single slide
│   ├── extract_notes.py     # export speaker notes to Markdown
│   ├── sync_slide_numbers.py# reserve/close slide-number slots
│   ├── list_layouts.py      # list slide layouts in a deck
│   └── helpers/             # internal modules (imported by the scripts, not run directly)
└── template/                # files copied into each scaffolded project
    ├── build_deck.py
    └── lib/                 # design.py + shapes.py helper library
```

## Requirements

- Python 3.10+
- [`uv`](https://docs.astral.sh/uv/) for environment and command running
- Python packages (installed in the **project** you scaffold, not this repo):
  - `python-pptx>=1.0.0`
  - `cairosvg>=2.0`
  - `pillow-heif>=1.0`
  - `defusedxml` (optional; used by some helper scripts)
- Optional: [LibreOffice](https://www.libreoffice.org/) (`soffice`) for
  `extract_slide.py --screenshot` rendering.

## Quick start

Replace `<repo>` with the path to this checkout.

```bash
# 1. Scaffold a project from your deck
uv run python <repo>/scripts/scaffold.py \
  --target "deck.pptx" \
  --output-dir my-deck \
  --project-name my-deck \
  --output-filename my-deck

# 2. Generate slide code for every slide (accepts 14 | 8-12 | 4,5,9)
uv run python <repo>/scripts/generate_slides.py \
  --target "deck.pptx" \
  --project-dir my-deck \
  --slides 1-20

# 3. Build the deck (run from the project root)
uv run --directory my-deck python build_deck.py --target "deck.pptx"

# 4. Inspect any slide without a screenshot
uv run python <repo>/scripts/extract_slide.py "deck.pptx" 10 --verbose
```

After the first build, the working target becomes `my-deck/out/my-deck.pptx`.
Edit `slides/*.py`, rebuild, and regenerate a slide from the fresh output to keep
the code canonical. See [`SKILL.md`](./SKILL.md) for the full workflow, including
partial updates and slide-number syncing.

## Supported features

Rectangles and preset shapes, lines/arrows/connectors, solid/gradient/pattern/
theme fills, rich text (runs, fonts, colors, bullets, alignment, hyperlinks,
Office Math), shape effects (shadow/glow/reflection/soft-edge), images (with
cropping; SVG/WebP/HEIC conversion) and freeform vector shapes, tables with
merged cells, charts, embedded video, true group shapes with group-relative
child positioning, slide backgrounds, speaker notes, and hidden slides.

Anything truly exotic (SmartArt, animations, transitions, custom geometry that
can't be vectorized) is emitted as a `# TODO` comment in the generated slide.

## Using it as a Claude Code skill

Point Claude Code at this directory as a skill. Claude reads
[`SKILL.md`](./SKILL.md) and drives the scripts through the round-trip workflow
automatically. The skill triggers on requests like *"migrate this pptx to
python-pptx"*, *"sync my deck changes to code"*, or *"what's on slide 10?"*.

## License

Copyright 2026 Yuan Tang.

Licensed under the [Apache License, Version 2.0](./LICENSE). You may not use
this project except in compliance with the License.
