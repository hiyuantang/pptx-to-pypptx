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

## Installation

This repo **is** the skill — [`SKILL.md`](./SKILL.md) sits at its root — so
installing it just means cloning the repo into a directory where Claude Code
discovers skills. There's no build step, and the skill has no dependencies of
its own (its Python packages are installed later, per deck project — see
[Requirements](#requirements)). The commands below are copy-paste friendly for a
human or an AI agent: hand Claude Code this repo's URL and ask it to install the
skill.

> **"Scope" means two different things here — don't conflate them:**
> - **Where the skill lives** (this section): a Claude Code *skills* directory —
>   **user** scope (all your projects) or **project** scope (one repo).
> - **Where the Python packages live** ([Requirements](#requirements)): inside
>   each **deck project you scaffold**, never in this repo.

**User scope — available in every project on your machine (recommended):**

```bash
git clone https://github.com/hiyuantang/pptx-to-pypptx.git \
  ~/.claude/skills/pptx-to-pypptx
```

Update anytime with `git -C ~/.claude/skills/pptx-to-pypptx pull`.

**Project scope — one repo only, shareable with collaborators:**

```bash
# run from the root of the repo you want the skill in
git clone https://github.com/hiyuantang/pptx-to-pypptx.git \
  .claude/skills/pptx-to-pypptx
```

To commit it so teammates get it on pull, drop the nested git history first
(`rm -rf .claude/skills/pptx-to-pypptx/.git`) so your repo tracks the files
directly instead of as a submodule, then commit `.claude/skills/`.

Either scope, the folder must be named `pptx-to-pypptx` with `SKILL.md` at its
root — the clones above do that. Start a new Claude Code session; the skill is
discovered automatically and triggers on requests like *"migrate this pptx to
python-pptx"*, *"sync my deck changes to code"*, or *"what's on slide 10?"*,
then reads [`SKILL.md`](./SKILL.md) and runs the round-trip for you. (No git?
Copy the repo folder to the same location instead.)

For distributing to many people, packaging this as a Claude Code plugin (shared
through a marketplace) is the scalable route; for personal or team use, the
clone above is the standard, simplest path.

## Repository layout

```
pptx-to-pypptx/
├── SKILL.md                 # skill entry point / full workflow docs
├── references/
│   └── SLIDE_FORMAT.md      # conventions for generated slide files
├── scripts/                 # command-line tools (run with uv)
│   ├── scaffold.py          # create a new project from a target deck
│   ├── generate_slides.py   # generate/refresh slides/sNN_*.py from a deck
│   ├── detect_project.py    # list existing projects
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
- Optional: [LibreOffice](https://www.libreoffice.org/) (`soffice`) for
  `extract_slide.py --screenshot` rendering.

## Quick start

Prefer to run the tools by hand instead of through the skill/agent? Replace
`<repo>` with wherever you cloned this repo (e.g.
`~/.claude/skills/pptx-to-pypptx` if you installed it at user scope).

```bash
# 1. Scaffold a project from your deck (the built deck is named after the output dir)
uv run python <repo>/scripts/scaffold.py \
  --target "deck.pptx" \
  --output-dir my-deck

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

## License

Copyright 2026 Yuan Tang.

Licensed under the [Apache License, Version 2.0](./LICENSE). You may not use
this project except in compliance with the License.
