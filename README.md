# pptx-to-pypptx

Turn an existing PowerPoint deck (`.pptx`) into a maintainable
[`python-pptx`](https://python-pptx.readthedocs.io/) project, edit the generated
code, and regenerate the deck — a clean round-trip loop between binary slides and
readable Python.

This repository is packaged as a [Claude Code](https://claude.com/claude-code)
**skill** ([`SKILL.md`](./SKILL.md)), but the scripts under [`scripts/`](./scripts)
are plain command-line tools and can be run on their own.

## Demo

https://github.com/user-attachments/assets/246ab7ae-b65a-44b0-83a7-6546c5fc0a19

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
  slides — the agent auto-detects what you changed in PowerPoint, so you never
  have to say which slides.
- **Inspect** any slide's shapes/positions/text (and optionally render a PNG)
  without opening PowerPoint.

## Installation

This repo **is** the skill (`SKILL.md` is at its root), so installing it just
means cloning it into a Claude Code skills directory. Its Python dependencies
install separately, per deck project (see [Requirements](#requirements)).

**User scope** — available in every project on your machine:

```bash
git clone https://github.com/hiyuantang/pptx-to-pypptx.git ~/.claude/skills/pptx-to-pypptx
```

**Project scope** — available only in one repo:

```bash
git clone https://github.com/hiyuantang/pptx-to-pypptx.git .claude/skills/pptx-to-pypptx
```

Add `.claude/skills/pptx-to-pypptx/` to that repo's `.gitignore` so the clone
stays a local install. Then start a new Claude Code session — the skill is
auto-discovered and triggers on requests like *"migrate this pptx to
python-pptx"* or *"what's on slide 10?"*.

## Updating

Ask the agent to **"upgrade the pptx-to-pypptx skill."** It pulls the latest
version, checks the diff, and re-baselines your deck projects only if the
tooling changed. (It's a git clone, so `git -C <clone-dir> pull` also works.)

## Requirements

- Python 3.10+
- [`uv`](https://docs.astral.sh/uv/) for environment and command running
- Python packages (the agent installs these into each deck **project**, not this
  repo): `python-pptx>=1.0.0`, `cairosvg>=2.0`, `pillow-heif>=1.0`
- Optional: [LibreOffice](https://www.libreoffice.org/) for rendering slides to
  images.

## Using it

You don't run any commands yourself — you talk to your coding agent (Claude
Code) and it drives the tools for you. Point it at a deck and ask in plain
language:

- *"Migrate `talk.pptx` into a python-pptx project."*
- *"What's on slide 10?"*
- *"Change the title on slide 3 to 'Results' and make it navy."*
- *"Add a slide after slide 12 titled 'Next steps.'"*
- *"Delete slide 7."*
- *"Put `logo.png` in the top-right corner of the title slide."*
- *"Roll back to the previous build."*
- *"Upgrade the pptx-to-pypptx skill."*

The agent scaffolds the project, writes one Python file per slide, rebuilds the
`.pptx`, and keeps code and deck in sync. [`SKILL.md`](./SKILL.md) is the
playbook it follows.

### Editing in PowerPoint syncs automatically

Edit `out/<name>.pptx` in PowerPoint, save, and just keep working with the agent
— you don't ask for a sync or say which slides you changed. Next time you have it
work on the deck, it detects your edits and regenerates only the changed slides
back into code. It's deck→code only, so it never rebuilds or overwrites the file
you just saved. Changed masters, layouts, or the theme? Those don't auto-sync —
ask the agent to refresh them.

### Your project

The agent creates a project folder (you choose the name):

```
my-deck/
├── slides/            # one .py file per slide — the deck's content
├── lib/
│   ├── design.py      # colors, fonts, spacing — the deck's theme
│   ├── shapes.py      # drawing engine — the agent's toolbox
│   └── base.pptx      # template shell (layouts/theme) the deck is built from
├── assets/            # images, video, GIFs, SVGs used by the slides
├── out/
│   └── my-deck.pptx   # the built deck — open, edit, and share this one
├── backup/            # the last 10 builds, auto-saved on every rebuild
└── build_deck.py      # generated plumbing
```

- **Open, edit, and share `out/<name>.pptx`.** That's the finished deck. Edit it
  in PowerPoint and save, and your changes sync back into code automatically —
  see [Editing in PowerPoint](#editing-in-powerpoint-syncs-automatically).
- **Add media** by dropping images, video, or GIFs into `assets/` (or just hand
  the agent the file path) and asking the agent to place them on a slide.
- **Roll back anytime** — every successful build is saved to `backup/` as
  `backup_<timestamp>.pptx` (the last 10 are kept), so you can ask the agent to
  restore an earlier one.
- **Prefer to edit the code yourself?** You don't have to, but you can: the deck
  is real Python. `slides/*.py` (each slide's content) and `lib/design.py` (theme
  colors and fonts) are safe to change by hand. Leave `build_deck.py`,
  `lib/shapes.py`, and `lib/base.pptx` alone — they're the engine — and don't rename files in
  `slides/` (their names set the slide order); ask the agent to add, remove, or
  reorder slides instead.
- **Your original `.pptx` is never modified** — it stays as a read-only
  reference.

## Supported features

Rectangles and preset shapes, lines/arrows/connectors, solid/gradient/pattern/
theme fills, rich text (runs, fonts, colors, bullets, alignment, hyperlinks,
Office Math), shape effects (shadow/glow/reflection/soft-edge), images (with
cropping; SVG/WebP/HEIC conversion) and freeform vector shapes, tables with
merged cells, charts, embedded video, true group shapes with group-relative
child positioning, slide backgrounds, speaker notes, and hidden slides.

Anything truly exotic (SmartArt, animations, transitions, custom geometry that
can't be vectorized) is emitted as a `# TODO` comment in the generated slide.

## Repository layout

```
pptx-to-pypptx/
├── SKILL.md                 # skill entry point / full workflow docs
├── references/
│   └── SLIDE_FORMAT.md      # conventions for generated slide files
├── scripts/                 # command-line tools (run with uv)
│   ├── scaffold.py          # create a new project from a target deck
│   ├── generate_slides.py   # generate/refresh slides/sNN_*.py from a deck
│   ├── autosync.py          # auto-sync code when the deck is edited in PowerPoint
│   ├── detect_project.py    # list existing projects
│   ├── extract_slide.py     # dump/screenshot a single slide
│   ├── extract_notes.py     # export speaker notes to Markdown
│   ├── sync_slide_numbers.py# reserve/close slide-number slots
│   ├── list_layouts.py      # list slide layouts in a deck
│   └── helpers/             # internal modules (imported by the scripts, not run directly)
└── template/                # files copied into each scaffolded project
    ├── build_deck.py
    └── lib/                 # design.py + shapes.py + roundtrip_state.py helper library
```

## License

Copyright 2026 Yuan Tang.

Licensed under the [Apache License, Version 2.0](./LICENSE). You may not use
this project except in compliance with the License.
