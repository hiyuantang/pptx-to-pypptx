# Slide file format

Each slide is one `slides/sNN_description.py` file. `build_deck.py` imports them in **sorted filename order** and calls `add_slide(prs, slide, n)`.

## Required shape

```python
TITLE = 'Slide title'   # optional, used by build_deck.py
LAYOUT = 0              # required; layout index from the target deck
                        # run `list_layouts.py --target <target.pptx>` to see indices

def add_slide(prs, slide, n):
    """Draw the slide. n is the 1-based deck slide number."""
```

## The rule

**Use helpers from `lib/shapes.py`.** Never hand-craft XML or call low-level `python-pptx` APIs.

Grep a helper to read its docstring and signature:

```bash
grep -n -A 20 "^def add_box" lib/shapes.py
```

## Available helpers

- `add_shape` — any preset shape (oval, chevron, etc.)
- `add_box` — colored rectangle/rounded box
- `add_text` — text box
- `add_label` — simple centered label
- `add_line` / `add_arrow` / `add_connector` — lines and connectors (support `rotation=`; `add_connector` also takes `preset=` for an exact bent/curved variant)
- `connect_shapes` — attach a connector's ends to shapes (`stCxn`/`endCxn`) so bent connectors reroute like the source
- `add_callout` — pre-styled callout box
- `add_image` — image from `assets/`
- `add_chart` — chart
- `add_movie` — embedded video from `assets/`
- `add_custom_table` — table with per-cell control
- `add_group` / `set_group_bounds` — group shapes
- `add_background` — slide background
- `add_notes` — speaker notes
- `add_comment` — PowerPoint comment thread (see **Comments** below)
- `set_slide_hidden` — hide/unhide the slide (`set_slide_hidden(slide)` to hide)

## Chrome

Draw chrome as normal shapes:

- `shapes.add_box(..., slide_number=n)`
- `shapes.add_box(..., name='Footer' / 'Date' / 'Header' / 'Title')`

## Hidden slides

A slide hidden in the source deck (`<p:sld show="0">`) is detected on
generation and emitted as `shapes.set_slide_hidden(slide)` at the top of
`add_slide`. To hide/unhide a slide by hand, add or remove that call — visible
is the default, so there is no `show="1"`. `extract_slide.py` flags hidden
slides (`[HIDDEN]` in text output, `"hidden": true` in `--json`).

## Don'ts

- No `main` block, no `prs.save()`.
- No manual filename renumbering; use `sync_slide_numbers.py`.
- Don't suppress `# TODO` comments without flagging them.
- Don't hand-edit the fenced comment region (see below) — `autosync.py` owns it.

## Comments

PowerPoint comments live at the end of `add_slide`, inside a fence:

```python
    # --- comments (managed; edit in PowerPoint, synced automatically) ---
    shapes.add_comment(slide, author='Gumb, Chris', created='2025-02-08T22:55:48.036',
        text='This is confusing. We should probably just call V the vocabulary size...')
    shapes.add_comment(slide, author='Yuan Tang', created='2025-02-09T15:16:04.798', resolved=True,
        text='Does something seem strange about this equation?',
        replies=[
            {'author': 'Yuan Tang', 'created': '2025-02-09T15:16:48.984', 'text': 'Or is this a consequence of...'},
        ])
    # --- end comments ---
```

`add_comment` draws nothing — like `add_notes`, it writes to a part python-pptx
owns separately. The call carries only what a human wrote; thread GUIDs, change
cookies, and author identities are derived or looked up at build time (see
`lib/comments.py` and `lib/authors.json`).

**Read this region, don't write it.** It is how you see what reviewers asked for
on a slide you're editing. `autosync.py` replaces everything between the two fence
lines from the deck on every task, so hand edits there are overwritten — but code
outside the fence is never touched. To add your own comment, use
`add_comment.py`, which writes the call for you (with `pending=True` until the
next build).
