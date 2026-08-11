# Lecture Notes

Turn a deck's speaker notes and instructional visuals into learner-facing Markdown. Treat this as a close editorial transformation, not a summary and not a transcript dump.

- [Output contract](#output-contract)
- [Prepare the source](#prepare-the-source)
- [Transform the prose](#transform-the-prose)
- [Organize the Markdown](#organize-the-markdown)
- [Select instructional visuals](#select-instructional-visuals)
- [Extract and prepare assets](#extract-and-prepare-assets)
- [Quality gate](#quality-gate)

## Output contract

- Deliver one Markdown file. Use the user's path; otherwise write `lecture-notes.md` beside the deck project.
- Put final visual assets in a sibling `lecture-notes-assets/` folder.
- Link assets with portable relative paths such as `![Single-neuron computation](lecture-notes-assets/single-neuron-computation.png)`.
- Produce Markdown, not HTML. Use standard headings, paragraphs, lists, tables, fenced code, and LaTeX math where appropriate.
- Keep the source deck and speaker notes unchanged unless the user separately asks to edit them.
- Use the deck and its notes as the content source. Do not add outside facts or silently repair a substantive claim through research unless the user asks.
- Deliver only final assets in `lecture-notes-assets/`; keep inventories, renders, and comparison scratch files outside it.

## Prepare the source

1. Identify the exact deck and the lecture or unit boundaries. Do not combine separate lectures merely because they share one deck.
2. If the deck belongs to a generated project, run `autosync.py` first so the slide code and notes reflect the current PowerPoint file.
3. Choose the current working deck: `out/<name>.pptx` for a generated project, or the user-supplied `.pptx` otherwise. Export its speaker notes directly; do not scaffold a project just to read notes:

   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/extract_notes.py \
     --target "<working-deck.pptx>" --output <scratch-dir>/speaker-notes.md
   ```

   Use `--project-dir <project-dir>` only when no current `.pptx` output exists and the generated slide files are the source of truth.

4. Read the speaker notes in order and inspect every corresponding slide.
5. Inspect slide structure and media when needed:

   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/extract_slide.py \
     "<deck.pptx>" all --verbose --json
   ```

6. Extract original embedded visual candidates and their provenance into scratch space:

   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/extract_lecture_assets.py \
     "<working-deck.pptx>" --slides <lecture-range> \
     --output-dir <scratch-dir>/candidate-assets \
     --manifest <scratch-dir>/lecture-assets.json
   ```

   The manifest maps each deduplicated asset to its slide usages, bounds, crop/luminance transforms, transparency, exact repeats, and composite render candidates. It flags formats and slide treatments that require a region render. It does not decide which progressive state to keep.

7. Build a scratch coverage ledger with one row per substantive teaching point and one row per instructional visual. Track the source slide(s), intended section, chosen asset, and any deliberate omission. Do not place this ledger in the final Markdown or asset folder.

Speaker notes are the primary narrative source. Slide text, equations, and visuals resolve references and verify notation; they are not permission to replace the notes with a new outline.

## Transform the prose

Preserve the original teaching sequence, terminology, definitions, equations, examples, contrasts, caveats, and conclusions. The result should feel extremely close to the speaker notes in subject matter while reading as a compact course text.

Apply these edits:

- Change lecturer-centered subjects into concept-centered prose. Replace “I will show” or “let's look at” with constructions such as “This section examines” or “The model uses.”
- Retain an occasional inclusive “we” when it genuinely guides a derivation or shared observation. Do not mechanically remove every first-person plural.
- Remove delivery scaffolding: greetings, roadmap chatter, timing remarks, slide directions, rhetorical applause lines, filler, false starts, and repeated punchlines.
- Remove in-lecture quizzes, polls, knowledge checks, and answer-choice interactions. Do not reproduce their prompts, options, pauses, answer reveals, scoring, or quiz-only visuals. If an interaction introduces substantive teaching content, state that concept or explanation directly; otherwise omit it. Keep ordinary worked questions and examples when their reasoning is itself instructional.
- Resolve deictic language. Replace vague “this,” “here,” “that one,” or “on the right” with the named idea, object, equation, or figure.
- Repair obvious spoken slips or transcription errors only when the slide or nearby notes make the intended wording unambiguous. Flag substantive ambiguity instead of guessing.
- Merge spoken repetition, but preserve the teaching function. A repeated definition may become one definition; a later contrast, consequence, or recap remains.
- Convert a spoken comparison into prose, bullets, or a small table when that makes the relationship clearer.
- Keep explanatory motivation around equations. Do not reduce a derivation to formulas alone.
- Prefer semi-formal sentences and precise terminology over conversational performance language. Avoid making the prose stiff, impersonal, or encyclopedic.

The transformation should normally stay close in length and concept coverage. Compression comes from removing performance language and repetition, not from deleting instruction.

## Organize the Markdown

- Use the lecture title as the single `#` heading.
- Organize by concepts with `##` and, when useful, `###` headings. Do not create one section per slide and do not expose slide numbers as the learner-facing structure.
- Open with a short orienting paragraph only when the notes provide that orientation.
- Keep paragraphs cohesive and moderately short. Use lists for real sets, steps, dimensions, conditions, or contrasts—not as the default prose form.
- Write inline math as `$...$` and display math as `$$...$$`. Preserve variable names, subscripts, hats, Greek symbols, and matrix dimensions exactly.
- Introduce each visual in the prose and place it immediately after the passage it supports. Do not append an unexplained image gallery.
- Give every image concise, meaningful alt text. The surrounding prose must still communicate the instructional takeaway for a learner who cannot see the image.
- For a complex visual whose important labels or relationships are not already conveyed in the prose, add a short Markdown paragraph beginning `**Figure description.**`. Do not repeat the same description around every image or depend on HTML-only disclosure widgets.
- End with a takeaway only when the speaker notes contain a genuine synthesis; do not add a generic recap to every lecture.

Use this shape as a guide, not a mandatory template:

```markdown
# Lecture title

Brief orientation grounded in the speaker notes.

## First concept

Semi-formal explanation with preserved notation, motivation, and detail.

![Meaningful description](lecture-notes-assets/first-concept.png)

## Consequence or comparison

- First precise relationship
- Second precise relationship
```

## Select instructional visuals

Start from all visuals in the lecture, then exclude only those that do not help the learner. Keep diagrams, charts, worked examples, annotated equations, model outputs, meaningful screenshots, and other visuals referenced or explained by the notes. Omit decorative backgrounds, logos, repeated chrome, ornamental photos, navigation, empty layout elements, and quiz prompts, answer-choice screens, or answer-reveal states.

Treat adjacent or repeated visuals as a build family and classify the change before selecting assets:

| Change between states | Default selection |
|---|---|
| Additive progressive reveal of the same idea | Keep the final complete state only. |
| Highlight, pointer, or emphasis change only | Keep the clean or final state unless the emphasis performs a teaching function in the prose. |
| Different notation, model, case, result, or counterexample | Keep each state that supports a distinct teaching point. |
| A later state removes or replaces an earlier component to teach a contrast | Keep both states when the comparison matters. |
| Exact or near-exact repetition across slides | Keep one copy and link it once at the best explanatory point. |

Similarity alone is not enough to discard an image. Two nearly identical diagrams may represent different models; several visually progressive slides may communicate only one final idea. Decide from the speaker notes and teaching purpose.

## Extract and prepare assets

For each selected visual, use the highest-fidelity source available:

1. **Single embedded picture:** use the lossless candidate produced by `extract_lecture_assets.py`. Preserve its native alpha channel and resolution. Do not use a slide screenshot when the source image is available and already matches its visible slide treatment.
2. **Vector artwork:** prefer a transparent SVG when it renders faithfully in the target Markdown environment. Otherwise export a high-resolution PNG with alpha.
3. **Composite visualization made from slide shapes:** export the logical group or its tight bounds from a scratch copy or asset-only slide, excluding the slide background, title, footer, and unrelated objects. Prefer a transparent background; never alter the source deck for asset extraction.
4. **Chart, table, equation, or composite that cannot be separated safely:** use a tightly cropped high-resolution render. An opaque neutral background is acceptable when removing it would erase white marks, labels, or other intended content.

Native Office Math can display correctly in PowerPoint while appearing blank in a LibreOffice render when the equation's compatibility fallback is unavailable. Inspect every rendered equation before using it as an asset. If a renderer drops the math, do not deliver the blank callout or capture the migrated deck as a substitute: render the original source with a verified fallback, or preserve the equation as Markdown math and omit the redundant image.

Prepare a source image or a selected slide region as a PNG:

```bash
# Preserve/crop an extracted raster and remove only a flat edge-connected background
uv run python <pptx-to-pypptx-dir>/scripts/prepare_lecture_asset.py \
  --input-image <scratch-dir>/candidate-assets/<source.png> \
  --output lecture-notes-assets/<semantic-name>.png \
  --transparent --trim --padding 20

# Render a composite region; bounds are x,y,w,h in slide inches
uv run python <pptx-to-pypptx-dir>/scripts/prepare_lecture_asset.py \
  --target "<working-deck.pptx>" --slide <N> --bounds <x,y,w,h> \
  --output lecture-notes-assets/<semantic-name>.png --dpi 300 \
  --transparent --trim --padding 20
```

Use `--transparent` only for a flat background around the visual. The helper removes matching pixels only when they are connected to an outer edge; enclosed white labels and equation fills remain opaque. It refuses ambiguous backgrounds and images with fewer than three opaque corners unless `--background '#RRGGBB'` is explicit. If the image already has useful alpha, omit `--transparent`. If the result is damaged or the background is not flat, rerun without `--transparent` and keep a tight opaque crop.

Transparency is a preference, not permission to damage a visual. Never key out all white pixels or flatten an existing alpha channel. Inspect transparent assets on both light and dark backgrounds so white strokes, formulas, and labels remain visible.

Additional asset rules:

- Crop tightly but leave consistent breathing room around the teaching object.
- Preserve readable labels at the Markdown page's expected display width. Re-export rather than enlarging a blurry crop.
- Keep meaningful source color and annotation semantics.
- Use semantic lowercase filenames such as `matrix-dimensions.png` or `perceptron-without-activation.svg`; avoid raw names such as `image17.png` when renaming does not break provenance.
- Avoid duplicate files. If one asset supports two nearby passages, place it once where it best anchors the explanation.
- Do not use absolute paths, data URIs, remote edX asset URLs, or files outside `lecture-notes-assets/` in the final Markdown.

## Quality gate

Before delivery:

1. Compare the completed note against the coverage ledger. Account for every substantive point as prose, math, a visual, or a deliberate non-instructional omission.
2. Confirm the concept order still follows the speaker notes unless a small reordering clearly improves written coherence without changing the teaching logic.
3. Remove lecturer performance language, in-lecture quiz interactions, unresolved slide references, transcription artifacts, and unsupported additions.
4. Verify every equation, symbol, dimension, example, and contrast against the slide and speaker notes.
5. Open every final asset, inspect transparency and legibility, and confirm progressive-build choices match the prose.
6. Resolve every relative image link and confirm that the Markdown refers only to files that exist in `lecture-notes-assets/`.
7. Read the Markdown once as a learner. It should be self-contained, semi-formal, and recognizably the same lecture—not a shortened substitute for it.

Run the deterministic link and asset check, then resolve every error and inspect its transparency warnings:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/validate_lecture_notes.py \
  lecture-notes.md --assets-dir lecture-notes-assets
```
