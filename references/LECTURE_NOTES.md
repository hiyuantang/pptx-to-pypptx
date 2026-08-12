# Lecture Notes

Turn a deck's speaker notes into learner-facing Markdown, then add only the diagrams and images that materially support the prose. Treat this as a close editorial transformation, not a summary, transcript dump, or slide-image conversion.

- [Output contract](#output-contract)
- [Build a slide-linked source](#build-a-slide-linked-source)
- [Transform the prose first](#transform-the-prose-first)
- [Organize the Markdown](#organize-the-markdown)
- [Select visuals from the edited prose](#select-visuals-from-the-edited-prose)
- [Extract and prepare assets](#extract-and-prepare-assets)
- [Finalize the learner-facing Markdown](#finalize-the-learner-facing-markdown)
- [Quality gate](#quality-gate)

## Output contract

- Deliver one final Markdown file. Use the user's path; otherwise write `lecture-notes.md` beside the deck project.
- Put final visual assets in a sibling `lecture-notes-assets/` folder.
- Link assets with portable relative paths such as `![Single-neuron computation](lecture-notes-assets/single-neuron-computation.png)`.
- Produce Markdown, not HTML. Use standard headings, paragraphs, lists, tables, fenced code, and LaTeX math where appropriate.
- Keep the source deck and speaker notes unchanged unless the user separately asks to edit them.
- Use the deck and its notes as the content source. Do not add outside facts or silently repair a substantive claim through research unless the user asks.
- Deliver only selected diagrams, images, and related visual callouts in `lecture-notes-assets/`. Keep full-slide previews, inventories, manifests, renders, and comparison files in scratch space.
- Do not target an image count or turn each slide into a picture. A good section may need no asset.

## Build a slide-linked source

1. Identify the exact deck and lecture or unit boundaries. Do not combine separate lectures merely because they share one deck.
2. If the deck belongs to a generated project, run `autosync.py` first so the notes reflect the current PowerPoint file.
3. Choose the current working deck: `out/<name>.pptx` for a generated project, or the user-supplied `.pptx` otherwise.
4. Export the speaker notes and a temporary reference image for every selected slide into scratch space:

   ```bash
   uv run python <pptx-to-pypptx-dir>/scripts/extract_notes.py \
     --target "<working-deck.pptx>" --slides <lecture-range> \
     --output <scratch-dir>/speaker-notes.md \
     --slide-images-dir <scratch-dir>/slide-images
   ```

   The Markdown keeps the real deck number in each `## Slide N: ...` source heading, adds a `lecture-source-slide` marker, and places the matching temporary slide image inside a marked preview block. These images are visual references for editing; they are never final lecture-note assets.

   Use `--project-dir <project-dir>` only when no current `.pptx` exists and the generated slide files are the source of truth. That mode exports numbered notes but cannot render slide references without a built deck.

5. Read the notes and matching preview together, in order. Speaker notes are the primary narrative source. Slide text, equations, and visuals resolve references and verify notation; they are not permission to replace the notes with a new outline.
6. Copy the source Markdown to a scratch draft. Keep source-slide provenance while editing, for example `<!-- source-slides: 12-15 -->`, so every passage can still be checked against the deck. The finalizer removes these owned markers later.

## Transform the prose first

Finish the prose pass before deciding what to extract. Work section by section, or at paragraph/teaching-claim granularity when a section spans several visual ideas.

Preserve the original teaching sequence, terminology, definitions, equations, examples, contrasts, caveats, and conclusions. The result should feel extremely close to the speaker notes in subject matter while reading as a compact course text.

Apply these edits:

- Change lecturer-centered subjects into concept-centered prose. Replace “I will show” or “let's look at” with constructions such as “This section examines” or “The model uses.”
- Retain an occasional inclusive “we” when it genuinely guides a derivation or shared observation. Do not mechanically remove every first-person plural.
- Remove delivery scaffolding: greetings, roadmap chatter, timing remarks, slide directions, rhetorical applause lines, filler, false starts, and repeated punchlines.
- Remove in-lecture quizzes, polls, knowledge checks, and answer-choice interactions. Do not reproduce their prompts, options, pauses, answer reveals, scoring, or quiz-only visuals. If an interaction introduces substantive teaching content, state that concept directly; otherwise omit it. Keep ordinary worked questions and examples when their reasoning is instructional.
- Resolve deictic language. Replace vague “this,” “here,” “that one,” or “on the right” with the named idea, object, equation, or figure.
- Repair obvious spoken slips or transcription errors only when the slide or nearby notes make the intended wording unambiguous. Flag substantive ambiguity instead of guessing.
- Merge spoken repetition while preserving its teaching function. A repeated definition may become one definition; a later contrast, consequence, or recap remains.
- Blend ordinary slide bullets and standalone text into the prose. Use a Markdown list only for a real set, sequence, dimensions, conditions, or contrast. Never preserve bullets merely by capturing them in an image.
- Preserve equations as Markdown math when the notation can be represented faithfully. Keep explanatory motivation around equations rather than reducing a derivation to formulas alone.
- Prefer semi-formal sentences and precise terminology over conversational performance language. Avoid making the prose stiff, impersonal, or encyclopedic.

The transformation should normally stay close in length and concept coverage. Compression comes from removing performance language and repetition, not from deleting instruction.

## Organize the Markdown

- Replace the `# Speaker Notes: ...` source title with the lecture title as the single `#` heading.
- Replace the temporary `## Slide N: ...` structure with concept-based `##` and, when useful, `###` headings. Do not expose slide numbers as learner-facing structure.
- Open with a short orienting paragraph only when the notes provide that orientation.
- Keep paragraphs cohesive and moderately short. Let bullets be part of the written explanation, not visual assets.
- Write inline math as `$...$` and display math as `$$...$$`. Preserve variable names, subscripts, hats, Greek symbols, and matrix dimensions exactly.
- Introduce each selected visual in the prose and place it immediately after the passage it supports. Do not append an unexplained image gallery.
- Give every image concise, meaningful alt text. The surrounding prose must still communicate the instructional takeaway for a learner who cannot see the image.
- For a complex visual whose important relationships are not already conveyed in the prose, add a short Markdown paragraph beginning `**Figure description.**`.
- End with a takeaway only when the speaker notes contain a genuine synthesis; do not add a generic recap to every lecture.

## Select visuals from the edited prose

Work outward from the edited prose, not inward from every slide. At each section—or for a finer-grained teaching claim—ask whether a spatial or visual relationship materially improves understanding.

Keep only:

- a diagram, model, process, architecture, chart, map, annotated equation, or other visual relationship explained by the notes;
- a meaningful source image, model output, or screenshot that the prose discusses; or
- callout labels, arrows, highlights, or captions whose spatial attachment to the diagram or image carries meaning.

Omit:

- temporary full-slide previews and ordinary slide screenshots;
- titles, divider slides, agendas, paragraphs, bullet lists, quotations, or standalone text that belongs in Markdown;
- a text box or callout whose meaning does not depend on where it points—rewrite it into the prose instead;
- decorative backgrounds, logos, repeated chrome, ornamental photos, navigation, and empty layout elements;
- quiz prompts, answer choices, answer-reveal states, and quiz-only visuals; and
- redundant visuals that restate what the prose already communicates clearly.

There is no one-image-per-slide or one-image-per-section quota. Some slides contribute only prose. Some multi-slide builds contribute one final diagram. A dense teaching claim may justify more than one visual when each shows a genuinely different case or result.

Treat adjacent or repeated diagrams as a build family:

| Change between states | Default selection |
|---|---|
| Additive progressive reveal of the same idea | Keep the final complete state only. |
| Highlight, pointer, or emphasis change only | Keep the clean or final state unless the emphasis performs a teaching function in the prose. |
| Different notation, model, case, result, or counterexample | Keep each state that supports a distinct teaching point. |
| A later state removes or replaces an earlier component to teach a contrast | Keep both states when the comparison matters. |
| Exact or near-exact repetition across slides | Keep one copy and link it once at the best explanatory point. |

Similarity alone is not enough to discard a diagram, but slide count is never a reason to keep one.

## Extract and prepare assets

Extract assets only from the slides selected during the prose pass. Use the highest-fidelity source available.

### Embedded pictures

Extract original embedded candidates and provenance from the relevant slides:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/extract_lecture_assets.py \
  "<working-deck.pptx>" --slides <selected-slides> \
  --output-dir <scratch-dir>/candidate-assets \
  --manifest <scratch-dir>/lecture-assets.json
```

Use a lossless candidate when it already matches the visible treatment. Preserve native alpha and resolution. If it is cropped, rotated, adjusted, or part of a composite, the manifest explains why a rendered selection may be more faithful.

Prepare an extracted raster conservatively:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/prepare_lecture_asset.py \
  --input-image <scratch-dir>/candidate-assets/<source.png> \
  --output lecture-notes-assets/<semantic-name>.png \
  --transparent --trim --padding 20
```

### Native-shape diagrams with arrows, text, and callouts

First inspect the slide's IDs and object hierarchy:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/extract_slide.py \
  "<working-deck.pptx>" <slide-number> --json --verbose
```

Then select the diagram objects and only the labels/callouts that belong to it:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/prepare_lecture_asset.py \
  --target "<working-deck.pptx>" --slide <slide-number> \
  --shape-ids <diagram-id,arrow-id,label-id,callout-id> \
  --output lecture-notes-assets/<semantic-name>.png \
  --dpi 300 --padding 20
```

`--shape-ids` is the deterministic equivalent of selecting those objects in PowerPoint and copying them. It keeps their native PPTX relationships, removes every unselected slide object, covers inherited background/master chrome, renders the same selection once on black and once on white, reconstructs clean alpha from the two mattes, and trims to the resulting artwork. This avoids the colored antialias fringe left by single-color chroma keying. Selecting a group ID keeps the complete group; selecting child IDs keeps only those children and their required group ancestry. Always inspect the result on light and dark backgrounds.

### Region-render fallback

Use a tightly bounded region only when the visual cannot be selected by shape ID without breaking it:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/prepare_lecture_asset.py \
  --target "<working-deck.pptx>" --slide <slide-number> --bounds <x,y,w,h> \
  --output lecture-notes-assets/<semantic-name>.png --dpi 300 \
  --transparent --trim --padding 20
```

A region render can accidentally capture neighboring bullets or text, so prefer `--shape-ids`. An opaque tight crop is acceptable when transparency would erase intended white marks, formulas, labels, or a non-flat background. Never use the temporary whole-slide reference as this fallback.

Native Office Math can display correctly in PowerPoint while appearing blank in LibreOffice when its compatibility fallback is unavailable. Inspect every rendered equation. If rendering drops the math, preserve it as Markdown math or use a verified source render; never deliver a blank callout.

Additional asset rules:

- Crop tightly but leave consistent breathing room around the teaching object.
- Preserve readable labels at the Markdown page's expected display width. Re-export rather than enlarging a blurry crop.
- Keep meaningful source color and annotation semantics.
- Preserve SVG when it renders faithfully; otherwise use a high-resolution PNG.
- Use semantic lowercase filenames such as `matrix-dimensions.png`; avoid raw names such as `image17.png`.
- Avoid duplicate files. Place a reused asset once where it best anchors the explanation.
- Do not use absolute paths, data URIs, remote asset URLs, or files outside `lecture-notes-assets/` in the final Markdown.

Transparency is a preference, not permission to damage a visual. The helper removes only flat, edge-connected background pixels; it does not globally key out every white pixel or flatten an existing alpha channel.

## Finalize the learner-facing Markdown

The edited draft may retain temporary slide previews and hidden `lecture-source-slide` or `source-slides` markers for verification. Once the prose is concept-organized and final assets are placed, remove those source-only elements deterministically:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/finalize_lecture_notes.py \
  <scratch-dir>/lecture-notes.draft.md \
  --output lecture-notes.md
```

The finalizer removes marked preview blocks and owned provenance comments. It refuses to write a file if `# Speaker Notes:` or `## Slide N: ...` remains, because deleting the number alone would hide an unedited slide-by-slide structure rather than fix it.

## Quality gate

Before delivery:

1. Account for every substantive teaching point as prose, math, a selected visual, or a deliberate non-instructional omission.
2. Confirm the concept order still follows the speaker notes unless a small reordering clearly improves written coherence without changing the teaching logic.
3. Remove lecturer performance language, in-lecture quiz interactions, unresolved slide references, transcription artifacts, and unsupported additions.
4. Verify every equation, symbol, dimension, example, and contrast against the slide and speaker notes.
5. Confirm no final asset is a temporary full-slide preview, bullet-list capture, text-only slide, title/divider, or decorative slide image.
6. Open every final asset. Inspect legibility, transparency on light and dark backgrounds, and whether its callouts are spatially relevant.
7. Resolve every relative image link and confirm that the Markdown refers only to files that exist in `lecture-notes-assets/`.
8. Read the Markdown once as a learner. It should be self-contained, semi-formal, and recognizably the same lecture—not a shortened substitute for it.

Run the deterministic link and asset check, then resolve every error and inspect its transparency warnings:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/validate_lecture_notes.py \
  lecture-notes.md --assets-dir lecture-notes-assets
```

The validator rejects temporary `slide-images/slide_N.png` references so source previews cannot leak into the final notes.
