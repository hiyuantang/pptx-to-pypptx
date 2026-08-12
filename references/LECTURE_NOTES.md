# Lecture Notes

Turn a deck's speaker notes into learner-facing Markdown, then add only the diagrams and images that materially support the prose. Treat this as a close editorial transformation, not a summary, transcript dump, or slide-image conversion.

- [Output contract](#output-contract)
- [Build a slide-linked source](#build-a-slide-linked-source)
- [Transform the prose first](#transform-the-prose-first)
- [Organize the Markdown](#organize-the-markdown)
- [Select visuals from the edited prose](#select-visuals-from-the-edited-prose)
- [Audit visual coverage before extraction](#audit-visual-coverage-before-extraction)
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

## Audit visual coverage before extraction

After the prose pass and initial visual selection, create a scratch-only visual coverage ledger. This is the completeness check that prevents selective extraction from becoming accidental under-illustration. It is not part of the final Markdown.

Review every temporary slide preview and account for:

- each substantive teaching claim whose source contains a process, architecture, model, spatial relationship, annotated result, worked visual example, or application image; and
- each multi-slide visual build family, using its final complete state plus any genuinely distinct contrast states.

Use one ledger row per claim or visual family:

| Prose section / teaching claim | Source slide(s) or build family | Candidate visual | Disposition | Reason / QA result |
|---|---|---|---|---|
| Concept heading and the relationship being taught | Real deck numbers | Diagram, image, or none | `keep`, `omit`, or `retry` | Instructional role, redundancy reason, extraction issue, or final asset name |

Apply these rules:

- `keep` means the visual contributes a relationship or concrete example that prose alone does not communicate as efficiently. Record its semantic filename and intended Markdown placement after extraction.
- `omit` requires a specific reason such as decorative, quiz-only, ordinary text moved into prose, exact duplicate, or genuinely redundant spatial content. “There are already enough images” is never a reason.
- `retry` means the visual is instructionally useful but the first extraction is clipped, illegible, opaque when it should be transparent, or otherwise unfaithful. It remains unresolved until repaired or changed to a justified `omit` after reasonable extraction attempts.
- A prose sentence that merely names the boxes in a pipeline or the stages in a model does not automatically make the source diagram redundant. Preserve a useful visual when the ordering, connection, grouping, direction, highlighting, or label attachment is part of the lesson.
- A low or high asset count is not evidence of quality. Completeness comes from reviewing all claims and build families, while selectivity comes from the recorded dispositions.

Do not finalize while a ledger row remains `retry`. Keep the ledger in scratch space for final review and delete neither its omission reasons nor its opaque-asset justifications until delivery is accepted.

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

Do not interpret a bad first render as evidence that the source diagram is unnecessary. For a useful diagram that renders poorly:

1. inspect the final complete state in its slide-build family;
2. inspect the slide's object hierarchy and distinguish top-level group IDs from child IDs;
3. retry with only the diagram, connectors, labels, and spatial callouts, omitting titles, prose, and chrome;
4. compare the asset with the source preview at the expected Markdown display width on both light and dark backgrounds; and
5. keep it as `retry` in the coverage ledger until it is faithful, or record why the source cannot be rendered cleanly and preserve its teaching content in prose or math.

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

True alpha is required for diagrams assembled from native shapes, arrows, labels, and callouts. A fully opaque canvas around a native-shape selection is an extraction failure: retry the shape selection or matte render rather than accepting a black, white, or slide-colored rectangle. Inspect edge antialiasing and all labels on both light and dark backgrounds.

Judge contrast from explicit light- and dark-background composites, not from the transparent PNG alone: image viewers may display transparency as black, white, or a checkerboard. Distinguish the transparent exterior from retained diagram fills. White text inside an opaque blue box, for example, keeps the same contrast on either page background and must not be rejected merely because the asset has alpha. If a label instead floats directly over transparency, or sits on a semitransparent source panel whose contrast changes with the page, retry with its contrast-preserving backing shape; omit the visual only when a faithful, legible composite still cannot be produced.

An embedded screenshot, photograph, UI panel, or source image whose background is intrinsic to the visual may remain opaque. Crop it tightly, keep slide chrome out, record the reason in the coverage ledger, and explicitly allowlist it in the strict validator. For other extracted rasters, transparency remains a preference rather than permission to erase intended white marks, formulas, labels, or non-flat backgrounds. The helper removes only flat, edge-connected background pixels; it does not globally key out every white pixel or flatten an existing alpha channel.

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

1. Review the completed visual coverage ledger. Account for every substantive teaching point and visual build family as prose, math, a selected visual, or a deliberate non-instructional omission; no row may remain `retry`.
2. Confirm the concept order still follows the speaker notes unless a small reordering clearly improves written coherence without changing the teaching logic.
3. Remove lecturer performance language, in-lecture quiz interactions, unresolved slide references, transcription artifacts, and unsupported additions.
4. Verify every equation, symbol, dimension, example, and contrast against the slide and speaker notes.
5. Confirm no final asset is a temporary full-slide preview, bullet-list capture, text-only slide, title/divider, or decorative slide image.
6. Open every final asset. Inspect legibility at expected display width, transparency on light and dark backgrounds, and whether its callouts are spatially relevant. Confirm that no native-shape diagram has an opaque canvas.
7. Resolve every relative image link and confirm that the Markdown refers only to files that exist in `lecture-notes-assets/`.
8. Read the Markdown once as a learner. It should be self-contained, semi-formal, and recognizably the same lecture—not a shortened substitute for it.

Run the deterministic link and asset check in strict-transparency mode. Omit `--allow-opaque` when every asset has alpha. Repeat it only for deliberately opaque screenshots, photographs, or intrinsic panels recorded in the coverage ledger:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/validate_lecture_notes.py \
  lecture-notes.md --assets-dir lecture-notes-assets \
  --strict-transparency \
  --allow-opaque application-screenshot.png \
  --allow-opaque source-photograph.jpg
```

The validator rejects temporary `slide-images/slide_N.png` references so source previews cannot leak into the final notes. In strict mode it also rejects every opaque raster that is not exactly allowlisted, and rejects stale allowlist entries that no longer identify a linked opaque asset. Resolve every error and warning before delivery. If the only remaining warning is that the document has no image links, the coverage ledger must demonstrate that the source lecture truly has no instructionally useful visual.
