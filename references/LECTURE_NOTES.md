# Lecture Notes

Turn a deck's speaker notes and instructional visuals into learner-facing notes that can replace attending the lecture for study purposes. Treat this as a close editorial transformation, not a summary, transcript dump, slide-image conversion, or newly designed explainer.

- [Output contract](#output-contract)
- [Definition of success](#definition-of-success)
- [Failure modes to avoid](#failure-modes-to-avoid)
- [Build a slide-linked source](#build-a-slide-linked-source)
- [Transform the prose first](#transform-the-prose-first)
- [Audit content coverage](#audit-content-coverage)
- [Organize the Markdown](#organize-the-markdown)
- [Select visuals from the edited prose](#select-visuals-from-the-edited-prose)
- [Audit visual coverage before extraction](#audit-visual-coverage-before-extraction)
- [Extract and prepare assets](#extract-and-prepare-assets)
- [Finalize the learner-facing Markdown](#finalize-the-learner-facing-markdown)
- [Optional standalone HTML](#optional-standalone-html)
- [Quality gate](#quality-gate)

## Output contract

- Deliver one canonical Markdown file. Use the user's path; otherwise write `lecture-notes.md` beside the deck project.
- Put final visual assets in a sibling `lecture-notes-assets/` folder.
- Link assets with portable relative paths such as `![Single-neuron computation](lecture-notes-assets/single-neuron-computation.png)`.
- Use standard headings, paragraphs, lists, tables, fenced code, and LaTeX math where appropriate. Raw HTML `<img>` tags are allowed when they are needed to preserve a source-relative display width; the image source must still be a portable relative path under `lecture-notes-assets/`.
- When the user requests HTML, deliver it as a companion to the canonical Markdown. Make the HTML standalone by embedding the final image bytes; do not replace or redraw slide visuals with HTML/CSS shapes.
- Keep the source deck and speaker notes unchanged unless the user separately asks to edit them.
- Use the deck and its notes as the content source. Do not add outside facts or silently repair a substantive claim through research unless the user asks.
- Deliver only selected diagrams, images, and related visual callouts in `lecture-notes-assets/`. Keep full-slide previews, inventories, manifests, renders, and comparison files in scratch space.
- Do not target an image count or turn each slide into a picture. A good section may need no asset.

## Definition of success

The completed notes pass five tests:

1. **Replacement test.** A learner can understand the lecture without watching the video or opening the slides. The document contains the reasoning, explanation, examples, caveats, and conclusions—not merely the topic names.
2. **Coverage test.** Every substantive teaching unit in the speaker notes is present, merged with an exact repetition, or deliberately excluded for one of the narrow non-instructional reasons below. A shorter document is not automatically a better document.
3. **Visual test.** Every instructional spatial relationship and concrete visual example is represented by an extracted source visual unless it is genuinely redundant with accessible prose, math, a real Markdown table, or a real Markdown list.
4. **Fidelity test.** Extracted visuals preserve the PowerPoint objects, font rendering, geometry, colors, meaningful internal fills, and source-relative size. The notes do not invent replacement diagrams.
5. **Delivery test.** The Table of Contents, alt text, links, transparency, light/dark legibility, image sizing, and any requested standalone HTML have all been verified from the delivered files.

The governing rule is **complete teaching coverage with selective representation**. Selectivity decides whether an idea is best expressed as prose, Markdown structure, math, or a source visual; it never licenses deleting instruction.

## Failure modes to avoid

| Failed approach | Why it fails | Required correction |
|---|---|---|
| Treating the assignment as summarization | Explanations, motivations, examples, and caveats disappear even though headings remain | Restore the missing instruction and use concise editing only for delivery scaffolding and exact repetition |
| Defending a large word-count gap with slide titles or headings | Structural labels do not account for missing narrative explanation | Compare narrative speaker-note prose with narrative lecture-note prose and audit every material unexplained gap |
| Assuming prose makes a diagram redundant because it names the boxes | Ordering, direction, grouping, arrows, highlights, or attachment may still be part of the lesson | Keep the source visual when the relationship is taught spatially |
| Keeping only a small hand-picked visual set | Visually rich lectures become under-illustrated | Review every slide and every build family in a visual coverage ledger; asset count is never the criterion |
| Always keeping only the final build state | Earlier cases, errors, before/after states, or contrasts can be distinct teaching points | Keep the final state plus every intermediate state that supports a distinct claim or is referenced by the prose |
| Recreating a source diagram in HTML, CSS, SVG, or a plotting library | The result changes source geometry, typography, and emphasis | Extract the actual PowerPoint objects or source raster; HTML only lays out and embeds them |
| Rendering through a substitute application after its fonts or layout diverge | Text wraps, font metrics, math, and object placement no longer match PowerPoint | Export isolated scratch selections through Microsoft PowerPoint when application fidelity matters |
| Accepting a black, white, or slide-colored rectangle around native shapes | The slide canvas has leaked into the asset | Re-export the selected objects on paired black/white mattes and reconstruct true exterior alpha |
| Globally deleting black or white pixels | Legitimate text, equations, screenshots, and interface panels are destroyed | Remove only the exterior slide canvas; preserve meaningful internal fills and edge-connected source content |
| Calling a PNG “transparent” without compositing it | White or dark labels can disappear on a page of the same color | Inspect explicit light and dark composites at the delivered display width and preserve a local backing or minimal source-derived underlay where needed |
| Letting every image expand to the page width | Small source objects become visually misleading and blurry | Store the source-relative PowerPoint width and render responsively at that width |
| Linking external images from an HTML deliverable | The file is not actually portable or standalone | Embed the exact final asset bytes and verify the embedded count and payloads |

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

The transformation should stay close in narrative length and complete in concept coverage. Compression comes only from removing performance language, quiz mechanics, and exact repetition—not from deleting explanation. Never discard a rationale, analogy, worked example, counterexample, interpretation, limitation, caveat, consequence, or transition in reasoning merely because a shorter statement remains true.

## Audit content coverage

Keep source-slide provenance until the prose is complete, then perform a source-to-notes audit before extracting final visuals. Use paragraph-level provenance or a scratch ledger such as:

| Source slide(s) | Teaching unit | Destination in notes | Disposition | Reason |
|---|---|---|---|---|
| Real deck numbers | Definition, explanation, example, derivation, caveat, or conclusion | Final concept heading | `retain`, `merge`, or `omit` | Exact destination, duplicate passage, or narrow non-instructional omission |

Apply these rules:

- `retain` is the default for substantive instruction. Rephrase it for written reading, but preserve its full teaching function.
- `merge` is valid only when the destination preserves every distinct idea from the merged passages. Repeated wording is not automatically repeated teaching: a later repetition may add a contrast, consequence, example, or synthesis.
- `omit` is limited to greetings, timing remarks, navigation, false starts, filler, quiz mechanics, rhetorical prompts, and exact repetition with no new teaching function. Record the reason; “too detailed,” “already long,” or “summary” is not sufficient.
- Slide titles, repeated headings, page numbers, and temporary preview captions are structural text. Exclude them when comparing narrative length, but never use them to explain away missing speaker-note prose.
- Compare narrative word counts after the first prose pass. Word-count parity does not prove coverage, but a material unexplained decrease is an audit trigger. Trace the removed passages and classify them; restore anything that does not fit an allowed omission or faithful merge.
- Read the final notes against the source in order. The source-to-notes audit is complete only when every instructional passage has a destination or a recorded narrow omission.

Do not start from a target length. The source determines how long a stand-alone lecture replacement needs to be.

## Organize the Markdown

- Replace the `# Speaker Notes: ...` source title with the lecture title as the single `#` heading.
- Replace the temporary `## Slide N: ...` structure with concept-based `##` and, when useful, `###` headings. Do not expose slide numbers as learner-facing structure.
- Open with a short orienting paragraph only when the notes provide that orientation.
- Add `## Table of Contents` immediately after the title's opening paragraph and before the first content section. Use a nested Markdown link list: include every learner-facing `##` section in document order and indent each `###` subsection beneath its parent. Link to the renderer's GitHub-style heading fragment, preserve the visible heading text exactly, and exclude the title and the Table of Contents heading itself. Rebuild the list after any heading edit; do not leave stale, missing, or reordered anchors.
- Keep paragraphs cohesive and moderately short. Let bullets be part of the written explanation, not visual assets.
- Write inline math as `$...$` and display math as `$$...$$`. Preserve variable names, subscripts, hats, Greek symbols, and matrix dimensions exactly.
- Introduce each selected visual in the prose and place it immediately after the passage it supports. Do not append an unexplained image gallery.
- Give every image concise, meaningful alt text. The surrounding prose must still communicate the instructional takeaway for a learner who cannot see the image.
- For a complex visual whose important relationships are not already conveyed in the prose, add a short Markdown paragraph beginning `**Figure description.**`.
- End with a takeaway only when the speaker notes contain a genuine synthesis; do not add a generic recap to every lecture.

## Select visuals from the edited prose

Work outward from the edited prose, not inward from every slide. At each section—or for a finer-grained teaching claim—ask whether a spatial or visual relationship materially improves understanding.

Choose the representation before judging redundancy:

- Use native Markdown for genuine prose, real bullet or numbered lists, real tables, fenced code, and equations that can be represented faithfully in LaTeX math.
- Use an extracted PowerPoint visual for instructional photographs, screenshots, application or model outputs, charts and plots, timelines, architectures, pipelines, arrows, spatial groupings, annotated examples, meaningful highlight states, and other compositions whose meaning depends on placement or appearance.
- Keep source text inside an extracted visual when it labels, points to, annotates, or is geometrically attached to that visual. Move independent paragraphs, ordinary bullets, and free-standing definitions into Markdown.
- Do not redraw a retained source visual in HTML, CSS, SVG, Python, or another diagram system merely because recreation seems easier. The task is to preserve the lecture's visual teaching, not reinterpret its design.

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
- exact visual duplicates and visuals with no instructional spatial relationship.

There is no one-image-per-slide or one-image-per-section quota. Some slides contribute only prose. Some multi-slide builds contribute one final diagram. A dense teaching claim may justify more than one visual when each shows a genuinely different case or result. A low asset count is a warning only when the ledger shows unexplained visual teaching—not a goal to optimize.

Treat adjacent or repeated diagrams as a build family:

| Change between states | Default selection |
|---|---|
| Additive progressive reveal of the same idea | Keep the final complete state only. |
| Highlight, pointer, or emphasis change only | Keep the clean or final state unless the emphasis performs a teaching function in the prose. |
| Different notation, model, case, result, or counterexample | Keep each state that supports a distinct teaching point. |
| A later state removes or replaces an earlier component to teach a contrast | Keep both states when the comparison matters. |
| An earlier state is explicitly interpreted or referenced in the speaker notes | Keep that state even when a later final state exists. |
| Exact or near-exact repetition across slides | Keep one copy and link it once at the best explanatory point. |

Similarity alone is not enough to discard a diagram, but slide count is never a reason to keep one.

## Audit visual coverage before extraction

After the prose pass and initial visual selection, create a scratch-only visual coverage ledger. This is the completeness check that prevents selective extraction from becoming accidental under-illustration. It is not part of the final Markdown. Review the complete slide sequence; do not build the ledger only from visuals noticed during prose editing.

Review every temporary slide preview and account for:

- each substantive teaching claim whose source contains a photograph, screenshot, interface state, chart, plot, timeline, process, architecture, model, spatial relationship, annotated result, worked visual example, or application image; and
- each multi-slide visual build family, using its final complete state plus any genuinely distinct contrast states.

Use one ledger row per claim or visual family:

| Prose section / teaching claim | Source slide(s) or build family | Candidate visual | Disposition | Reason / QA result |
|---|---|---|---|---|
| Concept heading and the relationship being taught | Real deck numbers | Diagram, image, or none | `keep`, `omit`, or `retry` | Instructional role, redundancy reason, extraction issue, or final asset name |

Apply these rules:

- `keep` means the visual contributes a relationship or concrete example that prose alone does not communicate as efficiently. Record its semantic filename and intended Markdown placement after extraction.
- `omit` requires a specific reason such as decorative, quiz-only, ordinary text moved into prose, exact duplicate, or genuinely redundant spatial content. “There are already enough images” is never a reason.
- `retry` means the visual is instructionally useful but the first extraction is clipped, illegible, opaque when it should be transparent, or otherwise unfaithful. It remains unresolved until repaired or changed to a justified `omit` after reasonable extraction attempts.
- For a diagram, pipeline, architecture, process, progressive state, before/after comparison, task head, or model output represented in text, an `omit` reason must state `no instructional spatial relationship` and explain why. Merely saying prose, a list, or a table contains its labels is not enough.
- Rendering or extraction failure is never an `omit` reason. Resolve it through the fidelity escalation, or leave the row as `retry` and report the delivery blocker.
- A prose sentence that merely names the boxes in a pipeline or the stages in a model does not automatically make the source diagram redundant. Preserve a useful visual when the ordering, connection, grouping, direction, highlighting, or label attachment is part of the lesson.
- A screenshot, photograph, chart, or model output discussed as evidence is not decorative merely because its general topic can be described in prose. Keep the concrete source example when the lecture asks the learner to inspect it.
- When a build family contains different inputs, predictions, errors, corrections, or before/after results, treat those states as separate candidates rather than as progressive duplicates.
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
5. keep it as `retry` in the coverage ledger until it is faithful; do not convert a rendering failure into an omission.

### Rendering-fidelity escalation

The extraction method is not successful merely because it produced a PNG. Compare the rendered selection with the PowerPoint slide for font family and weight, line breaks, symbol coverage, object geometry, z-order, crop, and effects.

If LibreOffice or another substitute renderer changes typography, equations, spacing, or geometry, do not compensate by retyping labels or rebuilding the diagram in HTML/Pillow. Use Microsoft PowerPoint as the renderer when it is available:

1. create scratch copies of the deck; never modify the source deck;
2. isolate the exact object IDs for each retained visual, preserving groups and z-order;
3. cover the inherited slide canvas with black in one scratch deck and white in the other;
4. export both scratch decks to PDF through Microsoft PowerPoint;
5. rasterize matching pages at high resolution, recover alpha from the paired mattes, trim to the artwork, and retain transparent padding; and
6. inspect a contact sheet of every output on light, dark, and checkerboard backgrounds.

This workflow preserves the source font renderer and native PowerPoint layout while still removing the slide canvas. Confirm the exported PDF's embedded font names when a prior attempt showed font substitution. If PowerPoint is unavailable and the substitute render is visibly wrong, keep the ledger row as `retry` and report the fidelity limit rather than delivering a redraw as if it were the source.

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

### Preserve the source display size

Do not default every image to `width: 100%`. Preserve how large the selected object appeared on the PowerPoint canvas:

1. calculate the union width of the selected PowerPoint objects and divide it by the slide width; or, for a 300-DPI selection render, divide the non-transparent pixel width by 300 to recover inches and then divide by the slide width;
2. record that percentage with the asset provenance; and
3. use a responsive raw HTML tag in the Markdown when ordinary image syntax would expand the asset:

```html
<img src="lecture-notes-assets/model-flow.png"
     alt="Model flow from token inputs through the network to the output distribution"
     width="63.50%"
     style="display: block; width: 63.50%; max-width: 100%; height: auto; margin: 1.6rem auto;" />
```

The percentage comes from the source, not visual guesswork. `max-width: 100%` allows narrow screens to shrink the image without enlarging it beyond its intended page footprint. If the source object itself was too small to read, make a tighter faithful extraction; do not solve blur by arbitrarily scaling a loose crop.

True alpha is required for diagrams assembled from native shapes, arrows, labels, and callouts. A fully opaque canvas around a native-shape selection is an extraction failure: retry the shape selection or matte render rather than accepting a black, white, or slide-colored rectangle. Inspect edge antialiasing and all labels on both light and dark backgrounds.

Judge contrast from explicit light- and dark-background composites at the delivered display width, not from the transparent PNG alone: image viewers may display transparency as black, white, or a checkerboard. Distinguish the transparent exterior from retained diagram fills. White text inside an opaque blue box, for example, keeps the same contrast on either page background and must not be rejected merely because the asset has alpha. If a label instead floats directly over transparency, or sits on a semitransparent source panel whose contrast changes with the page, first retain its source backing shape when one exists. If no local backing exists because the label relied on the slide canvas, add only the minimum source-derived edge underlay needed for light/dark legibility, keep the original PowerPoint artwork above it, and preserve transparent exterior pixels. Omit the visual only when a faithful, legible composite still cannot be produced.

An embedded screenshot, photograph, UI panel, clock face, or source image whose internal background is intrinsic to the visual may remain internally opaque. Crop it tightly, keep slide chrome out, and record the reason in the coverage ledger. It may still have a transparent exterior and therefore need no opaque allowlist. Use `--allow-opaque` only when the delivered raster has no transparency at all. For other extracted rasters, transparency remains a preference rather than permission to erase intended white marks, formulas, labels, or non-flat backgrounds. The helper removes only flat, edge-connected background pixels; it does not globally key out every white or black pixel or flatten an existing alpha channel.

## Finalize the learner-facing Markdown

The edited draft may retain temporary slide previews and hidden `lecture-source-slide` or `source-slides` markers for verification. Once the prose is concept-organized and final assets are placed, remove those source-only elements deterministically:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/finalize_lecture_notes.py \
  <scratch-dir>/lecture-notes.draft.md \
  --output lecture-notes.md
```

The finalizer removes marked preview blocks and owned provenance comments. It refuses to write a file if `# Speaker Notes:` or `## Slide N: ...` remains, because deleting the number alone would hide an unedited slide-by-slide structure rather than fix it.

## Optional standalone HTML

When the user requests HTML, keep `lecture-notes.md` as the canonical editable source and generate `lecture-notes.html` as a companion. The HTML is a document rendering, not a second authoring pass: do not shorten prose, reorder content, replace equations, or synthesize new visuals during conversion.

With Pandoc available, a suitable baseline is:

```bash
pandoc lecture-notes.md \
  --from markdown+raw_html+tex_math_dollars \
  --to html5 --standalone --embed-resources --mathml \
  --resource-path="$(dirname lecture-notes.md)" \
  --output lecture-notes.html
```

Add a local responsive stylesheet when presentation quality matters, but preserve the raw image widths from the Markdown. Do not draw substitute charts, arrows, boxes, timelines, or diagrams in HTML.

Verify the final HTML mechanically:

- every Markdown image has one embedded `data:image/...;base64,...` payload in the same document order;
- each decoded payload is byte-identical to its final file in `lecture-notes-assets/`;
- no image source still depends on `lecture-notes-assets/`, an absolute path, a remote URL, or a temporary preview;
- every source-relative width and responsive `max-width`/`height` rule survives conversion;
- headings, Table of Contents links, math, tables, code, alt text, and figure descriptions render correctly; and
- the file opens with networking disabled and without the assets folder beside it.

## Quality gate

Before delivery:

1. Review the source-to-notes content audit. Every substantive speaker-note passage must have a destination or a narrow recorded omission; investigate every material unexplained narrative word-count gap.
2. Confirm the concept order still follows the speaker notes unless a small reordering clearly improves written coherence without changing the teaching logic.
3. Read for missing teaching function, not keyword presence. Definitions, motivations, reasoning steps, examples, counterexamples, interpretations, caveats, limitations, and conclusions must survive even when nearby wording was merged.
4. Remove only lecturer performance language, in-lecture quiz mechanics, unresolved slide references, transcription artifacts, exact repetition, and unsupported additions.
5. Verify every equation, symbol, dimension, example, contrast, and numerical claim against the slide and speaker notes.
6. Review the completed visual coverage ledger. Account for every substantive visual teaching claim and every build family as a selected source visual or a specific justified omission; no row may remain `retry`.
7. Confirm no final asset is a temporary full-slide preview, bullet-list capture, text-only slide, title/divider, decorative slide image, or HTML-created substitute for a source visual.
8. Open every final asset. Inspect source-object completeness, font and geometry fidelity, legibility at the expected display width, transparency on light and dark backgrounds, and whether its callouts are spatially relevant. Confirm that no native-shape diagram has an opaque slide canvas.
9. Confirm every image's displayed width is source-derived and responsive rather than an arbitrary full-page default.
10. Resolve every relative image link and confirm that the Markdown refers only to files that exist in `lecture-notes-assets/` and has concise nonempty alt text.
11. Click or deterministically verify every Table of Contents entry. Confirm it matches one `##` or `###` content heading, uses the correct generated anchor, preserves document order, and nests subsections beneath their parent section.
12. If standalone HTML was requested, verify embedded-image count, byte identity, offline portability, retained sizing, and rendering of math/structure.
13. Read the delivered artifact once as a learner who has neither the video nor the slides. It should be self-contained, semi-formal, fully taught, and recognizably the same lecture—not a shortened substitute for it.

Run the deterministic link and asset check in strict-transparency mode. Omit `--allow-opaque` when every asset has alpha. Repeat it only for deliberately opaque screenshots, photographs, or intrinsic panels recorded in the coverage ledger:

```bash
uv run python <pptx-to-pypptx-dir>/scripts/validate_lecture_notes.py \
  lecture-notes.md --assets-dir lecture-notes-assets \
  --strict-transparency \
  --visual-ledger <scratch-dir>/visual-coverage-ledger.md \
  --allow-opaque application-screenshot.png \
  --allow-opaque source-photograph.jpg
```

The validator rejects temporary `slide-images/slide_N.png` references so source previews cannot leak into the final notes. It also rejects unresolved ledger retries, rendering failures recorded as omissions, diagrammatic text substitutions without the required spatial justification, and disagreement between retained assets and Markdown links. In strict mode it rejects every opaque raster that is not exactly allowlisted, including stale allowlist entries. Resolve every error and warning before delivery. If the only remaining warning is that the document has no image links, the coverage ledger must demonstrate that the source lecture truly has no instructionally useful visual.
