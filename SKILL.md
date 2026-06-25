---
name: image-ppt-to-editable-ppt
description: Use when the user asks to convert image-only PPT pages, HTML slide images, screenshots, or generated 4K slide images into an editable .pptx; asks to cut whole cards, boxes, modules, diagrams, icons, or targets into animation-ready PowerPoint layers; or reports that PPT cutouts are missing, misaligned, too fine-grained, or badly cropped.
---

# Image PPT To Editable PPT

## Core Rule

Act on the exact deck/image in front of you. Preserve the source visual system and original assets whenever possible. Do not redesign icons, cards, diagrams, charts, or logos unless the user explicitly asks.

When an element is hard to isolate cleanly, keep the whole visual group as one PNG, erase only the embedded text, then add editable text boxes over it.

When the user asks for animation-friendly chunks such as "each small box", "each card", "module", "whole target", or "do not split it so fine", use **module cutout mode**: keep each visual card/group as one opaque source-derived PNG layer and do not decompose its internal icons, text, and borders.

When the user explicitly asks to "整张生成", "重新生成整张图片", use `image-generation`, or make a high-fidelity image deck instead of editable layers, use **whole-slide image-generation mode**: generate one complete 4K PNG per slide and insert each image full-page into a new PPTX. Do not patch-edit on top of the old slide, do not cover old content with overlays, and do not overwrite the source PPTX.

## Required Skills And Tools

- Use `presentations:Presentations` for PPTX work and preview/export QA.
- Use `image-generation` for source image extraction, 4K image handling, local repair, and text/image fidelity rules.
- On Windows, use PowerPoint COM for final visual export when available. Always close COM objects.
- Prefer deterministic local repair with Pillow/PowerShell before regenerating a whole image.

## Workflow

1. **Inventory sources**
   - Locate the source PPTX, HTML, screenshots, or 4K generated slide images.
   - Export or identify one 4K source image per slide, ideally `3840x2160` for 16:9.
   - Save originals and work in a run folder under the project output area.

2. **Whole-slide image-generation mode**
   - Trigger this when the user wants complete regenerated slide images, especially for thesis/report decks where visual polish matters more than internal editability.
   - If a content deck and a style/template deck are both provided, export both to slide PNGs. Use the content slide as the semantic reference and the closest template slide as the visual-style reference.
   - Generate each slide as `3840x2160` PNG with `quality high` and opaque background. Put exact titles, Chinese labels, table values, metric numbers, and English tokens in the prompt.
   - Insert each generated PNG as a full-slide picture in a new PPTX with the same 16:9 page size. Keep a manifest recording source slide, template reference, prompt summary, output PNG, and dimensions.
   - QA the final PPTX by exporting it through PowerPoint when available, building a contact sheet, and visually checking text-heavy slides.

3. **Build the first editable deck**
   - Use full-slide preview images only as reference, not as final editable content.
   - Add editable text boxes for all user-facing text.
   - Extract visual elements as PNGs with transparent backgrounds.
   - Keep coordinates proportional to the source slide: 4K pixels map to 13.333 x 7.5 in at 288 px/in.

4. **Choose extraction strategy per element**
   - **Simple icon / shape:** recut from the 4K source with generous margins, transparentize near-white background, then tighten to the colored/non-background pixels.
   - **Chart / diagram / workflow:** keep as a larger source-derived PNG if splitting would damage lines, gradients, or icons.
   - **Complex group with text:** keep the visual group, erase embedded text with white/transparent patches, then overlay editable text.
   - **Animation module / card:** keep the entire card or small box as one opaque PNG. White-fill the same region on the base slide, then place the module crop back at the exact coordinates as an independent picture layer named `MODULE_Sxx_xx`.
   - **Identity assets:** preserve the user/source-provided asset; do not approximate or redraw.

5. **Module cutout mode**
   - First auto-detect likely card/module boxes from the full-slide image, but treat this only as a draft.
   - If the user wants control or the boxes are ambiguous, generate a browser red-box calibration page. Let the user drag, resize, add, delete, and save boxes to `adjusted_modules.json`.
   - The red-box file is valid only for the exact slide images used to create it. If slides were regenerated, trimmed, reordered, replaced, or visually edited, first build/identify a fresh full-image PPT for the current slide set and rerun `module_cutouts.py prepare` or `review` on that PPT. Do not scale or remap an older `adjusted_modules.json` unless the images are pixel-identical except for uniform resolution scaling.
   - Build the final PPTX from `adjusted_modules.json`: one white-filled base layer per slide plus one independent picture layer per module.
   - Do not overwrite the source PPTX. Auto-suffix the output when the requested output path already exists.
   - Keep `adjusted_modules.json` and the review page if they are useful for later edits; delete transient crop assets and exported preview images after QA.

6. **Repair bad cutouts**
   - Use the annotated browser/PPT location to identify the matching slide and shape id.
   - Inspect the current PNG. If an edge, icon head, star point, circle, or leader line is missing, recut from the 4K source with a wider crop.
   - If a cutout leaves text ghosts, erase a larger text region or use opaque white patches on white panels.
   - If a complete icon still looks wrong in PPT, check placement: the object may be touching a card border rather than being missing.

7. **Generate review HTML**
   - Export every PPT slide to `3840x2160` PNG.
   - Render an HTML review page using each full-slide preview as the background.
   - Overlay transparent hotspots for shapes so the user can annotate bad cutouts without relying on imperfect HTML text layout.
   - Add cache-busting query strings to preview images after every refresh.

8. **Verify before finishing**
   - View the affected slide export, not only the isolated PNG.
   - For whole-slide image-generation mode, verify all generated PNGs are `3840x2160`, the PPT slide count matches the generated image count, and the exported PPT contact sheet has no blank or wrong-aspect slides.
   - Check icons are complete, text is not duplicated, source labels are not ghosting, and overlays do not overlap.
   - Verify preview PNG dimensions.
   - For module cutout mode, verify layer counts: one `WHITEFILL_BASE_Sxx` layer per slide and one `MODULE_Sxx_xx` picture per saved module box.
   - Confirm the local review URL returns HTTP 200 if a review server is active.
   - Delete scratch preview/debug folders you created.

## Bundled Scripts

Use these scripts from the skill directory when helpful:

- `scripts/recut_from_source.py`: crop a source slide region, remove near-white background, optionally erase boxes, and tighten around foreground pixels.
- `scripts/replace_picture_shape.ps1`: replace a picture shape in a PPTX while preserving position and z-order, with optional position/size adjustment.
- `scripts/export_review_html.ps1`: export slide previews and create a review HTML page with shape hotspots.
- `scripts/module_cutouts.py`: prepare auto-detected module boxes, serve an adjustable red-box review page, and build a PPTX with white-filled base layers plus independent module picture layers.

## Example Commands

```powershell
$skill = "C:\Users\User\.codex\skills\image-ppt-to-editable-ppt"

# Recut a damaged icon from a 4K slide source.
python "$skill\scripts\recut_from_source.py" `
  --source "E:\path\slide_04_4k.png" `
  --out "E:\path\assets\s4_target.png" `
  --box "360,1720,720,2070" `
  --mode orange `
  --pad 28

# Replace the damaged icon shape in the editable PPT.
PowerShell -ExecutionPolicy Bypass -File "$skill\scripts\replace_picture_shape.ps1" `
  -Pptx "E:\path\deck_editable.pptx" `
  -Slide 4 `
  -ShapeId 38 `
  -Image "E:\path\assets\s4_target.png"

# Refresh the browser-review HTML.
PowerShell -ExecutionPolicy Bypass -File "$skill\scripts\export_review_html.ps1" `
  -Pptx "E:\path\deck_editable.pptx" `
  -OutDir "E:\path\html_review"

# Module cutout mode: make a red-box review page for whole cards/small boxes.
python "$skill\scripts\module_cutouts.py" review `
  --pptx "E:\path\source_image_deck.pptx" `
  --work-dir "E:\path\outputs\module_cutouts" `
  --port 8791

# After the user clicks Save in the review page, build the animation-friendly PPTX.
python "$skill\scripts\module_cutouts.py" build `
  --pptx "E:\path\source_image_deck.pptx" `
  --boxes "E:\path\outputs\module_cutouts\adjusted_modules.json" `
  --out "E:\path\source_image_deck_module_cutouts.pptx" `
  --work-dir "E:\path\outputs\module_cutouts"
```

## Pitfalls From Prior Work

- Do not hand-draw replacement icons when a source icon exists.
- Do not answer a module-animation request by extracting only text, or by splitting each card into tiny icons/text fragments.
- Do not use transparent background removal for full cards unless the user asks; opaque card crops preserve shadows, antialiasing, text, and borders better.
- Do not reuse old module coordinates after whole-slide regeneration, slide deletion, slide reordering, or content/layout changes. Stale boxes create missing modules, wrong crops, and visible white gaps; rerun the red-box review on the current deck.
- Do not crop exactly to visible bounds; add generous margins, then tighten after background removal.
- Transparent pixels may still carry RGB remnants; verify with the PPT export, not just isolated image viewers.
- For white PPT panels, opaque white erase patches can be better than transparent erase patches because PowerPoint can expose antialiasing remnants.
- Do not trust browser HTML text rendering as proof of PPT text layout. Use slide preview backgrounds plus hotspots.
- On Windows PowerShell, do not paste bash heredocs like `python - <<'PY'`; use a PowerShell here-string piped into Python.
- Do not leave debug folders such as `s3_star_debug` or temporary screenshots after QA.
