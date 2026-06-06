---
name: image-ppt-to-editable-ppt
description: Convert image-only PPT/HTML slide visuals into editable PowerPoint decks with source-preserving cutouts, editable text overlays, and module-level animation layers. Use when the user asks to turn picture-based PPT pages, slide screenshots, HTML slide images, or generated 4K presentation images into an editable .pptx; when they ask to cut whole cards/small boxes/modules for PowerPoint animation; or when extracted PPT icons/elements are badly cut.
---

# Image PPT To Editable PPT

## Core Rule

Act on the exact deck/image in front of you. Preserve the source visual system and original assets whenever possible. Do not redesign icons, cards, diagrams, charts, or logos unless the user explicitly asks.

When an element is hard to isolate cleanly, keep the whole visual group as one PNG, erase only the embedded text, then add editable text boxes over it.

When the user asks for animation-friendly chunks such as "each small box", "each card", "module", "whole target", or "do not split it so fine", use **module cutout mode**: keep each visual card/group as one opaque source-derived PNG layer and do not decompose its internal icons, text, and borders.

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

2. **Build the first editable deck**
   - Use full-slide preview images only as reference, not as final editable content.
   - Add editable text boxes for all user-facing text.
   - Extract visual elements as PNGs with transparent backgrounds.
   - Keep coordinates proportional to the source slide: 4K pixels map to 13.333 x 7.5 in at 288 px/in.

3. **Choose extraction strategy per element**
   - **Simple icon / shape:** recut from the 4K source with generous margins, transparentize near-white background, then tighten to the colored/non-background pixels.
   - **Chart / diagram / workflow:** keep as a larger source-derived PNG if splitting would damage lines, gradients, or icons.
   - **Complex group with text:** keep the visual group, erase embedded text with white/transparent patches, then overlay editable text.
   - **Animation module / card:** keep the entire card or small box as one opaque PNG. White-fill the same region on the base slide, then place the module crop back at the exact coordinates as an independent picture layer named `MODULE_Sxx_xx`.
   - **Identity assets:** preserve the user/source-provided asset; do not approximate or redraw.

4. **Module cutout mode**
   - First auto-detect likely card/module boxes from the full-slide image, but treat this only as a draft.
   - If the user wants control or the boxes are ambiguous, generate a browser red-box calibration page. Let the user drag, resize, add, delete, and save boxes to `adjusted_modules.json`.
   - Build the final PPTX from `adjusted_modules.json`: one white-filled base layer per slide plus one independent picture layer per module.
   - Do not overwrite the source PPTX. Auto-suffix the output when the requested output path already exists.
   - Keep `adjusted_modules.json` and the review page if they are useful for later edits; delete transient crop assets and exported preview images after QA.

5. **Repair bad cutouts**
   - Use the annotated browser/PPT location to identify the matching slide and shape id.
   - Inspect the current PNG. If an edge, icon head, star point, circle, or leader line is missing, recut from the 4K source with a wider crop.
   - If a cutout leaves text ghosts, erase a larger text region or use opaque white patches on white panels.
   - If a complete icon still looks wrong in PPT, check placement: the object may be touching a card border rather than being missing.

6. **Generate review HTML**
   - Export every PPT slide to `3840x2160` PNG.
   - Render an HTML review page using each full-slide preview as the background.
   - Overlay transparent hotspots for shapes so the user can annotate bad cutouts without relying on imperfect HTML text layout.
   - Add cache-busting query strings to preview images after every refresh.

7. **Verify before finishing**
   - View the affected slide export, not only the isolated PNG.
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
$skill = "C:\Users\User\.agents\skills\image-ppt-to-editable-ppt"

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
- Do not crop exactly to visible bounds; add generous margins, then tighten after background removal.
- Transparent pixels may still carry RGB remnants; verify with the PPT export, not just isolated image viewers.
- For white PPT panels, opaque white erase patches can be better than transparent erase patches because PowerPoint can expose antialiasing remnants.
- Do not trust browser HTML text rendering as proof of PPT text layout. Use slide preview backgrounds plus hotspots.
- On Windows PowerShell, do not paste bash heredocs like `python - <<'PY'`; use a PowerShell here-string piped into Python.
- Do not leave debug folders such as `s3_star_debug` or temporary screenshots after QA.
