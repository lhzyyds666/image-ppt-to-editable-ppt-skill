# Image PPT To Editable PPT Skill

Convert image-only PowerPoint decks, HTML slide exports, and 4K slide screenshots into editable PPTX files while preserving the original visual system.

This skill is especially useful when a deck is mostly one flattened image per slide but the user still needs animation-ready objects. It supports:

- source-preserving image cutouts
- editable text overlays when appropriate
- module/card-level cutouts for PowerPoint animations
- a browser-based red-box review page for adjusting module crop regions
- rebuilt PPTX files with one white-filled base layer plus independent module image layers

## Files

- `SKILL.md` - agent instructions and workflows
- `scripts/recut_from_source.py` - recut and repair source-derived element PNGs
- `scripts/replace_picture_shape.ps1` - replace a PPT picture shape while keeping position and z-order
- `scripts/export_review_html.ps1` - export PPT previews and shape hotspot review HTML
- `scripts/module_cutouts.py` - prepare, review, and build module-level cutout PPTX files

## Module Cutout Workflow

```powershell
$skill = "C:\Users\User\.codex\skills\image-ppt-to-editable-ppt"

python "$skill\scripts\module_cutouts.py" review `
  --pptx "E:\path\source_image_deck.pptx" `
  --work-dir "E:\path\outputs\module_cutouts" `
  --port 8791
```

Adjust the red boxes in the browser and click **Save**. Then build the animation-friendly PPTX:

```powershell
python "$skill\scripts\module_cutouts.py" build `
  --pptx "E:\path\source_image_deck.pptx" `
  --boxes "E:\path\outputs\module_cutouts\adjusted_modules.json" `
  --out "E:\path\source_image_deck_module_cutouts.pptx" `
  --work-dir "E:\path\outputs\module_cutouts"
```

The output deck names layers like `WHITEFILL_BASE_S01` and `MODULE_S01_01`, making modules easy to select and animate in PowerPoint.

## Requirements

- Python with `python-pptx` and `Pillow`
- Optional but recommended: `opencv-python` and `numpy` for automatic module box detection
- Windows PowerShell and PowerPoint for final visual export QA
