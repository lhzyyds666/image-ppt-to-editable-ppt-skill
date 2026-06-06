#!/usr/bin/env python3
"""Recut a slide element from a source image and make the background transparent."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Tuple

from PIL import Image


Box = Tuple[int, int, int, int]


def parse_box(value: str) -> Box:
    parts = [int(p.strip()) for p in value.split(",")]
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("box must be x0,y0,x1,y1")
    x0, y0, x1, y1 = parts
    if x1 <= x0 or y1 <= y0:
        raise argparse.ArgumentTypeError("box must have x1>x0 and y1>y0")
    return x0, y0, x1, y1


def color_hit(r: int, g: int, b: int, mode: str) -> bool:
    if mode == "orange":
        return r > 145 and 35 < g < 175 and b < 115 and r > g + 25
    if mode == "navy":
        return b < 120 and g < 135 and r < 140
    if mode == "dark":
        return r < 190 and g < 190 and b < 190
    if mode == "nonwhite":
        return not (r > 232 and g > 220 and b > 205)
    return not (r > 238 and g > 228 and b > 214)


def near_background(r: int, g: int, b: int) -> bool:
    if r > 232 and g > 218 and b > 200:
        return True
    if r > 215 and g > 180 and b > 145 and (r - g) < 85:
        return True
    return False


def erase_boxes(image: Image.Image, boxes: Iterable[Box], fill: Tuple[int, int, int, int]) -> None:
    for box in boxes:
        image.paste(fill, box)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--out", required=True, type=Path)
    parser.add_argument("--box", required=True, type=parse_box, help="Crop box x0,y0,x1,y1 in source pixels")
    parser.add_argument("--mode", default="auto", choices=["auto", "orange", "navy", "dark", "nonwhite"])
    parser.add_argument("--pad", type=int, default=20, help="Padding around detected foreground pixels")
    parser.add_argument("--no-tighten", action="store_true", help="Keep the full crop instead of tightening to foreground")
    parser.add_argument("--erase", action="append", type=parse_box, default=[], help="Erase box inside the crop; repeatable")
    parser.add_argument("--erase-fill", choices=["transparent", "white"], default="transparent")
    args = parser.parse_args()

    source = Image.open(args.source).convert("RGBA")
    crop = source.crop(args.box).convert("RGBA")
    px = crop.load()
    w, h = crop.size
    foreground = []

    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            hit = color_hit(r, g, b, args.mode)
            if hit:
                foreground.append((x, y))
            elif near_background(r, g, b):
                px[x, y] = (255, 255, 255, 0)

    fill = (255, 255, 255, 0) if args.erase_fill == "transparent" else (255, 255, 255, 255)
    erase_boxes(crop, args.erase, fill)

    if foreground and not args.no_tighten:
        xs = [p[0] for p in foreground]
        ys = [p[1] for p in foreground]
        pad = max(args.pad, 0)
        tight = (
            max(min(xs) - pad, 0),
            max(min(ys) - pad, 0),
            min(max(xs) + pad + 1, w),
            min(max(ys) + pad + 1, h),
        )
        crop = crop.crop(tight)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    crop.save(args.out)
    print(f"saved {args.out} {crop.size[0]}x{crop.size[1]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
