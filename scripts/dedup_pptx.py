#!/usr/bin/env python3
"""
dedup_pptx.py
Merge duplicate slides in an existing .pptx deck without touching the source
video. Slides are compared by the perceptual hash of their embedded frame
image; a slide is dropped when its image (optionally cropped to the main
content region) is nearly identical to the previously kept slide's.

Optional text-based pass: a slide may also be dropped when its notes text is
nearly identical to the previous kept slide's (--text-similarity).

Usage:
  python scripts/dedup_pptx.py input.pptx output.pptx \
      --image-similarity 0.95 --crop-top 0.15 --crop-bottom 0.08
"""
import argparse
import hashlib
import re
import sys
from io import BytesIO
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation

from video_to_pptx import hash_similarity, text_similarity


def image_hash_pil(img: Image.Image, size: int = 16, crop=None) -> str:
    """Perceptual hash of a PIL image; `crop` = (top, bottom, left, right)
    fractions trimmed from each edge before hashing (overlay removal)."""
    img = img.convert("L")
    if crop:
        w, h = img.size
        top, bottom, left, right = crop
        img = img.crop((int(w * left), int(h * top),
                        int(w * (1 - right)), int(h * (1 - bottom))))
    img = img.resize((size, size), Image.Resampling.LANCZOS)
    arr = np.array(img, dtype=np.float32)
    bits = (arr > arr.mean()).astype(np.uint8)
    return hashlib.sha1(bits.tobytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description="Deduplicate an existing pptx")
    parser.add_argument("input_pptx", type=Path)
    parser.add_argument("output_pptx", type=Path)
    parser.add_argument("--image-similarity", type=float, default=0.95,
                        help="Drop a slide when its frame-image hash similarity to "
                             "the previous kept slide is >= this (0 disables)")
    parser.add_argument("--crop-top", type=float, default=0.0)
    parser.add_argument("--crop-bottom", type=float, default=0.0)
    parser.add_argument("--crop-left", type=float, default=0.0)
    parser.add_argument("--crop-right", type=float, default=0.0)
    parser.add_argument("--text-similarity", type=float, default=0.0,
                        help="Additionally drop a slide when its notes text "
                             "similarity to the previous kept slide is >= this (0 disables)")
    args = parser.parse_args()

    crop = (args.crop_top, args.crop_bottom, args.crop_left, args.crop_right)
    crop = crop if any(crop) else None

    src = Presentation(str(args.input_pptx))
    total = len(src.slides)

    # Collect each slide's frame image + notes.
    slides = []
    for s in src.slides:
        img = None
        for sh in s.shapes:
            if sh.shape_type == 13:  # PICTURE
                img = Image.open(BytesIO(sh.image.blob)).convert("RGB")
                break
        notes = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        slides.append({"image": img, "notes": notes})

    keep_idx, last_hash, last_notes = [], None, None
    for i, sl in enumerate(slides):
        drop = False
        if args.image_similarity > 0 and sl["image"] is not None and last_hash is not None:
            h = image_hash_pil(sl["image"], crop=crop)
            if hash_similarity(h, last_hash) >= args.image_similarity:
                drop = True
        if not drop and args.text_similarity > 0 and last_notes is not None:
            if text_similarity(sl["notes"], last_notes) >= args.text_similarity:
                drop = True
        if drop:
            continue
        keep_idx.append(i)
        if args.image_similarity > 0 and sl["image"] is not None:
            last_hash = image_hash_pil(sl["image"], crop=crop)
        if args.text_similarity > 0:
            last_notes = sl["notes"]

    out = Presentation()
    out.slide_width = src.slide_width
    out.slide_height = src.slide_height
    blank = out.slide_layouts[6]
    for i in keep_idx:
        s = src.slides[i]
        slide = out.slides.add_slide(blank)
        for sh in s.shapes:
            if sh.shape_type == 13:
                img = sh.image
                slide.shapes.add_picture(BytesIO(img.blob), sh.left, sh.top,
                                         sh.width, sh.height)
                break
        if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip():
            slide.notes_slide.notes_text_frame.text = s.notes_slide.notes_text_frame.text

    out.save(str(args.output_pptx))
    print(f"{total} slides -> {len(keep_idx)} slides "
          f"(image-similarity={args.image_similarity}, crop={crop}, "
          f"text-similarity={args.text_similarity})")
    print(f"Saved {args.output_pptx}")


if __name__ == "__main__":
    main()
