# video-to-pptx

[English](README.md) | [中文](README.zh-CN.md)

Convert video courseware/lectures into `.pptx` slides.

```
video → ffmpeg extract frames → perceptual-hash dedup → PaddleOCR → python-pptx
```

## Install

```bash
pip install python-pptx opencv-python pillow numpy paddlepaddle paddleocr
# or fallback
pip install easyocr
```

The OCR layer auto-detects the installed PaddleOCR major version (2.x `ocr()` vs 3.x `predict()`), so either works.

Also install ffmpeg:
- Windows: `winget install Gyan.FFmpeg`
- macOS: `brew install ffmpeg`
- Linux: `sudo apt install ffmpeg`

## Usage

```bash
python scripts/video_to_pptx.py input.mp4 output.pptx --interval 10 --similarity 0.92
```

| flag | default | note |
|------|---------|------|
| `--interval` | 10 | seconds between frame captures |
| `--sample-similarity` | 0.5 | capture gate: skip a frame whose hash similarity to the previous captured frame is above this (0 disables) |
| `--content-select` | on | group frames by their underlying/base page, keep the most text-complete / image-complete frame per page (see below) |
| `--base-sim` | 0.96 | pixel similarity above which two frames count as the same underlying page |
| `--overlay-thresh` | 0.08 | pixel difference (fraction of 255) beyond which a cell counts as overlaid/color-changed vs the group base |
| `--image-weight` | 0.5 | how strongly image completeness (no overlay/color-change) counts when picking the representative, besides text completeness |
| `--region-grid` | 8 | (only with `--no-content-select`) divide frames into grid x grid regions for adjacent-frame dedup |
| `--region-similarity` | 0.5 | (only with `--no-content-select`) per-region hash similarity above which a region counts as "the same" |
| `--region-ratio` | 0.8 | (only with `--no-content-select`) drop a frame when more than this fraction of its regions match the previous kept frame (0 disables) |
| `--similarity` | 0.92 | (only with `--no-content-select`) global perceptual hash threshold; raise to keep fewer similar frames |
| `--ocr` | paddleocr | `paddleocr` or `easyocr` |
| `--keep-frames` | off | keep extracted frame images |
| `--crop-top/bottom/left/right` | 0 | fractions of the frame to ignore during image dedup (overlay removal: title bars, taskbar, player controls) |
| `--text-similarity` | 0 (off) | merge slides whose OCR text similarity ≥ this; "same main content => duplicate" |
| `--min-text-height` | 0.02 | drop OCR lines smaller than this fraction of frame height (removes menus / title bars / watermark text); 0 keeps all |
| `--text-top` / `--text-bottom` | 0 | ignore these fractions of frame height at top/bottom when keeping OCR text (ribbon / taskbar / watermark overlay removal) |

## How content selection works (default)

1. Extract a frame every `--interval` seconds; skip frames whose hash similarity
   to the previous captured frame is above `--sample-similarity`.
2. OCR every remaining frame.
3. Group frames whose underlying page is the same: two frames are in the same
   group when their full-frame pixel similarity is ≥ `--base-sim`.
4. Per group, keep the single frame with the most complete text *or* the most
   complete image (least overlaid / least color-changed), scored as
   `text_len / max_text_len + --image-weight * (1 - deviation)`.

This keeps one clean representative per distinct screen and drops intermediate
states, popups, and highlight/selection frames. Pass `--no-content-select` to
fall back to the older region + global perceptual-hash dedup.

## Tuning

- Too many pages → raise `--base-sim` (e.g. 0.97); too few → lower it (0.94)
- Representative looks overlaid/color-changed → raise `--image-weight` or lower `--overlay-thresh`
- Missing slides → lower `--similarity` / `--base-sim` or reduce `--interval`
- Animated transitions / camera movement → use `--interval 5` first to preview
- Screen recordings with fixed overlays (taskbar / title bars / watermark) → pass `--crop-top` / `--crop-bottom` to ignore those strips when judging duplicates
- Slides that differ only by cursor movement or tiny UI changes → enable `--text-similarity 0.9` to merge pages whose main content (OCR text) is the same

## Deduplicate an existing pptx (no re-OCRing)

If you already generated a deck and want to drop slides whose main content is the
same (ignoring overlays such as title bars / taskbars / watermarks), run the
post-processor directly on the pptx — it compares each slide's embedded frame
image with a perceptual hash, using `--crop-*` to ignore overlay strips:

```bash
python scripts/dedup_pptx.py input.pptx output.pptx \
    --image-similarity 0.95 --text-similarity 0.5 \
    --crop-top 0.15 --crop-bottom 0.08 --crop-right 0.10 \
    --region-top 0.15 --region-bottom 0.92 --region-right 0.90
```

| flag | default | note |
|------|---------|------|
| `--image-similarity` | 0.95 | drop a slide when its frame-image hash similarity to the previous kept slide is ≥ this (0 disables) |
| `--crop-top/bottom/left/right` | 0 | ignore these frame-edge fractions before comparing images (title bar / taskbar / side panel) |
| `--text-similarity` | 0.5 | drop a slide when most of its OCR text is the same as the previous kept slide's (≥ this); 0 disables the OCR pass |
| `--region-top/bottom/right` | 0.15/0.92/0.90 | ignore OCR text outside the main content region (ribbon / taskbar / side panel) |
| `--min-text-height` | 0.01 | ignore OCR lines smaller than this fraction of frame height |
| `--min-conf` | 0.3 | ignore OCR lines with confidence below this |

A slide is dropped when *either* the image hash is ≥ `--image-similarity` (same
screen after ignoring overlay strips) *or* most of its main-content text is ≥
`--text-similarity` (same content shown slightly differently). The OCR pass
requires `easyocr`.

## Install as Claude Code skill

```bash
mkdir -p ~/.claude/skills ~/.claude/scripts
cp skills/video-to-pptx.md ~/.claude/skills/
cp scripts/video_to_pptx.py ~/.claude/scripts/
```
