# video-to-pptx

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
python scripts/video_to_pptx.py input.mp4 output.pptx --interval 2 --similarity 0.92
```

| flag | default | note |
|------|---------|------|
| `--interval` | 2 | seconds between frame captures |
| `--similarity` | 0.92 | perceptual hash threshold; raise to keep fewer similar frames |
| `--ocr` | paddleocr | `paddleocr` or `easyocr` |
| `--keep-frames` | off | keep extracted frame images |
| `--crop-top/bottom/left/right` | 0 | fractions of the frame to ignore during dedup (overlay removal: title bars, taskbar, player controls) |
| `--text-similarity` | 0 (off) | merge slides whose OCR text similarity ≥ this; "same main content => duplicate" |
| `--min-text-height` | 0.02 | drop OCR lines smaller than this fraction of frame height (removes menus / title bars / watermark text); 0 keeps all |
| `--text-top` / `--text-bottom` | 0 | ignore these fractions of frame height at top/bottom when keeping OCR text (ribbon / taskbar / watermark overlay removal) |

## Tuning

- Too many duplicate slides → raise `--similarity` to 0.95+
- Missing slides → lower `--similarity` to 0.85–0.88 or reduce `--interval`
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
    --image-similarity 0.95 --crop-top 0.15 --crop-bottom 0.08
```

| flag | default | note |
|------|---------|------|
| `--image-similarity` | 0.95 | drop a slide when its frame-image hash similarity to the previous kept slide is ≥ this (0 disables) |
| `--crop-top/bottom/left/right` | 0 | ignore these frame-edge fractions before comparing images |
| `--text-similarity` | 0 (off) | additionally drop a slide when its notes text is ≥ this similar to the previous kept slide |

## Install as Claude Code skill

```bash
mkdir -p ~/.claude/skills ~/.claude/scripts
cp skills/video-to-pptx.md ~/.claude/skills/
cp scripts/video_to_pptx.py ~/.claude/scripts/
```
