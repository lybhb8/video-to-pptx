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

## Tuning

- Too many duplicate slides → raise `--similarity` to 0.95+
- Missing slides → lower `--similarity` to 0.85–0.88 or reduce `--interval`
- Animated transitions / camera movement → use `--interval 5` first to preview

## Install as Claude Code skill

```bash
mkdir -p ~/.claude/skills ~/.claude/scripts
cp skills/video-to-pptx.md ~/.claude/skills/
cp scripts/video_to_pptx.py ~/.claude/scripts/
```
