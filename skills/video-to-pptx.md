---
name: video-to-pptx
version: 1.0.0
description: Convert video courseware/lectures into a PowerPoint (.pptx) by extracting frames with ffmpeg, OCR with PaddleOCR/easyocr, deduplicating similar frames, and building slides with python-pptx.
trigger: |
  把视频转成 PPT
  视频课件转 pptx
  从视频提取课件幻灯片
---

# 视频课件转 PPTX

## Goal

将视频课件或录播按关键帧拆分为 `.pptx` 幻灯片，每页嵌入关键帧图片，备注区写入识别出的文字。

## Prerequisites

- Python 3.9+
- ffmpeg 已安装并在 PATH 中
  - Windows: `winget install Gyan.FFmpeg`
  - macOS: `brew install ffmpeg`
  - Linux: `sudo apt install ffmpeg`

## Install dependencies

```bash
pip install python-pptx opencv-python pillow numpy paddlepaddle paddleocr
```

如果 PaddleOCR 安装失败，可换用 easyocr：

```bash
pip install easyocr
```

## Workflow

1. 确认用户提供的视频路径和期望输出文件名。
2. 如果 `~/.claude/scripts/video_to_pptx.py` 不存在，按下方“Write the script”创建。
3. 运行脚本：
   ```bash
   python ~/.claude/scripts/video_to_pptx.py input.mp4 output.pptx --interval 2 --similarity 0.92
   ```
4. 报告生成的幻灯片数量，并提示用户检查去重效果。

## Write the script

将以下脚本保存为 `~/.claude/scripts/video_to_pptx.py`：

```python
#!/usr/bin/env python3
"""
video_to_pptx.py
Convert a video lecture/courseware into a .pptx deck.

Steps:
  1. Extract frames at a regular interval with ffmpeg.
  2. OCR each frame (PaddleOCR preferred, fallback to easyocr).
  3. Deduplicate visually similar frames via perceptual hashing.
  4. Build one slide per unique frame, embedding the frame image and
     appending recognized text to the slide notes.

Usage:
  python ~/.claude/scripts/video_to_pptx.py input.mp4 output.pptx \
      --interval 2 --similarity 0.92
"""
import argparse
import hashlib
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.util import Inches


def run(cmd):
    print(f"$ {' '.join(cmd)}")
    subprocess.run(cmd, check=True)


def extract_frames(video_path: Path, out_dir: Path, interval_sec: float):
    out_dir.mkdir(parents=True, exist_ok=True)
    fps = f"1/{interval_sec}"
    cmd = [
        "ffmpeg", "-y", "-i", str(video_path),
        "-vf", f"fps={fps},scale='min(iw,1920)':-1",
        "-q:v", "2",
        str(out_dir / "frame_%06d.jpg"),
    ]
    run(cmd)
    return sorted(out_dir.glob("frame_*.jpg"))


def image_hash(path: Path, size: int = 16) -> str:
    img = Image.open(path).convert("L").resize((size, size), Image.Resampling.LANCZOS)
    arr = np.array(img, dtype=np.float32)
    mean = arr.mean()
    bits = (arr > mean).astype(np.uint8)
    return hashlib.sha1(bits.tobytes()).hexdigest()


def hash_similarity(a: str, b: str) -> float:
    bin_a = bin(int(a, 16))[2:].zfill(160)
    bin_b = bin(int(b, 16))[2:].zfill(160)
    dist = sum(x != y for x, y in zip(bin_a, bin_b))
    return 1 - dist / 160.0


def deduplicate(frame_paths, threshold: float):
    kept, hashes = [], []
    for p in frame_paths:
        h = image_hash(p)
        if any(hash_similarity(h, e) >= threshold for e in hashes):
            continue
        kept.append(p)
        hashes.append(h)
    return kept


def ocr_text(path: Path, engine: str) -> str:
    if engine == "paddleocr":
        from paddleocr import PaddleOCR
        ocr = getattr(ocr_text, "_paddle", None)
        if ocr is None:
            ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
            ocr_text._paddle = ocr
        result = ocr.ocr(str(path), cls=True)
        texts = []
        if result and result[0]:
            for line in result[0]:
                texts.append(line[1][0])
        return "\n".join(texts)

    import easyocr
    reader = getattr(ocr_text, "_easyocr", None)
    if reader is None:
        reader = easyocr.Reader(["ch_sim", "en"])
        ocr_text._easyocr = reader
    return "\n".join(txt for (_, txt, _) in reader.readtext(str(path)))


def build_pptx(frame_paths, texts, output: Path, size=(13.333, 7.5)):
    prs = Presentation()
    prs.slide_width = Inches(size[0])
    prs.slide_height = Inches(size[1])
    blank = prs.slide_layouts[6]

    for frame, text in zip(frame_paths, texts):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(
            str(frame), Inches(0), Inches(0),
            width=prs.slide_width, height=prs.slide_height
        )
        if text.strip():
            slide.notes_slide.notes_text_frame.text = text.strip()

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    print(f"Saved {output} with {len(frame_paths)} slides.")


def main():
    parser = argparse.ArgumentParser(description="Convert video courseware to PPTX")
    parser.add_argument("input_video", type=Path)
    parser.add_argument("output_pptx", type=Path)
    parser.add_argument("--interval", type=float, default=2.0,
                        help="Seconds between extracted frames")
    parser.add_argument("--similarity", type=float, default=0.92,
                        help="Perceptual-hash similarity threshold (0-1)")
    parser.add_argument("--ocr", choices=["paddleocr", "easyocr"], default="paddleocr")
    parser.add_argument("--keep-frames", action="store_true")
    args = parser.parse_args()

    if not args.input_video.exists():
        sys.exit(f"Input not found: {args.input_video}")

    work = Path(tempfile.mkdtemp(prefix="v2pptx_"))
    try:
        frames_dir = work / "frames"
        print(f"Extracting frames every {args.interval}s...")
        frames = extract_frames(args.input_video, frames_dir, args.interval)

        print(f"Deduplicating {len(frames)} frames (threshold={args.similarity})...")
        unique = deduplicate(frames, args.similarity)

        print(f"OCR {len(unique)} unique frames with {args.ocr}...")
        texts = [ocr_text(f, args.ocr) for f in unique]

        build_pptx(unique, texts, args.output_pptx)

        if args.keep_frames:
            keep = args.output_pptx.with_suffix(".frames")
            shutil.copytree(frames_dir, keep, dirs_exist_ok=True)
            print(f"Frames kept at {keep}")
    finally:
        if not args.keep_frames:
            shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
```

## Run

```bash
python ~/.claude/scripts/video_to_pptx.py input.mp4 output.pptx --interval 2 --similarity 0.92
```

参数：
- `input_video`: 输入视频路径
- `output_pptx`: 输出 PPTX 路径
- `--interval`: 每隔多少秒抽一帧（默认 2）
- `--similarity`: 感知哈希相似度阈值（默认 0.92），越高越严格
- `--ocr`: `paddleocr` 或 `easyocr`
- `--keep-frames`: 保留中间抽帧目录

## Tuning notes

- 重复页过多 → 提高 `--similarity`（如 0.95）
- 真实翻页被合并 → 降低 `--similarity`（如 0.88）或减小 `--interval`
- 动画、转场、摄像头人像可能导致误判，可先用 `--interval 5` 粗跑观察

## Troubleshooting

- `ffmpeg: command not found` → 安装 ffmpeg 并加入 PATH。
- PaddleOCR 初始化失败 / CUDA 报错 → 换 `--ocr easyocr` 或安装 CPU 版 paddle。
- 生成 PPTX 为空 → 检查视频是否可读，`--interval` 是否过大。
