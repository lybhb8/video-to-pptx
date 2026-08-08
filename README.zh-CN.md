# video-to-pptx

[English](README.md) | [中文](README.zh-CN.md)

将视频课件/讲座转换为 `.pptx` 幻灯片。

```
视频 → ffmpeg 提取帧 → 感知哈希去重 → PaddleOCR → python-pptx
```

## 安装

```bash
pip install python-pptx opencv-python pillow numpy paddlepaddle paddleocr
# 或备选
pip install easyocr
```

OCR 层会自动检测已安装的 PaddleOCR 主版本（2.x `ocr()` 或 3.x `predict()`），两个版本都能正常工作。

另外需要安装 ffmpeg：
- Windows: `winget install Gyan.FFmpeg`
- macOS: `brew install ffmpeg`
- Linux: `sudo apt install ffmpeg`

## 用法

```bash
python scripts/video_to_pptx.py input.mp4 output.pptx --interval 10 --similarity 0.92
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--interval` | 10 | 帧提取间隔（秒） |
| `--sample-similarity` | 0.5 | 采样门控：若当前帧与上一帧哈希相似度超过此值则跳过（0 禁用） |
| `--content-select` | on | 按底层/基础页面对帧进行分组，每页保留文字最完整 / 图片最干净的一帧（见下方说明） |
| `--base-sim` | 0.96 | 像素相似度阈值，超过此值的两帧被视为同一底层页面 |
| `--overlay-thresh` | 0.08 | 像素差异（以 255 为单位的比例）超过此值时，该单元格视为被遮挡/颜色变化 |
| `--image-weight` | 0.5 | 选择代表帧时，图片完整度（无遮挡/颜色变化）在评分中的权重，与文字完整度共同决定 |
| `--region-grid` | 8 | （仅 `--no-content-select` 模式）将帧划分为 grid×grid 区域用于相邻帧去重 |
| `--region-similarity` | 0.5 | （仅 `--no-content-select` 模式）单区域哈希相似度阈值，超过即视为"相同" |
| `--region-ratio` | 0.8 | （仅 `--no-content-select` 模式）当帧中超过此比例的区域与前帧匹配时，丢弃该帧（0 禁用） |
| `--similarity` | 0.92 | （仅 `--no-content-select` 模式）全局感知哈希阈值；调高以保留更少的相似帧 |
| `--ocr` | paddleocr | `paddleocr` 或 `easyocr` |
| `--keep-frames` | off | 保留提取的帧图片 |
| `--crop-top/bottom/left/right` | 0 | 图片去重时忽略的画面边缘比例（去除叠加层：标题栏、任务栏、播放器控件） |
| `--text-similarity` | 0（关闭） | 合并 OCR 文本相似度 ≥ 此值的幻灯片；"主要内容相同 → 视为重复" |
| `--min-text-height` | 0.02 | 丢弃高度小于画面比例此值的 OCR 文本行（去除菜单/标题栏/水印文字）；0 保留全部 |
| `--text-top` / `--text-bottom` | 0 | 保留 OCR 文本时忽略画面顶部/底部的此比例高度（去除工具栏/任务栏/水印叠加层） |

## 内容选择原理（默认模式）

1. 每隔 `--interval` 秒提取一帧；若当前帧与上一帧哈希相似度超过
   `--sample-similarity`，则跳过该帧。
2. 对剩余帧进行 OCR。
3. 将底层页面相同的帧分组：两帧完整画面像素相似度 ≥ `--base-sim` 时
   归为同一组。
4. 每组保留文字最完整*或*图片最完整（遮挡/颜色变化最少）的那一帧，
   评分公式为 `文字长度/最大文字长度 + --image-weight * (1 - 偏离度)`。

这样每个独立画面只保留一个干净的代表帧，丢弃中间状态、弹窗和高亮/选中帧。
传入 `--no-content-select` 可回退到旧的区域 + 全局感知哈希去重模式。

## 调参指南

- 页面太多 → 调高 `--base-sim`（如 0.97）；页面太少 → 调低（如 0.94）
- 代表帧有遮挡/颜色变化 → 调高 `--image-weight` 或调低 `--overlay-thresh`
- 漏掉幻灯片 → 调低 `--similarity` / `--base-sim` 或减少 `--interval`
- 动画过渡/镜头移动 → 先用 `--interval 5` 预览
- 录屏有固定叠加层（任务栏/标题栏/水印）→ 传入 `--crop-top` / `--crop-bottom` 在判断重复时忽略这些区域
- 幻灯片仅因光标移动或微小 UI 变化而不同 → 启用 `--text-similarity 0.9` 合并主要内容（OCR 文字）相同的页面

## 对已有 pptx 去重（无需重新 OCR）

如果已经生成了幻灯片，想删除主要内容相同的页面（忽略标题栏/任务栏/水印等叠加层），可以直接对 pptx 文件运行后处理器——它会使用感知哈希比较每张幻灯片的内嵌帧图片，通过 `--crop-*` 忽略叠加区域：

```bash
python scripts/dedup_pptx.py input.pptx output.pptx \
    --image-similarity 0.95 --text-similarity 0.5 \
    --crop-top 0.15 --crop-bottom 0.08 --crop-right 0.10 \
    --region-top 0.15 --region-bottom 0.92 --region-right 0.90
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--image-similarity` | 0.95 | 当幻灯片帧图片哈希与上一张保留的幻灯片相似度 ≥ 此值时丢弃（0 禁用） |
| `--crop-top/bottom/left/right` | 0 | 比较图片前忽略这些画面边缘比例（标题栏/任务栏/侧边栏） |
| `--text-similarity` | 0.5 | 当幻灯片大部分 OCR 文字与上一张保留的幻灯片相同时丢弃（≥ 此值）；0 禁用 OCR 环节 |
| `--region-top/bottom/right` | 0.15/0.92/0.90 | 忽略主要内容区域之外的 OCR 文字（工具栏/任务栏/侧边栏） |
| `--min-text-height` | 0.01 | 忽略高度小于画面比例此值的 OCR 文本行 |
| `--min-conf` | 0.3 | 忽略置信度低于此值的 OCR 文本行 |

一张幻灯片在满足以下*任一*条件时被丢弃：帧图片哈希 ≥ `--image-similarity`（忽略叠加区域后为同一画面），*或者*主要内容文字大部分 ≥ `--text-similarity`（相同内容，显示略有不同）。OCR 环节需要 `easyocr`。

## 安装为 Claude Code 技能

```bash
mkdir -p ~/.claude/skills ~/.claude/scripts
cp skills/video-to-pptx.md ~/.claude/skills/
cp scripts/video_to_pptx.py ~/.claude/scripts/
```
