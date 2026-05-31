---
name: docx2images
description: Convert Word (.docx) documents to mobile-optimized multi-page images for 微头条/头条号/小红书 content publishing. Use when user needs to turn a document into a carousel of images for mobile social media platforms, or asks for "doc转图片", "word转图片", "微头条配图" or similar.
---

# docx2images

## 前置条件

- 依赖: `pip3 install python-docx Pillow`
- 运行环境: macOS (字体路径硬编码为 macOS 系统字体)
- 脚本路径: `scripts/word2images.py`

## 用法

```bash
python3 scripts/word2images.py <input.docx>
python3 scripts/word2images.py <input.docx> -o ./my_output -f jpg
```

- `-o` / `--output-dir`: 输出目录 (默认: 脚本同级 `output_images/`)
- `-f` / `--format`: 图片格式 `png` 或 `jpg` (默认: png)

## 输出

- 1242×2208 竖版 (iPhone 6/7/8 Plus 逻辑分辨率, 移动端主流适配)
- 暖白仿纸护眼底色, 大号字体 (34px), 舒适行距
- 标题自动对应 H1/H2/H3 层级, 带装饰条/分割线
- 有序/无序列表自动识别
- 智能分页, 不截断句子, 底部带页码
- 文件命名: `{文档名}_p01.png`, `{文档名}_p02.png` ...

## 局限性

- 仅支持 `.docx` 格式
- 只提取纯文本内容, 不处理表格/图片/文本框
- 复杂排版文档建议在 Word 中清理格式后再转换
