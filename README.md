# docx2images

Word 文档 → 多张手机端图片

一键将 `.docx` 文档转换为竖版多图，适配 **微头条 / 头条号 / 小红书** 发布。

## 使用方法

### 1. 安装依赖

```bash
pip3 install python-docx Pillow
```

### 2. 转换文档

```bash
cd skills/docx2images
python3 scripts/word2images.py 你的文档.docx
```

即可在 `output_images/` 目录下得到 `你的文档_p01.png`、`你的文档_p02.png` ... 按页码顺序排列，直接上传发布。

### 高级选项

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-o` / `--output-dir` | 输出目录 | 脚本同级 `output_images/` |
| `-f` / `--format` | 图片格式 (`png` / `jpg`) | `png` |

示例：
```bash
python3 scripts/word2images.py 文章.docx -o ./发布用 -f jpg
```

## 图片规格

| 项目 | 参数 |
|------|------|
| 尺寸 | 1242×2208 (竖版) |
| 底色 | 暖白仿纸 (247, 243, 236) |
| 正文字号 | 34px |
| 标题层级 | H1 / H2 / H3 装饰区分 |
| 列表 | 有序 / 无序自动识别 |
| 分页 | 智能断页，不截断句子 |
| 页码 | 底部居中 (第 x/总 y 页) |

## 局限性

- 仅支持 `.docx`
- 仅提取纯文本，不处理 **表格 / 图片 / 文本框**
- 文档排版复杂的，建议先在 Word 中清理格式再转换

## 打包文件

`docx2images.skill` 可分享给他人直接安装使用。
