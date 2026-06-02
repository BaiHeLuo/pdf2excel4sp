# PDF Intelligent Parser MVP

这是一个面向船舶修理规格书 PDF 的 Python MVP，提供：

- PDF 页面级文本抽取
- 编号结构保留与修复
- 基础表格转文本
- 章节识别与分块输出
- 结构化 JSON 导出

## 安装

```bash
cd "c:\Users\admin\Desktop\workspace\coding area\pdf_intelligent_parser"
pip install -e .
```

如果需要 OCR 兜底：

```bash
pip install -e .[ocr]
```

## 使用

```bash
pdf-intelligent-parser input.pdf --output output.json
```

也可以直接通过模块运行：

```bash
python -m pdf_intelligent_parser input.pdf --output output.json
```

## 输出内容

输出 JSON 包含：

- `pages`：页级文本、表格、编号保留后的内容
- `sections`：按编号识别出的章节
- `chunks`：适合 LLM 的分块结果，附带上下文提示

## MVP 边界

这个版本优先保证可运行和结构清晰，不追求对所有复杂 PDF 的完美版面还原。对于扫描件、复杂跨页表格和极端混排页面，后续可以继续增强 OCR 和表格布局策略。
