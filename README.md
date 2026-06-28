# Extract Articles

把包含多篇文章的双语 Word 文档按 `Heading 1` 标题拆分成单篇文章，并分别输出 Word 和 PDF。

## 功能

- 按 Word 中的 `Heading 1` 段落识别每篇文章标题
- 输出目录自动分为 `word/` 和 `pdf/`
- 中文标题使用微软雅黑，英文标题使用 Cooper Black
- 英文正文：Times New Roman，四号，两端对齐，绿色背景
- 中文正文：楷体，四号，单倍行距，蓝色背景
- 页边距：上下左右 1.27cm
- PDF 导出优先使用 LibreOffice；Windows 下可使用本机 Microsoft Word
- 文件名会自动替换不适合路径和 Office 自动化的引号字符

## 使用

安装 uv 后运行：

```powershell
uv run python .\extract_articles.py "输入文档.docx"
```

输出路径必须是文件夹，程序会在其中创建 `word/` 和 `pdf/` 两个子文件夹。

只导出前 1 篇用于预览：

```powershell
uv run python .\extract_articles.py "输入文档.docx" --max-articles 1 -o ".\preview"
```

只导出 Word，不导出 PDF：

```powershell
uv run python .\extract_articles.py "输入文档.docx" --no-pdf
```

## 测试

```powershell
uv run python -c "from test_extract_articles import test_split_and_export_two_articles; test_split_and_export_two_articles(); print('self-check passed')"
```

