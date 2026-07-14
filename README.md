# FinanceTool

PyQt5 财务文件工具，包含 PDF 合并、Excel 合并、A4 PDF 拆分和报关单提取。

## 安装

```powershell
python -m pip install -r requirements.txt
```

`easyocr` 仅用于扫描版报关单。首次使用前需准备中文简体和英文模型；程序不会在提取过程中自动下载模型。

## 图形界面

```powershell
python main.py
```

在“报关单提取”标签中添加或拖入 PDF，选择是否启用扫描件 OCR，然后点击“提取到Excel”。结果包含“报关单汇总”“商品明细”和“提取结果”三张表；带有 OCR 或字段异常的记录会标记为“需复核”。

## 命令行

```powershell
python customs_declaration_extractor.py "docs\6月报关" -o "报关单提取.xlsx"
```

命令行默认在存在“需复核”记录时返回退出码 1，但仍会写出 Excel。人工核对流程或批量导出可使用：

```powershell
python customs_declaration_extractor.py "docs\6月报关" -o "报关单提取.xlsx" --allow-review
```

禁用 OCR 时，扫描件会明确标记为“需要OCR”，不会静默生成错误数据：

```powershell
python customs_declaration_extractor.py input.pdf --no-ocr -o output.xlsx
```

## 测试

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest -q -m "not ocr"

$env:RUN_OCR_TESTS = "1"
python -m pytest -q -m ocr
Remove-Item Env:RUN_OCR_TESTS
```
