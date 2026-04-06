# QQY 翻译器 (SCI 风格界面)

这是一个拥有现代、专业界面的多格式文档翻译工具，基于 Python 开发。它使用免费的翻译 API（Google Translate），支持批量处理 Excel、Word 和 PDF 文档。

## 功能特点
- **多格式支持**: 
  - Excel (`.xlsx`): 保留表格结构，仅翻译文本。
  - Word (`.docx`): 逐段翻译，保留基本格式。
  - PDF (`.pdf`): **智能转换为 Word 后翻译**，最大程度保留排版并生成可编辑文档。
- **SCI 风格界面**: 简洁、专业的 `flatly` 主题。
- **自动检测**: 支持源语言自动检测。
- **多语言支持**: 支持 Google Translate 所有可用语言。
- **进度跟踪**: 实时显示翻译进度和日志。
- **独立运行**: 可打包为无需 Python 环境的 `.exe` 文件。

## 安装依赖

确保已安装 Python 3.8+。在终端中运行以下命令安装所需库：

```bash
pip install -r requirements.txt
```

*注意：PDF 转换功能依赖 `pdf2docx` 库。*

## 运行程序

直接运行 `main.py` 启动程序：

```bash
python main.py
```

## 打包为 EXE

要生成独立的 Windows 可执行文件，请运行以下命令：

```bash
pyinstaller --noconsole --onefile --name "QQY_Translator" main.py
```

打包完成后，可执行文件 `QQY_Translator.exe` 将位于 `dist` 文件夹中。

## 注意事项
- **PDF 翻译**: 目前仅支持**文本型 PDF**（即可以选择复制文字的 PDF）。如果是扫描版 PDF（图片），程序无法提取文本进行翻译。
- **API 限制**: 由于使用免费 API，大量翻译可能会触发速率限制。程序内置了简单的重试和延迟机制。
- **网络需求**: 运行时需要连接互联网。
