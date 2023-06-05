# ParaCompare

This is a python Code use for two document comparation.

## 要求/Requirements


- Python 3.7+
- 第三方库：os, difflib, docx, PIL, pytesseract

确保在运行代码之前，已经安装了所需的Python版本和第三方库。

## 实现方式/Method used.

该代码实现了一个OCR（光学字符识别）处理程序，用于比较文档之间的差异并标记编辑位置。它使用Python编写，并利用Tesseract OCR引擎进行图像转换。主要步骤包括：

1. 图像OCR转换：将输入的图像文件转换为文本。
2. 读取Word文档：读取指定的Word文档文件，并获取文本内容。
3. 文本预处理：对文本进行预处理，去除空格、替换特殊字符等。
4. 查找编辑位置：比较原始文本和OCR转换后的文本，找到它们之间的差异，并记录编辑位置。
5. 标记差异并保留格式：在Word文档中标记差异的位置，并保留原始文本的格式。
6. 保存修改后的文档：将修改后的Word文档保存到指定路径。

## 使用方法

提供使用代码的指南和示例。说明如何准备输入数据，并描述如何运行代码。例如：

1. 准备输入数据：
   - 将要处理的图像文件放置在指定路径。
   - 将要处理的Word文档文件放置在指定路径。

2. 安装依赖项：
   - 确保已安装Python 3.7+。
   - 安装所需的第三方库：os, difflib, docx, PIL, pytesseract。

3. 运行代码：
   - 在代码中指定图像文件路径、Word文档文件路径和修改后的文档保存路径。
   - 在命令行或终端中运行代码。

```python
python code.py
```

示例：

```python
if __name__ == '__main__':
    main(image_paths=["2.png"], word_document_paths=["2.docx"],
         modified_word_document_path="modified_document.docx")
```

这将处理名为`2.png`的图像文件和名为`2.docx`的Word文档文件，并将修改后的文档保存为`modified_document.docx`。

请确保已按照要求安装依赖项，并根据您的具体情况准备输入数据。运行代码后，您将在指定的路径找到修改后的Word文档。
