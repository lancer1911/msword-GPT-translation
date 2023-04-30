# msword-GPT-translation

[English](./README.md)

`msword-GPT-translation` 是一个使用 ChatGPT API 或 GPT-4 API 进行翻译的简单工具，可以实现 Microsoft Word 文档（`.docx` 格式）的翻译。它可以保留原始文档中的格式，并允许用户自定义字体。本工具适用于需要将 Word 文档翻译成其他语言的用户。

## 安装

首先，确保您已经安装了 Python 3。然后，在终端中运行以下命令：

```bash
pip install -r requirements.txt
```

这将安装本工具所需的所有依赖。

## 配置

在使用本工具之前，您需要创建一个名为 `settings.cfg` 的配置文件。您可以从 `settings.cfg.template` 文件开始，复制并重命名为 `settings.cfg`。然后，将您的 OpenAI API 密钥添加到配置文件中的相应位置。

```
[openai]
api_key = sk-
```

此外，您可以根据需要自定义其他设置，如源语言、目标语言等。

## 使用

要使用 `msword-GPT-translation`，只需运行以下命令：

```bash
python msword-GPT-translation.py sample.docx
```

其中，`sample.docx` 是您需要翻译的 Word 文档。

如果您没有在命令行中指定输入文件，程序会弹出一个文件选择对话框，让您浏览并选择需要翻译的文件。

翻译完成后，结果将保存在与输入文件同一目录下，文件名为 "原文件名-translated.docx"。

## 注意

由于 OpenAI API 有速率限制，请注意不要过快地连续发起请求。如果遇到速率限制错误，程序会自动等待一段时间后重试。请确保您已经在您的 OpenAI 账户中设置了合适的速率限制。

## 示例

在本仓库中，我们提供了一个名为 `sample.docx` 的样本 Word 文档。您可以使用此文件来测试 `msword-GPT-translation` 工具。
