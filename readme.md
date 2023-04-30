# msword-GPT-translation

[中文版本](./readme-zh.md)

`msword-GPT-translation` is a simple tool that uses ChatGPT API or GPT-4 API for translating Microsoft Word documents (`.docx` format). It preserves the formatting of the original document and allows users to customize the font. This tool is suitable for users who need to translate Word documents into other languages.

## Installation

First, make sure you have Python 3 installed. Then, run the following command in your terminal:

```bash
pip install -r requirements.txt
```

This will install all the dependencies required for this tool.

## Configuration

Before using the tool, you need to create a configuration file named `settings.cfg`. You can start with the `settings.cfg.template` file, copying and renaming it to `settings.cfg`. Then, add your OpenAI API key to the appropriate place in the configuration file.

```
[openai]
api_key = sk-
```

Additionally, you can customize other settings as needed, such as source language, target language, and more.

## Usage

To use `msword-GPT-translation`, simply run the following command:

```bash
python msword-GPT-translation.py sample.docx
```

Where `sample.docx` is the Word document you need to translate.

If you don't specify an input file in the command line, the program will prompt a file selection dialog, allowing you to browse and choose the file you need to translate.

Once the translation is complete, the result will be saved in the same directory as the input file, with the filename "original_filename-translated.docx".

## Note

As the OpenAI API has rate limits, please be careful not to send requests too quickly in succession. If a rate limit error is encountered, the program will automatically wait for a while and retry. Make sure you have set appropriate rate limits in your OpenAI account.

## Example

In this repository, we provide a sample Word document named `sample.docx`. You can use this file to test the `msword-GPT-translation` tool.