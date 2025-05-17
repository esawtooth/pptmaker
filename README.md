# pptmaker

A tool for converting PowerPoint presentations into a YAML representation.

This project uses `python-pptx` to parse PPTX files and `openai` to generate
textual descriptions for images found within slides. Tables are converted to
Markdown and all text content is preserved. The resulting YAML file contains a
list of slides with their titles, texts, images and tables.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python -m pptmaker input.pptx output.yaml
```

The OpenAI API key must be provided via the `OPENAI_API_KEY` environment
variable if image descriptions are required.
