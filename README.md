# docforge

Command-line tool that writes Word documents from **your Word templates**, a folder of
**research files**, and a **prompt** describing what you want. The result is a `.docx` that keeps
the template's styles, headers, footers and headings; Claude writes what goes under each heading.

## Install (Windows / PowerShell)

Requires Python 3.11+ (`winget install Python.Python.3.12`).

```powershell
cd docforge
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
$env:ANTHROPIC_API_KEY = "sk-ant-..."      # key from https://platform.claude.com
```

To keep the key between sessions: `setx ANTHROPIC_API_KEY "sk-ant-..."` and open a new terminal.

## Use

```powershell
docforge init C:\Docs\MyReport --sample    # creates the project folder (+ an example to try)
cd C:\Docs\MyReport
# 1. copy your templates (.docx / .dotx) into templates\
# 2. copy research files into research\
# 3. edit prompt.md
docforge inspect                           # check how your templates are read
docforge build --dry-run                   # see what would be sent, without calling the model
docforge build                             # writes output\<template>-<timestamp>.docx
```

Project folder:

```
MyReport\
  templates\   Word templates
  research\    .pdf .docx .xlsx .md .txt .csv .json .html, images (subfolders are fine)
  prompt.md    your instructions
  output\      generated .docx, plus a .json with the model's raw content
```

### The prompt

Write in plain language: which template to use, what the document is about, who reads it, how to
use the research (what matters most, what to ignore), language, tone and length. The template is
picked by `--template`, otherwise by a file name mentioned in the prompt, otherwise the model
chooses among the templates in `templates\`.

### How templates are read

- **Sections** are found from Word heading styles (Heading 1-9, including localized and custom
  styles based on them). Headings are never changed.
- **Text under a heading** is treated as guidance for that section ("Describe the background...").
  It is given to the model and replaced by the generated content. If the text is fixed
  boilerplate (e.g. a legal notice), the model leaves the section as it is.
- **Placeholders** such as `[Project Name]`, `{{date}}`, `<Customer>` or `«Author»` anywhere in
  the document, including headers and footers, are filled in. Ones the model cannot determine
  are left untouched.
- Generated content can be paragraphs, sub-headings, bullet and numbered lists, and tables, using
  the template's own styles (`List Bullet`, `List Number`, the template's table style).
- Where the research does not cover what a section asks for, the text contains a visible
  `[MISSING: ...]` marker, and the gaps are listed in the terminal after the build.

Run `docforge inspect` to see exactly what was detected in a template.

### Useful options for `build`

| Option | Purpose |
|---|---|
| `-t NAME` | use this template |
| `-p FILE` / `-m "text"` | another prompt file / instructions on the command line |
| `-o FILE` | output file |
| `--body-style "Body Text"` | paragraph style for generated text, if the detected one is wrong |
| `--from-json output\x.json` | re-render a previous result (after hand-editing the JSON, or with another `--body-style`) without calling the model |
| `--provider`, `--base-url`, `--model`, `--effort`, `--max-tokens` | model settings, see below |

## Choosing the model / API

By default docforge uses the Claude API (`claude-opus-5-5`, effort `high`) with the key in
`ANTHROPIC_API_KEY`. It can instead use **any OpenAI-compatible API** (OpenAI, Azure OpenAI,
OpenRouter, Mistral, Ollama, LM Studio, vLLM, llama.cpp, ...).

Settings are read from, in order: command-line options, environment variables, then
`docforge.toml` in the project folder (`docforge init` creates one with commented examples).

| `docforge.toml` | Option | Environment variable | Meaning |
|---|---|---|---|
| `provider` | `--provider` | `DOCFORGE_PROVIDER` | `anthropic` or `openai`; defaults to `openai` when a base URL is set |
| `base_url` | `--base-url` | `DOCFORGE_BASE_URL` | API address, e.g. `http://localhost:11434/v1` |
| `model` | `--model` | `DOCFORGE_MODEL` | model name; required for `openai` |
| `api_key_env` | | | name of the environment variable holding the key (default `DOCFORGE_API_KEY`, then `OPENAI_API_KEY`) |
| `effort` | `--effort` | `DOCFORGE_EFFORT` | reasoning effort; only sent to OpenAI-compatible APIs when set |
| `max_tokens` | `--max-tokens` | `DOCFORGE_MAX_TOKENS` | output limit; only sent to OpenAI-compatible APIs when set |
| `json_mode` | `--json-mode` | `DOCFORGE_JSON_MODE` | `auto` (default), `schema`, `object` or `none` |
| `send_images` | `--no-images` | | set `false` for models without vision |

The API key itself is never stored in the file. Local servers that need no key work without one.

Examples:

```powershell
# Ollama on this machine
docforge build --base-url http://localhost:11434/v1 --model qwen3:32b

# OpenAI
$env:DOCFORGE_API_KEY = "sk-..."
docforge build --provider openai --model gpt-5
```

Differences with an OpenAI-compatible API:

- PDFs are converted to text locally (scanned PDFs without a text layer are skipped and reported).
- With `json_mode = "auto"`, docforge asks for schema-constrained JSON and steps down to plain
  JSON mode, then to prompt-only, if the server rejects the request.
- Quality depends on the model. Small local models may ignore guidance text or return incomplete
  sections; check the "Sections the model did not return" line after a build.
- Local servers often default to a small context window (Ollama: a few thousand tokens), which
  silently drops research. Raise it in the server's settings.

## Limitations

- Content inside Word content controls and text boxes is not touched.
- Text before the first heading (title page) is only changed through placeholders.
- All research is sent in one request: PDFs are limited to about 32 MB / 600 pages in total, and
  very large collections can exceed the model's context window. Nothing is silently cut; the
  build stops with an error instead.
- If the template has a table of contents, Word asks to update fields when the document is opened;
  answer Yes.

## Development

```powershell
pip install -e ".[dev]"
pytest
```
