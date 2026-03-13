# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

w0rdF0rmat is a Python tool that automates formatting of academic papers in Word (.docx) format. It identifies document structures (titles, abstracts, keywords, numbered/Chinese-numbered sections) and applies formatting rules from YAML templates. The project is bilingual (Chinese + English) in comments, UI, and README. It is currently **Windows-only** due to the `pywin32` dependency.

## Key Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run CLI (processes ./test/test.docx → ./test/output.docx)
python main.py

# Run GUI (PyQt6 interface)
python run_gui.py
```

There is no automated test suite, linter, or CI pipeline. Testing is manual: place a `.docx` in `test/`, run `main.py`, and inspect the output.

## Architecture

### Processing Pipeline

```
Input .docx
    → Document.__init__() parses structure via 3-strategy fallback:
        1. Style-based: reads paragraph.style.name (e.g. "Title", "Heading 1", "标题")
        2. Traditional: keyword matching ("摘要", "Abstract", numbered headings)
        3. AI-assisted: sends text to OpenAI API (only if enabled in config)
    → Produces internal model: title, abstract, keywords, sections dict
    → WordFormatter.format() applies FormatSpec rules from YAML template
    → Document.save() writes output .docx
```

### Core Modules

- **`src/core/document.py`** — `Document` class wraps `python-docx`, owns the 3-strategy parsing pipeline. Stores parsed structure as `self.title`, `self.abstract`, `self.keywords`, `self.sections` (dict of section name → list of paragraphs).
- **`src/core/formatter.py`** — `WordFormatter` takes a `Document` + `ConfigManager`, loads the format template, applies font/size/alignment/spacing/margins to each element type.
- **`src/core/format_spec.py`** — Defines formatting specification data structures (the schema between YAML templates and the formatter).
- **`src/core/format_validator.py`** — Validates format specifications before applying.
- **`src/core/ai_assistant.py`** — `DocumentAI` class, optional OpenAI integration for ambiguous structures. Requires API key via `.env`.

### Configuration

- **`src/config/config.yaml`** — Global settings: AI toggle, model selection, template path.
- **`src/core/presets/default.yaml`** — Default formatting template defining per-element styles.
- **`src/config/config_manager.py`** — `ConfigManager` merges global config (`config.yaml`) with user config (`~/.w0rdF0rmat/config.json`). Also handles per-project `format_template.json`.

### GUI (`src/gui/`)

- `main_window.py` — PyQt6 `QMainWindow`, hosts three pages.
- `pages/document_page.py` — File selection/loading.
- `pages/format_page.py` — Template selection and customization.
- `pages/preview_page.py` — Rendered preview (uses PyMuPDF/WebEngine).
- `run_gui.py` — Entry point with 3-level icon fallback.

### Section Detection Heuristics

The traditional parser recognizes sections by:
- Arabic numbered headings (`1.`, `2.`, etc.)
- Chinese numbered headings (`一、`, `二、`, etc.)
- Chinese academic keywords (`引言`, `研究方法`, `实验`, `结果`, `讨论`, `结论`, `参考文献`)

## Dependencies

Core: `python-docx`, `openai`, `python-dotenv`, `pyyaml`. GUI adds: `PyQt6`, `PyQt6-WebEngine`, `Pillow`, `PyMuPDF`, `pywin32`. Python ≥ 3.8 required.
