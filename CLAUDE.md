# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Word document formatting tool (w0rdF0rmat) designed to automate the formatting of academic papers. It intelligently identifies document structures (titles, abstracts, keywords, sections, etc.) and applies specified formatting requirements. The tool supports multiple parsing methods (style-based, traditional parsing, and AI-assisted analysis) and provides flexible configuration options.

## Key Commands

### Running the Application

- **Run CLI version**: `python main.py`
  - Processes the test document at `./test/test.docx`
  - Saves output to `./test/output.docx`

- **Run GUI version**: `python run_gui.py`
  - Launches a PyQt6 GUI interface for interactive document formatting
  - Handles icon loading gracefully with fallback mechanisms

### Dependencies

Install dependencies with:
```bash
pip install -r requirements.txt
```

Key dependencies:
- `python-docx` for Word document processing
- `PyQt6` for GUI interface
- `openai` for AI assistance features
- `PyYAML` for configuration files

## Architecture Overview

### Core Components

1. **Document Processing (`src/core/`)**
   - `document.py`: Main document class that parses Word files using multiple strategies
   - `formatter.py`: Applies formatting rules to document elements
   - `format_spec.py`: Defines formatting specifications and rules
   - `ai_assistant.py`: Optional AI-powered document analysis

2. **Configuration System (`src/config/`)**
   - `config_manager.py`: Manages both global config (`~/.w0rdF0rmat/config.json`) and project-specific settings
   - `config.yaml`: Global configuration file for AI and formatting options
   - `presets/default.yaml`: Default formatting template with detailed style definitions

3. **GUI Interface (`src/gui/`)**
   - `main_window.py`: Main application window
   - `pages/`: Contains document, format, and preview pages
   - `components/`: Reusable GUI components like loading indicators

### Document Parsing Strategies

The tool uses a hierarchical approach to document parsing:

1. **Style-based parsing**: First attempts to parse using built-in Word styles
2. **Traditional parsing**: Falls back to keyword-based text analysis
3. **AI-assisted parsing**: Uses AI for complex document structures when enabled

### Formatting System

Formatting specifications are defined in YAML files with support for:
- Text formatting (font, size, bold, italic)
- Paragraph alignment and spacing
- Page margins and layout
- Table formatting with custom styles
- Image positioning and captions
- Table of contents configuration

## Configuration Details

### Global Config Location
- User configurations are stored at `~/.w0rdF0rmat/config.json`
- Project-specific templates are saved as `format_template.json` in the project directory

### Configuration Options
- AI assistant can be enabled/disabled in `config.yaml`
- Templates can be switched between default and custom
- User-defined formats are persisted across sessions

## Key Files and Directories

- `main.py`: CLI entry point that processes test documents
- `run_gui.py`: GUI application launcher
- `src/config/config.yaml`: Global configuration
- `src/core/presets/default.yaml`: Default formatting template
- `test/test.docx`: Sample document for testing
- `src/gui/`: Complete PyQt6 GUI implementation

## Development Notes

- The GUI handles missing icons gracefully by creating default icons
- Configuration system automatically creates necessary directories
- AI features are optional and require proper API key configuration
- The tool preserves document structure while applying formatting