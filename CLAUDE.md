# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Create and activate virtual environment (Windows)
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run all tests
pytest

# Run a single test file
pytest tests/test_example.py

# Run a single test by name
pytest tests/test_example.py::test_function_name
```

## Project Structure

```
ClaudeProjects/
├── src/          # Application source code
├── tests/        # pytest test files
├── requirements.txt
└── CLAUDE.md
```

## Conventions

- Source modules live in `src/`; import them in tests as `from src.module import ...`
- Test files are named `test_*.py` and live in `tests/`
- Always ask before running `pip install` or modifying system settings
