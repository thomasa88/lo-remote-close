# Introduction

This is a small script for closing Libreoffice documents via a shell on Linux. It can be used for closing documents remotely, using SSH.

# Example

```bash
# Invoke using uv run lo-remote-close.py or .venv/bin/python3 lo-remote-close.py .

$ .venv/bin/python3 lo-remote-close.py --help
usage: lo-remote-close.py [-h] {list,close} ...

positional arguments:
  {list,close}
    list        List open documents
    close       Close documents matching the given substrings

options:
  -h, --help    show this help message and exit

$ .venv/bin/python3 lo-remote-close.py list
book.odt
calculations.ods (Modified)
User manual.docx (Modified)

$ .venv/bin/python3 lo-remote-close.py close calc
"calculations.ods" is modified. Not closing.

# Save and close
$ .venv/bin/python3 lo-remote-close.py close -s calc
Saved "calculations.ods".
Closed "calculations.ods".

# Force close
$ .venv/bin/python3 lo-remote-close.py close -f calc
Closed "calculations.ods".
```

# Installation

First install the uv package manager: https://github.com/astral-sh/uv?tab=readme-ov-file#installation .

```bash
uv sync

# Set up symlinks to uno.py files. Fixes the error: "Are you sure that uno has been imported?"
uv run oooenv cmd-link -a
```
