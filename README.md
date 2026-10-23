# Introduction

This is a small script for closing Libreoffice documents via a shell on Linux. It can be used for closing documents remotely, using SSH.

[OOO Development Tools](https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html) is used to control Libreoffice.

# Example

```bash
# Invoke using "uv run lo-remote-close" (if installed with uv),
# or simply "lo-remote-close" (if installed with pipx)

$ uv run lo-remote-close --help
usage: lo-remote-close.py [-h] {list,close} ...

positional arguments:
  {list,close}
    list        List open documents
    close       Close documents matching the given substrings

options:
  -h, --help    show this help message and exit

$ uv run lo-remote-close list
book.odt
calculations.ods (Modified)
User manual.docx (Modified)

$ uv run lo-remote-close close calc
"calculations.ods" is modified. Not closing.

# Save and close
$ uv run lo-remote-close close -s calc
Saved "calculations.ods".
Closed "calculations.ods".

# Force close
$ uv run lo-remote-close close -f calc
Closed "calculations.ods".
```

# Installation

Clone the repository. 

Install the uv package manager: https://github.com/astral-sh/uv?tab=readme-ov-file#installation .

Get dependencies:
```bash
uv sync

# Set up symlinks to uno.py files. Fixes the error: "Are you sure that uno has been imported?"
uv run oooenv cmd-link -a
```

Alternatively, install with [pipx](https://pipx.pypa.io/): From the repository root, run
```bash
pipx install --system-site-packages .
```
Using the `--system-site-packages` option ensures that LibreOffice's modules, such as `uno.py`, are accessible in the resulting virtual environment.
