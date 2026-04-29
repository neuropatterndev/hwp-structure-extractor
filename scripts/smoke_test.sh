#!/usr/bin/env bash
set -euo pipefail
python -m pip install -e ".[dev]"
python -m pytest
hwp-parse --help >/dev/null
