#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
if [ ! -x ".venv/bin/python" ]; then
  echo "Creating a local Python environment..."
  python3 -m venv .venv
fi
.venv/bin/python -m pip install -r requirements.txt
exec .venv/bin/python server.py
