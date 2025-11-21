#!/bin/bash
set -e

# Activate virtual environment if exists, else create one
if [ ! -d "venv" ]; then
    python3 -m venv venv
fi

source venv/bin/activate

pip install -r requirements.txt

python main.py
