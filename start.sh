#!/bin/bash
echo "================================================"
echo "  VT to EcoTEA Conversion Tool"
echo "================================================"
echo ""

# Create virtual environment if it doesn't exist
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
fi

# Activate virtual environment
source .venv/bin/activate

# Install dependencies
echo "Checking dependencies..."
pip install -r requirements.txt --quiet

echo "Starting tool..."
python app.py
