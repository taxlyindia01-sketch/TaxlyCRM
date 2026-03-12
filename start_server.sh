#!/bin/bash
# Taxly CRM — Backend Startup Script (Mac / Linux)

echo ""
echo "  ╔══════════════════════════════════════╗"
echo "  ║     Taxly CRM — Backend Startup      ║"
echo "  ╚══════════════════════════════════════╝"
echo ""

# Move to project root (wherever this script lives)
cd "$(dirname "$0")/backend" || { echo "[ERROR] Cannot find backend/ folder"; exit 1; }

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "[ERROR] python3 not found. Install Python 3.10+ first."
    exit 1
fi

echo "[1/3] Installing / updating packages..."
pip3 install -r requirements.txt -q

echo "[2/3] Packages ready."
echo "[3/3] Starting Taxly backend on http://127.0.0.1:8000 ..."
echo ""
echo "  ✅ Open VS Code Live Server on frontend/index.html"
echo "  ✅ OR visit: http://localhost:8000"
echo "  Press Ctrl+C to stop"
echo ""

python3 -m uvicorn server:app --host 127.0.0.1 --port 8000 --reload
