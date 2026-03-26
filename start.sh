#!/bin/bash
# ============================================================
# DDC Generator V4F — Script de démarrage
# ============================================================
# Usage : bash start.sh
# ============================================================

set -e

echo ""
echo "╔══════════════════════════════════════╗"
echo "║     DDC Generator V4F — Démarrage    ║"
echo "╚══════════════════════════════════════╝"
echo ""

# 1. Vérifier Ollama
if ! command -v ollama &> /dev/null; then
    echo "❌ Ollama non trouvé. Installe-le sur https://ollama.com"
    exit 1
fi

echo "✓ Ollama trouvé"

# 2. Vérifier le modèle
MODEL="${OLLAMA_MODEL:-mistral-nemo}"
if ! ollama list | grep -q "$MODEL"; then
    echo "⚠ $MODEL absent. Téléchargement..."
    ollama pull "$MODEL"
fi
echo "✓ Modèle $MODEL prêt"

# 3. Lancer Ollama en background si pas déjà actif
if ! curl -s http://localhost:11434/api/version > /dev/null 2>&1; then
    echo "→ Démarrage Ollama..."
    ollama serve &
    sleep 3
fi
echo "✓ Ollama actif sur :11434"

# 4. Installer les dépendances Python
echo ""
echo "→ Vérification des dépendances Python..."
pip3 install -q fastapi uvicorn python-multipart pdfplumber ollama lxml python-pptx 2>/dev/null
echo "✓ Dépendances OK"

# 5. Vérifier les assets PPT
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TEMPLATE="$SCRIPT_DIR/TEMPLATE.pptx"
if [ -f "$TEMPLATE" ]; then
    echo "✓ Asset template trouvé : $TEMPLATE"
else
    echo "ℹ Aucun template PPT requis pour le nouveau moteur (generation par code)"
fi

# 6. Lancer ngrok (optionnel)
echo ""
echo "→ Pour exposer l'API sur internet (optionnel) :"
echo "   ngrok http 8000"
echo ""

# 7. Démarrer le backend
echo "→ Démarrage du backend FastAPI..."
echo ""
echo "┌─────────────────────────────────────────┐"
echo "│  Backend  : http://localhost:8000        │"
echo "│  Frontend : ouvre index.html dans Chrome │"
echo "│  Password : v4f2025 (modifiable dans     │"
echo "│             main.py → PASSWORD)          │"
echo "└─────────────────────────────────────────┘"
echo ""

cd "$SCRIPT_DIR/backend"
uvicorn main:app --reload --port 8000 --host 0.0.0.0
