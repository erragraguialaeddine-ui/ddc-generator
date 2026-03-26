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

# 2. Lancer Ollama si pas déjà actif
if ! curl -s http://localhost:11434/api/version > /dev/null 2>&1; then
    echo "→ Démarrage Ollama..."
    ollama serve &
    sleep 3
fi
echo "✓ Ollama actif sur :11434"

# 3. Vérifier le modèle
MODEL="${OLLAMA_MODEL:-mistral-nemo}"
if ! ollama list | grep -q "$MODEL"; then
    echo "⚠ $MODEL absent. Téléchargement..."
    ollama pull "$MODEL"
fi
echo "✓ Modèle $MODEL prêt"

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

# 6. Lancer ngrok en background + publier l'URL
echo ""
if command -v ngrok &> /dev/null; then
    # Vérifier si ngrok tourne déjà
    NGROK_URL=$(curl -s http://127.0.0.1:4040/api/tunnels 2>/dev/null | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['tunnels'][0]['public_url'])" 2>/dev/null || echo "")
    if [ -z "$NGROK_URL" ]; then
        echo "→ Démarrage ngrok..."
        ngrok http 8000 --log=false &
        NGROK_PID=$!
        sleep 4
        NGROK_URL=$(curl -s http://127.0.0.1:4040/api/tunnels 2>/dev/null | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['tunnels'][0]['public_url'])" 2>/dev/null || echo "")
    else
        echo "✓ ngrok déjà actif"
    fi
    if [ -n "$NGROK_URL" ]; then
        echo "✓ ngrok actif : $NGROK_URL"
        # Mettre à jour config.json et pousser sur GitHub Pages
        echo "{\"api_url\": \"$NGROK_URL\"}" > "$SCRIPT_DIR/docs/config.json"
        cd "$SCRIPT_DIR"
        git add docs/config.json && git commit -m "update ngrok url" --quiet && git push --quiet 2>/dev/null &
        echo "✓ URL publiée sur GitHub Pages"
    else
        echo "⚠ ngrok démarré mais URL non récupérée"
    fi
else
    echo "ℹ ngrok non installé — accès local uniquement"
    echo "  Pour installer : brew install ngrok"
fi

# 7. Démarrer le backend
echo ""
echo "→ Démarrage du backend FastAPI..."
echo ""
echo "┌──────────────────────────────────────────────┐"
echo "│  Backend  : http://localhost:8000             │"
if [ -n "$NGROK_URL" ]; then
echo "│  Public   : $NGROK_URL"
echo "│  Site     : https://erragraguialaeddine-ui.github.io/ddc-generator/"
fi
echo "│  Password : v4f2025                          │"
echo "└──────────────────────────────────────────────┘"
echo ""

cd "$SCRIPT_DIR/backend"
python3 -m uvicorn main:app --reload --port 8000 --host 0.0.0.0
