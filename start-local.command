#!/bin/bash
cd "$(dirname "$0")"
echo "Arrancando servidor local en http://localhost:3001 ..."
LOCAL_NO_AUTH=1 \
PORT=3001 \
APP_DATA_DIR="$(pwd)/data" \
PUPPETEER_CACHE_DIR="$HOME/.cache/puppeteer" \
node server.js
