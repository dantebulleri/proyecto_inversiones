#!/bin/bash
# =============================================
# MotoFlip - Script de inicio
# =============================================
echo ""
echo "  MotoFlip - Gestion de Flipping de Motos"
echo "  ========================================="
echo ""

# Ir al directorio del script
cd "$(dirname "$0")"

PORT=8080

# Verificar si el puerto esta en uso
if lsof -i :$PORT > /dev/null 2>&1; then
    PORT=8081
fi

echo "  Iniciando servidor en http://localhost:$PORT"
echo "  Presiona Ctrl+C para detener."
echo ""

# Abrir navegador automaticamente
if command -v xdg-open &> /dev/null; then
    sleep 1 && xdg-open "http://localhost:$PORT" &
elif command -v open &> /dev/null; then
    sleep 1 && open "http://localhost:$PORT" &
fi

python3 -m http.server $PORT
