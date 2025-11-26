#!/bin/bash
cd "/Users/quentinpro/Desktop/ByCommute - Code/Test"
git add .
git commit -m "Mise à jour du code - $(date +%Y-%m-%d)"
git push origin main
echo "✅ Modifications poussées sur GitHub"
echo "📦 Repository: https://github.com/bycommute/abris-velos-generator"

