#!/usr/bin/env bash
# Устанавливает шрифты VTB Group для точной подгонки текста и рендера (LibreOffice) на Linux/macOS.
# Использование: install_fonts.sh [папка_со_шрифтами]   (по умолчанию — assets/fonts скилла)
# Шрифты лицензированы ParaType для ВТБ и в репозиторий не входят: положите .otf/.ttf в assets/fonts.
set -euo pipefail
SRC="${1:-$(cd "$(dirname "$0")/.." && pwd)/assets/fonts}"
if [ "$(uname)" = "Darwin" ]; then DST="$HOME/Library/Fonts"; else DST="$HOME/.local/share/fonts/vtb"; fi
mkdir -p "$DST"
n=0
while IFS= read -r -d '' f; do cp "$f" "$DST/"; n=$((n+1)); done < <(find "$SRC" -type f \( -iname 'VTB*.otf' -o -iname 'VTB*.ttf' \) -print0)
[ "$n" -eq 0 ] && { echo "Шрифты VTB*.otf/ttf не найдены в $SRC"; exit 1; }
command -v fc-cache >/dev/null && fc-cache -f "$DST" >/dev/null || true
echo "Установлено файлов: $n → $DST"
command -v fc-match >/dev/null && fc-match "VTB Group Book" || true
