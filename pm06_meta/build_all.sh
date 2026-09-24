#!/bin/bash
# Пересборка pm06_meta из исходных выгрузок ветки rawdata.
# Файлы скачиваются по одному, обрабатываются и сразу удаляются —
# все выгрузки (≈700 МБ) на диске одновременно не нужны.
#   pm06_meta/build_all.sh <рабочая_папка>
set -u
HERE=$(cd "$(dirname "$0")" && pwd)
WORK=${1:?укажите рабочую папку}
BASE="https://raw.githubusercontent.com/victorekuznetsov/TOPO/rawdata"
mkdir -p "$WORK"
one(){ name=$1; shift
  [ -f "$WORK/$name.json" ] && { echo "$name: уже есть"; return; }
  ext=zip; [[ "$1" == *.xlsx ]] && ext=xlsx
  tmp="$WORK/_src.$ext"; urls=(); for p in "$@"; do urls+=("$BASE/$p"); done
  for t in 1 2 3; do
    curl -sSfL "${urls[@]}" > "$tmp" && python3 "$HERE/build_pm06_meta.py" extract "$tmp" "$name" "$WORK" && break
    echo "$name: повтор $t"; sleep $((t*4))
  done
  rm -f "$tmp"; }
for y in 2022 2023 2024 2025 2026; do one 1100_$y M06_1100_$y.zip.001 M06_1100_$y.zip.002 M06_1100_$y.zip.003; done
one 1100_2027 M06_1100_2027.zip.001 M06_1100_2027.zip.002
for y in 2022 2023 2024; do one 1200_$y M06_1200_$y.xlsx; done
one 1200_2025 M06_1200_2025.zip.001; one 1200_2026 M06_1200_2026.zip.001; one 1200_2027 M06_1200_2027.xlsx
one 1300_2023 M06_1300_2023.xlsx; one 1300_2024 M06_1300_2024.xlsx; one 1300_2025 M06_1300_2025.zip.001
one 1300_2026 M06_1300_2026.zip.001 M06_1300_2026.zip.002; one 1300_2027 M06_1300_2027.zip.001
for y in 2022 2023 2024; do one 1400_$y M06_1400_$y.xlsx; done
one 1400_2025 M06_1400_2025.zip.001; one 1400_2026 M06_1400_2026.zip.001 M06_1400_2026.zip.002; one 1400_2027 M06_1400_2027.xlsx
python3 "$HERE/build_pm06_meta.py" assemble "$WORK" "$HERE"
echo ГОТОВО
