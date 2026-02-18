#!/bin/bash
# cleanup-docx.sh - Remove intermediate .docx files after HWPX generation
# Called by Quarto post-render hook

for marker in *.docx.hwpx-cleanup; do
  [ -f "$marker" ] || continue
  docx_file=$(cat "$marker")
  if [ -f "$docx_file" ]; then
    rm -f "$docx_file"
    echo "[hwpx] Removed intermediate $docx_file"
  fi
  rm -f "$marker"
done
