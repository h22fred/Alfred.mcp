#!/bin/bash
# Run this before committing installer changes to regenerate CHECKSUMS.txt
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
shasum -a 256 "$ROOT_DIR/Setup_macOS.command" "$ROOT_DIR/Setup_Windows.bat" > "$SCRIPT_DIR/CHECKSUMS.txt"
echo "Updated CHECKSUMS.txt:"
cat "$SCRIPT_DIR/CHECKSUMS.txt"
