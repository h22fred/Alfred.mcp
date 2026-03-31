#!/bin/bash
# Run this before committing Setup.command changes to regenerate CHECKSUMS.txt
shasum -a 256 Setup.command > CHECKSUMS.txt
echo "Updated CHECKSUMS.txt:"
cat CHECKSUMS.txt
