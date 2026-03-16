#!/bin/zsh
set -euo pipefail

echo "== Xcode Cloud: verifying JavaScript dependencies before xcodebuild =="

if [ ! -d "node_modules/@capacitor/filesystem" ] || [ ! -d "node_modules/@capacitor/share" ]; then
  echo "Capacitor plugin packages missing, reinstalling dependencies."
  npm install --no-package-lock
fi

echo "== Xcode Cloud: refreshing Capacitor iOS sync before xcodebuild =="
npm run ios:sync
