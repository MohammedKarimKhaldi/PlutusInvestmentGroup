#!/bin/zsh
set -euo pipefail

echo "== Xcode Cloud: preparing JavaScript dependencies =="

if ! command -v npm >/dev/null 2>&1; then
  echo "npm is not available on this runner."
  exit 1
fi

npm install --no-package-lock

echo "== Xcode Cloud: syncing Capacitor iOS project =="
npm run ios:sync
