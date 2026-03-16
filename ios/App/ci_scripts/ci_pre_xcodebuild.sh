#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"

echo "== Xcode Cloud: verifying JavaScript dependencies before xcodebuild =="
echo "Repository root: ${REPO_ROOT}"

cd "${REPO_ROOT}"

install_dependencies() {
  if [ -f package-lock.json ] && git ls-files --error-unmatch package-lock.json >/dev/null 2>&1; then
    if npm ci; then
      return
    fi

    echo "package-lock.json is out of sync, falling back to npm install."
  fi

  npm install --no-package-lock
}

if [ ! -d "node_modules/@capacitor/filesystem" ] || [ ! -d "node_modules/@capacitor/share" ]; then
  echo "Capacitor plugin packages missing, reinstalling dependencies."
  install_dependencies
fi

echo "== Xcode Cloud: refreshing Capacitor iOS sync before xcodebuild =="
npm run ios:sync
