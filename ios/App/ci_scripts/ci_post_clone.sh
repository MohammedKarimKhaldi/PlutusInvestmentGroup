#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"

echo "== Xcode Cloud: preparing JavaScript dependencies =="
echo "Repository root: ${REPO_ROOT}"

cd "${REPO_ROOT}"

if ! command -v npm >/dev/null 2>&1; then
  echo "npm is not available on this runner."
  exit 1
fi

install_dependencies() {
  if [ -f package-lock.json ]; then
    if npm ci; then
      return
    fi

    echo "package-lock.json is out of sync, falling back to npm install."
  fi

  npm install --no-package-lock
}

install_dependencies

echo "== Xcode Cloud: syncing Capacitor iOS project =="
npm run ios:sync
