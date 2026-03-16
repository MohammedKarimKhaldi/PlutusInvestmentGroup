#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"

echo "== Xcode Cloud: verifying JavaScript dependencies before xcodebuild =="
echo "Repository root: ${REPO_ROOT}"

cd "${REPO_ROOT}"

ensure_node_runtime() {
  local brew_bin=""
  local node_formula="${XCLOUD_NODE_FORMULA:-node@20}"

  export PATH="/opt/homebrew/bin:/usr/local/bin:/opt/homebrew/opt/node/bin:/usr/local/opt/node/bin:/opt/homebrew/opt/${node_formula}/bin:/usr/local/opt/${node_formula}/bin:${PATH}"

  for brew_bin in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [ -x "${brew_bin}" ]; then
      eval "$("${brew_bin}" shellenv)"
      break
    fi
  done

  if command -v node >/dev/null 2>&1 && command -v npm >/dev/null 2>&1; then
    echo "Using Node $(node -v) and npm $(npm -v)"
    return
  fi

  if command -v brew >/dev/null 2>&1; then
    echo "Node.js is not on PATH, installing ${node_formula} with Homebrew."
    export HOMEBREW_NO_AUTO_UPDATE=1
    export HOMEBREW_NO_INSTALL_CLEANUP=1

    if ! brew list "${node_formula}" >/dev/null 2>&1; then
      if ! brew install "${node_formula}"; then
        echo "Homebrew could not install ${node_formula}, retrying with node."
        node_formula="node"
        brew install "${node_formula}"
      fi
    fi

    export PATH="$(brew --prefix "${node_formula}")/bin:${PATH}"
  fi

  if ! command -v node >/dev/null 2>&1 || ! command -v npm >/dev/null 2>&1; then
    echo "Node.js/npm are unavailable after bootstrap."
    exit 1
  fi

  echo "Using Node $(node -v) and npm $(npm -v)"
}

ensure_node_runtime

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
