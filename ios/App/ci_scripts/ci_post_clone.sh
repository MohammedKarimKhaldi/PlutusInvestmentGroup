#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/../../.." && pwd)"

echo "== Xcode Cloud: preparing JavaScript dependencies =="
echo "Repository root: ${REPO_ROOT}"

cd "${REPO_ROOT}"

configure_npm() {
  local npm_registry="${XCLOUD_NPM_REGISTRY:-${NPM_CONFIG_REGISTRY:-https://registry.npmjs.org/}}"

  export NPM_CONFIG_REGISTRY="${npm_registry}"
  export npm_config_registry="${npm_registry}"
  export NPM_CONFIG_AUDIT=false
  export npm_config_audit=false
  export NPM_CONFIG_FUND=false
  export npm_config_fund=false
  export NPM_CONFIG_FETCH_RETRIES="${NPM_CONFIG_FETCH_RETRIES:-5}"
  export npm_config_fetch_retries="${NPM_CONFIG_FETCH_RETRIES}"
  export NPM_CONFIG_FETCH_RETRY_MINTIMEOUT="${NPM_CONFIG_FETCH_RETRY_MINTIMEOUT:-20000}"
  export npm_config_fetch_retry_mintimeout="${NPM_CONFIG_FETCH_RETRY_MINTIMEOUT}"
  export NPM_CONFIG_FETCH_RETRY_MAXTIMEOUT="${NPM_CONFIG_FETCH_RETRY_MAXTIMEOUT:-120000}"
  export npm_config_fetch_retry_maxtimeout="${NPM_CONFIG_FETCH_RETRY_MAXTIMEOUT}"

  echo "Using npm registry ${NPM_CONFIG_REGISTRY}"
}

ensure_xcode_cloud_workspace_selection() {
  if [ "${CI_XCODE_CLOUD:-}" != "TRUE" ]; then
    return
  fi

  local selected_container="${CI_XCODE_PROJECT:-}"

  if [ -n "${selected_container}" ] && [[ "${selected_container}" != *.xcworkspace ]]; then
    echo "Xcode Cloud is configured to build ${selected_container}, but this repo requires ios/App/App.xcworkspace."
    echo "Capacitor dependencies are integrated through CocoaPods, so building App.xcodeproj will fail to resolve the Capacitor modules."
    echo "In App Store Connect, edit or recreate the workflow so it uses ios/App/App.xcworkspace with scheme App."
    exit 1
  fi
}

ensure_node_runtime() {
  local brew_bin=""
  local required_node_major="${XCLOUD_NODE_MAJOR:-22}"
  local node_formula="${XCLOUD_NODE_FORMULA:-node@${required_node_major}}"
  local current_node_major=""

  export PATH="/opt/homebrew/bin:/usr/local/bin:/opt/homebrew/opt/node/bin:/usr/local/opt/node/bin:/opt/homebrew/opt/${node_formula}/bin:/usr/local/opt/${node_formula}/bin:${PATH}"

  for brew_bin in /opt/homebrew/bin/brew /usr/local/bin/brew; do
    if [ -x "${brew_bin}" ]; then
      eval "$("${brew_bin}" shellenv)"
      break
    fi
  done

  if command -v node >/dev/null 2>&1 && command -v npm >/dev/null 2>&1; then
    current_node_major="$(node -p 'process.versions.node.split(".")[0]')"

    if [ "${current_node_major}" -ge "${required_node_major}" ]; then
      echo "Using Node $(node -v) and npm $(npm -v)"
      return
    fi

    echo "Node $(node -v) is too old; need Node >=${required_node_major}."
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

  current_node_major="$(node -p 'process.versions.node.split(".")[0]')"
  if [ "${current_node_major}" -lt "${required_node_major}" ]; then
    echo "Node $(node -v) is still too old after bootstrap; need Node >=${required_node_major}."
    exit 1
  fi

  echo "Using Node $(node -v) and npm $(npm -v)"
}

ensure_node_runtime
ensure_xcode_cloud_workspace_selection
configure_npm

run_npm_install() {
  local command_name="$1"
  shift
  local log_file

  log_file="$(mktemp -t xcloud-npm-install.XXXXXX.log)"
  if "$@" 2>&1 | tee "${log_file}"; then
    rm -f "${log_file}"
    return 0
  fi

  if grep -Eq 'ENOTFOUND|EAI_AGAIN|getaddrinfo|network request to https?://.+' "${log_file}"; then
    echo "npm could not reach ${NPM_CONFIG_REGISTRY} from Xcode Cloud."
    echo "This is a network/DNS issue during dependency installation, not a source-code compile error."
    echo "Try rerunning the build. If it keeps failing, set XCLOUD_NPM_REGISTRY to a reachable registry mirror in the workflow Environment settings."
    echo "Failed command: ${command_name}"
    rm -f "${log_file}"
    exit 1
  fi

  rm -f "${log_file}"
  return 1
}

install_dependencies() {
  if [ -f package-lock.json ] && git ls-files --error-unmatch package-lock.json >/dev/null 2>&1; then
    if run_npm_install "npm ci" npm ci --no-audit --fund=false; then
      return
    fi

    echo "package-lock.json is out of sync, falling back to npm install."
  fi

  run_npm_install "npm install --no-package-lock" npm install --no-package-lock --no-audit --fund=false
}

ensure_cocoapods() {
  if command -v pod >/dev/null 2>&1; then
    echo "Using CocoaPods $(pod --version)"
    return
  fi

  if command -v brew >/dev/null 2>&1; then
    echo "CocoaPods is not on PATH, installing cocoapods with Homebrew."
    export HOMEBREW_NO_AUTO_UPDATE=1
    export HOMEBREW_NO_INSTALL_CLEANUP=1
    brew install cocoapods
    export PATH="/opt/homebrew/bin:/usr/local/bin:${PATH}"
  fi

  if ! command -v pod >/dev/null 2>&1; then
    echo "CocoaPods is unavailable after bootstrap."
    exit 1
  fi

  echo "Using CocoaPods $(pod --version)"
}

install_dependencies
ensure_cocoapods

echo "== Xcode Cloud: syncing Capacitor iOS project =="
npm run ios:sync
