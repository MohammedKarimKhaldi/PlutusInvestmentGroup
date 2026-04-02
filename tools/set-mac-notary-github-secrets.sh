#!/usr/bin/env bash

set -euo pipefail

if ! command -v gh >/dev/null 2>&1; then
  echo "GitHub CLI (gh) is required."
  exit 1
fi

if [ "$#" -lt 2 ] || [ "$#" -gt 3 ]; then
  echo "Usage: $0 /path/to/AuthKey_<KEY_ID>.p8 <KEY_ID> [ISSUER_ID]"
  exit 1
fi

api_key_path="$1"
key_id="$2"
issuer_id="${3:-}"

if [ ! -f "$api_key_path" ]; then
  echo "API key file not found: $api_key_path"
  exit 1
fi

base64 < "$api_key_path" | gh secret set APPLE_NOTARY_API_KEY_BASE64
printf '%s' "$key_id" | gh secret set APPLE_NOTARY_API_KEY_ID
if [ -n "$issuer_id" ]; then
  printf '%s' "$issuer_id" | gh secret set APPLE_NOTARY_API_ISSUER
else
  gh secret delete APPLE_NOTARY_API_ISSUER >/dev/null 2>&1 || true
fi

echo "Updated GitHub notarization secrets."
