#!/bin/sh
# Fetch academy template at boot when secret-file mount is unavailable (Render 1MB limit).
# Expects deploy key at /etc/secrets/ppt_assets_deploy_key (read-only deploy key for
# cxr542/ppt-academizer-assets) unless TEMPLATE_PPTX already points to an existing file.
set -eu

DEST="${TEMPLATE_PPTX:-/tmp/academy-template.pptx}"
REPO_SSH="${TEMPLATE_ASSETS_REPO:-git@github.com:cxr542/ppt-academizer-assets.git}"
KEY_PATH="${TEMPLATE_ASSETS_DEPLOY_KEY:-/etc/secrets/ppt_assets_deploy_key}"

if [ ! -f "$DEST" ]; then
  if [ ! -f "$KEY_PATH" ]; then
    echo "ppt-academizer: template missing ($DEST) and deploy key not found ($KEY_PATH)" >&2
  else
    echo "ppt-academizer: fetching academy template from $REPO_SSH"
    mkdir -p /root/.ssh
    cp "$KEY_PATH" /root/.ssh/id_ed25519
    chmod 600 /root/.ssh/id_ed25519
    if [ ! -f /root/.ssh/known_hosts ] || ! grep -q github.com /root/.ssh/known_hosts 2>/dev/null; then
      ssh-keyscan -t ed25519,rsa github.com >> /root/.ssh/known_hosts 2>/dev/null || true
    fi
    rm -rf /tmp/ppt-academizer-assets
    git clone --depth 1 "$REPO_SSH" /tmp/ppt-academizer-assets
    cp /tmp/ppt-academizer-assets/academy-template.pptx "$DEST"
    rm -rf /tmp/ppt-academizer-assets
    echo "ppt-academizer: template ready at $DEST"
  fi
fi

export TEMPLATE_PPTX="$DEST"
exec "$@"
