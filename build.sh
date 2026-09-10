#!/bin/bash
# Build script for the sp-new-library-uncheck-navigation SPFx solution.
#
# Usage: ./build.sh [--increment-version]
#
# Selects the Node version from .nvmrc, optionally bumps the version across
# package.json / package-lock.json / config/package-solution.json, then runs the
# heft production build (lint, compile, bundle, package-solution).

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m'

INCREMENT_VERSION=false
while [[ $# -gt 0 ]]; do
  case "$1" in
    --increment-version) INCREMENT_VERSION=true; shift ;;
    --help)
      echo "Usage: ./build.sh [--increment-version]"
      exit 0 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

# Select the Node version from .nvmrc; aborts if it cannot be activated.
source "$(dirname "${BASH_SOURCE[0]}")/scripts/use-node.sh" || exit 1

# The local npm shim (Socket Firewall) needs network to its own API; bypass it
# for the build, which does not install anything.
export SFW_BYPASS=1

if [ "$INCREMENT_VERSION" = true ]; then
  echo -e "${BLUE}Incrementing version...${NC}"
  ./increment-version.sh
fi

echo "Using Node version: $(node --version)"

PACKAGE_PATH="sharepoint/solution/sp-new-library-uncheck-navigation.sppkg"
rm -f "$PACKAGE_PATH"

echo "Building solution..."
npm run build

if [ ! -f "$PACKAGE_PATH" ]; then
  echo -e "${RED}Error: Package not found at $PACKAGE_PATH${NC}"
  exit 1
fi

echo -e "${GREEN}Build complete!${NC} Package is available at: $PACKAGE_PATH"
