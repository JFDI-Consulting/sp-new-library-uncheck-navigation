#!/bin/bash

# Script to increment version numbers across all project manifests.
#
# Algorithm:
# 1. Read versions from all extant files
# 2. Normalise to a common 3-part (MAJOR.MINOR.PATCH) base
# 3. Take the highest version, increment PATCH
# 4. Write back to ALL files (package.json and package-lock.json as X.Y.Z,
#    the SharePoint/Teams manifests as X.Y.0.Z)
#
# This ensures files never drift out of sync, even if a previous run
# partially failed (e.g. sandbox blocking sed on one file).
#
# package-lock.json is edited with node rather than sed, deliberately. Its own
# version appears twice (the root object and `packages[""]`) among ~2850 other
# "version" lines belonging to dependencies, so a global substitution would
# rewrite a dependency the moment one happened to share the project's version
# number. Parsing the JSON and setting the two known fields cannot make that
# mistake, and is also correct for lockfileVersion 1, where `packages` does not
# exist and only one field is present. Round-tripping through
# JSON.stringify(…, null, 2) reproduces npm's own formatting byte for byte
# (verified against this repo's 1.4 MB lockfile), so the diff stays minimal.

set -e

echo "Incrementing version numbers..."

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m'

# File paths
PACKAGE_JSON="package.json"
PACKAGE_LOCK_JSON="package-lock.json"
PACKAGE_SOLUTION_JSON="config/package-solution.json"
TEAMS_MANIFEST_JSON="teams/manifest.json"

# Read a version out of a JSON file without regex. `key` is "root" for the
# top-level version, or "packages" for the packages[""] entry a modern lockfile
# carries. Prints an empty string when the field is absent.
read_json_version() {
    JSON_FILE="$1" JSON_KEY="$2" node -e '
        const fs = require("fs");
        const doc = JSON.parse(fs.readFileSync(process.env.JSON_FILE, "utf8"));
        const node = process.env.JSON_KEY === "packages"
            ? (doc.packages && doc.packages[""])
            : doc;
        process.stdout.write(String((node && node.version) || ""));
    '
}

# --- Read current versions ---

if [ ! -f "$PACKAGE_JSON" ]; then
    echo -e "${RED}Error: $PACKAGE_JSON not found!${NC}"
    exit 1
fi

if [ ! -f "$PACKAGE_SOLUTION_JSON" ]; then
    echo -e "${RED}Error: $PACKAGE_SOLUTION_JSON not found!${NC}"
    exit 1
fi

SKIP_TEAMS=true
if [ -f "$TEAMS_MANIFEST_JSON" ]; then
    SKIP_TEAMS=false
fi

SKIP_LOCK=true
if [ -f "$PACKAGE_LOCK_JSON" ]; then
    SKIP_LOCK=false
    # Erroring here rather than skipping is the point of this whole guard: a
    # silently un-bumped lockfile is exactly the drift this script exists to
    # prevent, and it is invisible until a build turns out not to be
    # reproducible. This runs under build.sh, which sources scripts/use-node.sh
    # first and aborts there if the .nvmrc version could not be activated.
    if ! command -v node > /dev/null 2>&1; then
        echo -e "${RED}Error: $PACKAGE_LOCK_JSON exists but node is not on PATH — cannot update it safely.${NC}"
        echo -e "${RED}Refusing to bump the other files and leave the lockfile behind.${NC}"
        exit 1
    fi
fi

# Extract version strings
PKG_VER=$(grep -oP '"version":\s*"\K[^"]+' "$PACKAGE_JSON")
SOL_VER=$(grep -oP '"version":\s*"\K[^"]+' "$PACKAGE_SOLUTION_JSON" | head -n 1)

TEAMS_VER=""
if [ "$SKIP_TEAMS" = false ]; then
    TEAMS_VER=$(grep -oP '"version":\s*"\K[^"]+' "$TEAMS_MANIFEST_JSON")
fi

LOCK_VER=""
LOCK_PKG_VER=""
if [ "$SKIP_LOCK" = false ]; then
    LOCK_VER=$(read_json_version "$PACKAGE_LOCK_JSON" root)
    if [ -z "$LOCK_VER" ]; then
        echo -e "${RED}Error: could not read a version from $PACKAGE_LOCK_JSON${NC}"
        exit 1
    fi
    # Read packages[""] too, and feed it into the highest-version scan below.
    # The two are written together here so this script cannot desync them, but
    # an outside edit can — and reading only the root would then let a HIGHER
    # packages[""] sit out the comparison and the version go backwards. Empty
    # on lockfileVersion 1, which has no packages map.
    LOCK_PKG_VER=$(read_json_version "$PACKAGE_LOCK_JSON" packages)
fi

echo -e "${BLUE}Current versions:${NC}"
echo "  package.json:          $PKG_VER"
if [ "$SKIP_LOCK" = false ]; then
    if [ -n "$LOCK_PKG_VER" ] && [ "$LOCK_PKG_VER" != "$LOCK_VER" ]; then
        echo "  package-lock.json:     $LOCK_VER (root), $LOCK_PKG_VER (packages[\"\"]) — inconsistent, both will be set"
    else
        echo "  package-lock.json:     $LOCK_VER"
    fi
fi
echo "  package-solution.json: $SOL_VER"
if [ "$SKIP_TEAMS" = false ]; then
    echo "  teams/manifest.json:   $TEAMS_VER"
fi

# --- Normalise all versions to 3-part MAJOR.MINOR.PATCH ---
# package.json:          X.Y.Z     -> X.Y.Z
# package-lock.json:     X.Y.Z     -> X.Y.Z
# package-solution.json: X.Y.0.Z   -> X.Y.Z  (drop the build number)
# teams/manifest.json:   X.Y.0.Z   -> X.Y.Z  (drop the build number)

normalise_to_3part() {
    local ver="$1"
    IFS='.' read -ra parts <<< "$ver"
    local count=${#parts[@]}
    if [ "$count" -eq 4 ]; then
        # 4-part: X.Y.BUILD.Z -> take X.Y.Z (skip BUILD at index 2)
        echo "${parts[0]}.${parts[1]}.${parts[3]}"
    elif [ "$count" -eq 3 ]; then
        echo "$ver"
    else
        echo -e "${RED}Error: unexpected version format '$ver'${NC}" >&2
        exit 1
    fi
}

PKG_NORM=$(normalise_to_3part "$PKG_VER")
SOL_NORM=$(normalise_to_3part "$SOL_VER")

TEAMS_NORM=""
if [ "$SKIP_TEAMS" = false ]; then
    TEAMS_NORM=$(normalise_to_3part "$TEAMS_VER")
fi

LOCK_NORM=""
LOCK_PKG_NORM=""
if [ "$SKIP_LOCK" = false ]; then
    LOCK_NORM=$(normalise_to_3part "$LOCK_VER")
    if [ -n "$LOCK_PKG_VER" ]; then
        LOCK_PKG_NORM=$(normalise_to_3part "$LOCK_PKG_VER")
    fi
fi

# --- Find the highest version ---
# Compare by converting to a sortable integer: MAJOR*1000000 + MINOR*1000 + PATCH

version_to_int() {
    IFS='.' read -ra p <<< "$1"
    echo $(( ${p[0]} * 1000000 + ${p[1]} * 1000 + ${p[2]} ))
}

HIGHEST_NORM="$PKG_NORM"
HIGHEST_INT=$(version_to_int "$PKG_NORM")

SOL_INT=$(version_to_int "$SOL_NORM")
if [ "$SOL_INT" -gt "$HIGHEST_INT" ]; then
    HIGHEST_NORM="$SOL_NORM"
    HIGHEST_INT="$SOL_INT"
fi

if [ "$SKIP_TEAMS" = false ] && [ -n "$TEAMS_NORM" ]; then
    TEAMS_INT=$(version_to_int "$TEAMS_NORM")
    if [ "$TEAMS_INT" -gt "$HIGHEST_INT" ]; then
        HIGHEST_NORM="$TEAMS_NORM"
        HIGHEST_INT="$TEAMS_INT"
    fi
fi

if [ "$SKIP_LOCK" = false ] && [ -n "$LOCK_NORM" ]; then
    LOCK_INT=$(version_to_int "$LOCK_NORM")
    if [ "$LOCK_INT" -gt "$HIGHEST_INT" ]; then
        HIGHEST_NORM="$LOCK_NORM"
        HIGHEST_INT="$LOCK_INT"
    fi
fi

if [ "$SKIP_LOCK" = false ] && [ -n "$LOCK_PKG_NORM" ]; then
    LOCK_PKG_INT=$(version_to_int "$LOCK_PKG_NORM")
    if [ "$LOCK_PKG_INT" -gt "$HIGHEST_INT" ]; then
        HIGHEST_NORM="$LOCK_PKG_NORM"
        HIGHEST_INT="$LOCK_PKG_INT"
    fi
fi

# --- Increment the highest version ---

IFS='.' read -ra H <<< "$HIGHEST_NORM"
NEW_MAJOR="${H[0]}"
NEW_MINOR="${H[1]}"
NEW_PATCH=$(( ${H[2]} + 1 ))

NEW_SEMVER="$NEW_MAJOR.$NEW_MINOR.$NEW_PATCH"
NEW_FOURPART="$NEW_MAJOR.$NEW_MINOR.0.$NEW_PATCH"

echo ""
echo -e "${GREEN}New versions:${NC}"
echo "  package.json:          $NEW_SEMVER"
if [ "$SKIP_LOCK" = false ]; then
    echo "  package-lock.json:     $NEW_SEMVER"
fi
echo "  package-solution.json: $NEW_FOURPART"
if [ "$SKIP_TEAMS" = false ]; then
    echo "  teams/manifest.json:   $NEW_FOURPART"
fi

# --- Write new versions ---

sed -i "s/\"version\": \"$PKG_VER\"/\"version\": \"$NEW_SEMVER\"/" "$PACKAGE_JSON"
sed -i "0,/\"version\": \"$SOL_VER\"/s//\"version\": \"$NEW_FOURPART\"/" "$PACKAGE_SOLUTION_JSON"

if [ "$SKIP_TEAMS" = false ]; then
    sed -i "s/\"version\": \"$TEAMS_VER\"/\"version\": \"$NEW_FOURPART\"/" "$TEAMS_MANIFEST_JSON"
fi

# See the header note for why this is node and not sed. Both of the lockfile's
# own version fields are set; `packages[""]` is absent on lockfileVersion 1 and
# is simply skipped there.
if [ "$SKIP_LOCK" = false ]; then
    LOCK_FILE="$PACKAGE_LOCK_JSON" NEW_VER="$NEW_SEMVER" node -e '
        const fs = require("fs");
        const path = process.env.LOCK_FILE;
        const version = process.env.NEW_VER;
        const lock = JSON.parse(fs.readFileSync(path, "utf8"));
        lock.version = version;
        if (lock.packages && lock.packages[""]) {
            lock.packages[""].version = version;
        }
        fs.writeFileSync(path, JSON.stringify(lock, null, 2) + "\n");
    '
fi

# --- Verify writes succeeded ---

VERIFY_PKG=$(grep -oP '"version":\s*"\K[^"]+' "$PACKAGE_JSON")
VERIFY_SOL=$(grep -oP '"version":\s*"\K[^"]+' "$PACKAGE_SOLUTION_JSON" | head -n 1)

FAILED=false
if [ "$VERIFY_PKG" != "$NEW_SEMVER" ]; then
    echo -e "${RED}ERROR: package.json version write failed (got $VERIFY_PKG, expected $NEW_SEMVER)${NC}"
    FAILED=true
fi
if [ "$VERIFY_SOL" != "$NEW_FOURPART" ]; then
    echo -e "${RED}ERROR: package-solution.json version write failed (got $VERIFY_SOL, expected $NEW_FOURPART)${NC}"
    FAILED=true
fi
if [ "$SKIP_TEAMS" = false ]; then
    VERIFY_TEAMS=$(grep -oP '"version":\s*"\K[^"]+' "$TEAMS_MANIFEST_JSON")
    if [ "$VERIFY_TEAMS" != "$NEW_FOURPART" ]; then
        echo -e "${RED}ERROR: teams/manifest.json version write failed (got $VERIFY_TEAMS, expected $NEW_FOURPART)${NC}"
        FAILED=true
    fi
fi
if [ "$SKIP_LOCK" = false ]; then
    # Checked structurally, and BOTH fields, so half a write cannot pass.
    VERIFY_LOCK=$(read_json_version "$PACKAGE_LOCK_JSON" root)
    if [ "$VERIFY_LOCK" != "$NEW_SEMVER" ]; then
        echo -e "${RED}ERROR: package-lock.json version write failed (got $VERIFY_LOCK, expected $NEW_SEMVER)${NC}"
        FAILED=true
    fi
    VERIFY_LOCK_PKG=$(read_json_version "$PACKAGE_LOCK_JSON" packages)
    if [ -n "$VERIFY_LOCK_PKG" ] && [ "$VERIFY_LOCK_PKG" != "$NEW_SEMVER" ]; then
        echo -e "${RED}ERROR: package-lock.json packages[\"\"] version write failed (got $VERIFY_LOCK_PKG, expected $NEW_SEMVER)${NC}"
        FAILED=true
    fi
fi

if [ "$FAILED" = true ]; then
    echo -e "${RED}Version update FAILED — files may be out of sync!${NC}"
    exit 1
fi

echo ""
echo -e "${GREEN}✓ Version numbers updated successfully!${NC}"
echo ""
echo "Summary:"
echo "  package.json:          $PKG_VER → $NEW_SEMVER"
if [ "$SKIP_LOCK" = false ]; then
    echo "  package-lock.json:     $LOCK_VER → $NEW_SEMVER"
fi
echo "  package-solution.json: $SOL_VER → $NEW_FOURPART"
if [ "$SKIP_TEAMS" = false ]; then
    echo "  teams/manifest.json:   $TEAMS_VER → $NEW_FOURPART"
fi
