#!/bin/bash
# scripts/use-node.sh — select the Node version this project builds with.
#
# Source it, don't execute it:   . "$(dirname "$0")/scripts/use-node.sh"
#
# The required version comes from .nvmrc, so there is one source of truth rather
# than a hardcoded number in every build script.
#
# Why this is more than `nvm use`:
#
# A version manager reporting success does NOT mean `node` on PATH changed.
# fnm activates a version by rewriting a symlink inside $FNM_MULTISHELL_PATH,
# which only works if that directory is on PATH — something `eval "$(fnm env)"`
# arranges. When FNM_MULTISHELL_PATH is *inherited* from a parent shell whose
# PATH entry did not come with it (common in CI, tmux, and tool-spawned shells),
# fnm writes the symlink into a directory nothing will ever look in, prints
# "Using Node v22.17.1", and exits 0 — while `node` keeps resolving to whatever
# the ambient PATH holds. The old `nvm use 22.17.1` line hit exactly that: gulp
# ran under v24 and SPFx aborted, after the release script had already bumped
# the version.
#
# So every strategy below is followed by an actual `node --version` check, and
# the script fails loudly if none of them worked. Verify, never assume.

# --------------------------------------------------------- required version
__use_node_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

if [ ! -f "${__use_node_root}/.nvmrc" ]; then
    echo "use-node.sh: ${__use_node_root}/.nvmrc not found; cannot determine the required Node version." >&2
    return 1 2>/dev/null || exit 1
fi

# Tolerate "22.17.1", "v22.17.1", and trailing whitespace.
REQUIRED_NODE="$(tr -d '[:space:]' < "${__use_node_root}/.nvmrc")"
REQUIRED_NODE="${REQUIRED_NODE#v}"

__node_is_required() {
    command -v node >/dev/null 2>&1 && [ "$(node --version 2>/dev/null)" = "v${REQUIRED_NODE}" ]
}

# Prepend a directory to PATH if it actually contains a node binary.
__try_node_bin_dir() {
    [ -x "$1/node" ] || return 1
    PATH="$1:$PATH"
    export PATH
    __node_is_required
}

# ------------------------------------------------------------- strategy 1/4
# Already correct — nothing to do. Covers shells whose profile is set up right.
if ! __node_is_required; then

    # --------------------------------------------------------- strategy 2/4
    # nvm, if a real one is installed. Sourcing is done in a subshell-safe way;
    # `nvm use` failing is not fatal, we just fall through to the next strategy.
    if [ -s "${NVM_DIR:-$HOME/.nvm}/nvm.sh" ]; then
        # shellcheck disable=SC1090
        . "${NVM_DIR:-$HOME/.nvm}/nvm.sh" >/dev/null 2>&1 || true
        if command -v nvm >/dev/null 2>&1 || type nvm >/dev/null 2>&1; then
            nvm use "$REQUIRED_NODE" >/dev/null 2>&1 || true
        fi
    fi
fi

if ! __node_is_required; then

    # --------------------------------------------------------- strategy 3/4
    # fnm. Re-evaluate `fnm env` unconditionally: an inherited-but-stale
    # FNM_MULTISHELL_PATH is precisely the failure this script exists to survive,
    # so we must not treat "the variable is set" as "fnm is wired up".
    __use_node_fnm=""
    if command -v fnm >/dev/null 2>&1; then
        __use_node_fnm="$(command -v fnm)"
    else
        for __use_node_cand in \
            "${FNM_DIR:-$HOME/.local/share/fnm}/fnm" \
            "$HOME/.local/share/fnm/fnm" \
            "$HOME/.fnm/fnm" \
            "/usr/local/bin/fnm" \
            "/opt/homebrew/bin/fnm"
        do
            [ -x "$__use_node_cand" ] && { __use_node_fnm="$__use_node_cand"; break; }
        done
        unset __use_node_cand
    fi

    if [ -n "$__use_node_fnm" ]; then
        unset FNM_MULTISHELL_PATH
        eval "$("$__use_node_fnm" env --shell bash)" >/dev/null 2>&1 || true
        "$__use_node_fnm" use "$REQUIRED_NODE" >/dev/null 2>&1 || true
    fi
fi

if ! __node_is_required; then

    # --------------------------------------------------------- strategy 4/4
    # Last resort: put the version's bin directory on PATH directly. This needs
    # no shell integration at all, so it works where every `... env` dance fails.
    for __use_node_dir in \
        "${FNM_DIR:-$HOME/.local/share/fnm}/node-versions/v${REQUIRED_NODE}/installation/bin" \
        "$HOME/.local/share/fnm/node-versions/v${REQUIRED_NODE}/installation/bin" \
        "$HOME/.fnm/node-versions/v${REQUIRED_NODE}/installation/bin" \
        "${NVM_DIR:-$HOME/.nvm}/versions/node/v${REQUIRED_NODE}/bin"
    do
        __try_node_bin_dir "$__use_node_dir" && break
    done
    unset __use_node_dir
fi

# ------------------------------------------------------------------- verdict
if ! __node_is_required; then
    echo "" >&2
    echo "ERROR: Node v${REQUIRED_NODE} (from .nvmrc) could not be activated." >&2
    echo "       node is currently: $(command -v node >/dev/null 2>&1 && node --version || echo 'not found')" >&2
    echo "" >&2
    echo "       SPFx refuses to build on an unsupported Node version, so stopping" >&2
    echo "       here rather than letting gulp fail further in." >&2
    echo "" >&2
    echo "       Install it with:  fnm install ${REQUIRED_NODE}   (or: nvm install ${REQUIRED_NODE})" >&2
    echo "" >&2
    unset __use_node_root __use_node_fnm REQUIRED_NODE
    unset -f __node_is_required __try_node_bin_dir
    return 1 2>/dev/null || exit 1
fi

unset __use_node_root __use_node_fnm
unset -f __node_is_required __try_node_bin_dir
