# CLAUDE.md

@AGENTS.md

## Claude Code specifics

- The repo-wide rules live in `AGENTS.md` (imported above). This file only adds
  what is specific to working here with Claude Code.
- Prefer the `Edit` tool for `package.json` and `config/*.json`. Some sandboxed
  shells mount those paths read-only and `sed -i`/`node -e` writes fail with
  `EROFS`; `increment-version.sh` uses `sed -i` and will stop half-way in that
  case. Finish the bump by hand (all three files must agree) and continue the
  release steps manually: build, commit `release: vX.Y.0.Z`, tag, push, `gh release create`.
- If `npm` cannot resolve the registry inside a sandbox with a package-manager
  shim, set `SFW_BYPASS=1`.
- If the working tree has no `origin` remote (sandboxes may block `.git/config`
  writes), push with the explicit URL
  `https://github.com/JFDI-Consulting/sp-new-library-uncheck-navigation.git`.
- Running the headed proof needs a display and, behind an authenticating proxy,
  `HTTPS_PROXY` set; `e2e/prove.js` parses it. Do not convert the proof to
  headless without checking that the Microsoft sign-in still works.
- Do not "tidy" `docs/proof/*.png`; they are the evidence referenced from the
  README and CHANGELOG. Regenerate them by running the proof, not by hand.
- When asked to change behaviour, re-read `docs/DECISIONS.md` first. Several
  obvious-looking improvements (Site Settings link, property bag storage,
  tenant-wide deployment) were tried or analysed and rejected with evidence.
