# AGENTS.md

Instructions for AI coding agents (and new humans) working in this repository.
`CLAUDE.md` imports this file; keep repo-specific guidance here.

## What this is

An SPFx 1.23 **Application Customizer** (no web parts) that unticks "Show in site
navigation" in SharePoint's modern Create list / Create document library panels,
plus a per-site on/off switch shown on Site contents. Read `README.md` first,
then `docs/ARCHITECTURE.md`. Decisions that look odd usually have an ADR in
`docs/DECISIONS.md` with the tenant evidence; check there before "fixing" them.

## Non-negotiables

- **Node 22 only.** SPFx 1.23 refuses Node 24. `.nvmrc` holds the version;
  `scripts/use-node.sh` selects it and verifies `node --version`. If `npm run build`
  fails immediately, check the Node version before anything else.
- **Never commit `sharepoint/solution/*.sppkg`**, `dist/`, `lib/`, `release/`,
  `temp/`. They are gitignored; releases attach the `.sppkg` to GitHub.
- **Do not add `Microsoft.SharePoint.SiteSettings` custom actions** or any
  `UrlAction` to `sharepoint/assets/elements.xml`. The package will validate in the
  catalog and then fail to install on every site (ADR-003). The only allowed element
  is the `ClientSideExtension.ApplicationCustomizer` action.
- **Do not switch on `skipFeatureDeployment`** (tenant-wide deployment) without
  first replacing the settings store; the panel's Save depends on a per-site custom
  action (ADR-004).
- **Keep React and Fluent UI out of the customizer's static imports.** They load
  through the dynamic `import('./SettingsUi')` in `_syncSettingsUi` so every other
  page load stays at ~16 KB. Check `dist/` sizes after touching imports.
- **`requiresCustomScript` stays `false`** in the manifest, and nothing may depend
  on custom script being enabled. Default modern sites are NoScript.
- Component id `d31c6f18-3a0d-462b-b677-c09314fbf3e6` and solution id
  `3d32ef0c-4b53-4eb9-b342-10def9451d2e` are referenced from `SettingsService.ts`,
  `serve.json`, `elements.xml` and the tenant. Never regenerate them.

## Build, test, release

```bash
. scripts/use-node.sh          # or ./build.sh, which sources it
npm run build                  # lint (warnings fail the release build) + jest + bundle + package
npm run e2e                    # headed Playwright proof against a real tenant; see README
./scripts/release              # bump, build, commit "release: vX.Y.0.Z", tag, GitHub Release
```

- Write the `## [X.Y.0.Z]` CHANGELOG entry before running the release script.
- The version bump touches `package.json` (X.Y.Z), `package-lock.json` and
  `config/package-solution.json` (X.Y.0.Z). Keep them aligned; the script does.
- There are no meaningful unit tests; the proof is `e2e/prove.js`. Any change to
  matching, iframe handling, the settings UI or its lifecycle must be re-proven
  on a tenant, and the "Last verified" line in `README.md` updated.
- The proof can hit a spurious 30 s locator timeout. Rerun once before
  investigating.

## Tenant for proofs

The maintainers use `g53.sharepoint.com`: app catalog `/sites/appcatalog`, test
site `/sites/UncheckNavTest` (app installed), root site as the no-app control.
Deploy with CLI for Microsoft 365 as in `docs/OPERATIONS.md`. When iterating on
the same solution version, SharePoint will not re-provision assets on
`app upgrade`; uninstall and install instead (30 s). A failed install shows as
`AppStatus: 6` and its cause is only visible via
`POST /_api/web/GetAppInstanceById('<instance-id>')/GetErrorDetails`.

## Code conventions

- TypeScript with explicit types (SPFx ESLint profile is enforced; `null` needs
  an eslint-disable with a reason, as in `SettingsService.ts`).
- Fluent UI **8** (`@fluentui/react`), React 17, matching SPFx 1.23. Do not add
  Fluent UI 9; the create-list panel uses it, we do not.
- User-facing strings go in `src/extensions/uncheckSiteNavigation/loc/en-us.js`
  and `myStrings.d.ts`. No hardcoded UI text in components.
- Test hooks for Playwright are `data-automation-id="jfdi-unav-*"`; keep them
  stable or update `e2e/prove.js` in the same change.
- Commit messages: imperative subject, body explains why. Releases are
  `release: vX.Y.0.Z`.

## Where things live

| Path | Purpose |
|------|---------|
| `src/extensions/uncheckSiteNavigation/UncheckSiteNavigationApplicationCustomizer.ts` | The customizer: DOM observation, iframe hooking, retry, settings UI lifecycle |
| `src/extensions/uncheckSiteNavigation/SettingsService.ts` | Reads/writes `enabled` on the custom action via REST (web scope, then site scope) |
| `src/extensions/uncheckSiteNavigation/Settings{Ui,Bar,Panel}.tsx` | Lazy-loaded React UI |
| `sharepoint/assets/elements.xml` | Per-site custom action provisioned by the feature |
| `sharepoint/assets/ClientSideInstance.xml` | Tenant-wide variant (unused today, see ADR-004) |
| `e2e/prove.js` | Headed Playwright proof; `docs/proof/` holds its screenshots |
| `scripts/release`, `increment-version.sh`, `build.sh` | Release tooling, copied from sp-calendar-planner |
| `docs/` | Architecture, ADRs, operations runbook, lessons learned |
