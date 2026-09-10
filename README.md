# sp-new-library-uncheck-navigation

SharePoint Framework (SPFx) **Application Customizer** that defaults the
**"Show in site navigation"** checkbox to **unchecked** in the modern
*Create list* and *Create document library* panels, with a per-site on/off
switch that site owners control from Site contents.

SharePoint has no server-side setting for this default. The customizer works on
the rendered page: it observes DOM changes, finds the checkbox whose accessible
label matches a configured string, and clicks it once. A user who wants the new
list or library in the navigation ticks the box again as before; the customizer
never overrides a deliberate choice.

| I am a… | Start here |
|---------|------------|
| Site owner who wants to switch it on or off | [Using the switch](#using-the-switch) |
| Tenant or site admin deploying it | [Deploying](#deploying) and [docs/OPERATIONS.md](docs/OPERATIONS.md) |
| Developer maintaining it | [Building and releasing](#building-and-releasing), [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md), [docs/DECISIONS.md](docs/DECISIONS.md), [docs/LESSONS.md](docs/LESSONS.md) |
| AI coding agent | [AGENTS.md](AGENTS.md) |

## Using the switch

On **Site contents**, users with *Manage Web* permission (site owners) see a
one-line status bar at the top of the page saying what the current default is,
with a **Change** link. It opens a panel with a single toggle; **Save** writes
the choice back to the site.

![Settings panel](docs/proof/settings-panel-disable.png)

The panel can also be opened by URL, for linking from a help page or an intranet
landing page of your own:

```
https://tenant.sharepoint.com/sites/YourSite/_layouts/15/viewlsts.aspx?jfdiUncheckNav=settings
```

A change applies immediately in the tab that saved it. Other users and tabs pick
it up on their next page load, once SharePoint's cached page data expires
(typically a few minutes). Users without Manage Web who open the deep link get a
read-only view of the panel.

## Deploying

1. Download the `.sppkg` from the latest [GitHub Release](https://github.com/JFDI-Consulting/sp-new-library-uncheck-navigation/releases).
2. Upload it to the tenant App Catalog (or a site collection App Catalog) and deploy it.
   Leave "Make this solution available to all sites" **unticked**; see
   [Tenant-wide deployment](#tenant-wide-deployment) for why.
3. Add the app to each site collection where you want the behaviour. Adding the
   app activates a feature that registers the customizer's custom action on the
   site.

With CLI for Microsoft 365:

```bash
m365 spo app add --filePath sp-new-library-uncheck-navigation.sppkg --overwrite
m365 spo app deploy --id <app-id>
m365 spo app install --id <app-id> --siteUrl https://tenant.sharepoint.com/sites/YourSite
m365 spo app upgrade --id <app-id> --siteUrl https://tenant.sharepoint.com/sites/YourSite   # later versions
```

To remove the behaviour, remove the app from the site, or switch it off from the
settings panel. Full runbook, troubleshooting and CLI recipes for admin teams:
[docs/OPERATIONS.md](docs/OPERATIONS.md).

### Configuration

The custom action accepts optional `ClientSideComponentProperties`:

| Property  | Type       | Default | Purpose |
|-----------|------------|---------|---------|
| `enabled` | `boolean`  | `true`  | Switch the behaviour off for the site without removing the app. This is what the settings panel writes. |
| `labels`  | `string[]` | `["Show in site navigation", "Show list in site navigation", "Show library in site navigation"]` | Label texts to match (case-insensitive). **Required on non-English tenants**; add the localised label text. |
| `debug`   | `boolean`  | `false` | Log each match and setting change to the browser console. |

Defaults live in `sharepoint/assets/elements.xml`. To change them after
deployment:

```bash
m365 spo customaction list --webUrl https://tenant.sharepoint.com/sites/YourSite    # find the Id
m365 spo customaction set --webUrl https://tenant.sharepoint.com/sites/YourSite \
  --id <Id> --clientSideComponentProperties '{"enabled":false,"labels":["Im Websitenavigation anzeigen"]}'
```

The panel merges `enabled` into whatever is already stored, so `labels` and
`debug` survive a toggle. A reinstall of the app resets the properties to the
package defaults.

### Tenant-wide deployment

The package ships `ClientSideInstance.xml`, so it *can* be deployed tenant-wide
(every existing and future modern site) by setting `skipFeatureDeployment` and
ticking "Make this solution available to all sites". **Do not do this with the
current version.** In that mode there is no per-site custom action, so the
settings panel has nothing to write to and Save fails with an explanatory
message. A site-local store that works tenant-wide (a hidden settings list) is
designed but not built; see [docs/DECISIONS.md](docs/DECISIONS.md#adr-004-per-site-install-not-tenant-wide-deployment-for-now).

### Why is the switch not on the Site Settings page?

Because it cannot be, on a standard modern site:

- App packages cannot declare a custom action with location
  `Microsoft.SharePoint.SiteSettings`. The feature schema rejects it at install
  time, and the catalog rejects `~site` URLs.
- Adding such a link afterwards with REST, PnP or CLI returns 403 on sites with
  custom script disabled (the default), even for tenant admins.
- Site Settings is a classic page where application customizers do not run, so
  the panel could not live there anyway.

Site contents is where lists and libraries are created, so the setting sits
next to the behaviour it controls. Details and evidence in
[docs/DECISIONS.md](docs/DECISIONS.md#adr-003-settings-entry-point-on-site-contents-not-site-settings).

## Building and releasing

Requires Node 22.x (`.nvmrc`; `scripts/use-node.sh` selects it and verifies).

```bash
npm install
npm run build          # lint + compile + bundle + package
./build.sh             # same, after switching to the right Node version
```

The package is written to `sharepoint/solution/sp-new-library-uncheck-navigation.sppkg`
and is never committed.

```bash
./scripts/release              # bump version, build, commit, tag vX.Y.0.Z, publish GitHub Release
./scripts/release --no-publish # bump and build only; prints the manual steps
```

Write the `## [X.Y.0.Z]` entry in `CHANGELOG.md` **before** running it; the script
refuses to publish without one. The version is bumped across `package.json`,
`package-lock.json` and `config/package-solution.json` by `increment-version.sh`,
which aligns them to the highest version found and increments the patch. Solution
versions are `X.Y.0.Z`; npm versions are `X.Y.Z`.

### Local debugging

```bash
npm start
```

Then open the debug URL printed by the tool (it targets `_layouts/15/viewlsts.aspx`,
the Site contents page) and create a list or library. `config/serve.json` sets
`debug: true` so matches are logged to the console.

### End-to-end proof

`e2e/prove.js` is a headed Playwright script that:

1. opens Site contents on a site with the app installed, walks **New → List** and
   **New → Document library**, and asserts the checkbox is unchecked;
2. repeats on a control site without the app and asserts it is checked;
3. uses the status bar to switch the customizer **off**, asserts the checkbox is
   then left checked, switches it back **on** through the deep link, and asserts
   it is unchecked again;
4. reaches Site contents by client-side navigation from the home page and asserts
   the status bar appears there and disappears again on leaving.

Screenshots land in `docs/proof/`.

```bash
npx playwright install chromium
TEST_SITE=https://tenant.sharepoint.com/sites/WithApp \
CONTROL_SITE=https://tenant.sharepoint.com \
npm run e2e
```

Sign in once in the browser window; the profile persists in `~/.cache/pw-sp-profile`.
Behind an authenticating HTTP proxy, set `HTTPS_PROXY` and the script passes it
to the browser.

Last verified 2026-09-10 on tenant g53.sharepoint.com (site `/sites/UncheckNavTest`,
package version 1.0.0.3): `PROOF: PASS`.

| Scenario | List | Document library |
|----------|------|------------------|
| Site with app | unchecked | unchecked |
| Control site without app | checked | checked |
| Site with app, switched off in the panel | checked | (not exercised) |
| Site with app, switched back on via deep link | unchecked | (not exercised) |
| Status bar via client-side navigation | shown on Site contents, hidden elsewhere | |

The proof occasionally hits a 30-second locator timeout on the tenant; rerun
before treating a failure as real.

## Caveats

- This relies on the label text and structure of Microsoft's UI. If Microsoft
  changes the label, update the `labels` property; if they change the control
  type, the customizer may need a code change.
- The customizer only runs on modern pages. Classic list-creation pages are unaffected.
- The panel's own strings are English only (`loc/en-us.js`). On non-English
  tenants the `labels` property must also be set, or the customizer silently
  matches nothing; the status bar reports the setting, not whether a match has
  happened.
- The settings bar only appears on Site contents and only to users who can
  manage the web. Saving needs the same permission.

## Documentation map

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): how the customizer, the settings UI and the storage work.
- [docs/DECISIONS.md](docs/DECISIONS.md): architecture decision records, including rejected options and the evidence.
- [docs/OPERATIONS.md](docs/OPERATIONS.md): runbook for admin teams: deploy, upgrade, configure, troubleshoot.
- [docs/LESSONS.md](docs/LESSONS.md): lessons learned building and proving this on a real tenant.
- [CHANGELOG.md](CHANGELOG.md): release history.
- [AGENTS.md](AGENTS.md) / [CLAUDE.md](CLAUDE.md): instructions for AI coding agents working in this repo.
