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

A change applies immediately in the saving tab. Other tabs and users refresh on
navigation after their **60-second session cache** expires. Opening the panel
always reads fresh settings. Users without settings-list write permission get
a read-only panel through the deep link.

## Deploying

The package supports tenant-level deployment with site-local settings on modern
NoScript sites. Tenant-wide deployment registers the customizer for modern
sites. No settings list is created until an owner first saves a change.

```bash
m365 spo app add --filePath sharepoint/solution/sp-new-library-uncheck-navigation.sppkg --overwrite
m365 spo app deploy --id <catalog-app-id> --skipFeatureDeployment
```

The App Catalog equivalent is **Make this solution available to all sites**.
A per-site app installation is unnecessary. Deployment provisions a Tenant Wide
Extensions entry; verify there is exactly one matching entry after upgrades.
The [verification report](docs/LIST-SETTINGS-VERIFICATION.md) records the pilot and its cleanup.

Before migrating existing installations, preserve explicit disabled settings
in the hidden list and remove duplicate per-site registrations. Follow the
[migration and deployment runbook](docs/OPERATIONS.md).

### Configuration

The extension registration accepts optional `ClientSideComponentProperties`:

| Property  | Type       | Default | Purpose |
|-----------|------------|---------|---------|
| `enabled` | `boolean` | `true` | Legacy/default value, used only when no settings list exists. The hidden-list value takes precedence. |
| `labels`  | `string[]` | `["Show in site navigation", "Show list in site navigation", "Show library in site navigation"]` | Label texts to match (case-insensitive). **Required on non-English tenants**; add the localised label text. |
| `debug`   | `boolean`  | `false` | Log each match and setting change to the browser console. |

Defaults live in `sharepoint/assets/ClientSideInstance.xml` (tenant registration)
and `elements.xml` (per-site registration). For an existing tenant registration:

```bash
m365 spo tenant applicationcustomizer list --output json
m365 spo tenant applicationcustomizer set --id <tenant-extension-item-id> \
  --clientSideComponentProperties '{"debug":false,"labels":["Im Websitenavigation anzeigen"]}'
```

Include all registration properties you want to retain when replacing that JSON.

The panel stores `Enabled` in the hidden `Lists/JfdiUnavSettings` list, independently
of these properties. Existing `labels` and `debug` are not rewritten by Save.
The setting survives upgrades, reinstalls and removal of the app.

### Tenant-wide deployment

Tenant deployment and site settings are independent. Each web, including a
subsite, can have its own hidden list. Owners have Full Control; associated
Members and Visitors receive Read. Additional managers can be explicitly granted
list edit permission. Manage Web alone does not override the list ACL; initial
provisioning also requires Manage Lists and Manage Permissions.

An uncached page makes one asynchronous settings read. Valid results, including
an absent list, are cached per web and signed-in user for 60 seconds in
`sessionStorage`. Storage-disabled browsers fall back to REST. A failed read
leaves the SharePoint checkbox unchanged and is not cached. An existing empty
or malformed settings list is an error, not the enabled default.

See [the storage design](docs/HIDDEN-LIST-DESIGN.md) and
[measured storage performance](docs/HIDDEN-LIST-PERFORMANCE.md).

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

`e2e/prove.js` is a Playwright script (headed or `HEADLESS=1`) that:

1. opens Site contents on a site with the customizer registered, walks **New → List** and
   **New → Document library**, and asserts the checkbox is unchecked;
2. repeats on a control site without a registration and asserts it is checked;
3. uses the status bar to switch the customizer **off**, asserts the checkbox is
   then left checked, switches it back **on** through the deep link, and asserts
   it is unchecked again;
4. reaches Site contents by client-side navigation from the home page and asserts
   the status bar appears there and disappears again on leaving;
5. verifies the real library cache, expiry and failed-read recovery.

Screenshots land in `docs/proof/`.

```bash
npx playwright install chromium
# Set HEADLESS=1 when running without a display.
TEST_SITE=https://tenant.sharepoint.com/sites/WithApp \
CONTROL_SITE=https://tenant.sharepoint.com \
npm run e2e
```

Sign in once in the browser window; the profile persists in `~/.cache/pw-sp-profile`.
Behind an authenticating HTTP proxy, set `HTTPS_PROXY` and the script passes it
to the browser.

Last verified 2026-09-11 on tenant g53.sharepoint.com (site `/sites/UncheckNavTest`,
package version 1.0.0.4): `PROOF: PASS`. [Runtime evidence](docs/proof/runtime-list-settings.json)
records a 77 ms cold settings read, zero settings requests on the cached library
reload, and an 83 ms refresh after expiry. All 14 Jest tests and the production
build passed. Tenant-wide activation also passed on two sites without per-site
app installations; see the [verification report](docs/LIST-SETTINGS-VERIFICATION.md).

| Scenario | List | Document library |
|----------|------|------------------|
| Site using centrally deployed assets | unchecked | unchecked |
| Control site without registration | checked | checked |
| Site switched off in the panel, after reload | checked | checked |
| Site switched back on via deep link | unchecked | (not exercised) |
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
  manage the web. Saving also requires settings-list edit permission; first Save requires provisioning permissions.

## Documentation map

- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md): how the customizer, the settings UI and the storage work.
- [docs/LIST-SETTINGS-VERIFICATION.md](docs/LIST-SETTINGS-VERIFICATION.md): runtime, cache and tenant-wide proof with remaining validation limits.
- [docs/HIDDEN-LIST-DESIGN.md](docs/HIDDEN-LIST-DESIGN.md): tenant-wide settings store, cache contract and reproducible performance experiment.
- [docs/DECISIONS.md](docs/DECISIONS.md): architecture decision records, including rejected options and the evidence.
- [docs/OPERATIONS.md](docs/OPERATIONS.md): runbook for admin teams: deploy, upgrade, configure, troubleshoot.
- [docs/LESSONS.md](docs/LESSONS.md): lessons learned building and proving this on a real tenant.
- [CHANGELOG.md](CHANGELOG.md): release history.
- [AGENTS.md](AGENTS.md) / [CLAUDE.md](CLAUDE.md): instructions for AI coding agents working in this repo.
