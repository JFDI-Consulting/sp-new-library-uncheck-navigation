# sp-new-library-uncheck-navigation

SharePoint Framework (SPFx) **Application Customizer** that defaults the
**"Show in site navigation"** checkbox to **unchecked** in the modern
*Create list* and *Create document library* panels.

SharePoint has no server-side setting for this default, so the customizer works
on the rendered page: it observes DOM mutations, finds the checkbox (or Fluent UI
toggle) whose accessible label matches a configured string, and clicks it if it
is checked. Each control is only targeted once, so a user who deliberately
re-ticks the box is not overridden.

Two details matter for how it does that:

- The modern create experience is hosted in a **same-origin iframe**
  (`/_layouts/15/createlist.aspx?dlg=true`) where SPFx extensions do not load.
  The customizer therefore also observes every same-origin iframe it can reach
  from the page and re-hooks them whenever they navigate.
- The document-library panel renders its checkbox before React has attached its
  handlers, so a single early click is dropped. After clicking, the customizer
  verifies the state and retries with a short backoff (150 ms to 1.5 s).

## Build

Requires Node 22.x (see `engines` in `package.json`).

```bash
npm install
npm run build          # lint + compile + bundle + package
```

The package is written to `sharepoint/solution/sp-new-library-uncheck-navigation.sppkg`.

## Deploy

1. Upload the `.sppkg` to the tenant App Catalog (or a site collection App Catalog).
2. Add the app to each site collection where you want the behaviour.
   Adding the app activates a feature that registers the Application Customizer
   custom action on the site.

To remove the behaviour, remove the app from the site.

## Configuration

The custom action accepts optional `ClientSideComponentProperties`:

| Property | Type       | Default                                                                                          | Purpose |
|----------|------------|--------------------------------------------------------------------------------------------------|---------|
| `labels` | `string[]` | `["Show in site navigation", "Show list in site navigation", "Show library in site navigation"]` | Label texts to match (case-insensitive). Add localised variants for non-English UI. |
| `debug`  | `boolean`  | `false`                                                                                          | Log each match to the browser console. |

Defaults live in `sharepoint/assets/elements.xml` and `ClientSideInstance.xml`.
To change them after deployment, update the custom action's
`ClientSideComponentProperties` with PnP PowerShell or CLI for Microsoft 365.

## Local debugging

```bash
npm start
```

Then open the debug URL printed by the tool (it targets `_layouts/15/viewlsts.aspx`,
the Site contents page) and create a list or library.

## End-to-end proof

`e2e/prove.js` is a headed Playwright script that opens Site contents on a site
with the app installed, walks **New → List** and **New → Document library**, and
asserts the checkbox is unchecked. It then repeats on a control site without the
app and asserts the checkbox is checked. Screenshots land in `docs/proof/`.

```bash
npx playwright install chromium
TEST_SITE=https://tenant.sharepoint.com/sites/WithApp \
CONTROL_SITE=https://tenant.sharepoint.com \
npm run e2e
```

Sign in once in the browser window; the profile persists in `~/.cache/pw-sp-profile`.
If you run behind an authenticating HTTP proxy, set `HTTPS_PROXY` and the script
passes it to the browser.

Last verified 2026-09-10 on tenant g53.sharepoint.com (site `/sites/UncheckNavTest`,
package version 1.0.0.1): `PROOF: PASS`.

| Scenario | List | Document library |
|----------|------|------------------|
| Site with app | unchecked | unchecked |
| Control site without app | checked | checked |

## Caveats

- This relies on the label text and structure of Microsoft's UI. If Microsoft
  changes the label, update the `labels` property; if they change the control
  type, the customizer may need a code change.
- The customizer only runs on modern pages. Classic list-creation pages are unaffected.
