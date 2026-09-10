# Operations runbook

For tenant admins and site admins deploying, configuring and supporting the
customizer. Commands use [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
(`m365`), signed in as a SharePoint admin. PnP PowerShell equivalents exist for
every step.

## What it does, in one paragraph

On sites where the app is installed, the "Show in site navigation" box in the
modern *Create list* and *Create document library* panels starts **unticked**.
Users can still tick it. Site owners can switch the behaviour off and on per
site from a status bar on **Site contents**. Nothing is installed on users'
machines; nothing runs outside the browser; no custom script is required.

## Requirements

- SharePoint Online, modern sites. Classic pages are unaffected.
- Tenant App Catalog (or a site collection App Catalog).
- To install on a site: site collection admin or site owner with permission to
  add apps. To use the switch: *Manage Web* (site owners).
- Tenant UI language: English out of the box. Other languages need the `labels`
  property set (see [Configuration](#configuration)).

## Install

```bash
# 1. Upload the package from the GitHub Release and deploy it (do NOT tick "make available to all sites")
m365 spo app add --filePath sp-new-library-uncheck-navigation.sppkg --overwrite
m365 spo app list                                  # note the ID (a GUID) for the next steps
m365 spo app deploy --id <app-id>

# 2. Add it to a site collection
m365 spo app install --id <app-id> --siteUrl https://tenant.sharepoint.com/sites/YourSite
```

Installation takes 10 to 60 seconds. Verify:

```bash
m365 spo app instance list --siteUrl https://tenant.sharepoint.com/sites/YourSite   # AppStatus 4 = Installed
m365 spo customaction list --webUrl https://tenant.sharepoint.com/sites/YourSite \
  --output json | grep -E '"(Title|ClientSideComponentId|ClientSideComponentProperties)"'
```

You should see one action titled `UncheckSiteNavigation` with component id
`d31c6f18-3a0d-462b-b677-c09314fbf3e6`. Then open Site contents as a site owner:
the status bar appears at the top, and **New → List** shows the box unticked.

## Upgrade

```bash
m365 spo app add --filePath sp-new-library-uncheck-navigation.sppkg --overwrite
m365 spo app deploy --id <app-id>
m365 spo app upgrade --id <app-id> --siteUrl https://tenant.sharepoint.com/sites/YourSite
```

Repeat the last line per site, or script it over `m365 spo site list`. An
upgrade **resets the custom action's properties to the package defaults**, so
any site that had switched the behaviour off, or configured `labels`, will be
back to defaults. Record per-site settings before a bulk upgrade if that
matters (see [Reading and setting the switch by script](#reading-and-setting-the-switch-by-script)).

Sites only pick up new assets when the solution version changes. Re-uploading
the same version and running `upgrade` is a no-op; uninstall and install instead.

## Remove

```bash
m365 spo app uninstall --id <app-id> --siteUrl https://tenant.sharepoint.com/sites/YourSite --force
```

Removing the app deletes its custom action, and the create panels return to
SharePoint's standard. Retracting from the catalog (`m365 spo app retract`)
removes it everywhere at once.

## Configuration

The custom action's `ClientSideComponentProperties` JSON holds three optional keys:

| Key | Default | Meaning |
|-----|---------|---------|
| `enabled` | `true` | `false` switches the behaviour off for the site. The settings panel writes this. |
| `labels` | English label texts | Accessible label texts to match, case-insensitive. Add the localised text on non-English tenants, e.g. `"Im Websitenavigation anzeigen"`. |
| `debug` | `false` | Log matches and setting changes to the browser console (F12). |

### Reading and setting the switch by script

```bash
SITE=https://tenant.sharepoint.com/sites/YourSite
ID=$(m365 spo customaction list --webUrl $SITE --output json \
     | jq -r '.[] | select(.ClientSideComponentId=="d31c6f18-3a0d-462b-b677-c09314fbf3e6") | .Id')

# read
m365 spo customaction get --webUrl $SITE --id $ID --output json | jq -r .ClientSideComponentProperties

# switch off, keeping other keys
m365 spo customaction set --webUrl $SITE --id $ID \
  --clientSideComponentProperties '{"enabled":false}'

# German tenant
m365 spo customaction set --webUrl $SITE --id $ID \
  --clientSideComponentProperties '{"labels":["Im Websitenavigation anzeigen","In der Websitenavigation anzeigen"]}'
```

`customaction set` replaces the whole JSON, so include every key you want to
keep. The in-page panel merges instead.

## What site owners see

- **Site contents**, top of page: a grey one-line bar: "This site is set to
  untick 'Show in site navigation' for new lists and libraries." (or the
  SharePoint-standard wording when off) with a **Change** link.
- Clicking **Change** opens a right-hand panel with one toggle and Save/Cancel.
- Deep link for help pages: `<site>/_layouts/15/viewlsts.aspx?jfdiUncheckNav=settings`.
- Members do not see the bar. If they open the deep link they get a read-only
  panel with a notice.

Changes apply immediately in the saving user's tab and within a few minutes for
everyone else (SharePoint caches page configuration).

## Troubleshooting

| Symptom | Likely cause | Check / fix |
|---------|--------------|-------------|
| Box is still ticked on a site with the app | Switched off in the panel; or non-English UI with default `labels`; or app not actually installed | Open Site contents as an owner and read the bar. Check `labels`. `m365 spo app instance list` shows `AppStatus` 4. |
| Box ticked in *Document library* only, briefly | Normal: the library panel needs the retry (up to 1.5 s). If it stays ticked, Microsoft may have changed the control | Set `debug:true`, open F12, look for `[UncheckSiteNavigationApplicationCustomizer]` lines saying "could not be unchecked". Report with a screenshot. |
| No status bar on Site contents | User lacks Manage Web; or page is not `viewlsts.aspx`; or Top placeholder failed | Confirm the user is a site owner. Hard-refresh. With `debug:true` the console logs placeholder waits. |
| Save shows "You need to be a site owner" | 403 from SharePoint | Grant Manage Web (Owners group) or have an owner save. |
| Save shows "custom action was not found" | App deployed tenant-wide, or action removed manually | This version must be installed per site. See [Tenant-wide deployment](#tenant-wide-deployment). |
| Save shows "existing properties are not valid JSON" | Someone set malformed `ClientSideComponentProperties` by script | Fix with `customaction set` using valid JSON. |
| `app install` leaves `AppStatus: 6` and no custom action | Install failed (package feature rejected) | Get the reason: `POST <site>/_api/web/GetAppInstanceById('<instance-id>')/GetErrorDetails` (instance id from `app instance list`, field `AppId`). Uninstall, fix, redeploy, install. |
| `app deploy` says `ResourceNotFoundException` | Catalog still processing the upload, or the package failed validation | Wait 10 s and retry; check `IsValidAppPackage`/`AppPackageErrorMessage` on the "Apps for SharePoint" list item. |
| Setting reverted after an upgrade | Expected: upgrade resets properties | Re-apply by script or in the panel. |

## Tenant-wide deployment

Not supported by the current version. The package can technically be deployed
to all sites (`skipFeatureDeployment`), and the customizer would run
everywhere, but the per-site switch has nowhere to store its value in that mode
and Save fails. If tenant-wide coverage is needed, ask for the hidden-list
storage described in [DECISIONS.md, ADR-004](DECISIONS.md#adr-004-per-site-install-not-tenant-wide-deployment-for-now).

## Security and compliance notes

- The package contains only client-side code (TypeScript compiled to
  JavaScript, React, Fluent UI). No server components, no external calls: the
  only network requests are to the site's own `/_api/web/UserCustomActions`,
  made with the signed-in user's permissions.
- `requiresCustomScript` is false; nothing needs custom script enabled.
- `isDomainIsolated` is false and no API permissions are requested, so the
  package needs no approval in the API access page.
- No telemetry, no storage outside SharePoint, no data leaves the tenant.
- Publisher: JFDI Consulting Ltd (MPN 2339010). Source, releases and changelog:
  https://github.com/JFDI-Consulting/sp-new-library-uncheck-navigation
