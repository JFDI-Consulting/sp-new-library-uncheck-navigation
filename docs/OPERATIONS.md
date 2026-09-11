# Operations runbook

For tenant administrators and site owners deploying and operating the customizer.
Commands use [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
(`m365`), signed in as a SharePoint administrator.

## What it does

The centrally deployed customizer unticks **Show in site navigation** in modern
Create list and Create document library panels. Site owners can override that
default per web on Site contents. The setting is stored in a hidden SharePoint
list, so it works on NoScript sites and survives app updates and reinstalls.

## Requirements

- SharePoint Online and modern pages. Classic pages are unaffected.
- Tenant App Catalog and a SharePoint administrator for deployment and tenant
  registration.
- No custom script or API permission is required.
- The first user to save an override must be able to create and secure a list.
  An associated Owners-group member with Full Control normally has this
  permission. Manage Web alone does not necessarily permit list creation or
  permission changes.

## Migrate existing disabled sites

Before the tenant-wide deployment, preserve every existing per-site
custom-action `enabled:false` override. While a web has no settings list, that
legacy property remains its default. On every explicitly disabled site, open the
settings panel and Save its current false value. The panel provisions the hidden
list even when the value has not changed.

This migration must finish before deploying the tenant-wide package. Otherwise,
a site whose legacy action is removed can temporarily fall back to enabled.
Preserve any `labels` and `debug` configuration separately: those become the
properties in the tenant-wide registration.

## Deploy and automatically activate tenant-wide

The package has `skipFeatureDeployment: true` and includes
`ClientSideInstance.xml`. SharePoint uses that manifest to create the Tenant Wide
Extensions registration automatically when the package is deployed with
`--skipFeatureDeployment`. This is the standard SPFx extension deployment path.

Before changing the package, inventory both the legacy actions and the existing
tenant registration rows. Save the exact IDs; do not remove by component ID when
duplicates may exist.

```bash
PACKAGE=sharepoint/solution/sp-new-library-uncheck-navigation.sppkg
COMPONENT=d31c6f18-3a0d-462b-b677-c09314fbf3e6
TEST_SITE=https://g53.sharepoint.com/sites/UncheckNavTest

m365 spo applicationcustomizer list --webUrl "$TEST_SITE" --scope All --output json \
  > /tmp/unav-test-actions-before.json
m365 spo tenant applicationcustomizer list --output json \
  > /tmp/unav-tenant-actions-before.json
m365 spo tenant applicationcustomizer list --output json |
  jq --arg component "$COMPONENT" \
    '[.[] | select(.TenantWideExtensionComponentId == $component)]'

# Upload the final package only after the old test/pilot action is removed.
m365 spo app add --filePath "$PACKAGE" --overwrite
m365 spo app list --output json       # identify and inspect the exact app ID
m365 spo app deploy --id <app-id> --skipFeatureDeployment
```

If redeploying an already deployed solution version leaves no entry, confirm the
catalog reports `ContainsTenantWideExtension: true`, then register it explicitly:

```bash
m365 spo tenant applicationcustomizer add --title UncheckSiteNavigation \
  --clientSideComponentId "$COMPONENT" \
  --clientSideComponentProperties '{"debug":false}'
```

Run that only after confirming there are zero matching entries. This was needed
in the development tenant when the same version was first deployed without its
instance XML and then corrected; the CLI rejected registration until the XML
was restored to the package.

The final list query must show one enabled row with the component ID, location
`ClientSideExtension.ApplicationCustomizer`, and the expected properties. The
first Tenant Wide Extensions registration can take up to 20 minutes to apply.

```bash
m365 spo tenant applicationcustomizer list --output json |
  jq --arg component "$COMPONENT" \
    '[.[] | select(.TenantWideExtensionComponentId == $component)]'
```

Package updates can create duplicate Tenant Wide Extensions entries. After every
deploy, inventory rows again and retain exactly one intended enabled entry. If an
extra row was created, resolve its list-item `Id` from the inventory and remove
only that row:

```bash
m365 spo tenant applicationcustomizer remove --id <extra-tenant-extension-list-item-id> --force
```

Do not use `WebTemplate`, `Disabled`, or `HostProperties` to make a site-specific
test registration. Tenant Wide Extensions cannot target a SiteId or WebId;
`WebTemplate` applies to every matching web, `Disabled` disables the whole row,
and placeholder preallocation would leave unused space on pages other than Site
contents.

## Per-web setting, permissions and cache

The first authorized Save creates `Lists/JfdiUnavSettings` in that web. Normal
page loads never create a list. When the list is confirmed absent, the customizer
uses the registration's legacy/default `enabled` value, which defaults to true.

| Item | Value |
|---|---|
| List visibility | Hidden, not crawled, not in Quick Launch |
| Canonical row | `Title=configuration` |
| Identity marker | `JfdiUnavStoreId=d31c6f18-3a0d-462b-b677-c09314fbf3e6` |
| Setting | `Enabled` Boolean |
| History | SharePoint list version history |

The list breaks inherited permissions. Associated Owners receive Full Control;
associated Members and Visitors receive Read; the item inherits those
permissions. Other intended reader principals need explicit Read. Hiding a list,
`WriteSecurity`, or an Everyone-except-external-users grant does not provide the
required owners-only write boundary.

The customizer reads asynchronously, deduplicates requests within a page, and
keeps a 60-second per-user/per-web `sessionStorage` cache. Its cache key starts
`jfdi-unav:settings:v1:` and contains the encoded lower-case web URL and encoded login name.
It revalidates on relevant navigation and when the settings panel opens; it does
not refresh merely because the tab gains focus. A successful Save updates the
writer's cache. A confirmed absent list uses the default. A 403,
malformed/duplicate row, throttling or server failure fails closed: SharePoint's
checkbox remains unchanged and Site contents shows an error.

When a web acquires direct reader grants or new web groups after inheritance has
been broken, repair intended reader access without restoring inherited writes:

```bash
node scripts/grant-settings-readers.js \
  https://g53.sharepoint.com/sites/UncheckNavTest \
  <principal-id> [<principal-id>...]
```

With no principal IDs, the script targets associated Members and Visitors. It
validates the hidden list, canonical row and identity marker before adding Read;
it does not alter Owners or writer assignments.

## Configuration

Tenant registration properties configure labels and diagnostics. They are not the
per-web setting store.

| Property | Default | Meaning |
|---|---:|---|
| `labels` | English label texts | Accessible names to match, case-insensitive |
| `debug` | `false` | Browser-console diagnostics |
| `enabled` | `true` | Default only while this web has no settings list |

For a non-English tenant, update the exact tenant extension row and preserve
every required property.

```bash
m365 spo tenant applicationcustomizer set --id <tenant-extension-list-item-id> \
  --clientSideComponentProperties \
  '{"labels":["Im Websitenavigation anzeigen","In der Websitenavigation anzeigen"],"debug":false}'
```

## User experience

- On **Site contents**, a user with Manage Web sees the status bar and
  **Change**. The panel becomes editable only when that user can edit the
  settings list (or can provision it on a web where it is absent).
- Opening the panel re-reads the current value. Save uses an ETag, so a
  concurrent owner change is reported instead of silently overwritten.
- The deep link is
  `<site>/_layouts/15/viewlsts.aspx?jfdiUncheckNav=settings`.
- Users without list-write permission get a read-only panel.

## Troubleshooting

| Symptom | Check / fix |
|---|---|
| Box remains ticked | Confirm one intended registration/action, the setting, labels, and browser console with `debug:true`. |
| Box is unchanged and settings show an error | Validate list path, schema, marker, canonical row and list ACL. Do not take over or delete an unfamiliar list at the fixed path. |
| Owner cannot save first override | Use an account that can create the list and manage permissions, or pre-provision it. |
| Save reports a concurrent change | Reopen the panel, review the fresh value and save again. |
| Reader gets a settings error | Add the intended principal with the reader-repair script; a list 403 deliberately fails closed. |
| Extension runs twice | Remove the legacy per-site or pilot action. Keep exactly one tenant-wide registration. |
| Tenant activation is not immediate | The first row can take up to 20 minutes. Recheck its exact list-item ID. |

The tenant proof verified ACL construction, administrator read/write, ETag
conflicts, NoScript compatibility and cleanup. It did **not** include a login as
an ordinary Member or Visitor because those groups were empty on the test site.
Verify ordinary-reader access and writer denial before production rollout.

## Remove

Remove the tenant registration by exact list-item ID. This stops activation but
intentionally retains independently created settings lists and their history.
Delete a settings list only after identifying its web, backing up required
history, and confirming it should return to the registration default.

```bash
m365 spo tenant applicationcustomizer remove --id <tenant-extension-list-item-id> --force
```

Retract the package only after the tenant registration is gone; retraction
affects every consumer of the package.

## Security and compliance notes

- Configuration requests stay within the current SharePoint web; no data leaves
  the tenant.
- `requiresCustomScript` remains false and the setting store works on NoScript
  sites.
- The package has no API permissions and no telemetry.
- Publisher: JFDI Consulting Ltd (MPN 2339010). Source, releases and changelog:
  https://github.com/JFDI-Consulting/sp-new-library-uncheck-navigation
