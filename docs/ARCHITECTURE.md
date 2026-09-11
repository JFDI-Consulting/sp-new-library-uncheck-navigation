# Architecture

How the tenant-wide customizer works. For the decision record and storage
performance evidence, see [DECISIONS.md](DECISIONS.md) and
[HIDDEN-LIST-PERFORMANCE.md](HIDDEN-LIST-PERFORMANCE.md).

## Components

```
.sppkg (tenant-deployable; ClientSideInstance.xml included)
└── component d31c6f18-… made available tenant-wide
    └── SharePoint automatically creates one Tenant Wide Extensions row

Browser (every modern page matching that registration)
└── UncheckSiteNavigationApplicationCustomizer
    ├── asynchronous per-web settings reader + 60-second session cache
    ├── MutationObservers on document.body and same-origin iframes
    ├── label matcher + click-and-verify retry
    └── Site contents only, lazily: SettingsUi
        └── SettingsBar + SettingsPanel (React 17 + Fluent UI 8)

SharePoint web
└── Lists/JfdiUnavSettings (created only by first authorized Save)
    └── Title=configuration, JfdiUnavStoreId=<component id>, Enabled=<Boolean>
```

## 1. Activation and settings resolution

`skipFeatureDeployment` makes the component assets available to SharePoint
sites. The packaged `ClientSideInstance.xml` tells SharePoint to create the
Tenant Wide Extensions registration when the tenant administrator deploys the
package with tenant-wide deployment enabled. The extension then runs on modern
pages matching that registration.

On initialization and relevant client-side navigation, the customizer resolves
the current web's setting outside the mutation-observer scan loop. It
deduplicates in-flight work and uses a 60-second per-user/per-web
`sessionStorage` entry:

```
jfdi-unav:settings:v1:<encoded-lowercase-web-url>:<encoded-login-name>
```

On a cache miss it makes one narrow REST request to the known list path. A valid
canonical item supplies `Enabled`. A confirmed missing list uses the
registration's legacy/default `enabled` property. A permissions, throttling,
server, schema or duplicate-item failure fails closed: the customizer leaves
SharePoint's checkbox unchanged.

Responses from a prior web or disposed customizer are discarded. Opening the
settings panel always requests a fresh value. A successful Save updates the
writer's cached value. The implementation does not refresh just because a tab
gains focus; the intentionally bounded stale window closes on relevant
navigation after cache expiry, panel opening, or writer save.

## 2. Unticking the checkbox

SharePoint has no server-side default for **Show in site navigation**, so the
customizer changes the rendered UI after settings resolve.

1. `onInit` normalizes configured labels, hooks `document.body` with a
   `MutationObserver`, and starts the asynchronous settings read.
2. Once a valid enabled/default value is available, scans are coalesced to one
   per animation frame. A scan hooks accessible same-origin iframes, queries
   checkbox/switch controls, resolves their accessible names and compares them
   with normalized labels.
3. Each match is processed once. `_uncheck` clicks a checked control, then
   verifies with retries after 150, 400, 800 and 1500 ms. The retries handle the
   document-library panel, whose Fluent UI checkbox can render before its React
   handlers are ready.
4. A valid disabled setting or failed settings read leaves the native checkbox
   alone. Observers stay attached so a successful owner save can take effect in
   the current tab.

The modern list and library panels load in a same-origin
`<iframe name="createListFrame">`. SPFx does not load inside that frame, but the
top-page customizer can observe it and re-hook it after each iframe navigation.

## 3. Per-web settings store

The store is a custom list at the stable web-relative path
`Lists/JfdiUnavSettings`, hidden from navigation and search. It has one
canonical row with:

| Field | Purpose |
|---|---|
| `Title` | Fixed key `configuration` |
| `JfdiUnavStoreId` | Component-ID marker that prevents taking over an unrelated list |
| `Enabled` | Boolean switch |
| `Modified`, `Editor`, history | Audit and version information |

The normal read selects only the canonical row identity and Boolean, filters the
fixed key, and limits the result to two rows so a duplicate is detected rather
than selected arbitrarily.

The first authorized Save provisions the list. Provisioning validates any object
at that path, configures its schema, breaks inheritance, grants associated Owners
Full Control and associated Members/Visitors Read, then writes the canonical
row. The item inherits list permissions. Hidden visibility is for usability
only; the list ACL is the write boundary. The reader-repair script validates the
marker and canonical row before adding an intended reader as Read, without
restoring inherited write permissions.

Save uses the item's ETag. A stale ETag produces a conflict instead of replacing
a newer owner choice. List absence is the only storage condition that maps to the
enabled default; an incomplete or inaccessible list is reported and leaves
behavior inactive.

## 4. Settings UI lifecycle

Modern SharePoint navigation does not rerun `onInit`, so the customizer refreshes
state after navigation and follows placeholder changes. On pages other than
`/_layouts/15/viewlsts.aspx`, it unmounts the settings UI and releases the Top
placeholder. On Site contents it lazy-loads the React/Fluent UI chunk only after
the setting has resolved and the placeholder is available.

The bar is shown to a user with Manage Web. The panel is editable only after the
service confirms list-edit permission, or the two bootstrap permissions when the
list is absent; the deep link can render it read-only for other users. Opening
revalidates the Boolean; an unchanged Save can still provision the list when a
legacy false override must be migrated. Save updates in-memory state, invalidates
the cache and schedules a rescan. The customizer strips only its own query
parameter with `history.replaceState`.

## 5. Packaging and deployment

- `config/package-solution.json` has `includeClientSideAssets: true`,
  `isDomainIsolated: false` and `skipFeatureDeployment: true`.
- `ClientSideInstance.xml` is included in package assets. This is the supported
  automatic registration mechanism for a tenant-wide SPFx extension.
- Tenant administrators inventory the generated Tenant Wide Extensions row after
  deployment and retain exactly one matching component entry. Package updates can
  create duplicates, which must be removed by their exact list-item IDs.
- Tenant Wide Extensions supports template filters, not a specific SiteId/WebId.
- The settings list is independent of app installation and upgrades. Removing an
  application registration does not delete a settings list.

## 6. Test hooks

Playwright locates the UI by `data-automation-id`:
`jfdi-unav-bar`, `jfdi-unav-change`, `jfdi-unav-panel`,
`jfdi-unav-toggle`, and `jfdi-unav-save`. Console logging is behind the
`debug` registration property and prefixed
`[UncheckSiteNavigationApplicationCustomizer]`.
