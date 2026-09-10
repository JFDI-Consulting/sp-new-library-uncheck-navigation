# Architecture

How the customizer works, end to end. For *why* it works this way, see
[DECISIONS.md](DECISIONS.md).

## Components

```
.sppkg
└── Feature e6602822-… (per-site)
    └── elements.xml: CustomAction Location=ClientSideExtension.ApplicationCustomizer
        ClientSideComponentId=d31c6f18-…  ClientSideComponentProperties={"debug":false}

Browser (every modern page on a site with the app)
└── UncheckSiteNavigationApplicationCustomizer (~16 KB)
    ├── MutationObservers on document.body and every same-origin iframe
    ├── label matcher + click-and-verify with backoff
    ├── SettingsService (REST, @microsoft/sp-http)
    └── on Site contents only, lazily:  chunk "uncheck-nav-settings" (~2 KB)
        └── SettingsUi → SettingsBar + SettingsPanel   (React 17 + Fluent UI 8, ~230 KB chunk)
```

## 1. Unticking the checkbox

SharePoint has no server-side default for "Show in site navigation", so the
customizer edits the rendered UI.

1. **`onInit`** normalises the `labels` property (default: the three English
   label variants), reads `enabled`, hooks `document.body` with a
   `MutationObserver` (childList, subtree, and the `checked`/`aria-checked`
   attributes) and schedules a scan.
2. **Scans are coalesced** to one per animation frame. A scan:
   - hooks any same-origin `<iframe>` it can reach (observer on the frame's
     `body`, plus a `load` listener so the frame is re-hooked after it
     navigates, because each navigation is a new document);
   - queries `input[type=checkbox], [role=switch], [role=checkbox]` in the
     document and recursively in hooked frames;
   - computes each control's accessible name (aria-label, aria-labelledby,
     `label[for]`, wrapping `label`, resolved against the control's own
     `ownerDocument`) and compares it, normalised, with the label list;
   - marks a matching control as processed (WeakSet) so it is only ever acted
     on once, then calls `_uncheck`.
3. **`_uncheck`** clicks the control if it is checked, then re-checks after
   150, 400, 800 and 1500 ms and clicks again if still checked. This matters
   for the document-library panel, whose Fluent UI 9 checkbox is in the DOM
   before React has attached its handlers, so the first click is dropped.
4. When `enabled` is false the scan returns before matching. The observers stay
   attached so a save in the settings panel takes effect without a reload.

**Why iframes:** the modern create experience is
`/_layouts/15/createlist.aspx?dlg=true&mode=sites` (list) or `mode=sitesDoclib`
(library), loaded into `<iframe name="createListFrame">`. SPFx extensions do not
load inside it, but it is same-origin, so the top-page customizer can reach its
document.

## 2. The per-site switch

### Storage

The switch is the `enabled` key inside the customizer's own
`ClientSideComponentProperties` on the web-scoped user custom action that the
feature provisioned. Reading costs nothing (SPFx hands the properties to the
customizer on every page). Writing is one REST MERGE:

```
GET  /_api/web/UserCustomActions?$filter=ClientSideComponentId eq guid'd31c6f18-…'
POST /_api/web/UserCustomActions('<Id>')   X-HTTP-Method: MERGE   IF-MATCH: *
     { "ClientSideComponentProperties": "{\"debug\":false,\"enabled\":false}" }
```

`SettingsService` merges into the existing JSON so `labels`/`debug` survive,
refuses to overwrite properties that are not a JSON object, falls back to
`/_api/site/UserCustomActions` for site-collection-scoped registrations, and
reports 403 as "you need to be a site owner". Manage Web is required by
SharePoint for the write; the UI additionally hides itself from users without it.

### UI lifecycle

Modern SharePoint navigates client-side without re-running `onInit`, so the UI
is driven by `_syncSettingsUi`, which is called from `onInit`,
`application.navigatedEvent` and `placeholderProvider.changedEvent`, and is
idempotent:

- not on `/_layouts/15/viewlsts.aspx` → unmount React and dispose the Top
  placeholder if present;
- on Site contents, user lacks Manage Web and there is no
  `?jfdiUncheckNav=settings` → do nothing;
- otherwise `tryCreateContent(PlaceholderName.Top)`; if the placeholder is not
  ready yet, return and wait for `changedEvent`;
- dynamically `import('./SettingsUi')`, render the bar (owners only) and the
  panel (opened immediately for the deep link), then strip only our query
  parameter with `history.replaceState` so a refresh does not reopen it.

The chunk split is deliberate: React and Fluent UI 8 would otherwise ride
along on every page view for every user on every site with the app.

### Panel

`SettingsPanel` is a Fluent UI 8 `Panel` with one `Toggle`, Save and Cancel.
Save is disabled until the value differs from the stored one; errors from
`SettingsService` are shown in a `MessageBar`; a mounted-ref guards against
state updates after the customizer disposes mid-save. After a successful save
the customizer updates its in-memory `enabled` and schedules a rescan.

## 3. Packaging and deployment

- `config/package-solution.json`: `includeClientSideAssets: true` (assets ship
  inside the `.sppkg`, served from the app catalog's ClientSideAssets library),
  `isDomainIsolated: false`, `skipFeatureDeployment` **not** set.
- Per-site install activates the feature, which creates the web-scoped custom
  action from `elements.xml`. Upgrading the app re-applies the feature and
  resets the action's properties to the package defaults.
- `ClientSideInstance.xml` is present for a future tenant-wide mode but is not
  used while `skipFeatureDeployment` is off.

## 4. Test hooks

Playwright locates the UI by `data-automation-id`:
`jfdi-unav-bar`, `jfdi-unav-change`, `jfdi-unav-panel`, `jfdi-unav-toggle`,
`jfdi-unav-save`. Console logging is behind the `debug` property and prefixed
`[UncheckSiteNavigationApplicationCustomizer]`.
