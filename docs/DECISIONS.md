# Architecture decision records

Short records of the choices that shape this project, with the evidence behind
them. Dates are when the decision was proven, not guessed. If you plan to
reverse one, re-run the experiment first; the platform behaviour may have moved.

- [ADR-001: Change the default by editing the rendered UI](#adr-001-change-the-default-by-editing-the-rendered-ui)
- [ADR-002: Store the per-site switch in the custom action's properties](#adr-002-store-the-per-site-switch-in-the-custom-actions-properties)
- [ADR-003: Settings entry point on Site contents, not Site Settings](#adr-003-settings-entry-point-on-site-contents-not-site-settings)
- [ADR-004: Per-site install, not tenant-wide deployment (for now)](#adr-004-per-site-install-not-tenant-wide-deployment-for-now)
- [ADR-005: Lazy-load the settings UI](#adr-005-lazy-load-the-settings-ui)
- [ADR-006: Headed Playwright against a real tenant as the test of record](#adr-006-headed-playwright-against-a-real-tenant-as-the-test-of-record)

---

## ADR-001: Change the default by editing the rendered UI

**Status:** accepted, 2026-09-10.

**Context.** SharePoint Online has no setting, list template property, site
script action or API that changes the default of "Show in site navigation" in
the modern create panels. The only levers are after the fact (remove the
navigation node once the list exists) or in the browser.

**Decision.** An Application Customizer observes the DOM, finds the checkbox by
its accessible label, and clicks it once, with verification and retry. It also
observes same-origin iframes because the create panel lives in
`createlist.aspx` inside `createListFrame`, where SPFx does not load.

**Consequences.** Fragile by nature: it depends on Microsoft's label text and
control semantics. Mitigated by matching on accessible name rather than DOM
structure, a configurable `labels` list, touching each control only once, and
an end-to-end proof that is cheap to rerun. Non-English tenants must configure
`labels`.

**Rejected.** Post-creation clean-up (a timer job or event receiver removing the
nav node) changes user-visible state after the user has seen it and needs
infrastructure outside SharePoint. A ListView Command Set cannot reach the
create panel.

---

## ADR-002: Store the per-site switch in the custom action's properties

**Status:** accepted, 2026-09-10.

**Context.** The switch must be per site, writable by site owners from the
browser, and must not require custom script (default modern sites are
NoScript; `requiresCustomScript` is false).

**Decision.** `enabled` lives in the `ClientSideComponentProperties` of the
customizer's own web-scoped user custom action. Proven from the browser on
`g53.sharepoint.com/sites/UncheckNavTest`: `MERGE /_api/web/UserCustomActions('<id>')`
returned 204 and the value round-tripped; a tenant admin and, by permission
model, any Manage Web user can do it.

**Consequences.** Zero read cost (SPFx passes the properties in); one REST call
to write; admins can flip it with `m365 spo customaction set`. Downsides: an
app upgrade or reinstall resets it to the package default; it does not exist in
tenant-wide deployment mode (ADR-004); no optimistic concurrency (`IF-MATCH: *`),
so two owners saving in the same second could overwrite each other's `labels`
edit, judged acceptable for a setting changed a handful of times per site.

**Rejected.**
- *Web property bag*: the natural home, but writes are blocked on NoScript
  sites (`Access denied`), which is every modern site by default.
- *Hidden list*: works on NoScript, but adds provisioning, a read on every page
  load, and a visible artefact. Kept as the design for tenant-wide mode
  (ADR-004).
- *Uninstalling the app*: already the blunt off switch; not owner-friendly.

---

## ADR-003: Settings entry point on Site contents, not Site Settings

**Status:** accepted, 2026-09-10. Supersedes the original plan of a Site
Settings link.

**Context.** The obvious home for the switch is a link under *Site
Administration* on `/_layouts/15/settings.aspx`. Three independent experiments
on the g53 tenant showed it cannot be done on a standard modern site:

1. `m365 spo customaction add --location Microsoft.SharePoint.SiteSettings …`
   as tenant admin → HTTP 403, `Attempted to perform an unauthorized operation`.
   This is the NoScript block on URL custom actions.
2. The same action declared in the package's `elements.xml` → the app catalog
   rejected the package: `Custom action urls must start with "http:", "https:",
   "~appWebUrl" or "~remoteAppUrl". The url "~site/…" is not in the right format.`
3. With an accepted token (`~appWebUrl`, `~remoteAppUrl`) the package validated
   but **every install failed**, leaving the app at `AppStatus 6` with no custom
   actions at all. `GetErrorDetails` returned: `Feature definition with Id
   e6602822-… failed validation … The 'Location' attribute is invalid - The
   value 'Microsoft.SharePoint.SiteSettings' is invalid according to its
   datatype 'CustomActionLocations' - The Enumeration constraint failed.`

Separately, Site Settings is a classic page, so the customizer could not render
a panel there even if the link existed.

**Decision.** The customizer renders a one-line status bar with a **Change**
link in the Top placeholder on Site contents (`/_layouts/15/viewlsts.aspx`) for
users with Manage Web, and opens the panel when the URL carries
`?jfdiUncheckNav=settings`. Site contents was chosen because it is where lists
and libraries are created and because SharePoint preserves the query string
there (it appends `&view=14` but keeps ours).

**Consequences.** Discoverable for owners without any custom script. The bar is
an extra element on one page for owners only. A Site Settings link remains
possible on sites where custom script is enabled, added by an admin script, but
Microsoft auto-reverts custom script after 24 hours, so this is not offered.

**Rejected.** Query-string-only (not discoverable); a dedicated site page with a
web part (needs provisioning on every site); enabling custom script.

---

## ADR-004: Per-site install, not tenant-wide deployment (for now)

**Status:** accepted, 2026-09-10; explicitly deferred by the product owner.

**Context.** Tenant-wide deployment (`skipFeatureDeployment: true`, "Make this
solution available to all sites") would apply the customizer to every existing
and future modern site with no per-site action. But it also removes the per-site
custom action that ADR-002 stores the switch in. The settings panel would render
and then fail on Save. `SettingsService` already reports this case with a
message pointing at the Tenant Wide Extensions list.

**Decision.** Ship per-site install only. Keep `ClientSideInstance.xml` so the
package is ready for the other mode, but document that it must not be turned on
with the current storage.

**Planned path if tenant-wide is wanted.** Store the switch site-locally in a
hidden list (`Hidden=true`, one item per key, created by the first Save with
role inheritance broken so members read and owners write), read once per page
load and cached in `sessionStorage` per web for a few minutes; treat a missing
list as enabled. This works in both modes, survives reinstalls, and keeps
owners-only control. Keep `enabled` on the custom action as an admin override
for per-site mode, with the list winning when present.

**Rejected.** Site designs with `installSolution` (new sites only); scripted
bulk install (misses future sites); site-created custom action (would load the
customizer twice in tenant-wide mode).

---

## ADR-005: Lazy-load the settings UI

**Status:** accepted, 2026-09-10, following code review.

**Context.** An Application Customizer's bundle is fetched and parsed on every
modern page view by every user on every site with the app. Adding React 17 and
Fluent UI 8 statically took the bundle from ~10 KB to ~235 KB for a UI that only
owners see on one page.

**Decision.** The customizer keeps only `sp-core-library`, `sp-application-base`,
`sp-page-context`, `sp-http` and `SettingsService` as static imports. The UI is
loaded with `import('./SettingsUi')` after the route and permission checks pass.
Measured: main bundle 16 KB, settings chunk 3 KB, vendor chunk 238 KB loaded only
on Site contents for owners.

**Consequences.** One extra request on Site contents for owners. The
`_syncSettingsUi` flow has to guard against a placeholder being disposed while
the import is in flight; it does.

---

## ADR-006: Headed Playwright against a real tenant as the test of record

**Status:** accepted, 2026-09-10.

**Context.** The behaviour depends on Microsoft's live UI, cross-frame DOM
access, SharePoint's client-side router and real permission checks. None of
that is reproducible in jest or a mock.

**Decision.** `e2e/prove.js` runs a headed Chromium with a persistent signed-in
profile against a test site with the app and a control site without it, and
asserts the checkbox state in both, the panel's disable/enable round trip, and
the status bar's behaviour under client-side navigation. `PROOF: PASS` plus the
screenshots in `docs/proof/` is the acceptance evidence for a release.

**Consequences.** Needs a display, a tenant, and a human sign-in once. The
proof is subject to tenant flakiness (one 30 s locator timeout was observed in
an otherwise passing run); rerun before treating a failure as real. Jest remains
in the build for lint-level checks only.
