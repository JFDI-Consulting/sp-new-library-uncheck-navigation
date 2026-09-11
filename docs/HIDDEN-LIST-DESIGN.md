# Hidden per-site settings list

Implementation design, 2026-09-11. This supersedes the storage recommendation
in `TENANT-STORAGE-RESEARCH.md`. The application now uses this store and a
60-second session cache. The original isolated benchmark remains documented
below; actual runtime validation is recorded separately.

## Store and lifetime

Use one custom list per SharePoint **web**, at the stable site-relative path
`Lists/JfdiUnavSettings`. This gives subsites independent settings too. Keep
`Hidden=true`, `NoCrawl=true`, `OnQuickLaunch=false`, attachments and folders off,
and version history on. These flags reduce clutter; the permissions below
provide the security boundary.

The list has one canonical item:

| Field | Meaning |
|---|---|
| `Title` | Fixed key `configuration`; indexed and unique |
| `Enabled` | Boolean; whether to untick the creation checkbox |
| `JfdiUnavStoreId` | Application component ID; validates the configuration row |
| Built-in `Modified`, `Editor`, version history | Who changed it and when |

Do not encode configuration in the list description or in a document. Use a
Boolean field, and use the list's server-relative URL to locate it; changing the
display title then does not break reads. No API discovery or schema retrieval
is needed in the normal read path.

The authoritative state survives application upgrades, reinstalls and browser
cache eviction. App removal does not delete an independently created list.
Administrator cleanup is an explicit operation.

## Permissions

Break inheritance at list level; keep the item inheriting the list. The default
ACL for a conventional site is:

| Principal | List permission |
|---|---|
| Associated Owners group | Full Control |
| Associated Members group | Read |
| Associated Visitors group | Read |
| Explicit additional configuration-writer group, if required | Item read/edit with dependencies; no permission-management rights |
| Other site reader principals, where required | Read, explicitly provisioned |

Site collection administrators retain their normal administrative access.
Resolve groups and role definitions by IDs/type rather than English names.
Remove the provisioner's automatic individual Full Control assignment after
the Owners grant is established, unless it is explicitly intended.

Do not retain inherited Edit/Contribute grants and do not rely on hiding the
list or the `WriteSecurity` setting to enforce owners-only writes. SharePoint's
REST API enforces the list ACL even when someone bypasses the settings panel.
The supported mechanism is [list-level role inheritance and assignments](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/set-custom-permissions-on-a-list-by-using-the-rest-interface).

This deliberately defines writers as Owners or explicitly designated
configuration administrators. Manage Web alone is not an automatic entitlement
to write to a separately secured list. The existing UI permission test therefore
needs adapting: show configuration to relevant owners/managers, and check actual
list edit permission when opening the panel. Provisioning additionally requires
Manage Lists and Manage Permissions, normally supplied by Full Control.

Membership changes inside associated groups propagate naturally. New direct
web grants and new web groups do not propagate after inheritance is broken.
Provide an idempotent administrative repair command that adds the intended
reader principals as Read without restoring inherited write access. Do not grant
Everyone except external users as a shortcut: that broadens access and omits
guests. Users with only unique-library access also need an explicit reader
strategy; a settings-list 403 must not be treated as the default enabled state.

## Provision only when required

For tenant-wide coverage with minimal per-site work, an absent list means the
documented default, enabled. Create nothing during ordinary page loads. The
first authorized Save provisions and secures the list, then stores the choice.
Sites that never override the default need no list.

Bootstrap sequence:

1. An owner opens the panel; check the required provisioning permissions.
2. Create the list at the fixed URL, empty and marked as this application's
   settings store. If it already exists, validate its identity/schema; do not
   silently take over an unrelated list.
3. Break inheritance, assign Owners and reader roles, and remove undesired
   grants. Verify the resulting ACL before writing configuration.
4. Add the Boolean field and unique key constraint; create the canonical item.
5. Save the requested value and update the current page's state.

An interrupted bootstrap or an existing list with no valid canonical item is
an incomplete configuration, not an absent-list default. Report and repair it.
On concurrent bootstrap conflicts, re-read and validate the existing list/item
rather than creating a second one. Do not use a SharePoint REST batch as if it
were an atomic transaction.

A tenant administrator can also pre-provision and validate these lists using
CLI for Microsoft 365. Tenant-wide SPFx activation itself does not provision
per-site feature assets: [tenant-scoped deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-scoped-deployment).

## Normal read and Save paths

One request obtains the setting from the known list URL:

```http
GET <web>/_api/web/GetList(@list)/items
    ?@list='<server-relative-web>/Lists/JfdiUnavSettings'
    &$select=Id,Enabled,JfdiUnavStoreId
    &$filter=Title eq 'configuration'
    &$top=2
Accept: application/json;odata=nometadata
```

URL-encode parameters in implementation. The two-row limit detects duplicates
rather than silently selecting an arbitrary setting. Require exactly one item
and a Boolean. Only a confirmed absent-list response uses the default; do not
turn 403, 429, 5xx, invalid JSON or an incomplete list into enabled=true.

On panel opening, fetch the current item and its ETag. Save performs a MERGE
with that ETag, then updates the in-memory value and saving tab's cache. A 412
means another owner saved meanwhile: report the conflict and require reopening the panel to refresh before retrying. The proof tests this conflict behavior. See
[list REST operations and ETags](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest).

## Performance and freshness strategy

Start the read asynchronously when the customizer initializes. Do not hold up
SharePoint's page rendering. Keep checkbox mutation inactive while the setting
is unresolved; enable scanning only after a valid true/default result. A valid
false or a failed read leaves SharePoint's checkbox unchanged. Continue to lazy
load React/Fluent UI only for the settings interface.

Deduplicate requests within a page and reuse the result across DOM observations,
iframe hooks and retry timers. Revalidate once on relevant client-side
navigation, including library navigation; discard responses for an old web.

The implemented cache is **60 seconds per user and web**, stored in
`sessionStorage`. It includes valid values and confirmed absence; failed reads
are never cached. Every navigation consults the cache and fetches when expired.
There is no timer that interrupts an open creation panel. A tab left idle picks
up external changes on its next navigation, or immediately when its settings
panel is opened. Other tabs can use stale settings until expiry and navigation.

Panel opening always bypasses the cache and obtains a fresh ETag. Successful
Save updates the saving tab's cache. Storage-disabled browsers still work via
REST. In-flight reads are deduplicated and cannot overwrite a later save.

## Reproducible tenant experiment

Run:

```bash
. scripts/use-node.sh
node e2e/prove-storage.js
```

The script uses the existing M365 connection and signed-in Playwright profile.
It creates a uniquely named temporary list on `/sites/UncheckNavTest`, applies
and verifies the three group ACLs, tests browser-cookie reads/writes and ETag
conflicts, measures latency, and deletes only that temporary list in cleanup.
It verifies NoScript and the property-bag exception were not changed.

Measurements use `fetch(cache: 'no-store')`, include response transfer and JSON
parsing, and record every sample. The single first read occurs after provisioning
and a write round trip; it is not a claim of a cold SharePoint server cache.
The warm series is 50 sequential reads; absent-list handling is 20 reads.
Session-cache timings are 30 batches of 1,000 reads, reported per operation.

Navigation tests alternate AB/BA order across 20 pairs of full Site contents
loads in the same browser context. A is unchanged; B starts an asynchronous
settings read at DOMContentLoaded. The measured page milestone is visibility
of the New button, not full page rendering. This approximates early customizer
initialization; it does not deploy or measure the future production customizer.

Budgets chosen before measurement: uncached warm p95 <=300 ms, page-load read
p95 <=500 ms, cached read p95 <=1 ms. These are provisional engineering budgets,
not Microsoft service guarantees. Paired navigation timings must be interpreted
against their variation; an asynchronous call does not establish zero network
contention or a zero millisecond page-load cost.

An additional 20 full loads of the actual Documents library measure the same
settings request during library loading. Results and remaining limits are in
the [performance report](HIDDEN-LIST-PERFORMANCE.md).

## Migration and release boundary

Before switching the package to tenant-wide deployment, export existing per-site
custom-action settings. Open the panel and Save the existing value to populate the hidden list for
explicit overrides, especially `enabled=false`; otherwise the absent-list default would unexpectedly
re-enable those sites. Preserve existing labels/debug configuration separately.
Once list storage is active, it is authoritative for the toggle.

Remove duplicate per-site extension registrations as part of the controlled
deployment transition, so a site never runs two competing customizer instances.
The package enables `skipFeatureDeployment` and includes `ClientSideInstance.xml`;
standard tenant deployment provisions the Tenant Wide Extensions registration.
Retain component/solution IDs and `requiresCustomScript=false`.

The actual panel round trip, list/library creation defaults, cached library
reload, expiry, SPA navigation and failed-read recovery now pass in the
[runtime proof](proof/runtime-list-settings.json). Ordinary Member/Visitor logins
and interrupted-bootstrap recovery remain separate validation gaps. The original
storage microbenchmark is not a substitute for these implementation tests.
