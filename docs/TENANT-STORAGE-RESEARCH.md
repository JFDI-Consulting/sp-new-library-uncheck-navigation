# Tenant-wide settings storage research

Researched and tenant-tested 2026-09-11. This is a proposal, not a deployment
change or a replacement for ADR-004.

Follow-up: the preferred direction is now the [hidden per-site list design](HIDDEN-LIST-DESIGN.md),
with a reproducible browser performance experiment. The alternatives below are
retained as research history.

## Recommendation

There are two defensible alternatives to a custom list on every site:

1. **Web property bag with Microsoft's NoScript exception:** smallest operational
   and browser footprint, but conditional on the required permission model.
2. **A central configuration service:** can enforce Manage Web exactly without
   changing NoScript policy, at the cost of a backend and authentication overhead.

No equally lightweight built-in alternative was established that guarantees
both unchanged NoScript policy and writes by every user with Manage Web alone.

## 1. Web property bag with the documented exception

Microsoft now permits property-bag updates without enabling custom script.
An administrator can enable
`AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled` either for a site
collection or for the tenant. This changes property-bag policy, not
`DenyAddAndCustomizePages`. See [Microsoft's custom-script guidance](https://learn.microsoft.com/en-us/sharepoint/allow-or-prevent-custom-script)
and [Set-SPOSite](https://learn.microsoft.com/en-us/powershell/module/microsoft.online.sharepoint.powershell/set-sposite?view=sharepoint-ps).

Store a namespaced setting in `Web.AllProperties`, scoped to the current web.
It needs no list, library, app installation, or second extension registration.
Read the selected property through same-origin REST; write through CSOM on Save.
The tenant experiment below used CLI for Microsoft 365, whose property-bag
commands use CSOM. A browser implementation can use a small CSOM request rather
than introduce a large client library.

**Permission qualification:** PnP documents the tenant exception as allowing
users who had Add and Customize Pages before NoScript removed that permission.
Manage Web is a separate permission, so the exception cannot be represented as
a proven Manage-Web-only solution. The administrator test below does not settle
this distinction. See [Set-PnPTenant parameter documentation](https://pnp.github.io/powershell/cmdlets/Set-PnPTenant.html#-allowwebpropertybagupdatewhendenyaddandcustomizepagesisenabled).

**Expected read cost:** one small same-origin request per fresh settings read.
This is a request-count estimate, not a measured latency advantage over a list.
Unlike the current custom-action properties, arbitrary property-bag values are
not automatically supplied to the customizer by SPFx.

**Fit:** preferred if the supported writers are ordinary site owners with the
necessary underlying permissions and the administrator accepts the policy
exception. A tenant setting avoids provisioning the exception per site.

## 2. Central configuration service

Proposed design: an Entra-protected API stores a Boolean keyed by tenant ID,
site-collection ID and web ID in Azure Table Storage. A single central SharePoint
list is another backing-store choice; neither requires per-site lists.

On Save, the API authenticates the caller, resolves the target web within the
allowed tenant, obtains delegated SharePoint access using On-Behalf-Of, and checks
the caller's effective Manage Web permission on that web before writing.
Authorization must run on the server for each write; hiding the toggle or
accepting a browser-supplied permission result is insufficient. Read authorization
should likewise require access to the target web. Azure storage remains private
to the service.

These are architectural recommendations built on Microsoft's supported
[SPFx AadHttpClient integration](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient-enterpriseapi)
and [On-Behalf-Of authentication](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-on-behalf-of-flow),
not a service implemented or tested during this research.

**Expected read cost:** one browser API request per fresh read, plus any token
acquisition, preflight, backend authorization and storage requests. Regional
hosting and bounded caching can help; cold starts and authentication mean it
must not be advertised as faster than SharePoint REST without measurement.

**Fit:** the strongest design when Manage Web alone is a strict requirement.
Costs include hosting, Entra configuration, permission consent and ongoing
service maintenance. NoScript does not restrict the service's database.

## Other candidates and their limitations

- **JSON file in an existing library:** avoids a new list and can use normal
  document REST operations on NoScript sites. However, file edit permissions
  govern writes, not Manage Web. Inheriting Site Assets permissions could let
  members change the setting; granting an Owners group access does not cover
  everyone independently granted Manage Web. Requires provisioning and ACL
  management or a backend, so it is not a third equivalent fit. See
  [SharePoint file REST operations](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest).
- **Folder property bag:** the root-folder experiment below also failed under
  NoScript; do not treat changing the bag's location as a bypass.
- **Storage Entities / tenant properties:** managed in the app catalog. Manage
  Web on an arbitrary consuming site grants no corresponding app-catalog write
  authority. Site-specific keys alone do not fix this. See
  [Microsoft tenant properties guidance](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties?tabs=sprest).
- **Current custom-action properties:** `SettingsService.ts` requires an existing
  web- or site-scoped registration. Tenant-wide activation supplies no such
  per-web record; creating another registration also risks duplicate execution.
  See [tenant-wide extension deployment](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions).
- **Browser storage:** useful as a cache, not shared authoritative state across
  users, devices or browser profiles.

## Page-load design for either alternative

Fetch settings asynchronously once per full page load and refresh for relevant
client-side navigation. Do not fetch on every DOM observation or iframe scan.
Key any memory/browser cache by tenant, site and web; include user identity for
permission-sensitive responses. Invalidate it after Save.

A short cache lifetime reduces requests but delays changes in other tabs and
users. If every library load must observe the latest saved state, revalidate
there; a time-based cache cannot promise that freshness. Cache a genuinely absent
setting separately from an authorization, network or service error.

Keep the checkbox mutation inactive until the initial setting resolves, then
scan if enabled. Otherwise the current initial `true` value could untick a box
before a stored `false` arrives. A failed read should leave SharePoint's checkbox
alone and expose a useful error in the settings UI. Preserve lazy UI imports.

## Tenant evidence and restoration

Tenant: `g53.sharepoint.com`; isolated target: `/sites/UncheckNavTest`.
Actor: the existing tenant-administrator CLI connection. Node 22 selected through
`scripts/use-node.sh`; CLI for Microsoft 365 11.10.0.

| Probe | Observed result |
|---|---|
| Initial site state | `DenyAddAndCustomizePages = 2`, property-bag exception `false` |
| Initial tenant exception | `false` |
| Write temporary web property with original policy | Access denied, `0x80070005` |
| Write temporary root-folder property with original policy | Access denied |
| Enable exception on test site only through `m365 request` / admin CSOM | Successful; operation complete |
| Re-read site policy | NoScript still `2`; exception `true` |
| Write `jfdiUnavResearch20260911 = false` to web property bag | Successful |
| Read temporary key | Returned `false` |
| Remove temporary key | Successful; subsequent read returned no value |
| Restore and re-read site policy | NoScript `2`; exception `false` |
| Re-read tenant exception | Still `false` |

Raw successful-experiment responses are in
[the evidence file](proof/tenant-storage-research-2026-09-11.json).
The global exception was never changed. No package was redeployed and the
customizer's existing setting was untouched.

**Validation limits:** no separate Manage-Web-only, member or reader identity was
tested. No browser latency benchmark or Playwright UI proof was run. Application
source and package configuration were not changed, so build/unit/UI tests were
not needed for this research artifact. Before implementing option 1, prove a
non-admin writer's permission boundary and reader access; before releasing either
option, prove the toggle, cross-user persistence, tenant-wide activation without
per-site installation, navigation freshness and failure behavior in Playwright.
