# List settings implementation verification

Verified 2026-09-11 on `g53.sharepoint.com`, package **1.0.0.4**.

## Result

The implemented customizer persists its toggle in a hidden per-web list and
uses a 60-second session cache scoped by web and signed-in user. Tenant-level
component deployment works without per-site app installation. NoScript remains
enabled and the property-bag exception remains disabled.

The production build, lint and **14 Jest tests** passed. Tests cover the initial
settings gate, old-navigation responses, cache reuse/expiry/isolation, encoded
web paths, failed reads, ETag conflicts, denied provisioning and reads that
race with Save.

## Actual browser proof

[Runtime evidence](proof/runtime-list-settings.json) records:

| Check | Result |
|---|---|
| Uncached settings GET on Documents library | 76.633 ms, one request |
| Same-tab library reload within cache lifetime | **Zero settings requests** |
| Reload after forced cache expiry | 83.478 ms, one request |
| Enabled List and Document library creation | Both unchecked |
| Disabled setting saved, then full reload | Both creation checkboxes remain checked |
| Re-enabled through the deep link | List checkbox unchecked |
| Save transport | Hidden-list writes; no UserCustomActions writes |
| SPA navigation | Bar appears on Site contents and disappears on leaving |
| Injected settings 503 | Checkbox unchanged; normal reads recover with HTTP 200 |
| Control without registration | Both creation checkboxes remain checked |

The two request timings are individual browser network samples, not percentile
estimates or a universal page-load guarantee. Cache expiry was forced by expiring
the real sessionStorage entry. The earlier [storage benchmark](HIDDEN-LIST-PERFORMANCE.md)
provides repeated measurements; it measured a smaller query before implementation.
The application does not await settings in SPFx `onInit`, so page rendering
continues while it fetches. Checkbox mutation waits for a valid result.

The final entry bundle is **30,054 bytes**, up from roughly 16 KB before list
storage. React/Fluent remain lazy-loaded: the settings UI chunk is 3,411 bytes,
and its separate shared UI dependency chunk is 237,612 bytes. These are uncompressed
build artifact sizes, not per-navigation network transfer claims.

## Provisioning and permissions

The first live Save created `Lists/JfdiUnavSettings` on `/sites/UncheckNavTest`.
[Deployment evidence](proof/list-settings-deployment.json) verifies:

- Hidden, NoCrawl, no attachments, and version history enabled.
- Unique list permissions: Owners Full Control; Members and Visitors Read.
- The singleton item inherits list permissions.
- Indexed/unique required Title; required Boolean Enabled and application marker.
- Final stored value is true.
- NoScript is 2; the property-bag update exception is false.

The reader-repair command was run against these existing grants and retained
them without changing write permissions. Normal Save does not reset the ACL or
remove explicitly granted configuration administrators.

## Tenant-wide activation

[Tenant registration evidence](proof/tenant-list-runtime.json) records the
current bundle loading on both root and test sites, with zero direct custom
actions and no app installation on either site. Both showed exactly one settings
bar and unticked the list-creation checkbox. The root picked up registration on
the second attempt; the test site did on the first.

The package includes `ClientSideInstance.xml` and supports standard tenant
activation. Because this solution version had already been deployed during
iteration, redeployment did not re-provision the tenant row; the proof used
`m365 spo tenant applicationcustomizer add` after confirming zero matching rows.
An earlier attempt without the instance XML was rejected by that CLI command,
which led to correcting the final package.

Cleanup removed the exact temporary tenant row and restored one pilot web
registration on `/sites/UncheckNavTest`. The catalog retains the new package,
and the test site's settings list remains to preserve its setting. No global
registration remains. Follow [operations](OPERATIONS.md) for a production rollout
and migration of legacy disabled settings.

## Limits and test-harness corrections

The authenticated test identity was a site collection administrator. The ACL
was verified, but ordinary Member/Visitor reads and write denial were **not tested
with separate logins**. Interrupted-bootstrap recovery also remains unproven;
existing incomplete stores fail closed and require administrator repair.

The browser profile initially served an older manifest on library pages. The
proof now warms the authenticated profile and tests in a fresh cookies-only
context. Site contents and library toolbars expose different selectors; those
were corrected. An initial deferred-response timeout escaped the harness; errors
are now handled and the final complete run passed. These corrections did not
change the measured cache results reported above.
