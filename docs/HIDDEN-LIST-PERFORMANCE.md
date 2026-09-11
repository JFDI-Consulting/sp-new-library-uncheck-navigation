# Hidden-list performance evidence

**Conclusion:** a single, narrowly selected hidden-list read is fast enough for
this customizer on the tested tenant. Use an asynchronous fresh read on each
relevant navigation, with request deduplication within the page. Cross-page
caching is optional, not a prerequisite for acceptable measured performance.

Tested 2026-09-11 on `g53.sharepoint.com/sites/UncheckNavTest`, Node 22.17.1,
headless Chromium with the existing authenticated profile, no network throttling.
NoScript remained enabled and the property-bag exception remained disabled.

## Latest complete run

| Measurement | Samples | Median | p95 | Maximum |
|---|---:|---:|---:|---:|
| Warm uncached settings GET | 50 | 76 ms | 133 ms | 521 ms |
| Missing-list GET (404) | 20 | 62 ms | 110 ms | 194 ms |
| Settings GET during Site contents load | 20 | 135 ms | 139 ms | 166 ms |
| Settings GET during Documents library load | 20 | 132 ms | 141 ms | 148 ms |

Each successful settings response contained **42 bytes of JSON**, excluding
HTTP headers and transport overhead. The request selected only `Id,Enabled`,
filtered for the canonical key, and limited results to two for duplicate
detection. Timings include network, body transfer and JSON parsing; HTTP cache
was disabled for these reads using `cache: 'no-store'`. Server-side SharePoint
caches were not flushed.

The provisional predeclared p95 budgets were 300 ms for warm reads and 500 ms
for reads during page loading. Both passed. A 521 ms warm-read outlier remains
in the evidence; this is not a claim that every request completes below 300 ms.

## Observed page-load impact

Twenty AB/BA pairs compared ordinary Site contents loads with otherwise similar
loads issuing a settings GET at DOMContentLoaded:

| New button visible | Median | p95 |
|---|---:|---:|
| Baseline | 787 ms | 962 ms |
| With asynchronous settings GET | 803 ms | 914 ms |

The median of the paired differences was +12 ms. Settings were ready before
the New button became visible in **20/20** instrumented loads. The small
median difference and variable tail show no substantial consistent slowdown
in this sample; they do not establish a zero-cost request or a fleet-wide
performance guarantee.

The earlier run corroborated the read cost: warm p95 94 ms and Site contents
concurrent-read p95 162 ms. Its baseline/instrumented New-ready medians were
839/837 ms, with one instrumented 2.4-second page outlier. That outlier is
retained in the first-run evidence rather than dropped from the comparison.

These are post-DOMContentLoaded concurrent-request measurements, not an A/B
deployment of the production customizer. Library samples fetched settings while
the actual default Documents library loaded; they did not exercise creation
panels, disabled-state UI or client-side navigation.

## Optional cache

A synthetic `sessionStorage` read/JSON-parse/expiry check averaged about
0.0005 ms per operation across batches of 1,000; p95 of batch averages was
0.0008 ms, with no network requests. This is only a microbenchmark, not the
future extension's cache or a user-visible timing guarantee. It demonstrates
that cache lookup CPU time is negligible compared with a network request.
Caching still trades freshness for fewer requests.

## Functional and ACL checks

- Created a hidden generic list while NoScript was enabled.
- Broke list inheritance; verified exactly Owners Full Control, Members Read,
  Visitors Read. Removed the automatic individual provisioner grant.
- Verified the item inherited its list permissions in the complete run.
- Saved false from the authenticated browser: HTTP 204, false read back.
- Tried an update with a stale ETag: HTTP 412, as expected.
- Saved true with the current ETag: HTTP 204, true read back.
- Deleted the temporary list and verified its ID was absent.
- Verified site policy before/after was unchanged.

The browser user was a site collection administrator. These checks establish
the ACL configuration and administrator write/read behavior, **not** an actual
ordinary-member 403 or reader login proof. Member and visitor groups were empty
on this test site. Those identity tests remain a release requirement.

## Artifacts and repeatability

- [Design and provisioning plan](HIDDEN-LIST-DESIGN.md)
- [Runnable experiment](../e2e/prove-storage.js)
- [Complete run: all samples, ACLs, policy and verified cleanup](proof/hidden-list-performance.json)
- [First run, including the page outlier and cleanup recovery](proof/hidden-list-performance-run1.json)

```bash
. scripts/use-node.sh
node e2e/prove-storage.js
```

The first measured run completed its benchmarks but list deletion returned 412
because the DELETE lacked `IF-MATCH`. Its exact temporary list was then deleted
with `IF-MATCH:*` and absence verified. The harness was corrected; the complete
repeat run exited 0 and verified automatic cleanup. An earlier locator-only
attempt timed out before creating a list; inspection corrected the New button
selector to its actual HTML name.

No application code, deployed package, existing site setting, or tenant policy
was changed. This evidence supports selecting the hidden list as the storage
design. Full implementation still needs the panel, provisioning/repair,
migration, permission identities and tenant-wide end-to-end proof described in
the design document.
