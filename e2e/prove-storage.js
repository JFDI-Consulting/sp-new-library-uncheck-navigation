// Isolated tenant storage experiment. Creates and deletes only its own temporary
// hidden list. Does not deploy or modify the customizer. Requires m365 login and
// the signed-in Playwright profile used by prove.js. Run with Node 22.
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const assert = require('assert/strict');

const site = process.env.TEST_SITE || 'https://g53.sharepoint.com/sites/UncheckNavTest';
if (new URL(site).hostname !== 'g53.sharepoint.com') throw new Error('This proof is scoped to g53.sharepoint.com');
const title = `JfdiUnavStorageProof${Date.now()}`;
const output = process.env.STORAGE_PROOF_OUTPUT || path.join(__dirname, '../docs/proof/hidden-list-performance.json');
const result = { timestamp: new Date().toISOString(), site, title,
  environment: { node: process.version, browser: 'Chromium, headless', network: 'existing host network/proxy; no throttling' },
  budgets: { warmReadP95Ms: 300, pageReadP95Ms: 500, cacheP95Ms: 1 },
  limits: ['Only administrator identity authenticated; ACL verification is not non-admin write-denial proof.',
    'Instrumented REST lookup; production customizer remains unchanged.',
    'Single tenant, client location and run; no fleet-wide latency guarantee.'] };
let listId;
let context;

function cli(args) {
  const raw = execFileSync('m365', [...args, '--output', 'json'], { encoding: 'utf8', timeout: 60000 });
  return raw.trim() ? JSON.parse(raw) : undefined;
}
function api(relative, method = 'get', body) {
  const args = ['request', '--url', `${site}/_api/${relative}`, '--method', method,
    '--accept', 'application/json;odata=nometadata'];
  if (method === 'delete') args.push('--if-match', '*');
  if (body !== undefined) args.push('--body', JSON.stringify(body), '--content-type', 'application/json;odata=nometadata');
  return cli(args);
}
function stats(samples, key = 'ms') {
  const values = samples.map(s => typeof s === 'number' ? s : s[key]).sort((a, b) => a - b);
  const p = n => values[Math.max(0, Math.ceil(n * values.length) - 1)];
  return { n: values.length, min: values[0], median: p(0.5), p95: p(0.95), max: values.at(-1) };
}
function proxyFromEnv() {
  const raw = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY;
  if (!raw) return undefined;
  const u = new URL(raw);
  return { server: `${u.protocol}//${u.hostname}:${u.port}`, username: decodeURIComponent(u.username),
    password: decodeURIComponent(u.password), bypass: process.env.NO_PROXY };
}

async function sampleReads(page, url, count, expectedStatus) {
  return page.evaluate(async ({ url, count, expectedStatus }) => {
    const samples = [];
    for (let i = 0; i < count; i++) {
      const start = performance.now();
      const response = await fetch(url, { cache: 'no-store', headers: { Accept: 'application/json;odata=nometadata' } });
      const text = await response.text();
      const data = JSON.parse(text);
      if (response.status !== expectedStatus) throw new Error(`Unexpected ${response.status}: ${text}`);
      if (response.ok && (data.value.length !== 1 || typeof data.value[0].Enabled !== 'boolean')) throw new Error('Invalid setting');
      samples.push({ ms: performance.now() - start, status: response.status, bytes: new TextEncoder().encode(text).length });
    }
    return samples;
  }, { url, count, expectedStatus });
}

(async () => {
  try {
    result.policyBefore = cli(['spo', 'tenant', 'site', 'get', '--url', site, '--query',
      '{NoScript:DenyAddAndCustomizePages,PropertyBagException:AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled}']);
    assert.equal(result.policyBefore.NoScript, 2);
    context = await chromium.launchPersistentContext(path.join(os.homedir(), '.cache/pw-sp-profile'), {
      headless: true, proxy: proxyFromEnv(), viewport: { width: 1400, height: 900 }
    });
    const page = context.pages()[0] || await context.newPage();
    await page.goto(`${site}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded', timeout: 45000 });
    await page.locator('button[name="New"]').first().waitFor({ timeout: 30000 });
    const groups = api('web?$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup');
    const roles = api('web/roledefinitions?$select=Id,RoleTypeKind,BasePermissions').value;
    const ownerId = groups.AssociatedOwnerGroup.Id;
    const readerIds = [groups.AssociatedMemberGroup.Id, groups.AssociatedVisitorGroup.Id];
    assert(ownerId && readerIds.every(Boolean), 'Proof needs standard associated groups');
    const fullControl = roles.find(r => r.RoleTypeKind === 5).Id;
    const read = roles.find(r => r.RoleTypeKind === 2).Id;
    const provisionStart = performance.now();
    listId = api('web/lists', 'post', { Title: title, Description: 'Temporary isolated Uncheck Navigation storage benchmark',
      BaseTemplate: 100, Hidden: true, NoCrawl: true, OnQuickLaunch: false, EnableAttachments: false, EnableVersioning: true }).Id;
    assert.match(listId, /^[0-9a-f-]{36}$/i);
    const list = `web/lists(guid'${listId}')`;
    api(`${list}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`, 'post');
    api(`${list}/roleassignments/addroleassignment(principalid=${ownerId},roledefid=${fullControl})`, 'post');
    for (const principal of readerIds) api(`${list}/roleassignments/addroleassignment(principalid=${principal},roledefid=${read})`, 'post');
    // breakroleinheritance gives the caller Full Control. Remove that explicit
    // grant after Owners are assigned; site collection admins still bypass ACLs.
    for (const assignment of api(`${list}/roleassignments?$select=PrincipalId`).value) {
      if (![ownerId, ...readerIds].includes(assignment.PrincipalId)) api(`${list}/roleassignments/getbyprincipalid(${assignment.PrincipalId})`, 'delete');
    }
    api(`${list}/fields`, 'post', { Title: 'Enabled', FieldTypeKind: 8 });
    const item = api(`${list}/items`, 'post', { Title: 'configuration', Enabled: true });
    result.itemInheritsList = !api(`${list}/items(${item.Id})?$select=HasUniqueRoleAssignments`).HasUniqueRoleAssignments;
    assert.equal(result.itemInheritsList, true);
    result.provisioningCliWallMs = performance.now() - provisionStart;
    const metadata = api(`${list}?$select=Id,Hidden,NoCrawl,HasUniqueRoleAssignments,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder`);
    result.list = metadata;
    result.acl = api(`${list}/roleassignments?$select=PrincipalId,RoleDefinitionBindings/Id,RoleDefinitionBindings/Name,RoleDefinitionBindings/BasePermissions&$expand=RoleDefinitionBindings`).value;
    assert.equal(metadata.Hidden, true);
    assert.equal(metadata.HasUniqueRoleAssignments, true);
    assert.equal(result.acl.length, 3);
    for (const assignment of result.acl) assert.deepEqual(assignment.RoleDefinitionBindings.map(r => r.Id),
      [assignment.PrincipalId === ownerId ? fullControl : read]);
    const query = `web/GetList(@list)/items?@list=${encodeURIComponent("'" + metadata.RootFolder.ServerRelativeUrl + "'")}&$select=Id,Enabled&$filter=Title%20eq%20'configuration'&$top=2`;
    const endpoint = `${site}/_api/${query}`;
    const missing = endpoint.replace(title, title + 'Missing');
    result.endpoint = endpoint;
    // Browser-cookie read/write proof, with exact ETag rather than IF-MATCH:*.
    result.roundTrip = await page.evaluate(async ({ endpoint, itemUrl }) => {
      const digestResponse = await fetch(`${location.origin}${new URL(itemUrl).pathname.split('/_api/')[0]}/_api/contextinfo`,
        { method: 'POST', headers: { Accept: 'application/json;odata=nometadata' } });
      if (!digestResponse.ok) throw new Error('Digest failed');
      const digest = (await digestResponse.json()).FormDigestValue;
      async function save(enabled, etag) {
        const r = await fetch(itemUrl, { method: 'POST', headers: { Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata', 'X-RequestDigest': digest, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': etag },
          body: JSON.stringify({ Enabled: enabled }) });
        return r.status;
      }
      async function read() { const r = await fetch(itemUrl, { cache: 'no-store', headers: { Accept: 'application/json;odata=minimalmetadata' } });
        if (!r.ok) throw new Error('Read failed'); return { etag: r.headers.get('etag'), data: await r.json() }; }
      const initial = await read();
      if (!initial.etag) throw new Error('Missing ETag');
      const disabledStatus = await save(false, initial.etag);
      const disabled = await read();
      const conflictStatus = await save(true, initial.etag);
      const enabledStatus = await save(true, disabled.etag);
      const final = await read();
      return { disabledStatus, disabled: disabled.data.Enabled, conflictStatus, enabledStatus, enabled: final.data.Enabled };
    }, { endpoint, itemUrl: `${site}/_api/${list}/items(${item.Id})` });
    assert.deepEqual(result.roundTrip, { disabledStatus: 204, disabled: false, conflictStatus: 412, enabledStatus: 204, enabled: true });
    result.firstRead = (await sampleReads(page, endpoint, 1, 200))[0];
    result.warmReads = await sampleReads(page, endpoint, 50, 200);
    result.missingReads = await sampleReads(page, missing, 20, 404);
    result.cache = await page.evaluate(({ endpoint }) => {
      const key = `jfdi-unav-proof:${endpoint}`;
      const value = JSON.stringify({ enabled: true, expires: Date.now() + 60000 });
      const old = sessionStorage.getItem(key);
      const batches = [];
      try {
        sessionStorage.setItem(key, value);
        for (let b = 0; b < 30; b++) { const start = performance.now();
          for (let i = 0; i < 1000; i++) { const data = JSON.parse(sessionStorage.getItem(key));
            if (!data.enabled || data.expires < Date.now()) throw new Error('Cache invalid'); }
          batches.push((performance.now() - start) / 1000); }
      } finally { if (old === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, old); }
      return { perReadBatchAveragesMs: batches, networkRequests: 0 };
    }, { endpoint });
    console.log('Direct reads, missing-list handling, ETag conflict and cache measured. Starting paired navigation samples.');
    result.navigations = [];
    // Alternating AB / BA order reduces systematic warmup/order bias. Each pair
    // shares a browser context. Request cache:no-store prevents HTTP cache wins.
    for (let pair = 0; pair < 20; pair++) {
      for (const enabled of pair % 2 ? [true, false] : [false, true]) {
        const p = await context.newPage();
        try {
          await p.addInitScript(({ endpoint, enabled }) => {
            document.addEventListener('DOMContentLoaded', () => {
              window.__storageProof = { enabled, dclMs: performance.now() };
              if (enabled) {
                const start = performance.now();
                window.__storageProof.pending = fetch(endpoint, { cache: 'no-store', headers: { Accept: 'application/json;odata=nometadata' } })
                  .then(async r => { const text = await r.text(); const data = JSON.parse(text);
                    if (!r.ok || data.value.length !== 1 || data.value[0].Enabled !== true) throw new Error(`Bad settings: ${r.status}`);
                    window.__storageProof.read = { ms: performance.now() - start, readyMs: performance.now(), status: r.status, bytes: new TextEncoder().encode(text).length };
                  }).catch(e => { window.__storageProof.error = e.message; });
              }
            }, { once: true });
          }, { endpoint, enabled });
          await p.goto(`${site}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded', timeout: 45000 });
          await p.locator('button[name="New"]').first().waitFor({ timeout: 30000 });
          const newReadyMs = await p.evaluate(() => performance.now());
          const data = await p.evaluate(async () => { await window.__storageProof.pending;
            const { pending, ...data } = window.__storageProof; return data; });
          if (data.error) throw new Error(data.error);
          result.navigations.push({ pair, ...data, newReadyMs });
        } finally { await p.close(); }
      }
      if ((pair + 1) % 5 === 0) console.log(`Navigation pairs: ${pair + 1}/20`);
    }
    const withRead = result.navigations.filter(n => n.enabled);
    const baseline = result.navigations.filter(n => !n.enabled);
    const documentLibrary = api('web/lists?$select=DefaultViewUrl&$filter=BaseTemplate%20eq%20101%20and%20Hidden%20eq%20false').value
      .find(l => l.DefaultViewUrl.includes('/Shared Documents/'));
    assert(documentLibrary, 'Proof requires the default Documents library');
    result.libraryUrl = new URL(documentLibrary.DefaultViewUrl, site).href;
    result.libraryReads = [];
    for (let i = 0; i < 20; i++) {
      const p = await context.newPage();
      try {
        await p.addInitScript(endpoint => {
          document.addEventListener('DOMContentLoaded', () => {
            const start = performance.now();
            window.__storageLibraryProof = fetch(endpoint, { cache: 'no-store', headers: { Accept: 'application/json;odata=nometadata' } })
              .then(async response => {
                const text = await response.text(); const data = JSON.parse(text);
                if (!response.ok || data.value.length !== 1 || data.value[0].Enabled !== true) throw new Error('Library settings read failed');
                return { ms: performance.now() - start, status: response.status, bytes: new TextEncoder().encode(text).length };
              });
          }, { once: true });
        }, endpoint);
        await p.goto(result.libraryUrl, { waitUntil: 'domcontentloaded', timeout: 45000 });
        result.libraryReads.push(await p.evaluate(() => window.__storageLibraryProof));
        await p.waitForLoadState('load', { timeout: 30000 });
      } finally { await p.close(); }
    }
    result.summary = { warm: stats(result.warmReads), missing: stats(result.missingReads),
      cache: stats(result.cache.perReadBatchAveragesMs), concurrentRead: stats(withRead.map(n => n.read)),
      libraryRead: stats(result.libraryReads),
      baselineNewReady: stats(baseline, 'newReadyMs'), withReadNewReady: stats(withRead, 'newReadyMs'),
      pairedNewReadyDifference: stats(withRead.map(n => n.newReadyMs - baseline.find(b => b.pair === n.pair).newReadyMs)),
      settingsReadyBeforeNew: withRead.filter(n => n.read.readyMs <= n.newReadyMs).length };
    result.budgetPass = result.summary.warm.p95 <= result.budgets.warmReadP95Ms &&
      result.summary.concurrentRead.p95 <= result.budgets.pageReadP95Ms && result.summary.libraryRead.p95 <= result.budgets.pageReadP95Ms &&
      result.summary.cache.p95 <= result.budgets.cacheP95Ms;
    console.log(JSON.stringify(result.summary, null, 2));
    console.log(`STORAGE PERFORMANCE BUDGET: ${result.budgetPass ? 'PASS' : 'FAIL'}`);
    if (!result.budgetPass) process.exitCode = 2;
  } catch (e) {
    result.error = e.stack;
    console.error(e.message);
    process.exitCode = 1;
  } finally {
    try {
      if (listId) {
        api(`web/lists(guid'${listId}')`, 'delete');
        const remaining = api(`web/lists?$select=Id&$filter=Id%20eq%20guid'${listId}'`).value;
        assert.equal(remaining.length, 0);
        result.cleanup = { deletedListId: listId, absenceVerified: true };
      }
      result.policyAfter = cli(['spo', 'tenant', 'site', 'get', '--url', site, '--query',
        '{NoScript:DenyAddAndCustomizePages,PropertyBagException:AllowWebPropertyBagUpdateWhenDenyAddAndCustomizePagesIsEnabled}']);
      assert.deepEqual(result.policyAfter, result.policyBefore);
    } catch (e) { result.cleanupError = e.message; console.error('Cleanup:', e.message); process.exitCode = 1; }
    if (context) await context.close();
    fs.writeFileSync(output, JSON.stringify(result, null, 2) + '\n');
    console.log(`Evidence: ${output}`);
  }
})();
