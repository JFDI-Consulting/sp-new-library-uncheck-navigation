// Bounded live proof of Tenant Wide Extensions registration, with exact rollback.
// Temporarily runs the customizer tenant-wide on the explicitly authorized g53 tenant.
const { chromium } = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');
const assert = require('assert');
const COMPONENT = 'd31c6f18-3a0d-462b-b677-c09314fbf3e6';
const SITE = 'https://g53.sharepoint.com/sites/UncheckNavTest';
const ROOT = 'https://g53.sharepoint.com';
const title = `UncheckNavigationProof${Date.now()}`;
const output = path.join(__dirname, '../docs/proof/tenant-list-runtime.json');
const evidence = { startedAt: new Date().toISOString(), title, sites: [] };
function cli(args) {
  const text = execFileSync('m365', [...args, '--output', 'json'], { encoding: 'utf8' }).trim();
  return text ? JSON.parse(text) : undefined;
}
function actions(web) { return cli(['spo', 'applicationcustomizer', 'list', '--webUrl', web, '--scope', 'All']).filter((a) => a.ClientSideComponentId === COMPONENT); }
function tenantActions() { return cli(['spo', 'tenant', 'applicationcustomizer', 'list']).filter((a) => a.TenantWideExtensionComponentId === COMPONENT); }
function proxy() {
  const raw = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY;
  if (!raw) return undefined;
  const url = new URL(raw);
  return { server: `${url.protocol}//${url.hostname}:${url.port}`, username: decodeURIComponent(url.username), password: decodeURIComponent(url.password), bypass: process.env.NO_PROXY };
}
(async () => {
  let context;
  let profileContext;
  let original;
  let removed = false;
  let added = false;
  let tenantId;
  try {
    assert.strictEqual(cli(['spo', 'tenant', 'appcatalogurl', 'get']), 'https://g53.sharepoint.com/sites/appcatalog', 'This proof is authorized only on g53');
    assert.deepStrictEqual(tenantActions(), [], 'Existing tenant registration must be reviewed before the proof');
    assert.deepStrictEqual(actions(ROOT), [], 'Root must not have a direct registration');
    const existing = actions(SITE);
    assert.strictEqual(existing.length, 1, 'Expected exactly one scoped pilot action');
    original = existing[0];
    assert.strictEqual(original.Scope, 3, 'Pilot must be web-scoped');
    evidence.originalAction = original;
    for (const web of [SITE, ROOT]) {
      const apps = cli(['spo', 'app', 'instance', 'list', '--siteUrl', web]);
      assert.ok(!apps.some((app) => app.ProductId === '3d32ef0c-4b53-4eb9-b342-10def9451d2e'), `${web} has a site app installation`);
    }
    cli(['spo', 'applicationcustomizer', 'remove', '--webUrl', SITE, '--id', original.Id, '--scope', 'Web', '--force']);
    removed = true;
    assert.deepStrictEqual(actions(SITE), []);
    cli(['spo', 'app', 'add', '--filePath', path.join(__dirname, '../sharepoint/solution/sp-new-library-uncheck-navigation.sppkg'), '--overwrite']);
    const apps = cli(['spo', 'app', 'list']).filter((app) => app.ProductId === '3d32ef0c-4b53-4eb9-b342-10def9451d2e');
    assert.strictEqual(apps.length, 1);
    // ClientSideInstance.xml causes this supported deployment to register the extension.
    added = true;
    cli(['spo', 'app', 'deploy', '--id', apps[0].ID, '--skipFeatureDeployment']);
    let rows = tenantActions();
    // Re-deploying an already deployed solution version may not re-provision its XML.
    // The supported CLI registration now validates the packaged tenant declaration.
    if (rows.length === 0) {
      cli(['spo', 'tenant', 'applicationcustomizer', 'add', '--title', title,
        '--clientSideComponentId', COMPONENT, '--clientSideComponentProperties', '{"debug":false}']);
      evidence.registrationMode = 'explicit CLI after redeploy of existing solution version';
      rows = tenantActions();
    } else {
      evidence.registrationMode = 'automatic deployment registration';
    }
    assert.strictEqual(rows.length, 1);
    tenantId = rows[0].Id || rows[0].ID;
    assert.ok(tenantId, 'Missing created tenant row ID');
    evidence.tenantRegistration = rows[0];
    profileContext = await chromium.launchPersistentContext(path.join(os.homedir(), '.cache/pw-sp-profile'), {
      headless: true, proxy: proxy(), viewport: { width: 1400, height: 900 }
    });
    const warmPage = profileContext.pages()[0] || await profileContext.newPage();
    await warmPage.goto(`${SITE}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded' });
    await warmPage.locator('button[name="New"]:visible').or(warmPage.getByRole('button', { name: /^New\b/i })).first().waitFor({ state: 'visible', timeout: 60000 });
    context = await profileContext.browser().newContext({ storageState: { cookies: await profileContext.cookies(), origins: [] }, viewport: { width: 1400, height: 900 } });
    const page = await context.newPage();
    const loadedBundles = [];
    page.on('response', (response) => {
      if (/uncheck-site-navigation-application-customizer_[a-z0-9]+\.js/.test(response.url())) loadedBundles.push(response.url());
    });
    evidence.loadedBundles = loadedBundles;
    for (const web of [ROOT, SITE]) {
      let shown = false;
      const attempts = [];
      for (let attempt = 0; attempt < 6 && !shown; attempt++) {
        await page.goto(`${web}/_layouts/15/viewlsts.aspx?unavProof=${Date.now()}`, { waitUntil: 'domcontentloaded' });
        shown = await page.locator('[data-automation-id="jfdi-unav-bar"]').waitFor({ state: 'visible', timeout: 20000 }).then(() => true, () => false);
        attempts.push({ attempt: attempt + 1, shown });
        console.log(`${web}: tenant extension visible=${shown}, attempt=${attempt + 1}`);
        if (!shown) await page.waitForTimeout(10000);
      }
      assert.ok(shown, 'Tenant registration did not become visible within the proof window');
      assert.strictEqual(await page.locator('[data-automation-id="jfdi-unav-bar"]').count(), 1, 'Duplicate customizer bar');
      await page.locator('button[name="New"]').click();
      await page.getByRole('menuitem', { name: /^List$/ }).click();
      const frame = page.frameLocator('iframe[name="createListFrame"]');
      const tile = frame.getByText(/^(List|Blank list)$/).first();
      await tile.waitFor({ state: 'visible', timeout: 15000 });
      await tile.click();
      const checkbox = frame.getByRole('checkbox', { name: /show (list )?in site navigation/i });
      await checkbox.waitFor({ state: 'visible', timeout: 30000 });
      await page.waitForTimeout(2000);
      assert.strictEqual(await checkbox.isChecked(), false);
      const screenshot = path.join(__dirname, '../docs/proof', web === ROOT ? 'tenant-root-list.png' : 'tenant-test-list.png');
      await page.screenshot({ path: screenshot });
      evidence.sites.push({ web, installedApp: false, directActions: actions(web).length, bars: 1, checked: false, attempts, screenshot });
    }
    const expectedBundle = fs.readdirSync(path.join(__dirname, '../dist')).find((name) => /^uncheck-site-navigation-application-customizer_[a-z0-9]+\.js$/.test(name));
    assert.ok(expectedBundle && loadedBundles.some((url) => url.endsWith(expectedBundle)), 'The current built runtime was not loaded');
    evidence.expectedBundle = expectedBundle;
    evidence.pass = true;
  } catch (error) {
    evidence.pass = false;
    evidence.error = error.stack || String(error);
    console.error(evidence.error);
    process.exitCode = 1;
  } finally {
    if (context) await context.close();
    if (profileContext) await profileContext.close();
    try {
      if (added) {
        // If a post-add lookup failed, identify only our unique titled row.
        const rows = tenantActions().filter((row) => tenantId ? String(row.Id || row.ID) === String(tenantId) : (row.Title === 'UncheckSiteNavigation' || row.Title === title));
        for (const row of rows) cli(['spo', 'tenant', 'applicationcustomizer', 'remove', '--id', String(row.Id || row.ID), '--force']);
        assert.deepStrictEqual(tenantActions(), [], 'Tenant proof registration remains');
      }
      if (removed) {
        assert.deepStrictEqual(actions(SITE), [], 'Unexpected pilot action prevents restoration');
        cli(['spo', 'applicationcustomizer', 'add', '--webUrl', SITE, '--scope', 'Web', '--title', original.Title,
          '--clientSideComponentId', COMPONENT, '--clientSideComponentProperties', original.ClientSideComponentProperties || '{}']);
        const restored = actions(SITE);
        assert.strictEqual(restored.length, 1);
        assert.strictEqual(restored[0].ClientSideComponentProperties, original.ClientSideComponentProperties);
        evidence.restoredPilotId = restored[0].Id;
      }
      evidence.cleanupVerified = true;
    } catch (error) {
      evidence.cleanupError = error.stack || String(error);
      process.exitCode = 1;
      console.error(evidence.cleanupError);
    }
    evidence.completedAt = new Date().toISOString();
    fs.writeFileSync(output, JSON.stringify(evidence, null, 2) + '\n');
    console.log(`Tenant proof: ${evidence.pass && evidence.cleanupVerified ? 'PASS' : 'FAIL'}; ${output}`);
  }
})();
