// Playwright proof: on a site with the app installed the 'Show in site navigation'
// checkbox must be unchecked for both List and Document library; on a control site without
// the app it must be checked. The settings panel then switches the customizer off and on,
// proving the hidden-list write, session cache, cache expiry and failed-read behavior.
// Sign in once in the browser window; the profile persists in ~/.cache/pw-sp-profile.
// Usage: HEADLESS=true TEST_SITE=... CONTROL_SITE=... node e2e/prove.js
const { chromium } = require('playwright');
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const os = require('os');

const SHOTS = path.join(__dirname, '..', 'docs', 'proof');
const SITE = process.env.TEST_SITE || 'https://g53.sharepoint.com/sites/UncheckNavTest';
const CONTROL = process.env.CONTROL_SITE || 'https://g53.sharepoint.com';
const PROFILE = path.join(os.homedir(), '.cache', 'pw-sp-profile');
const OUTPUT = process.env.RUNTIME_PROOF_OUTPUT || path.join(SHOTS, 'runtime-list-settings.json');
const TEST_SITE = SITE;
const CONTROL_SITE = CONTROL;
const LIBRARY_URL = process.env.TEST_LIBRARY_URL || `${TEST_SITE}/Shared%20Documents/Forms/AllItems.aspx`;
const HEADLESS = /^(1|true|yes)$/i.test(process.env.HEADLESS || 'false');
const FRESH_CONTEXT = HEADLESS && !/^(0|false|no)$/i.test(process.env.FRESH_CONTEXT || 'true');
const SIGN_IN_TIMEOUT_MS = Number(process.env.SIGN_IN_TIMEOUT_MS) || (HEADLESS ? 60 * 1000 : 10 * 60 * 1000);
const CACHE_KEY_PREFIX = `jfdi-unav:settings:v1:${encodeURIComponent(TEST_SITE.replace(/\/$/, '').toLowerCase())}:`;
function proxyFromEnv() {
  const raw = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY;
  if (!raw) return undefined;
  const u = new URL(raw);
  return { server: `${u.protocol}//${u.hostname}:${u.port}`, username: decodeURIComponent(u.username), password: decodeURIComponent(u.password), bypass: process.env.NO_PROXY };
}
const NAV_RE = /show (list |library )?in site navigation/i;

function newButton(page) {
  // Site Contents varies between an HTML name and a longer accessible label.
  return page.locator('button[name="New"]:visible, button[aria-label^="New."]:visible').first();
}

function libraryCreateButton(page) {
  return page.locator('button[data-automationid="newCommand"]:visible').first();
}

function decodedUrl(requestOrEntry) {
  let url = typeof requestOrEntry.url === 'function' ? requestOrEntry.url() : requestOrEntry.url;
  try { url = decodeURIComponent(url); } catch { /* retain the original URL */ }
  return url;
}

function isSettingsStoreRequest(requestOrEntry) {
  return /\/Lists\/JfdiUnavSettings/i.test(decodedUrl(requestOrEntry));
}

function isSettingsRead(requestOrEntry) {
  const method = typeof requestOrEntry.method === 'function' ? requestOrEntry.method() : requestOrEntry.method;
  return method === 'GET'
    && /\/_api\/web\/GetList\(@list\)\/items/i.test(decodedUrl(requestOrEntry))
    && isSettingsStoreRequest(requestOrEntry);
}

function isCustomActionRequest(request) {
  return /\/_api\/(web|site)\/UserCustomActions/i.test(request.url());
}

function trackSettingsNetwork(page) {
  const requests = [];
  const entries = new Map();
  page.on('request', (request) => {
    if (isSettingsStoreRequest(request) || isCustomActionRequest(request)) {
      const timing = request.timing();
      const entry = {
        method: request.method(),
        url: request.url(),
        startTime: timing.startTime,
        durationMs: undefined
      };
      requests.push(entry);
      entries.set(request, entry);
    }
  });
  page.on('requestfinished', (request) => {
    const entry = entries.get(request);
    if (!entry) return;
    const timing = request.timing();
    // Playwright reports responseEnd in milliseconds relative to startTime.
    entry.durationMs = timing.responseEnd >= 0 ? timing.responseEnd : undefined;
    entries.delete(request);
  });
  page.on('requestfailed', (request) => entries.delete(request));
  return requests;
}

async function siteCacheEntries(page) {
  return page.evaluate((prefix) => Object.keys(sessionStorage)
    .filter((key) => key.startsWith(prefix))
    .map((key) => ({ key, value: sessionStorage.getItem(key) })), CACHE_KEY_PREFIX);
}

async function clearSiteCache(page) {
  await page.evaluate((prefix) => {
    for (const key of Object.keys(sessionStorage)) {
      if (key.startsWith(prefix)) sessionStorage.removeItem(key);
    }
  }, CACHE_KEY_PREFIX);
}

async function expireSiteCache(page) {
  return page.evaluate((prefix) => {
    const keys = Object.keys(sessionStorage).filter((key) => key.startsWith(prefix));
    if (keys.length !== 1) throw new Error(`Expected one settings cache entry, found ${keys.length}`);
    const parsed = JSON.parse(sessionStorage.getItem(keys[0]));
    parsed.expiresAt = Date.now() - 1;
    sessionStorage.setItem(keys[0], JSON.stringify(parsed));
    return { key: keys[0], value: parsed };
  }, CACHE_KEY_PREFIX);
}

async function waitForSignIn(page) {
  console.log(`>>> Waiting for authenticated SharePoint UI (up to ${Math.round(SIGN_IN_TIMEOUT_MS / 1000)} seconds).`);
  const deadline = Date.now() + SIGN_IN_TIMEOUT_MS;
  while (Date.now() < deadline) {
    const onSpo = page.url().startsWith('https://g53.sharepoint.com');
    const newBtn = newButton(page);
    if (onSpo && await newBtn.isVisible().catch(() => false)) return;
    await page.waitForTimeout(1000);
  }
  const shot = path.join(SHOTS, 'signin-timeout.png');
  await page.screenshot({ path: shot, fullPage: true }).catch(() => {});
  throw new Error(`Timed out waiting for authenticated SharePoint UI at ${page.url()} (screenshot: ${shot})`);
}

async function closeOverlays(page) {
  for (let i = 0; i < 3; i++) {
    const closeBtn = page.getByRole('button', { name: /^(Close|Cancel)$/ }).first();
    if (await closeBtn.isVisible().catch(() => false)) {
      await closeBtn.click().catch(() => {});
      await page.waitForTimeout(500);
    } else {
      await page.keyboard.press('Escape').catch(() => {});
      await page.waitForTimeout(300);
    }
  }
}

async function openCreatePanel(page, menuItem) {
  await newButton(page).click();
  const mi = page.getByRole('menuitem', { name: new RegExp('^' + menuItem + '$', 'i') });
  await mi.waitFor({ state: 'visible', timeout: 15000 });
  await page.waitForTimeout(800);
  await mi.click({ force: true });
  await page.waitForTimeout(2000);
  // The create experience is usually hosted in a same-origin iframe (createlist.aspx).
  const iframe = page.locator('iframe[name="createListFrame"], iframe[src*="createlist.aspx"]').first();
  const inFrame = await iframe.isVisible().catch(() => false);
  const scope = inFrame ? page.frameLocator('iframe[name="createListFrame"], iframe[src*="createlist.aspx"]').first() : page;
  console.log(`  ${menuItem}: create UI ${inFrame ? 'is in createlist.aspx iframe' : 'is in the top document'}`);
  // Template chooser first ("List" / "Blank list" / "Blank library" tile), if shown.
  const tile = scope.getByText(/^(List|Blank list|Document library|Blank library|Library)$/).first();
  try {
    await tile.waitFor({ state: 'visible', timeout: 8000 });
    await page.waitForTimeout(500);
    await tile.click();
  } catch { /* went straight to the panel */ }
  const cb = scope.getByRole('checkbox', { name: NAV_RE }).first();
  await cb.waitFor({ state: 'visible', timeout: 30000 });
  // Give the customizer a moment to act after the panel renders.
  await page.waitForTimeout(2000);
  return cb;
}

async function checkSite(page, siteUrl, tag) {
  const results = {};
  await page.goto(`${siteUrl}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded' });
  await waitForSignIn(page);
  for (const item of ['List', 'Document library']) {
    const cb = await openCreatePanel(page, item);
    const checked = await cb.isChecked();
    const label = await cb.evaluate((el) => (el.closest('label') || el.parentElement).textContent.trim());
    const shot = path.join(SHOTS, `${tag}-${item.replace(/\s+/g, '-').toLowerCase()}.png`);
    await page.screenshot({ path: shot });
    results[item] = { label, checked, screenshot: shot };
    console.log(`[${tag}] ${item}: "${label}" checked=${checked}`);
    await closeOverlays(page);
  }
  return results;
}

// Settings panel proof: on the test site, switch the customizer off from the status bar on
// Site contents, confirm the checkbox is then left checked, then switch it back on via the
// query-string deep link and confirm it is unchecked again.
async function setEnabledViaPanel(page, siteUrl, enabled, how, network) {
  const url = how === 'query' ? `${siteUrl}/_layouts/15/viewlsts.aspx?jfdiUncheckNav=settings` : `${siteUrl}/_layouts/15/viewlsts.aspx`;
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await waitForSignIn(page);
  const bar = page.locator('[data-automation-id="jfdi-unav-bar"]');
  await bar.waitFor({ state: 'visible', timeout: 30000 });
  console.log(`  status bar: "${(await bar.innerText()).trim()}"`);
  if (how !== 'query') {
    await page.locator('[data-automation-id="jfdi-unav-change"]').click();
  }
  const toggle = page.locator('[data-automation-id="jfdi-unav-toggle"]');
  const saveButton = page.locator('[data-automation-id="jfdi-unav-save"]');
  await toggle.waitFor({ state: 'visible', timeout: 15000 });
  await page.waitForTimeout(500);
  const changed = (await toggle.isChecked()) !== enabled;
  if (!changed && await saveButton.isDisabled()) {
    await page.locator('[data-automation-id="jfdi-unav-panel"]').getByRole('button', { name: /^Cancel$/ }).last().click();
    await toggle.waitFor({ state: 'hidden', timeout: 15000 });
    console.log(`  enabled=${enabled} already set; no Save needed`);
    return { changed: false, saved: false, listWrites: [], customActionRequests: [] };
  }
  if (changed) {
    await toggle.click();
  }
  assert.strictEqual(await toggle.isChecked(), enabled, `Settings toggle did not change to ${enabled}`);
  await page.screenshot({ path: path.join(SHOTS, `settings-panel-${enabled ? 'enable' : 'disable'}.png`) });
  const requestStart = network.length;
  await saveButton.click();
  await toggle.waitFor({ state: 'hidden', timeout: 15000 });
  await page.waitForTimeout(750);
  const saveRequests = network.slice(requestStart);
  const customActionRequests = saveRequests.filter((entry) => /\/_api\/(web|site)\/UserCustomActions/i.test(entry.url));
  const listWrites = saveRequests.filter((entry) => entry.method !== 'GET' && isSettingsStoreRequest(entry));
  assert.strictEqual(customActionRequests.length, 0, 'Save wrote to UserCustomActions instead of the settings list');
  assert.ok(listWrites.length > 0, 'Save did not issue a hidden-list write');
  console.log(`  saved enabled=${enabled} via ${how}`);
  return { changed, saved: true, listWrites, customActionRequests };
}

async function checkListCheckbox(page, siteUrl, tag) {
  await page.goto(`${siteUrl}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded' });
  await waitForSignIn(page);
  const cb = await openCreatePanel(page, 'List');
  const checked = await cb.isChecked();
  await page.screenshot({ path: path.join(SHOTS, `${tag}-list.png`) });
  console.log(`[${tag}] List: checked=${checked}`);
  await closeOverlays(page);
  return checked;
}

// Reach Site contents by client-side navigation (no page load) and confirm the status bar
// appears, then leave by client-side navigation and confirm it goes away again.
async function checkPartialNavigation(page, siteUrl) {
  await page.goto(siteUrl, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(4000);
  const bar = page.locator('[data-automation-id="jfdi-unav-bar"]');
  const onHome = await bar.isVisible().catch(() => false);
  await page.getByRole('link', { name: /^Site contents$/ }).first().click();
  await bar.waitFor({ state: 'visible', timeout: 30000 }).catch(() => {});
  const onSiteContents = await bar.isVisible().catch(() => false);
  await page.screenshot({ path: path.join(SHOTS, 'partial-nav-site-contents.png') });
  await page.getByRole('link', { name: /^Home$/ }).first().click();
  await page.waitForTimeout(3000);
  const backOnHome = await bar.isVisible().catch(() => false);
  console.log(`[partial-nav] bar on home=${onHome}, after clicking Site contents=${onSiteContents}, back on home=${backOnHome}`);
  return { onHome, onSiteContents, backOnHome };
}

async function checkSettings(page, siteUrl, network) {
  const disableSave = await setEnabledViaPanel(page, siteUrl, false, 'bar', network);
  const whileDisabled = await checkSite(page, siteUrl, 'with-app-disabled');
  const enableSave = await setEnabledViaPanel(page, siteUrl, true, 'query', network);
  const afterReenable = await checkListCheckbox(page, siteUrl, 'with-app-reenabled');
  return { whileDisabled, afterReenable, disableSave, enableSave };
}

async function checkRuntimeCache(page, network) {
  await clearSiteCache(page);
  const coldStart = network.length;
  const coldResponse = page.waitForResponse((response) => isSettingsRead(response.request()), { timeout: 30000 })
    .then((response) => ({ response }), (error) => ({ error }));
  const [, coldObserved] = await Promise.all([
    (async () => {
      await page.goto(LIBRARY_URL, { waitUntil: 'domcontentloaded' });
      await libraryCreateButton(page).waitFor({ state: 'visible', timeout: 30000 });
    })(),
    coldResponse
  ]);
  if (coldObserved.error) throw coldObserved.error;
  const coldResult = coldObserved.response;
  const coldStatus = coldResult.status();
  await coldResult.finished();
  await page.waitForTimeout(1000);
  const coldReads = network.slice(coldStart).filter(isSettingsRead);
  const coldCache = await siteCacheEntries(page);
  assert.strictEqual(coldStatus, 200, 'Cold settings-list read did not return 200');
  assert.strictEqual(coldReads.length, 1, `Expected one cold settings read, found ${coldReads.length}`);
  assert.ok(Number.isFinite(coldReads[0].durationMs), 'Cold settings read timing was not recorded');
  assert.strictEqual(coldCache.length, 1, `Expected one populated cache entry, found ${coldCache.length}`);
  const coldValue = JSON.parse(coldCache[0].value);
  assert.strictEqual(coldValue.enabled, true, 'Cold read did not cache enabled=true');
  assert.ok(coldValue.expiresAt > Date.now(), 'Cold read cache entry is already expired');

  const cacheStart = network.length;
  await page.reload({ waitUntil: 'domcontentloaded' });
  await libraryCreateButton(page).waitFor({ state: 'visible', timeout: 30000 });
  await page.waitForTimeout(1500);
  const cachedReads = network.slice(cacheStart).filter(isSettingsRead);
  assert.strictEqual(cachedReads.length, 0, `Library reload bypassed session cache (${cachedReads.length} reads)`);

  const expired = await expireSiteCache(page);
  const expiryStart = network.length;
  const expiryResponse = page.waitForResponse((response) => isSettingsRead(response.request()), { timeout: 30000 })
    .then((response) => ({ response }), (error) => ({ error }));
  const [, expiryObserved] = await Promise.all([
    (async () => {
      await page.reload({ waitUntil: 'domcontentloaded' });
      await libraryCreateButton(page).waitFor({ state: 'visible', timeout: 30000 });
    })(),
    expiryResponse
  ]);
  if (expiryObserved.error) throw expiryObserved.error;
  const expiryResult = expiryObserved.response;
  const expiryStatus = expiryResult.status();
  await expiryResult.finished();
  await page.waitForTimeout(1000);
  const expiryReads = network.slice(expiryStart).filter(isSettingsRead);
  assert.strictEqual(expiryStatus, 200, 'Expired-cache refresh did not return 200');
  assert.strictEqual(expiryReads.length, 1, `Expected one read after cache expiry, found ${expiryReads.length}`);
  assert.ok(Number.isFinite(expiryReads[0].durationMs), 'Expired-cache settings read timing was not recorded');

  console.log(`[cache] cold reads=${coldReads.length}, cached reload reads=${cachedReads.length}, expired reads=${expiryReads.length}`);
  return {
    key: coldCache[0].key,
    cold: { status: coldStatus, reads: coldReads.length, durationMs: coldReads[0].durationMs, cached: coldValue },
    libraryReload: { reads: cachedReads.length },
    expiry: { forced: expired.value.expiresAt, status: expiryStatus, reads: expiryReads.length, durationMs: expiryReads[0].durationMs }
  };
}

async function checkReadFailureIsFailClosed(page) {
  await clearSiteCache(page);
  let intercepted = 0;
  const handler = async (route) => {
    if (isSettingsRead(route.request())) {
      intercepted++;
      await route.fulfill({ status: 503, contentType: 'application/json', body: JSON.stringify({ error: 'E2E injected settings failure' }) });
    } else {
      await route.continue();
    }
  };
  await page.route('**/*', handler);
  let checked;
  try {
    await page.goto(`${TEST_SITE}/_layouts/15/viewlsts.aspx`, { waitUntil: 'domcontentloaded' });
    await waitForSignIn(page);
    const cb = await openCreatePanel(page, 'List');
    checked = await cb.isChecked();
    await page.screenshot({ path: path.join(SHOTS, 'settings-read-failure-list.png') });
    await closeOverlays(page);
  } finally {
    await page.unroute('**/*', handler);
  }
  assert.ok(intercepted > 0, 'The injected settings-list failure was not exercised');
  assert.strictEqual(checked, true, 'A failed settings read did not leave SharePoint\'s checkbox unchanged');

  // Restore a valid, enabled cache entry so the proof leaves the browser in the normal state.
  await clearSiteCache(page);
  const recoveryResponse = page.waitForResponse((response) => isSettingsRead(response.request()), { timeout: 30000 })
    .then((response) => ({ response }), (error) => ({ error }));
  const [, recoveryObserved] = await Promise.all([
    (async () => {
      await page.goto(LIBRARY_URL, { waitUntil: 'domcontentloaded' });
      await libraryCreateButton(page).waitFor({ state: 'visible', timeout: 30000 });
    })(),
    recoveryResponse
  ]);
  if (recoveryObserved.error) throw recoveryObserved.error;
  const recoveryStatus = recoveryObserved.response.status();
  assert.strictEqual(recoveryStatus, 200, 'Settings read did not recover after removing the interception');
  console.log(`[fail-closed] intercepted reads=${intercepted}, checkbox checked=${checked}, recovery=${recoveryStatus}`);
  return { interceptedReads: intercepted, checked, recoveryStatus };
}

(async () => {
  let browser;
  let context;
  if (FRESH_CONTEXT) {
    // Reuse only authenticated cookies from the profile. A fresh browser context avoids
    // SharePoint's cached component manifest serving the previous package after deployment.
    const profileContext = await chromium.launchPersistentContext(PROFILE, {
      headless: true,
      proxy: proxyFromEnv(),
      args: ['--disable-blink-features=AutomationControlled']
    });
    let state;
    try {
      const warmPage = profileContext.pages()[0] || await profileContext.newPage();
      await warmPage.goto(`${TEST_SITE}/_layouts/15/viewlsts.aspx?jfdiAuthWarm=${Date.now()}`, {
        waitUntil: 'domcontentloaded', timeout: SIGN_IN_TIMEOUT_MS
      });
      await waitForSignIn(warmPage);
      state = await profileContext.storageState();
    } finally {
      await profileContext.close();
    }
    browser = await chromium.launch({ headless: true, proxy: proxyFromEnv(), args: ['--disable-blink-features=AutomationControlled'] });
    context = await browser.newContext({
      storageState: { cookies: state.cookies, origins: [] },
      viewport: { width: 1400, height: 900 }
    });
  } else {
    context = await chromium.launchPersistentContext(PROFILE, {
      headless: HEADLESS,
      proxy: proxyFromEnv(),
      viewport: { width: 1400, height: 900 },
      args: ['--disable-blink-features=AutomationControlled']
    });
  }
  const page = context.pages()[0] || await context.newPage();
  const network = trackSettingsNetwork(page);
  page.on('console', (m) => { if (/UncheckSiteNavigation/.test(m.text())) console.log('CONSOLE:', m.text()); });
  let ok = false;
  let summary = { testSite: TEST_SITE, controlSite: CONTROL_SITE, headless: HEADLESS, freshContext: FRESH_CONTEXT, startedAt: new Date().toISOString() };
  try {
    const initialState = await setEnabledViaPanel(page, TEST_SITE, true, 'query', network);
    const testResults = await checkSite(page, TEST_SITE, 'with-app');
    const controlResults = await checkSite(page, CONTROL_SITE, 'control-no-app');
    const settings = await checkSettings(page, TEST_SITE, network);
    const cache = await checkRuntimeCache(page, network);
    const partialNav = await checkPartialNavigation(page, TEST_SITE);
    const failClosed = await checkReadFailureIsFailClosed(page);
    summary = { ...summary, completedAt: new Date().toISOString(), initialState, withApp: testResults, control: controlResults, settings, cache, partialNav, failClosed };
    console.log(JSON.stringify(summary, null, 2));
    ok = Object.values(testResults).every((r) => r.checked === false)
      && Object.values(controlResults).every((r) => r.checked === true)
      && Object.values(settings.whileDisabled).every((r) => r.checked === true)
      && settings.afterReenable === false
      && partialNav.onHome === false && partialNav.onSiteContents === true && partialNav.backOnHome === false
      && cache.cold.reads === 1 && cache.libraryReload.reads === 0 && cache.expiry.reads === 1
      && failClosed.checked === true && failClosed.recoveryStatus === 200;
    console.log(ok ? 'PROOF: PASS' : 'PROOF: FAIL');
  } catch (e) {
    summary = { ...summary, completedAt: new Date().toISOString(), error: e.stack || e.message || String(e) };
    console.error('ERROR:', e.message);
    await page.screenshot({ path: path.join(SHOTS, 'error.png') }).catch(() => {});
  } finally {
    fs.mkdirSync(path.dirname(OUTPUT), { recursive: true });
    fs.writeFileSync(OUTPUT, JSON.stringify({ ...summary, pass: ok }, null, 2) + '\n');
    console.log(`Evidence: ${OUTPUT}`);
    await context.close();
    if (browser) await browser.close();
  }
  process.exit(ok ? 0 : 1);
})();
