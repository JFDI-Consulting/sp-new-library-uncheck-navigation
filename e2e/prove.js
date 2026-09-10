// Headed Playwright proof: on a site with the app installed the 'Show in site navigation'
// checkbox must be unchecked for both List and Document library; on a control site without
// the app it must be checked. Then the settings panel is used to switch the customizer off
// (checkbox stays checked) and back on (unchecked again). Sign in once in the browser window; the profile persists in
// ~/.cache/pw-sp-profile. Usage: TEST_SITE=... CONTROL_SITE=... node e2e/prove.js
const { chromium } = require('playwright');
const path = require('path');
const os = require('os');

const SHOTS = path.join(__dirname, '..', 'docs', 'proof');
const SITE = process.env.TEST_SITE || 'https://g53.sharepoint.com/sites/UncheckNavTest';
const CONTROL = process.env.CONTROL_SITE || 'https://g53.sharepoint.com';
const PROFILE = path.join(os.homedir(), '.cache', 'pw-sp-profile');
const TEST_SITE = SITE;
const CONTROL_SITE = CONTROL;
function proxyFromEnv() {
  const raw = process.env.HTTPS_PROXY || process.env.https_proxy || process.env.HTTP_PROXY;
  if (!raw) return undefined;
  const u = new URL(raw);
  return { server: `${u.protocol}//${u.hostname}:${u.port}`, username: decodeURIComponent(u.username), password: decodeURIComponent(u.password), bypass: process.env.NO_PROXY };
}
const NAV_RE = /show (list |library )?in site navigation/i;

async function waitForSignIn(page) {
  console.log('>>> If a Microsoft sign-in page is showing, sign in as jjadmin@g53.onmicrosoft.com in the browser window (up to 10 min).');
  const deadline = Date.now() + 10 * 60 * 1000;
  while (Date.now() < deadline) {
    const onSpo = page.url().startsWith('https://g53.sharepoint.com');
    const newBtn = page.locator('button[name="New"], button:has-text("New")').first();
    if (onSpo && await newBtn.isVisible().catch(() => false)) return;
    await page.waitForTimeout(1000);
  }
  throw new Error('Timed out waiting for sign-in / Site contents page');
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
  await page.locator('button[name="New"], button:has-text("New")').first().click();
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
async function setEnabledViaPanel(page, siteUrl, enabled, how) {
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
  await toggle.waitFor({ state: 'visible', timeout: 15000 });
  await page.waitForTimeout(500);
  if ((await toggle.isChecked()) !== enabled) {
    await toggle.click();
  }
  await page.screenshot({ path: path.join(SHOTS, `settings-panel-${enabled ? 'enable' : 'disable'}.png`) });
  await page.locator('[data-automation-id="jfdi-unav-save"]').click();
  await toggle.waitFor({ state: 'hidden', timeout: 15000 });
  console.log(`  saved enabled=${enabled} via ${how}`);
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

async function checkSettings(page, siteUrl) {
  await setEnabledViaPanel(page, siteUrl, false, 'bar');
  const whileDisabled = await checkListCheckbox(page, siteUrl, 'with-app-disabled');
  await setEnabledViaPanel(page, siteUrl, true, 'query');
  const afterReenable = await checkListCheckbox(page, siteUrl, 'with-app-reenabled');
  return { whileDisabled, afterReenable };
}

(async () => {
  const context = await chromium.launchPersistentContext(PROFILE, {
    headless: false,
    proxy: proxyFromEnv(),
    viewport: { width: 1400, height: 900 },
    args: ['--disable-blink-features=AutomationControlled']
  });
  const page = context.pages()[0] || await context.newPage();
  page.on('console', (m) => { if (/UncheckSiteNavigation/.test(m.text())) console.log('CONSOLE:', m.text()); });
  let ok = false;
  try {
    const testResults = await checkSite(page, TEST_SITE, 'with-app');
    const controlResults = await checkSite(page, CONTROL_SITE, 'control-no-app');
    const settings = await checkSettings(page, TEST_SITE);
    const partialNav = await checkPartialNavigation(page, TEST_SITE);
    const summary = { testSite: TEST_SITE, controlSite: CONTROL_SITE, withApp: testResults, control: controlResults, settings, partialNav };
    console.log(JSON.stringify(summary, null, 2));
    ok = Object.values(testResults).every((r) => r.checked === false)
      && Object.values(controlResults).every((r) => r.checked === true)
      && settings.whileDisabled === true && settings.afterReenable === false
      && partialNav.onHome === false && partialNav.onSiteContents === true && partialNav.backOnHome === false;
    console.log(ok ? 'PROOF: PASS' : 'PROOF: FAIL');
  } catch (e) {
    console.error('ERROR:', e.message);
    await page.screenshot({ path: path.join(SHOTS, 'error.png') }).catch(() => {});
  } finally {
    await context.close();
  }
  process.exit(ok ? 0 : 1);
})();
