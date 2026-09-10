// Headed Playwright proof: on a site with the app installed the 'Show in site navigation'
// checkbox must be unchecked for both List and Document library; on a control site without
// the app it must be checked. Sign in once in the browser window; the profile persists in
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
    const summary = { testSite: TEST_SITE, controlSite: CONTROL_SITE, withApp: testResults, control: controlResults };
    console.log(JSON.stringify(summary, null, 2));
    ok = Object.values(testResults).every((r) => r.checked === false)
      && Object.values(controlResults).every((r) => r.checked === true);
    console.log(ok ? 'PROOF: PASS' : 'PROOF: FAIL');
  } catch (e) {
    console.error('ERROR:', e.message);
    await page.screenshot({ path: path.join(SHOTS, 'error.png') }).catch(() => {});
  } finally {
    await context.close();
  }
  process.exit(ok ? 0 : 1);
})();
