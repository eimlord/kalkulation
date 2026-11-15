// Hilfsfunktionen, um eine per Link freigegebene Google-Tabelle im (headless) Browser zu bearbeiten.
const { chromium } = require('playwright');
const SHEET = process.env.GSHEET_ID || '1WPRZqluxXPPg9reSPBmSH1vB4Qzn4yZ8Ke16rYpXa44';

async function open() {
  const browser = await chromium.launch();
  const ctx = await browser.newContext({ viewport: { width: Number(process.env.VW || 1600), height: 1000 }, locale: 'de-AT' });
  await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: 'https://docs.google.com' });
  const page = await ctx.newPage();
  await page.goto(`https://docs.google.com/spreadsheets/d/${SHEET}/edit`, { waitUntil: 'domcontentloaded', timeout: 90000 });
  await page.waitForSelector('#t-name-box', { timeout: 90000 });
  await page.waitForTimeout(4000);
  return { browser, ctx, page };
}

async function goto(page, ref) {
  const box = page.locator('#t-name-box');
  await box.click();
  await box.fill(ref);
  await box.press('Enter');
  await page.waitForTimeout(400);
}

async function paste(page, text, html) {
  await page.evaluate(async ({ text, html }) => {
    const items = { 'text/plain': new Blob([text], { type: 'text/plain' }) };
    if (html) items['text/html'] = new Blob([html], { type: 'text/html' });
    await navigator.clipboard.write([new ClipboardItem(items)]);
  }, { text, html });
  await page.keyboard.press('Control+V');
  await page.waitForTimeout(1500);
}

async function formulaBar(page) {
  return (await page.locator('#t-formula-bar-input').innerText()).trim();
}

async function waitSaved(page, ms = 60000) {
  const t0 = Date.now();
  while (Date.now() - t0 < ms) {
    const s = await page.evaluate(() => document.body.innerText);
    if (/Alle Änderungen|gespeichert|Gespeichert/.test(s) && !/Wird gespeichert/.test(s)) return true;
    await page.waitForTimeout(1000);
  }
  return false;
}

async function csv(page, sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SHEET}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  return await page.evaluate(async (u) => (await fetch(u)).text(), url);
}

module.exports = { open, goto, paste, formulaBar, waitSaved, csv, SHEET };
