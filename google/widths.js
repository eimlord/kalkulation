// Setzt Spaltenbreiten in Pixeln über das Kontextmenü der Spaltenköpfe.
// Aufruf: VW=4200 [HIDE=A:A,W:W] node google/widths.js <Blatt> <A=45,B=55,...> [Basisbreite]
// Ausgeblendete Spalten werden vorher eingeblendet (sonst stimmen die Positionen nicht)
// und danach wieder ausgeblendet, wenn HIDE sie nennt.
const gs = require('./gs.js');
const sheet = process.argv[2];
const spec = Object.fromEntries(process.argv[3].split(',').map(x => x.split('=')).map(([c, w]) => [c, Number(w)]));
const base = Number(process.argv[4] || 100);
const ROWHEAD = 46, HEADY = 153;
const idx = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
const name = n => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

async function setWidth(page, x, px) {
  await page.mouse.click(x, HEADY, { button: 'right' });
  await page.waitForTimeout(800);
  await page.locator('.goog-menuitem:visible', { hasText: /Breite der Spalte|Größe der Spalte/ }).first().click();
  await page.waitForTimeout(900);
  const inp = page.locator('.modal-dialog input[type="text"]:visible, [role="dialog"] input:visible').first();
  await inp.fill(String(px));
  await page.locator('button:visible, [role="button"]:visible', { hasText: /^OK$/ }).first().click();
  await page.waitForTimeout(1200);
}

(async () => {
  const { browser, page } = await gs.open();
  const last = Math.max(...Object.keys(spec).map(idx));
  await gs.goto(page, `'${sheet}'!A1`);
  await gs.goto(page, `'${sheet}'!A:${name(last)}`);
  await page.waitForTimeout(600);
  await page.keyboard.press('Control+Shift+0');            // ausgeblendete Spalten einblenden
  await page.waitForTimeout(1200);
  await setWidth(page, ROWHEAD + base / 2, base);          // alle Spalten auf Basisbreite
  const widths = {}; for (let i = 1; i <= last; i++) widths[i] = base;
  for (let i = 1; i <= last; i++) {
    const want = spec[name(i)];
    if (!want || want === base) continue;
    let x = ROWHEAD; for (let j = 1; j < i; j++) x += widths[j];
    await gs.goto(page, `'${sheet}'!A1`);
    await page.mouse.click(x + widths[i] / 2, HEADY);       // Spalte markieren
    await page.waitForTimeout(400);
    await setWidth(page, x + widths[i] / 2, want);
    widths[i] = want;
    console.log(sheet, name(i), want);
  }
  for (const r of (process.env.HIDE || '').split(',').filter(Boolean)) {
    await gs.goto(page, `'${sheet}'!${r}`);
    await page.keyboard.press('Control+Alt+0');
    await page.waitForTimeout(1200);
    console.log(sheet, 'ausgeblendet', r);
  }
  await gs.goto(page, `'${sheet}'!A1`);
  await page.waitForTimeout(1500);
  if (process.env.SHOT) await page.screenshot({ path: process.env.SHOT });
  await browser.close();
})();
