// Befüllt die Google-Tabelle im Browser aus google/sheets.json (erzeugt von build_sheets.py).
// Aufruf: node google/fill.js google/sheets.json <schritt> [<schritt> ...]
//   create            Blätter anlegen (fehlende) in der richtigen Reihenfolge
//   paste:<Blatt>     Zellen und Formatierung einfügen
//   ops:<Blatt>       Checkboxen, Dropdowns, Zahlenformate, Fixierung
//   delete:<Blatt>    Blatt löschen (z. B. das leere Tabellenblatt1)
const fs = require('fs');
const gs = require('./gs.js');

const spec = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const steps = process.argv.slice(3);
const log = (...a) => console.log(new Date().toISOString().slice(11, 19), ...a);

function colName(n) { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function css(st) {
  if (!st) return '';
  const p = [];
  if (st.bold) p.push('font-weight:bold');
  if (st.italic) p.push('font-style:italic');
  if (st.bg) p.push(`background-color:${st.bg}`);
  if (st.color) p.push(`color:${st.color}`);
  if (st.size) p.push(`font-size:${st.size}pt`);
  return p.join(';');
}

async function menu(page, menuId, ...path) {
  await page.click(menuId);
  await page.waitForTimeout(600);
  for (let i = 0; i < path.length; i++) {
    const item = page.locator('.goog-menu:visible .goog-menuitem:visible', { hasText: path[i] }).first();
    if (i < path.length - 1) { await item.hover(); await page.waitForTimeout(700); }
    else { await item.click(); }
  }
  await page.waitForTimeout(1000);
}

async function sheetNames(page) {
  return await page.$$eval('.docs-sheet-tab .docs-sheet-tab-name', els => els.map(e => e.innerText.trim()));
}

async function create(page) {
  for (const s of spec) {
    const names = await sheetNames(page);
    if (names.includes(s.name)) { log('vorhanden:', s.name); continue; }
    await page.click('[aria-label="Tabellenblatt hinzufügen"]');
    await page.waitForTimeout(2000);
    const tab = page.locator('.docs-sheet-active-tab .docs-sheet-tab-name').first();
    await tab.dblclick();
    await page.waitForTimeout(600);
    await page.keyboard.press('Control+A');
    await page.keyboard.type(s.name);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(1500);
    log('angelegt:', s.name, '->', (await sheetNames(page)).join(', '));
  }
}

// Ausgeblendete Spalten einblenden: Beim Einfügen ab A1 rutscht sonst alles um die ausgeblendeten Spalten nach rechts
async function unhideCols(page, sheet, ncols) {
  await gs.goto(page, `'${sheet}'!A1`);
  await gs.goto(page, `'${sheet}'!A:${colName(ncols)}`);
  await page.waitForTimeout(600);
  await page.mouse.click(90, 153, { button: 'right' });
  await page.waitForTimeout(800);
  const it = page.locator('.goog-menuitem:visible', { hasText: /^Spalten einblenden$/ });
  if (await it.count()) { await it.first().click(); await page.waitForTimeout(1500); log(sheet, 'Spalten eingeblendet'); }
  else await page.keyboard.press('Escape');
}

// Blatt vollständig zurücksetzen: Inhalte, Formate, Datenvalidierung, bedingte Formatierung, Fixierung
async function resetSheet(page, s) {
  const ncols = Math.max(...s.rows.map(r => r.length)) + 4;
  const ref = `'${s.name}'!A1:${colName(ncols)}${s.rows.length + 20}`;
  await unhideCols(page, s.name, ncols - 4);   // ganze Spalten nur bis zur letzten vorhandenen Spalte
  await gs.goto(page, ref);
  await page.keyboard.press('Delete'); await page.waitForTimeout(1500);
  log('Inhalte gelöscht'); await menu(page, '#docs-format-menu', 'Formatierung entfernen'); log('Formatierung entfernt');
  // „Formatierung entfernen“ lässt das Zahlenformat „Nur Text“ stehen
  await gs.goto(page, ref); await menu(page, '#docs-format-menu', 'Zahl', 'Automatisch'); log('Zahlenformat Automatisch');
  // „Formatierung entfernen“ nimmt auch die Datenvalidierung mit; Reste über die Seitenleiste entfernen
  try {
    await gs.goto(page, ref);
    await page.locator('#docs-menubar [role="menuitem"]', { hasText: /^Daten$/ }).first().click();
    await page.waitForTimeout(700);
    await page.locator('.goog-menu:visible .goog-menuitem:visible', { hasText: 'Datenvalidierung' }).first().click();
    await page.waitForTimeout(2000);
    const rm = page.locator('.waffle-sidebar-container [role="button"]:visible', { hasText: /^Alle entfernen$/ });
    if (await rm.count()) { await rm.first().click(); await page.waitForTimeout(1500); log('Validierung entfernt'); }
    await page.locator('[role="button"][aria-label="Schließen"]:visible').first().click().catch(() => {});
    await page.waitForTimeout(800);
  } catch (e) { log('Validierung: übersprungen', e.message.split('\n')[0]); await page.keyboard.press('Escape'); }
  await gs.goto(page, ref);
  await menu(page, '#docs-format-menu', 'Bedingte Formatierung');
  await page.waitForTimeout(2000);
  const rules = page.locator('.waffle-conditionalformat-rule-list [role="option"]:visible');
  try {
    for (let i = 0; i < 20 && await rules.count(); i++) {
      await rules.first().hover({ timeout: 3000 }); await page.waitForTimeout(400);
      await rules.first().locator('.waffle-conditionalformat-delete-button-holder *').first().click({ force: true, timeout: 3000 });
      await page.waitForTimeout(1200);
    }
  } catch (e) { log('bedingte Formatierung: keine Regeln gelöscht'); }
  // ohne Regeln öffnet die Seitenleiste direkt eine neue Regel: abbrechen, damit keine leere Regel entsteht
  await page.locator('[role="button"]:visible', { hasText: /^Abbrechen$/ }).first().click({ timeout: 3000 }).catch(() => {});
  await page.waitForTimeout(800);
  await page.locator('[role="button"][aria-label="Schließen"]:visible').first().click({ timeout: 3000 }).catch(() => {});
  await page.waitForTimeout(800);
  await gs.goto(page, `'${s.name}'!A1`);
  for (const what of ['Keine Zeilen', 'Keine Spalten']) {
    try { await menu(page, '#docs-view-menu', 'Fixieren', what); }
    catch (e) { log('Fixierung:', what, 'übersprungen'); await page.keyboard.press('Escape'); }
  }
  log('zurückgesetzt:', s.name);
}

async function pasteSheet(page, s) {
  await unhideCols(page, s.name, Math.max(...s.rows.map(r => r.length)));
  const styles = {};
  for (const [r, c, st] of s.styles) styles[`${r},${c}`] = st;
  const rows = s.rows;
  const CH = 30;
  for (let start = 0; start < rows.length; start += CH) {
    const chunk = rows.slice(start, start + CH);
    const tsv = chunk.map(r => r.join('\t')).join('\n');
    const html = '<table>' + chunk.map((r, i) => '<tr>' + r.map((v, j) => {
      const st = css(styles[`${start + i + 1},${j + 1}`]);
      return `<td${st ? ` style="${st}"` : ''}>${esc(v)}</td>`;
    }).join('') + '</tr>').join('') + '</table>';
    await gs.goto(page, `'${s.name}'!A${start + 1}`);
    await gs.paste(page, tsv, html);
    await page.waitForTimeout(1500 + chunk.length * 40);
    log(`${s.name}: Zeilen ${start + 1}-${start + chunk.length} eingefügt`);
  }
}

async function dropdown(page, op) {
  await gs.goto(page, op.range);
  await menu(page, '#docs-insert-menu', 'Drop-down');
  await page.waitForTimeout(2000);
  if (op.source) {
    await page.click('[role="listbox"][aria-label="Kriterien"]');
    await page.waitForTimeout(800);
    await page.locator('.goog-menuitem:visible', { hasText: 'Drop-down (aus einem Bereich)' }).first().click();
    await page.waitForTimeout(1200);
    const inputs = page.locator('input.waffle-range-selection-input:visible');
    const inp = inputs.nth((await inputs.count()) - 1);
    await inp.click();
    await inp.fill(op.source.replace(/^=/, ''));
    await page.keyboard.press('Enter');
    await page.waitForTimeout(800);
  } else {
    for (let i = 0; i < op.options.length; i++) {
      let vals = page.locator('input[aria-label="Wert"]:visible');
      if ((await vals.count()) <= i) {
        await page.locator('[role="button"]:visible', { hasText: 'Weiteres Element hinzufügen' }).first().click();
        await page.waitForTimeout(500);
        vals = page.locator('input[aria-label="Wert"]:visible');
      }
      const v = vals.nth(i);
      await v.click();
      await v.fill(op.options[i]);
    }
  }
  await page.locator('[role="button"]:visible', { hasText: 'Fertig' }).first().click();
  await page.waitForTimeout(1500);
}

// Füllfarbe in der Seitenleiste „Bedingte Formatierung“ (nicht der Füllfarbe-Knopf der Symbolleiste)
async function cfFill(page, color) {
  await page.click('.waffle-conditionalformat-custom-toolbar [aria-label="Füllfarbe"]'); await page.waitForTimeout(700);
  await page.locator(`.docs-material-colorpalette-cell[aria-label="${!color || color === 'rot' ? 'Hellrot 3' : color}"]:visible`).first().click();
  await page.waitForTimeout(700);
}

// Vorhandene Regeln mit Standardfarbe (grün) im Bereich auf Rot umstellen
async function cfRecolor(page, ref) {
  await gs.goto(page, ref);
  await menu(page, '#docs-format-menu', 'Bedingte Formatierung');
  await page.waitForTimeout(2000);
  const green = page.locator('.waffle-conditionalformat-format-preview[aria-label*="#b7e1cd"]:visible');
  let n = 0;
  while (await green.count()) {
    await green.first().click(); await page.waitForTimeout(1500);
    await cfFill(page, 'rot');
    await page.locator('[role="button"]:visible', { hasText: /^Fertig$/ }).first().click();
    await page.waitForTimeout(1500);
    if (++n > 10) break;
  }
  await page.locator('[role="button"][aria-label="Schließen"]:visible').first().click().catch(() => {});
  await page.waitForTimeout(800);
  log('Regeln umgefärbt:', ref, n);
}

async function runOps(page, s) {
  const only = process.env.OPS_ONLY ? process.env.OPS_ONLY.split(',') : null;
  for (const op of s.ops) {
    if (only && !only.includes(op.op)) continue;
    try {
      if (op.op === 'currency' || op.op === 'percent') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await page.click(op.op === 'currency' ? '#t-num-fmt-currency' : '#t-num-fmt-percent');
        await page.waitForTimeout(900);
      } else if (op.op === 'text') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await menu(page, '#docs-format-menu', 'Zahl', 'Nur Text');
      } else if (op.op === 'checkbox') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await menu(page, '#docs-insert-menu', 'Kästchen');
        // Kästchen setzen vorhandene Werte auf FALSCH: gesetzte Häkchen aus der Vorgabe wiederherstellen
        const m = op.range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) || op.range.match(/^([A-Z]+)(\d+)$/);
        if (m) {
          const ci = c => { let n = 0; for (const ch of c) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
          const [c1, r1, c2, r2] = m[3] ? [ci(m[1]), +m[2], ci(m[3]), +m[4]] : [ci(m[1]), +m[2], ci(m[1]), +m[2]];
          for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
            const v = (s.rows[r - 1] || [])[c - 1];
            // Einfügen von „WAHR“ ergäbe Text; die Leertaste schaltet das (jetzt leere) Kästchen auf WAHR
            if (v === 'WAHR' || v === 'TRUE') { await gs.goto(page, `'${s.name}'!${colName(c)}${r}`); await page.keyboard.press('Space'); await page.waitForTimeout(800); }
          }
        }
      } else if (op.op === 'dropdown' || op.op === 'dropdown_range') {
        await dropdown(page, { ...op, range: `'${s.name}'!${op.range}` });
      } else if (op.op === 'freeze') {
        const cell = `'${s.name}'!${colName(op.cols || 1)}${op.rows || 1}`;
        await gs.goto(page, cell);
        if (op.rows) await menu(page, '#docs-view-menu', 'Fixieren', op.rows === 1 ? '1 Zeile' : `Bis zu Zeile ${op.rows}`);
        if (op.cols) await menu(page, '#docs-view-menu', 'Fixieren', op.cols === 1 ? '1 Spalte' : `Bis zu Spalte ${colName(op.cols)}`);
      } else if (op.op === 'hide') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await page.keyboard.press('Control+Alt+0');
        await page.waitForTimeout(1200);
      } else if (op.op === 'cf_formula') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await menu(page, '#docs-format-menu', 'Bedingte Formatierung');
        await page.waitForTimeout(2000);
        if (!(await page.locator('[role="listbox"][aria-label="Condition"]:visible').count())) {
          await page.locator('[role="button"]:visible', { hasText: 'Weitere Regel hinzufügen' }).first().click();
          await page.waitForTimeout(1500);
        }
        const rangeInput = page.locator('input.waffle-range-selection-input:visible').first();
        await rangeInput.click(); await rangeInput.fill(op.range); await page.keyboard.press('Enter');
        await page.waitForTimeout(600);
        await page.click('[role="listbox"][aria-label="Condition"]'); await page.waitForTimeout(700);
        await page.locator('.goog-menuitem:visible', { hasText: 'Benutzerdefinierte Formel ist' }).first().click();
        await page.waitForTimeout(800);
        const f = page.locator('input[aria-label="Wert oder Formel"]:visible').first();
        await f.click(); await f.fill(op.formula); await page.waitForTimeout(500);
        await cfFill(page, op.color);
        await page.locator('[role="button"]:visible', { hasText: /^Fertig$/ }).first().click();
        await page.waitForTimeout(1200);
        await page.locator('[role="button"][aria-label="Schließen"]:visible').first().click().catch(() => {});
        await page.waitForTimeout(800);
      } else if (op.op === 'copyfmt') {
        // leere Zelle samt Format (ohne Validierung) über einen Bereich kopieren, z. B. um Auswahlfelder zu entfernen
        await gs.goto(page, `'${s.name}'!${op.src}`); await page.keyboard.press('Control+C'); await page.waitForTimeout(500);
        await gs.goto(page, `'${s.name}'!${op.dst}`); await page.keyboard.press('Control+V'); await page.waitForTimeout(900);
      } else if (op.op === 'plain') {
        await gs.goto(page, `'${s.name}'!${op.range}`);
        await menu(page, '#docs-format-menu', 'Zahl', 'Automatisch');
      }
      log(s.name, op.op, op.range || '');
    } catch (e) {
      log('FEHLER', s.name, JSON.stringify(op), e.message.split('\n')[0]);
      await page.keyboard.press('Escape');
    }
  }
}

async function deleteSheet(page, name) {
  const tabs = page.locator('.docs-sheet-tab');
  let tab = null;
  for (let i = 0; i < await tabs.count(); i++) {
    const t = (await tabs.nth(i).locator('.docs-sheet-tab-name').innerText()).trim();
    if (t === name) { tab = tabs.nth(i); break; }
  }
  if (!tab) { log('nicht gefunden:', name); return; }
  await tab.click({ button: 'right' });
  await page.waitForTimeout(800);
  await page.locator('.goog-menu:visible .goog-menuitem:visible', { hasText: /^Löschen$/ }).first().click();
  await page.waitForTimeout(1000);
  const ok = page.locator('[role="dialog"] button:visible, .modal-dialog button:visible', { hasText: 'OK' });
  if (await ok.count()) await ok.first().click();
  await page.waitForTimeout(1500);
  log('gelöscht:', name);
}

(async () => {
  const { browser, page } = await gs.open();
  try {
    for (const step of steps) {
      const [kind, arg] = step.split(':');
      if (kind === 'create') await create(page);
      else if (kind === 'paste') await pasteSheet(page, spec.find(s => s.name === arg));
      else if (kind === 'ops') await runOps(page, spec.find(s => s.name === arg));
      else if (kind === 'cells') {
        // cells:<Blatt>:<Zelle>[,<Zelle>...] – einzelne Zellen aus der Vorgabe neu einfügen
        const [, name, list] = step.split(':');
        const s = spec.find(x => x.name === name);
        for (const ref of list.split(',')) {
          const m = ref.match(/^([A-Z]+)(\d+)$/);
          let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
          const v = s.rows[Number(m[2]) - 1][c - 1];
          await gs.goto(page, `'${name}'!${ref}`);
          await gs.paste(page, v, `<table><tr><td style="background-color:#f8f9fa">${esc(v)}</td></tr></table>`);
          log('neu:', name, ref);
        }
      }
      else if (kind === 'colblock') {
        // colblock:<Blatt>:<Spalte>:<von>:<bis> – Spaltenausschnitt aus der Vorgabe neu einfügen
        const [, name, colL, from, to] = step.split(':');
        const s = spec.find(x => x.name === name);
        let c = 0; for (const ch of colL) c = c * 26 + ch.charCodeAt(0) - 64;
        const styles = {}; for (const [r, cc, st] of s.styles) styles[`${r},${cc}`] = st;
        for (let a = Number(from); a <= Number(to); a += 40) {
          const b = Math.min(Number(to), a + 39);
          const vals = []; for (let r = a; r <= b; r++) vals.push([r, s.rows[r - 1][c - 1]]);
          const tsv = vals.map(([, v]) => v).join('\n');
          const html = '<table>' + vals.map(([r, v]) => { const st = css(styles[`${r},${c}`]); return `<tr><td${st ? ` style="${st}"` : ''}>${esc(v)}</td></tr>`; }).join('') + '</table>';
          await gs.goto(page, `'${name}'!${colL}${a}`);
          await gs.paste(page, tsv, html);
          await page.waitForTimeout(1500);
          log('Spalte neu:', name, `${colL}${a}:${colL}${b}`);
        }
      }
      else if (kind === 'reset') await resetSheet(page, spec.find(x => x.name === arg));
      else if (kind === 'del') { const ref = step.slice(4); await gs.goto(page, ref); await page.keyboard.press('Delete'); await page.waitForTimeout(1200); log('gelöscht:', ref); }
      else if (kind === 'check') {
        // check:<Blatt>!<Zelle>,… – Kästchen einfügen und auf WAHR schalten
        const [name, cells] = step.slice(6).split('!');
        for (const c of cells.split(',')) {
          await gs.goto(page, `'${name}'!${c}`); await menu(page, '#docs-insert-menu', 'Kästchen');
          await gs.goto(page, `'${name}'!${c}`); await page.keyboard.press('Space'); await page.waitForTimeout(800);
          log('Kästchen WAHR:', name, c);
        }
      }
      else if (kind === 'cfdel') {
        // cfdel:<Ref>|<Text> – Regeln der bedingten Formatierung löschen, deren Beschreibung den Text enthält
        const [ref, text] = step.slice(6).split('|');
        await gs.goto(page, ref);
        await menu(page, '#docs-format-menu', 'Bedingte Formatierung');
        await page.waitForTimeout(2000);
        const opt = page.locator('.waffle-conditionalformat-rule-list [role="option"]:visible', { hasText: text });
        let n = 0;
        while (await opt.count() && n < 10) {
          await opt.first().hover(); await page.waitForTimeout(500);
          const del = opt.first().locator('.waffle-conditionalformat-delete-button-holder [role="button"], .waffle-conditionalformat-delete-button-holder div').first();
          await del.click({ force: true }); await page.waitForTimeout(1500); n++;
        }
        await page.locator('[role="button"][aria-label="Schließen"]:visible').first().click().catch(() => {});
        log('Regeln gelöscht:', ref, text, n);
      }
      else if (kind === 'cfrot') { await cfRecolor(page, step.slice(6)); }
      else if (kind === 'plain') {
        // plain:<Blatt>!<Bereich> – Zahlenformat „Automatisch“ (z. B. wenn Formeln als Text gespeichert wurden)
        const ref = step.slice(6); await gs.goto(page, ref); await menu(page, '#docs-format-menu', 'Zahl', 'Automatisch'); log('Automatisch:', ref);
      }
      else if (kind === 'clear') {
        // clear:<Blatt> – Inhalte, Formate und Fixierung eines Blatts zurücksetzen
        await gs.goto(page, `'${arg}'!A1:AZ1000`);
        await page.keyboard.press('Delete'); await page.waitForTimeout(1500);
        await page.keyboard.press('Control+Backslash'); await page.waitForTimeout(1500);
        await menu(page, '#docs-view-menu', 'Fixieren', 'Keine Zeilen');
        await menu(page, '#docs-view-menu', 'Fixieren', 'Keine Spalten');
        log('geleert:', arg);
      }
      else if (kind === 'copyfmt') {
        // copyfmt:<Blatt>:<Quelle>:<Ziel>,<Ziel> – Zelle samt Format (ohne Validierung) kopieren
        const [, name, src, list] = step.split(':');
        for (const dst of list.split(',')) {
          await gs.goto(page, `'${name}'!${src}`); await page.keyboard.press('Control+C'); await page.waitForTimeout(500);
          await gs.goto(page, `'${name}'!${dst}`); await page.keyboard.press('Control+V'); await page.waitForTimeout(800);
          log('kopiert:', name, src, '->', dst);
        }
      }
      else if (kind === 'delete') { for (const n of arg.split(',')) await deleteSheet(page, n); }
      else if (kind === 'add') { await page.click('[aria-label="Tabellenblatt hinzufügen"]'); await page.waitForTimeout(2000); await page.locator('.docs-sheet-active-tab .docs-sheet-tab-name').first().dblclick(); await page.waitForTimeout(600); await page.keyboard.press('Control+A'); await page.keyboard.type(arg); await page.keyboard.press('Enter'); await page.waitForTimeout(1500); log('angelegt:', arg); }
    }
    await page.waitForTimeout(5000);
    if (process.env.SHOT) await page.screenshot({ path: process.env.SHOT });
  } catch (e) {
    // Bei Fehlern Bildschirm und offene Dialoge festhalten
    const dlg = await page.$$eval('.modal-dialog', els => els.filter(x => x.offsetParent).map(x => x.innerText)).catch(() => []);
    log('ABBRUCH:', e.message.split('\n')[0], dlg.length ? 'Dialog: ' + dlg.join(' | ') : '');
    if (process.env.ERRSHOT) await page.screenshot({ path: process.env.ERRSHOT });
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
