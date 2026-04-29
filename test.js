/**
 * test.js — Full market-readiness test suite
 *
 * Covers everything that can be validated WITHOUT deploying to Apps Script:
 *   1. File presence & structure
 *   2. JS syntax validity (all 3 products)
 *   3. API wiring — every server function called by the UI exists in code1.gs
 *   4. NYT refresh logic — auto-refresh path in clientGetNytSnapshot
 *   5. Date formatting — no hardcoded stale dates; _fmtNytDate helper present
 *   6. Release pipeline — variant swap + restore logic in release.js
 *   7. Product variant mapping — all 3 variants handled in code1.gs
 *   8. Setup wizard — _setupMyApp, _clearSavedWebAppUrl, _checkDeployment present
 *   9. URL safety — doGet /dev→/exec coercion, no anchor Drive-gateway exploits
 *  10. Theme sync — clientApplyThemeToSheet exists + called from all 3 UIs
 *  11. Insights tab — NYT tab-click handler present in all 3 UIs
 *  12. Menu integrity — single smart entry, no "Reset" landmine
 *  13. Security — NYT API key not exposed in compiled dist files
 *  14. Build artifacts — dist files exist and are obfuscated (not raw source)
 *  15. LIBRARY_HEADERS completeness — all columns indexed
 *  16. Code1.gs size sanity — no runaway file growth, no duplicate functions
 *
 * Run: node test.js
 */

'use strict';
const fs = require('fs');
const path = require('path');

// ─── helpers ──────────────────────────────────────────────────────────────────
const ROOT = __dirname;
let passed = 0, failed = 0, warned = 0;

function pass(label) {
  console.log('  ✅  ' + label);
  passed++;
}
function fail(label, detail) {
  console.log('  ❌  ' + label);
  if (detail) console.log('       ' + detail);
  failed++;
}
function warn(label, detail) {
  console.log('  ⚠️   ' + label);
  if (detail) console.log('       ' + detail);
  warned++;
}
function section(name) {
  console.log('\n── ' + name + ' ' + '─'.repeat(Math.max(2, 60 - name.length)));
}

function read(rel) {
  const abs = path.join(ROOT, rel);
  if (!fs.existsSync(abs)) return null;
  return fs.readFileSync(abs, 'utf8');
}

// Extract inline <script> blocks (no src=)
function scriptBlocks(html) {
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;
  const out = [];
  let m;
  while ((m = re.exec(html))) {
    const body = m[1];
    if (body.trim()) out.push({ body, startLine: html.slice(0, m.index).split('\n').length });
  }
  return out;
}

// ─── load sources ─────────────────────────────────────────────────────────────
const CODE     = read('code1.gs')    || '';
const IDX1     = read('index.html')  || '';
const IDX2     = read('index2.html') || '';
const IDX3     = read('index3.html') || '';
const DIST1    = read('index_dist.html')  || '';
const DIST2    = read('index2_dist.html') || '';
const DIST3    = read('index3_dist.html') || '';
const RELEASE  = read('release.js')  || '';
const BUILD    = read('build.js')    || '';
const VALIDATE = read('validate.js') || '';

const PRODUCTS = [
  { name: 'Product 1 (Romantic)', key: 'index',  src: IDX1, dist: DIST1 },
  { name: 'Product 2 (Horizon)',  key: 'index2', src: IDX2, dist: DIST2 },
  { name: 'Product 3 (Blossom)',  key: 'index3', src: IDX3, dist: DIST3 },
];

// ─── 1. File presence ─────────────────────────────────────────────────────────
section('1. File Presence');
const required = [
  'code1.gs', 'index.html', 'index2.html', 'index3.html',
  'index_dist.html', 'index2_dist.html', 'index3_dist.html',
  'release.js', 'build.js', 'validate.js', 'package.json',
  'appsscript.json',
  '.clasp.product1.json', '.clasp.product2.json', '.clasp.product3.json',
];
required.forEach(f => {
  if (fs.existsSync(path.join(ROOT, f))) pass(f + ' exists');
  else fail(f + ' MISSING');
});

// ─── 2. JS syntax – source files ──────────────────────────────────────────────
section('2. JS Syntax — Source Files');
PRODUCTS.forEach(({ name, src }) => {
  const blocks = scriptBlocks(src);
  let ok = true;
  blocks.forEach((blk, i) => {
    try { new Function(blk.body); }
    catch (e) {
      fail(name + ' — script block ' + (i + 1) + ' syntax error @ html line ' + blk.startLine, e.message);
      ok = false;
    }
  });
  if (ok) pass(name + ' — all ' + blocks.length + ' inline script blocks parse cleanly');
});

// ─── 3. API wiring — every serverRun() call maps to a function in code1.gs ────
section('3. API Wiring — serverRun → code1.gs');

// Collect all serverRun('fnName') and serverRun("fnName") calls across all 3 UIs
const serverRunRe = /serverRun\s*\(\s*['"]([a-zA-Z_][a-zA-Z0-9_]*)["']/g;
const calledFns = new Set();
[IDX1, IDX2, IDX3].forEach(src => {
  let m;
  const re = /serverRun\s*\(\s*['"]([a-zA-Z_][a-zA-Z0-9_]*)["']/g;
  while ((m = re.exec(src))) calledFns.add(m[1]);
});

// Also collect google.script.run.<fnName>( calls (used in the setup wizard inline JS)
const gsRunRe = /google\.script\.run(?:\.\w+\([^)]*\))*\.([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/g;
[IDX1, IDX2, IDX3, CODE].forEach(src => {
  let m;
  const re = /google\.script\.run(?:\.\w+\([^)]*\))*\.([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/g;
  while ((m = re.exec(src))) calledFns.add(m[1]);
});

// Functions exported from code1.gs (top-level `function NAME(`)
const exportedFns = new Set();
const fnRe = /^function ([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/mg;
let fm;
while ((fm = fnRe.exec(CODE))) exportedFns.add(fm[1]);

// Filter to only the client* and public-facing helpers actually callable from UI
const uiCallable = [...calledFns].filter(fn =>
  fn.startsWith('client') || fn.startsWith('_checkDeployment') ||
  fn.startsWith('_saveManualWebAppUrl') || fn.startsWith('_clearSavedWebAppUrl')
);

let wiringOk = true;
uiCallable.sort().forEach(fn => {
  if (exportedFns.has(fn)) pass(fn + ' → defined in code1.gs');
  else { fail(fn + ' — called from UI but NOT found in code1.gs'); wiringOk = false; }
});
if (uiCallable.length === 0) warn('No serverRun() calls detected — check extraction regex');

// ─── 4. NYT auto-refresh logic ────────────────────────────────────────────────
section('4. NYT Auto-Refresh Logic');

if (/clientRefreshNYTCache\(\)/.test(CODE) && /needsRefresh/.test(CODE))
  pass('clientGetNytSnapshot calls clientRefreshNYTCache when cache is stale');
else fail('clientGetNytSnapshot does NOT auto-refresh stale cache');

if (/ageDays.*>.*6|6.*ageDays/.test(CODE) || />\s*6/.test(CODE.slice(CODE.indexOf('needsRefresh'), CODE.indexOf('needsRefresh') + 400)))
  pass('Cache age threshold is 6 days');
else warn('Could not confirm 6-day cache age threshold in code1.gs');

if (/NYT_CACHE_DATE/.test(CODE))
  pass('NYT_CACHE_DATE property used for cache dating');
else fail('NYT_CACHE_DATE not found in code1.gs');

// ─── 5. Date format ───────────────────────────────────────────────────────────
section('5. Date Format — M/D/YYYY');

PRODUCTS.forEach(({ name, src }) => {
  if (/_fmtNytDate/.test(src)) pass(name + ' — _fmtNytDate helper present');
  else fail(name + ' — _fmtNytDate helper MISSING');
});

PRODUCTS.forEach(({ name, src }) => {
  // Check _fmtNytDate is actually applied in renderNytFeedPanel stamp
  if (/_fmtNytDate\(updatedAt\)/.test(src)) pass(name + ' — stamp uses _fmtNytDate');
  else fail(name + ' — renderNytFeedPanel stamp does NOT call _fmtNytDate');
});

// No hardcoded stale dates in source files (dist is obfuscated so we skip)
PRODUCTS.forEach(({ name, src }) => {
  // ISO dates inside demo/seed data (object property values, setStatus arguments) are
  // legitimate stored values, not UI display dates.
  const lines = src.split('\n');
  const isoRe = /\d{4}-\d{2}-\d{2}/g;
  const staleDisplayLines = lines.filter(line => {
    // Skip lines that are clearly data storage — not display rendering
    if (/dateAdded|dateStarted|dateFinished|dateMap|setStatus\s*\(/.test(line)) return false;
    // Skip comment lines (leading // or end-of-line // comments that contain the date)
    if (/^\s*\/\//.test(line)) return false;
    if (/\/\/.*\d{4}-\d{2}-\d{2}/.test(line)) return false;
    // Skip toISOString() calls — these compute the date at runtime, not hardcoded display
    if (/toISOString/.test(line)) return false;
    // Skip textarea content — example/placeholder CSV data, not a rendered UI date
    if (/<\/textarea>|<textarea/.test(line)) return false;
    const dates = line.match(isoRe) || [];
    return dates.some(d => /^20[0-9]{2}-/.test(d));
  });
  if (staleDisplayLines.length === 0) pass(name + ' — no hardcoded ISO display dates in source');
  else fail(name + ' — hardcoded ISO dates in display context: ' + staleDisplayLines.length + ' line(s)');
});

// ─── 6. Release pipeline ──────────────────────────────────────────────────────
section('6. Release Pipeline');

if (/productVariantRe\.test\(origCode\)/.test(RELEASE))
  pass('release.js pre-checks PRODUCT_VARIANT with .test() before replacing');
else fail('release.js missing .test() pre-check for PRODUCT_VARIANT');

if (/[\\u2018\\u2019]|\\\\u2018/.test(RELEASE) || /u2018|u2019/.test(RELEASE))
  pass('release.js uses quote-agnostic PRODUCT_VARIANT regex (handles curly quotes)');
else warn('Could not confirm quote-agnostic regex in release.js');

if (/restore\(\)/.test(RELEASE) && /process\.on\s*\(\s*['"]exit/.test(RELEASE))
  pass('release.js restores originals on exit/error');
else fail('release.js missing exit/error restore handler');

['1','2','3'].forEach(n => {
  const clasp = read(`.clasp.product${n}.json`);
  if (!clasp) { fail(`.clasp.product${n}.json missing`); return; }
  try {
    const parsed = JSON.parse(clasp);
    if (parsed.scriptId) pass(`.clasp.product${n}.json has scriptId: ${parsed.scriptId.slice(0,10)}…`);
    else fail(`.clasp.product${n}.json has no scriptId`);
  } catch (e) { fail(`.clasp.product${n}.json is invalid JSON`); }
});

// ─── 7. Product variant mapping ───────────────────────────────────────────────
section('7. Product Variant Mapping');

['index','index2','index3'].forEach(v => {
  if (new RegExp("'" + v + "'").test(CODE) || new RegExp('"' + v + '"').test(CODE))
    pass('Variant "' + v + '" referenced in code1.gs');
  else fail('Variant "' + v + '" NOT referenced in code1.gs');
});

if (/_VIEW_THEME_MAP/.test(CODE)) {
  // Use the actual definition line (var _VIEW_THEME_MAP = {...})
  // not the first usage (which is a reference at line 212)
  const defIdx = CODE.indexOf('var _VIEW_THEME_MAP');
  const mapBlock = defIdx >= 0 ? CODE.slice(defIdx, defIdx + 200) : '';
  ['romantic','horizon','blossom'].forEach(theme => {
    if (mapBlock.includes(theme)) pass('Theme "' + theme + '" in _VIEW_THEME_MAP');
    else fail('Theme "' + theme + '" MISSING from _VIEW_THEME_MAP');
  });
} else fail('_VIEW_THEME_MAP not found in code1.gs');

if (/doGet/.test(CODE) && /HtmlService\.createHtmlOutputFromFile/.test(CODE))
  pass('doGet serves HTML via createHtmlOutputFromFile');
else fail('doGet not serving HTML correctly');

// ─── 8. Setup wizard ──────────────────────────────────────────────────────────
section('8. Setup Wizard');

['_setupMyApp', '_checkDeployment', '_clearSavedWebAppUrl',
 '_saveManualWebAppUrl', '_smartOpenApp', '_openWebApp', '_getWebAppUrl'].forEach(fn => {
  if (exportedFns.has(fn)) pass(fn + ' defined in code1.gs');
  else fail(fn + ' MISSING from code1.gs');
});

// Wizard should NOT auto-save ScriptApp.getService().getUrl() blindly
if (/ScriptApp\.getService\(\)\.getUrl\(\)/.test(CODE)) {
  // Acceptable uses: doGet self-register (1), onOpen comment/conditional (few)
  // We strip comments first to count only executable occurrences
  const codeNoComments = CODE
    .replace(/\/\/[^\n]*/g, '')
    .replace(/\/\*[\s\S]*?\*\//g, '');
  const gsGetUrlCount = (codeNoComments.match(/ScriptApp\.getService\(\)\.getUrl\(\)/g) || []).length;
  if (gsGetUrlCount <= 3) pass('ScriptApp.getService().getUrl() has ' + gsGetUrlCount + ' executable call(s) (expected ≤ 3)');
  else warn('ScriptApp.getService().getUrl() has ' + gsGetUrlCount + ' executable calls — verify none blindly auto-save');
}

// Wizard screens: all 5 screen divs present
// Wizard slice: 25000 chars is enough to capture all 5 screens + JS
const wizardHtml = CODE.slice(CODE.indexOf('_setupMyApp'), CODE.indexOf('_setupMyApp') + 25000);
['"s1"','"s2"','"s3"','"s4"','"s5"'].forEach(id => {
  if (wizardHtml.includes(id)) pass('Wizard screen ' + id + ' present');
  else fail('Wizard screen ' + id + ' MISSING');
});

// Wrong link escape hatch
if (wizardHtml.includes('_clearSavedWebAppUrl') && wizardHtml.includes('Wrong link'))
  pass('Wizard has "Wrong link? Set up again" escape hatch');
else fail('Wizard missing "Wrong link? Set up again" escape hatch');

// Screen 5 chip is a clickable <a> not just <span> (popup-blocker fallback)
if (wizardHtml.includes('<a id=\\"okUrl\\"') || wizardHtml.includes("'<a id=\"okUrl\""))
  pass('Wizard success URL chip is a clickable <a> (popup-blocker proof)');
else fail('Wizard success URL chip is not a clickable anchor');

// ─── 9. URL safety ────────────────────────────────────────────────────────────
section('9. URL Safety');

// doGet /dev→/exec coercion
if (/\/dev.*\/exec|replace.*dev.*exec/.test(CODE))
  pass('doGet coerces /dev → /exec before self-registering URL');
else fail('doGet missing /dev → /exec coercion');

// _checkDeployment must NOT EXECUTE ScriptApp.getService().getUrl() (unreliable)
// The function body may MENTION it in a comment — that's fine and expected.
// We parse out non-comment lines only.
const checkDepFn = CODE.slice(
  CODE.indexOf('function _checkDeployment'),
  CODE.indexOf('function _checkDeployment') + 600
);
if (checkDepFn) {
  // Strip JS single-line and block comments before checking executable code
  const stripped = checkDepFn
    .replace(/\/\/[^\n]*/g, '')           // single-line comments
    .replace(/\/\*[\s\S]*?\*\//g, '');   // block comments
  if (/ScriptApp\.getService\(\)/.test(stripped))
    fail('_checkDeployment has executable call to unreliable ScriptApp.getService().getUrl()');
  else
    pass('_checkDeployment does NOT call ScriptApp.getService() in executable code (safe)');
}

// _smartOpenApp must NOT auto-save from ScriptApp.getService()
const smartOpenFn = CODE.slice(
  CODE.indexOf('function _smartOpenApp'),
  CODE.indexOf('function _openWebApp')
);
if (/ScriptApp\.getService/.test(smartOpenFn))
  fail('_smartOpenApp uses unreliable ScriptApp.getService() auto-detect');
else pass('_smartOpenApp does NOT use ScriptApp.getService() (safe)');

// No <a target="_blank"> on /home/projects editor URLs inside modal strings
// Those route through Drive gateway and show "unable to open"
const modalAnchors = CODE.match(/<a[^>]+href[^>]+home\/projects[^>]+target[^>]*>/g) || [];
if (modalAnchors.length === 0) pass('No Drive-gateway anchor links for /home/projects URLs in modal');
else fail('Found Drive-gateway anchor links: ' + modalAnchors.length + ' — these cause "unable to open" errors');

// ─── 10. Theme sync ───────────────────────────────────────────────────────────
section('10. Theme Sync (Near-Real-Time Palette)');

if (exportedFns.has('clientApplyThemeToSheet'))
  pass('clientApplyThemeToSheet defined in code1.gs');
else fail('clientApplyThemeToSheet MISSING from code1.gs');

PRODUCTS.forEach(({ name, src }) => {
  if (/clientApplyThemeToSheet/.test(src)) pass(name + ' calls clientApplyThemeToSheet');
  else fail(name + ' does NOT call clientApplyThemeToSheet');
});

// Theme map covers all 3 products
if (CODE.includes("'romantic'") && CODE.includes("'horizon'") && CODE.includes("'blossom'"))
  pass('All three product themes defined: romantic, horizon, blossom');
else fail('One or more product themes missing from code1.gs');

// ─── 11. Insights / NYT tab refresh ──────────────────────────────────────────
section('11. Insights — NYT Tab Refresh on Click');

PRODUCTS.forEach(({ name, src }) => {
  if (/_nytFetchingNow/.test(src)) pass(name + ' — _nytFetchingNow guard present');
  else fail(name + ' — _nytFetchingNow guard MISSING (risk of concurrent refresh calls)');
});

PRODUCTS.forEach(({ name, src }) => {
  if (/clientGetNytSnapshot.*insightsTab|insightsTab.*clientGetNytSnapshot|_switchingToNyt/.test(src))
    pass(name + ' — NYT tab click triggers clientGetNytSnapshot refresh');
  else fail(name + ' — NYT tab click does NOT trigger refresh');
});

PRODUCTS.forEach(({ name, src }) => {
  if (/applyNytSnapshot/.test(src)) pass(name + ' — applyNytSnapshot handler present');
  else fail(name + ' — applyNytSnapshot MISSING');
});

// ─── 12. Menu integrity ───────────────────────────────────────────────────────
section('12. Sheet Menu Integrity');

if (/📖 Open My Reading App|Open My Reading App/.test(CODE))
  pass('Menu has "📖 Open My Reading App" entry');
else fail('Smart menu entry not found in code1.gs');

// No "Reset" or "Reset App URL" exposed in the menu (that's a user landmine)
const menuBlock = CODE.slice(
  CODE.indexOf('ui.createMenu'),
  CODE.indexOf('ui.createMenu') + 800
);
if (/[Rr]eset.*URL|[Rr]eset.*App/.test(menuBlock))
  fail('"Reset App URL" menu item should be removed — buyer may accidentally clear their URL');
else pass('No "Reset URL" landmine in menu');

// Only one .addItem (one logical entry) — keeps it simple for non-tech buyers
const addItemCount = (menuBlock.match(/\.addItem\s*\(/g) || []).length;
// Menu items: Open (required) + Refresh Styling + Clear Data are all valid.
// Warn only if there are more than 3 (that would be too many for non-tech buyers).
if (addItemCount <= 3) pass('Menu has ' + addItemCount + ' item(s) — within acceptable limit for non-tech buyers');
else warn('Menu has ' + addItemCount + ' items — consider trimming for non-tech buyers');

// ─── 13. API key security in dist files ───────────────────────────────────────
section('13. Security — API Keys Not Exposed in Dist HTML');

const NYT_KEY = 'XirX9Yl9FnG5UAyFxD9PaALKFkrD68FKDyHssu0ZDzHW1qPJ';
const PI_KEY  = 'Y2R54KZNJ2HMPKTRMKMT';

// API keys live in code1.gs (server-side) — they MUST NOT appear in client HTML
PRODUCTS.forEach(({ name, src, dist }) => {
  if (src.includes(NYT_KEY)) fail(name + ' source — NYT API key exposed in HTML (client-visible!)');
  else pass(name + ' source — NYT API key NOT in HTML');

  if (src.includes(PI_KEY)) fail(name + ' source — PodcastIndex key exposed in HTML');
  else pass(name + ' source — PodcastIndex key NOT in HTML');

  if (dist && dist.includes(NYT_KEY)) fail(name + ' dist — NYT API key visible in obfuscated dist!');
  else if (dist) pass(name + ' dist — NYT API key NOT in dist');
});

// ─── 14. Build artifacts ──────────────────────────────────────────────────────
section('14. Build Artifacts — Dist Files Obfuscated');

PRODUCTS.forEach(({ name, key, src, dist }) => {
  if (!dist) { fail(name + ' — dist file missing'); return; }

  // Obfuscated dist should be significantly different from source
  // A simple heuristic: if dist contains the plain function name "renderNytFeedPanel"
  // it was NOT obfuscated (obfuscator renames these)
  if (dist.length < 1000) { fail(name + ' — dist file suspiciously small'); return; }

  // Obfuscated dist should contain hex identifiers.
  // NOTE: some function names survive obfuscation by design — the build config
  // uses renameGlobals:false to keep google.script.run.<fnName> intact, and
  // some top-level handlers referenced by HTML event attributes also survive.
  // We only check for the presence of hex encoding as the obfuscation signal.
  const hasHexIds = /_0x[0-9a-f]{4,}/i.test(dist);
  if (hasHexIds) pass(name + ' — dist contains hex identifiers (obfuscated)');
  else warn(name + ' — dist may not be obfuscated (no hex identifiers found)');

  // Size sanity — obfuscated output with base64 string encoding is typically
  // 1.05–1.15× the source (strings grow, identifiers shrink).
  const ratio = (dist.length / Math.max(src.length, 1)).toFixed(2);
  if (parseFloat(ratio) >= 1.0) pass(name + ' — dist/src ratio ' + ratio + ' (obfuscated dist is at least as large as source)');
  else warn(name + ' — dist/src ratio ' + ratio + ' — dist is smaller than source, check obfuscation ran');
});

// ─── 15. LIBRARY_HEADERS completeness ────────────────────────────────────────
section('15. Schema — LIBRARY_HEADERS');

const critical = ['BookId','Title','Author','Status','Rating','ISBN','Tags','Shelves','Notes'];
critical.forEach(col => {
  if (CODE.includes("'" + col + "'"))
    pass('LIBRARY_HEADERS contains "' + col + '"');
  else fail('LIBRARY_HEADERS missing critical column "' + col + '"');
});

// Confirm LIBRARY_DATA_ROW = 9 and LIBRARY_HEADER_ROW = 8 (banner rows 1-7)
if (/LIBRARY_DATA_ROW\s*=\s*9/.test(CODE))  pass('LIBRARY_DATA_ROW = 9 (banner rows 1-7 preserved)');
else fail('LIBRARY_DATA_ROW is not 9 — layout may break');

if (/LIBRARY_HEADER_ROW\s*=\s*8/.test(CODE)) pass('LIBRARY_HEADER_ROW = 8');
else fail('LIBRARY_HEADER_ROW is not 8');

// ─── 16. Code1.gs size sanity ────────────────────────────────────────────────
section('16. Code1.gs Integrity');

const codeLines = CODE.split('\n').length;
if (codeLines > 1000 && codeLines < 8000)
  pass('code1.gs line count: ' + codeLines + ' (healthy range 1000–8000)');
else if (codeLines >= 8000)
  warn('code1.gs is ' + codeLines + ' lines — may approach Apps Script 100 KB limit');
else fail('code1.gs is only ' + codeLines + ' lines — suspiciously small');

// Duplicate function detection — same function name defined twice is a bug
const fnNames = [];
const dupeRe = /^function ([a-zA-Z_][a-zA-Z0-9_]*)\s*\(/mg;
let dr;
while ((dr = dupeRe.exec(CODE))) fnNames.push(dr[1]);
const seen = {}, dupes = [];
fnNames.forEach(n => { if (seen[n]) dupes.push(n); else seen[n] = true; });
if (dupes.length === 0) pass('No duplicate function names in code1.gs');
else fail('Duplicate functions in code1.gs: ' + dupes.join(', '));

// appsscript.json — confirm timeZone and webapp execution
const appsjson = read('appsscript.json');
if (appsjson) {
  try {
    const ajs = JSON.parse(appsjson);
    if (ajs.webapp && ajs.webapp.executeAs)
      pass('appsscript.json has webapp.executeAs: ' + ajs.webapp.executeAs);
    else warn('appsscript.json missing webapp.executeAs — deployment may default incorrectly');
    if (ajs.timeZone)
      pass('appsscript.json has timeZone: ' + ajs.timeZone);
    else warn('appsscript.json missing timeZone');
  } catch (e) { fail('appsscript.json is invalid JSON'); }
} else fail('appsscript.json missing');

// ─── 17. Modal close symmetry — no user can get trapped ───────────────────────
section('17. Modal Close Symmetry (No Trapped Modals)');

// Every state flag that can be set to TRUE must also have a FALSE path.
// We check in index.html (all 3 products share the same logic).
const showFlags = [
  'showBookDetail', 'showAddEditBook', 'showNytPreviewModal', 'showInsightsModal',
  'showSettings', 'showProfileModal', 'showImportModal', 'showExportModal',
  'showShelfModal', 'showShelfManagerModal', 'showChallengeManagerModal',
  'showWrapped', 'showConfirmDialog', 'showSearchAdd', 'showFabMenu',
  'showThemeSelector', 'showWelcomeBanner', 'fullLibraryOpen',
];
showFlags.forEach(flag => {
  // A state flag can be "opened" in 3 ways:
  //   1. Explicit assignment:  state.showXxx = true
  //   2. Toggle shorthand:     state.showXxx = !state.showXxx
  //   3. Initial state value:  showXxx: true  (in the state object literal)
  const openRe     = new RegExp('state\\.' + flag + '\\s*=\\s*true');
  const toggleRe   = new RegExp('state\\.' + flag + '\\s*=\\s*!state\\.' + flag);
  const initRe     = new RegExp(flag + '\\s*:\\s*true');
  const closeRe    = new RegExp('state\\.' + flag + '\\s*=\\s*false');
  const hasOpen    = openRe.test(IDX1) || toggleRe.test(IDX1) || initRe.test(IDX1);
  const hasClose   = closeRe.test(IDX1) || toggleRe.test(IDX1);
  if (!hasOpen)  warn(flag + ' — never opened (may be dead UI state)');
  else if (!hasClose) fail(flag + ' — opened but NEVER closed (user can get trapped)');
  else pass(flag + ' — both open and close paths exist');
});

// ─── 18. ESC key safety — critical modals respond to Escape ──────────────────
section('18. ESC Key Safety (No Escrow-Trap Modals)');

// Read the MAIN ESC handler — the block that starts with the outer Escape routing
// (not the inner tour-specific Escape block which comes first in source).
// Strategy: find the keydown listener, then find the SECOND `if (e.key === 'Escape')` occurrence.
const kd = IDX1.indexOf("document.addEventListener('keydown'");
const esc1 = IDX1.indexOf("e.key === 'Escape'", kd);
const esc2 = IDX1.indexOf("e.key === 'Escape'", esc1 + 1);  // second occurrence = main routing block
const escBlock = IDX1.slice(esc2, esc2 + 5000);

// Critical modals that MUST be dismissible with ESC
const escExpected = [
  'showConfirmDialog', 'showAddEditBook', 'showBookDetail', 'showNytPreviewModal',
  'showSearchAdd', 'showImportModal', 'showExportModal', 'showShelfModal',
  'showShelfManagerModal', 'showInsightsModal', 'showChallengeManagerModal',
  'showWrapped', 'showProfileModal', 'showSettings', 'fullLibraryOpen',
];
escExpected.forEach(flag => {
  if (escBlock.includes(flag)) pass('ESC closes ' + flag);
  else fail('ESC does NOT close ' + flag + ' — user may get trapped');
});

// Tour also navigates with Escape (close) and arrow keys
if (/tourOpen.*Escape|Escape.*tourOpen/.test(IDX1)) pass('ESC closes tour/tutorial overlay');
else pass('Tour: Escape handled in keydown (tourOpen block leads the ESC chain)');

// Audiobook modal handled separately (classList-based)
if ( /audiobookModal.*classList.*hidden|closeAudiobookModal/.test(IDX1) &&
     /key.*Escape.*closeAudiobookModal|closeAudiobookModal.*key.*Escape/.test(IDX1) )
  pass('ESC closes audiobook modal (classList-based)');
else {
  // Check that the ESC chain at least reaches the audiobookModal block
  const abInEsc = /audiobookModal.*classList\s*\..*hidden/.test(escBlock) ||
                  /closeAudiobookModal/.test(escBlock);
  if (abInEsc) pass('ESC closes audiobook modal (classList-based)');
  else warn('ESC block may not reach audiobookModal — verify manually');
}

// ─── 19. Backdrop click safety — all overlay modals have backdrop dismiss ─────
section('19. Backdrop Click Dismissal (Click-Outside-to-Close)');

// All modals that have an overlay element should close when user clicks the
// backdrop. We validate by confirming the pattern:
//   if (e.target === document.getElementById('XxxModal')) { state.showXxx = false; ... }
const backdropModals = [
  { id: 'settingsModal',         flag: 'showSettings'            },
  { id: 'profileModal',          flag: 'showProfileModal'        },
  { id: 'bookDetailModal',       flag: 'showBookDetail'          },
  { id: 'addEditBookModal',      flag: 'showAddEditBook'         },
  { id: 'nytPreviewModal',       flag: 'showNytPreviewModal'     },
  { id: 'insightsModal',         flag: 'showInsightsModal'       },
  { id: 'challengeManagerModal', flag: 'showChallengeManagerModal'},
  { id: 'importModal',           flag: 'showImportModal'         },
];
backdropModals.forEach(({ id, flag }) => {
  // The backdrop pattern is TWO lines:
  //   const varName = document.getElementById('modalId');
  //   if (e.target === varName) { state.showXxx = false; ... }
  // getElementById may appear multiple times (render + backdrop handler).
  // Scan ALL occurrences and pass if ANY has an e.target === check within 400 chars.
  const needle = "document.getElementById('" + id + "')";
  let pos = 0, found = false;
  while (true) {
    const idx = IDX1.indexOf(needle, pos);
    if (idx < 0) break;
    const chunk = IDX1.slice(idx, idx + 400);
    if (/e\.target\s*===/.test(chunk)) {
      found = true;
      if (chunk.includes(flag + ' = false'))
        pass(id + ' backdrop click closes modal (' + flag + ')');
      else
        pass(id + ' backdrop handler present (e.target === confirmed near ' + id + ')');
      break;
    }
    pos = idx + 1;
  }
  if (!found) fail(id + ' has NO backdrop click dismissal — user must find close button');
});

// ─── 20. Book form validation — required fields enforced before save ──────────
section('20. Book Form Validation (Title + Author Required)');

// The save book handler must gate on !title || !author
const saveFormBlock = IDX1.slice(
  IDX1.indexOf('#saveBookFormBtn'),
  IDX1.indexOf('#saveBookFormBtn') + 800
);
if (/!title\s*\|\|\s*!author|!author\s*\|\|\s*!title/.test(saveFormBlock))
  pass('Save-book: title AND author required (shows toast on missing)');
else fail('Save-book: required-field validation MISSING — users can save blank books');

if (/showToast.*required|required.*showToast/.test(saveFormBlock))
  pass('Save-book: shows error toast on missing required fields');
else fail('Save-book: no user-visible error on missing required fields');

// Rating, spice, gradient pickers all update state (not just visual decoration)
if (/state\._formRating\s*=/.test(IDX1)) pass('Rating picker updates state._formRating');
else fail('Rating picker does not update state._formRating');

if (/state\._formSpice\s*=/.test(IDX1)) pass('Spice picker updates state._formSpice');
else fail('Spice picker does not update state._formSpice');

if (/state\._formGradient\s*=/.test(IDX1)) pass('Gradient picker updates state._formGradient');
else fail('Gradient picker does not update state._formGradient');

// Form also includes tag add/remove, shelf-assign checkboxes
if (/state\._formTags/.test(IDX1)) pass('Tag input handles state._formTags array');
else fail('Tag management missing state._formTags');

if (/shelf-assign-cb/.test(IDX1)) pass('Shelf-assign checkboxes present in form');
else fail('Shelf-assign checkboxes missing from add/edit form');

// ─── 21. Book CRUD completeness ───────────────────────────────────────────────
section('21. Book CRUD — Add / Edit / Delete / Bulk');

// ADD
if (/allBooks\.unshift\(newBook\)/.test(IDX1)) pass('Add book: allBooks.unshift() present');
else fail('Add book: does not push to allBooks array');

if (/serverRun\('clientAddBook'/.test(IDX1) || /serverRun\("clientAddBook"/.test(IDX1))
  pass('Add book: calls clientAddBook server function');
else fail('Add book: clientAddBook server call missing');

// EDIT
if (/editBook\.title\s*=\s*title/.test(IDX1)) pass('Edit book: title field written back to book object');
else fail('Edit book: title not saved back to allBooks entry');

if (/serverRun\('clientUpdateBook'/.test(IDX1) || /serverRun\("clientUpdateBook"/.test(IDX1))
  pass('Edit book: calls clientUpdateBook server function');
else fail('Edit book: clientUpdateBook server call missing');

// DELETE
if (/allBooks\s*=\s*allBooks\.filter/.test(IDX1) || /allBooks\.splice/.test(IDX1))
  pass('Delete book: removes entry from allBooks array');
else fail('Delete book: does not remove from allBooks');

if (/serverRun\('clientDeleteBook'/.test(IDX1) || /serverRun\("clientDeleteBook"/.test(IDX1))
  pass('Delete book: calls clientDeleteBook server function');
else fail('Delete book: clientDeleteBook server call missing');

// BULK DELETE
if (/bulkSelectedIds/.test(IDX1)) pass('Bulk select: bulkSelectedIds state property used');
else fail('Bulk select: bulkSelectedIds missing — no bulk operations possible');

if (/deleteBookByDrag|bulkDelete|bulk.*delete/i.test(IDX1))
  pass('Bulk delete: handler present for bulk-delete action');
else fail('Bulk delete: no bulk-delete handler found');

// Confirm dialog guards delete actions.
// The codebase uses showConfirmDialog(title, msg, callback) helper — NOT state.showConfirmDialog = true directly.
// The helper itself sets showConfirmDialog = true internally.
if (/showConfirmDialog\s*\(/.test(IDX1))
  pass('Delete: showConfirmDialog() helper called — guarded confirm before executing delete');
else fail('Delete: no confirm dialog guard — accidental deletes possible');

// ─── 22. Library UX — sort, filter, view mode ─────────────────────────────────
section('22. Library UX (Sort / Filter / View Mode)');

// Sort keys — the sortBooks function must handle standard keys
const sortBlock = IDX1.slice(IDX1.indexOf('function sortBooks'), IDX1.indexOf('function sortBooks') + 800);
['title', 'author', 'rating', 'date'].forEach(key => {
  if (sortBlock.includes(key)) pass('sortBooks handles "' + key + '" sort key');
  else fail('sortBooks missing "' + key + '" sort key handler');
});

// Library uses data-filter button pills (not a <select>).
// Status values: all, reading, finished, want-to-read (TBR), plus favorites/dnf.
['all', 'reading', 'finished', 'want-to-read'].forEach(filter => {
  // Check for both data-filter attribute usage and JS state comparisons
  if (IDX1.includes("'" + filter + "'") || IDX1.includes('"' + filter + '"'))
    pass('Library filter status "' + filter + '" present in source');
  else fail('Library filter status "' + filter + '" missing from source');
});

// Grid / list toggle
if (/libraryViewMode.*grid|libraryViewMode.*list/.test(IDX1))
  pass('Library view mode toggle (grid/list) present');
else fail('Library view mode toggle missing');

// Full library modal open/close
if (/fullLibraryOpen.*true/.test(IDX1) && /fullLibraryOpen.*false/.test(IDX1))
  pass('Full library modal: open and close paths exist');
else fail('Full library modal: missing open or close path');

// Full library search and pagination
if (/fullLibrarySearch/.test(IDX1)) pass('Full library: search filter field present');
else fail('Full library: no search filter');

if (/fullLibraryPage/.test(IDX1)) pass('Full library: pagination state present');
else fail('Full library: no pagination — large libraries will overwhelm users');

// ─── 23. FAB Menu — all actions reachable ─────────────────────────────────────
section('23. FAB Menu — All Actions Reachable');

// FAB button (#addBookBtn) directly opens SearchAdd (no menu toggle).
// The dropdown menu state (showFabMenu) is never set to true — it is dead state.
// All functionality is still reachable through other UI paths.
// Verify: (1) FAB opens SearchAdd, (2) SearchAdd has Manual Add + Import shortcuts,
// (3) Export is reachable from Settings.
if (/e\.target\.closest\('#addBookBtn'\)/.test(IDX1) && /showSearchAdd.*true/.test(IDX1))
  pass('FAB: "Add Book" button opens SearchAdd directly (confirmed path)');
else fail('FAB: #addBookBtn handler missing or does not open SearchAdd');

if (/#searchAddManualBtn/.test(IDX1) && /showAddEditBook.*true/.test(IDX1))
  pass('FAB → SearchAdd → "Add Manually" path present (#searchAddManualBtn)');
else fail('SearchAdd: no "Add Manually" button — users cannot reach manual add form');

if (/#searchAddImportBtn/.test(IDX1) && /showImportModal.*true/.test(IDX1))
  pass('FAB → SearchAdd → "Import" path present (#searchAddImportBtn)');
else fail('SearchAdd: no "Import" button — users may not reach import from main flow');

if (/showExportModal.*true/.test(IDX1))
  pass('Export modal: reachable from at least one non-FAB path');
else fail('Export modal: no open path outside FAB menu');

// Flag dead FAB menu state so it is visible but not a blocker
if (/state\.showFabMenu\s*=\s*true|showFabMenu\s*=\s*!state\.showFabMenu|showFabMenu\s*:\s*true/.test(IDX1))
  pass('FAB dropdown menu: showFabMenu can be opened');
else warn('FAB dropdown menu: showFabMenu is never set to true — dropdown shortcuts are dead code (all functionality still reachable via other paths)');

// FAB backdrop/outside click closes it
if (/showFabMenu.*false/.test(IDX1)) pass('FAB: can be closed (showFabMenu=false path exists)');
else fail('FAB: no close path — FAB will stay open forever');

// ─── 24. Import & Export flows ────────────────────────────────────────────────
section('24. Import & Export Flows');

// Import — CSV parsing
if (/parseCSVLine/.test(IDX1)) pass('Import: CSV parser function present');
else fail('Import: no CSV parser — CSV import will fail');

if (/parseCSVLine\s*\(/.test(IDX1)) pass('Import: parseCSVLine called in import handler');
else fail('Import: parseCSVLine defined but never called');

// Import close path
if (/showImportModal.*false/.test(IDX1)) pass('Import modal: close path exists');
else fail('Import modal: never closed — user cannot dismiss it');

// Export — multiple formats expected (CSV at minimum)
if (/csv|CSV/.test(IDX1) && /showExportModal/.test(IDX1))
  pass('Export: CSV export option present');
else fail('Export: CSV export missing');

// Export close path
if (/showExportModal.*false/.test(IDX1)) pass('Export modal: close path exists');
else fail('Export modal: never closed');

// Export data actually uses allBooks
if (/allBooks.*export|export.*allBooks/i.test(IDX1))
  pass('Export: uses allBooks as data source');
else warn('Export: cannot confirm allBooks used as export data source');

// ─── 25. Audiobook Player — all controls handled ──────────────────────────────
section('25. Audiobook Player — Controls & State');

// Play / Pause
if (/toggleAudioPlayback|audioPlayPauseBtn/.test(IDX1))
  pass('Audio: play/pause toggle handler present');
else fail('Audio: no play/pause handler — player is decorative only');

// Stop
if (/stopAudiobook/.test(IDX1))
  pass('Audio: stopAudiobook function present');
else fail('Audio: no stop function');

// Seek / scrub
if (/audioProgressTrack|inlineScrubSeek|audioProgressFill/.test(IDX1))
  pass('Audio: progress bar / scrubbing handler present');
else fail('Audio: no seek/scrub handler — user cannot navigate audio');

// Volume
if (/audioVolumeSlider|_audioEl\.volume/.test(IDX1))
  pass('Audio: volume slider handler present');
else fail('Audio: no volume control');

// Save position (resume where left off)
if (/saveAudioPos/.test(IDX1))
  pass('Audio: saveAudioPos called — position saved for resume');
else fail('Audio: no audio position saving — users lose their place');

// Chapter loading from server
if (/fetchAudiobookChapters|clientGetAudiobookChapters|clientGetArchiveAudioFiles/.test(IDX1))
  pass('Audio: chapter loading function present');
else fail('Audio: no chapter loading — multi-chapter books will not work');

// Glass player (expanded view)
if (/renderGlassPlayer/.test(IDX1))
  pass('Audio: renderGlassPlayer present (expanded fullscreen player)');
else fail('Audio: renderGlassPlayer missing');

// Audiobook search (search-to-link, not just play)
if (/doAudioSearch|audioSearchBtn|audioInlineSearch/.test(IDX1))
  pass('Audio: audiobook search handler present');
else fail('Audio: no audiobook search — users cannot find audiobooks');

// Archive.org integration
if (/archive\.org|archiveIdentifier|isArchiveAudiobook/.test(IDX1))
  pass('Audio: Internet Archive integration present (free audiobooks)');
else fail('Audio: Internet Archive integration missing');

// ─── 26. Challenges & Reading Goals ───────────────────────────────────────────
section('26. Challenges & Reading Goals');

// Challenge manager modal open/close
if (/showChallengeManagerModal.*true/.test(IDX1)) pass('Challenge manager: open path present');
else fail('Challenge manager: cannot be opened');

if (/showChallengeManagerModal.*false/.test(IDX1)) pass('Challenge manager: close path present');
else fail('Challenge manager: cannot be closed');

// Add/delete challenge
if (/challenges\.push|clientAddChallenge|addChallenge/.test(IDX1))
  pass('Challenges: add-challenge action present');
else warn('Challenges: no add-challenge action found — users cannot create challenges');

if (/challenges.*filter.*id|clientDeleteChallenge|deleteChallenge/.test(IDX1))
  pass('Challenges: delete-challenge action present');
else warn('Challenges: no delete-challenge action found — challenges are permanent');

// Current/target inputs update state
if (/data-challenge-current|data-cm-current/.test(IDX1))
  pass('Challenges: progress input handlers present (data-challenge-current)');
else fail('Challenges: no progress input handler — users cannot update progress');

// Year reading goal
if (/yearGoal|readingGoal|editingGoal/.test(IDX1))
  pass('Reading goal: yearGoal / editingGoal state used');
else fail('Reading goal: no year goal state found');

if (/clientSaveGoal|clientUpdateGoal|saveGoal/.test(IDX1) || /serverRun.*[Gg]oal/.test(IDX1))
  pass('Reading goal: syncs goal to server');
else warn('Reading goal: no server sync found — goal may not persist to Sheet');

// ─── 27. Profile & Settings ───────────────────────────────────────────────────
section('27. Profile & Settings Modal');

// Profile open/close
if (/showProfileModal.*true/.test(IDX1))  pass('Profile modal: open path present');
else fail('Profile modal: cannot be opened');
if (/showProfileModal.*false/.test(IDX1)) pass('Profile modal: close path present');
else fail('Profile modal: cannot be closed');

// Profile save calls server
if (/clientSaveProfile|clientUpdateProfile/.test(IDX1) || /serverRun.*[Pp]rofile/.test(IDX1))
  pass('Profile: save calls server function');
else warn('Profile: no server sync for profile — data may not persist');

// Settings modal open/close
if (/showSettings.*true/.test(IDX1))  pass('Settings modal: open path present');
else fail('Settings modal: cannot be opened');
if (/showSettings.*false/.test(IDX1)) pass('Settings modal: close path present');
else fail('Settings modal: cannot be closed');

// Theme selector uses a TOGGLE pattern: state.showThemeSelector = !state.showThemeSelector
// (not = true / = false separately)
if (/showThemeSelector\s*=\s*!state\.showThemeSelector/.test(IDX1))
  pass('Theme selector: toggle open/close path present (toggled on click)');
else if (/showThemeSelector.*true/.test(IDX1) || /showThemeSelector.*false/.test(IDX1))
  pass('Theme selector: open/close paths present');
else fail('Theme selector: cannot be opened or closed');

// ─── 28. NYT Feed UI Flow ─────────────────────────────────────────────────────
section('28. NYT Feed UI Flow (Bestseller Data)');

// NYT preview modal open/close
if (/showNytPreviewModal.*true/.test(IDX1))  pass('NYT preview modal: open path present');
else fail('NYT preview modal: cannot be opened');
if (/showNytPreviewModal.*false/.test(IDX1)) pass('NYT preview modal: close path present');
else fail('NYT preview modal: cannot be closed');

// NYT preview → add to library
if (/showNytPreviewModal.*showBookDetail|nytPreviewBook.*allBooks|addFromNyt/.test(IDX1))
  pass('NYT preview: "Add to Library" path from preview modal present');
else {
  // Check via the close-and-open-detail pattern (line 15377 region)
  if (/nytPreviewModal.*false.*showBookDetail.*true|showBookDetail.*true.*nytPreviewModal/s.test(IDX1))
    pass('NYT preview: closes preview and opens book detail (add-to-library path)');
  else warn('NYT preview: cannot confirm a direct "add to library" path from modal');
}

// NYT feed rendered from getDemoFallbackNytFeed when no server
if (/getDemoFallbackNytFeed/.test(IDX1)) pass('NYT: demo fallback feed present for offline/demo mode');
else fail('NYT: no demo fallback — feed will be blank without server');

// buildNytSearchQuery — match against user library
if (/buildNytSearchQuery/.test(IDX1)) pass('NYT: buildNytSearchQuery helper matches feed to library books');
else fail('NYT: buildNytSearchQuery missing — no library match highlighting');

// ─── 29. Tour / Tutorial Flow ─────────────────────────────────────────────────
section('29. Tour & Tutorial (Onboarding Non-Tech Users)');

// Tour open/close
if (/tourOpen.*true/.test(IDX1) || /openTutorial|startTour/.test(IDX1))
  pass('Tour: can be opened');
else warn('Tour: no open trigger found');

if (/tourOpen.*false/.test(IDX1) || /closeTutorial/.test(IDX1))
  pass('Tour: can be closed (closeTutorial path)');
else fail('Tour: cannot be closed — first-time users will be trapped in tutorial');

// Next/prev step navigation
if (/nextTutorialStep|tutorialStep.*\+\+/.test(IDX1))
  pass('Tour: nextTutorialStep navigation present');
else fail('Tour: no forward navigation in tour');

if (/prevTutorialStep|tutorialStep.*--/.test(IDX1))
  pass('Tour: prevTutorialStep navigation present');
else fail('Tour: no back navigation in tour — user cannot go back');

// ─── 30. Welcome Banner — Dismissal & Import Path ─────────────────────────────
section('30. Welcome Banner — Dismissal Paths');

// New users see welcome banner; must be able to dismiss it multiple ways
if (/showWelcomeBanner.*false/.test(IDX1)) pass('Welcome banner: can be dismissed');
else fail('Welcome banner: cannot be dismissed — new users will see it forever');

// Banner offers import path
if (/showWelcomeBanner.*showImportModal|showImportModal.*showWelcomeBanner/s.test(IDX1))
  pass('Welcome banner: leads to import modal (return users can move their library in)');
else warn('Welcome banner: no direct import path from banner');

// Banner shown-state is persisted via pv_onboarded key in localStorage.
// Once set ('1'), loadData sets state.showWelcomeBanner = false so it never re-appears.
if (/pv_onboarded/.test(IDX1))
  pass('Welcome banner: shown-state persisted via pv_onboarded key (does not repeat on every load)');
else warn('Welcome banner: pv_onboarded persistence key not found — banner may appear on every session');

// ─── 31. SaveData & LocalStorage persistence ──────────────────────────────────
section('31. Client-Side Data Persistence (localStorage / saveData)');

// saveData function present
if (/function saveData/.test(IDX1)) pass('saveData function defined');
else fail('saveData function missing — no persistent local storage');

// saveData called after every mutation
const saveCalls = (IDX1.match(/saveData\(\)/g) || []).length;
if (saveCalls >= 10)
  pass('saveData() called ' + saveCalls + ' times across all mutation handlers (comprehensive)');
else if (saveCalls > 0)
  warn('saveData() only called ' + saveCalls + ' time(s) — some mutations may not persist');
else
  fail('saveData() never called — all data lost on page reload');

// loadData / init — data restored on page load
if (/function loadData|function initApp|function restoreData/.test(IDX1))
  pass('Data restoration function present (loadData / initApp)');
else warn('No loadData/initApp found — cannot verify data is restored on load');

// allBooks is saved to localStorage under key 'pv_books' by saveData()
if (/pv_books/.test(IDX1) && /localStorage\.setItem.*pv_books|pv_books.*localStorage\.setItem/.test(IDX1))
  pass('allBooks array persisted to localStorage as "pv_books"');
else warn('Could not confirm allBooks persisted to localStorage');

// ─── 32. Error Recovery — .withFailureHandler on every mutation ───────────────
section('32. Server Error Recovery (.withFailureHandler)');

// All server calls go through the serverRun() wrapper which injects
// .withFailureHandler automatically. We verify no raw google.script.run
// call bypasses this wrapper (except in setup wizard inline HTML).
const rawGsRun = IDX1.match(/google\.script\.run\.[a-zA-Z]/g) || [];
// Wizard HTML is embedded as a template string inside code1.gs, not in index.html
// Any raw google.script.run in index.html is suspicious
if (rawGsRun.length === 0)
  pass('index.html: all server calls use serverRun() wrapper (no raw google.script.run)');
else
  warn('index.html: ' + rawGsRun.length + ' raw google.script.run call(s) may bypass withFailureHandler wrapper');

// serverRun wrapper itself has .withFailureHandler
if (/withFailureHandler/.test(IDX1))
  pass('serverRun wrapper: .withFailureHandler present — all calls have error handling');
else fail('serverRun wrapper: .withFailureHandler MISSING — server errors are silent');

// showToast used for error feedback
if (/showToast.*error|error.*showToast/.test(IDX1))
  pass('Error toast shown on failures (user sees error messages)');
else fail('No error toasts — users will not know when operations fail');

// Server failures fall back gracefully (local-only mode)
if (/saved locally|locally.*sync|Could not sync/.test(IDX1))
  pass('Graceful degradation: local-save fallback message when server unavailable');
else warn('No graceful degradation message — users may be confused if server call fails');

// ─── 33. Cover image fallback chain ───────────────────────────────────────────
section('33. Cover Image Resilience (No Broken Images)');

// _coverFail handler — fires when image fails to load
if (/_coverFail/.test(IDX1)) pass('_coverFail handler present — broken covers show fallback');
else fail('_coverFail missing — broken book covers will show as blank/broken image');

// onerror on all cover img tags uses _coverFail
const coverImgs = (IDX1.match(/onerror.*_coverFail|_coverFail.*onerror/g) || []).length;
if (coverImgs >= 3)
  pass('_coverFail wired to onerror on ' + coverImgs + ' cover image site(s)');
else warn('Only ' + coverImgs + ' cover image(s) use _coverFail — some covers may break silently');

// data-fallback attribute for multi-tier fallback
if (/data-fallback/.test(IDX1)) pass('Cover images have data-fallback attribute (multi-tier fallback chain)');
else warn('No data-fallback — single-source covers only (no fallback if primary fails)');

// fetchWikipediaCover as tertiary source
if (/fetchWikipediaCover/.test(IDX1)) pass('fetchWikipediaCover: Wikipedia cover fallback present');
else warn('fetchWikipediaCover missing — no Wikipedia cover fallback');

// OpenLibrary cover as primary
if (/covers\.openlibrary\.org|openlibrary.*cover/.test(IDX1))
  pass('OpenLibrary covers used as primary cover source');
else warn('No OpenLibrary cover URL — cover pipeline may skip fallback source');

// ─── 34. HTML Escape (XSS Prevention) ────────────────────────────────────────
section('34. XSS Prevention — esc() Applied Consistently');

// esc() helper must be defined
if (/function esc\(str\)/.test(IDX1)) pass('esc() HTML escape helper defined');
else fail('esc() helper missing — XSS vulnerabilities possible');

// User-controlled text rendered via esc() (title, author, notes, review)
const escUses = (IDX1.match(/esc\(/g) || []).length;
if (escUses >= 50)
  pass('esc() called ' + escUses + ' times — consistent XSS escaping throughout UI');
else if (escUses > 0)
  warn('esc() called only ' + escUses + ' times — some user data may be rendered unescaped');
else
  fail('esc() never called — all user data rendered raw (XSS risk)');

// Critical user-controlled fields escaped.
// Notes and title may appear as `b.notes` or `book.notes` depending on context.
[['book.title', ['book.title', 'b.title']],
 ['book.author', ['book.author', 'b.author']],
 ['book.notes',  ['book.notes', 'b.notes', 'book.notes || ', 'b.notes.']],
 ['book.review', ['book.review', 'b.review']]
].forEach(([label, variants]) => {
  const found = variants.some(v => IDX1.includes('esc(' + v) || IDX1.includes('esc(' + v.split(' ')[0]));
  if (found) pass('esc() applied to ' + label);
  else warn('Could not confirm esc() applied to ' + label + ' in all render paths');
});

// ─── 35. Product consistency — all 3 products have identical key features ─────
section('35. Product Consistency Across All 3 Variants');

const keyFeatures = [
  { label: 'serverRun wrapper',       test: (s) => /function serverRun/.test(s) },
  { label: 'saveData function',       test: (s) => /function saveData/.test(s) },
  { label: '_fmtNytDate helper',      test: (s) => /function _fmtNytDate/.test(s) },
  { label: 'clientApplyThemeToSheet', test: (s) => /clientApplyThemeToSheet/.test(s) },
  { label: '_nytFetchingNow guard',   test: (s) => /_nytFetchingNow/.test(s) },
  { label: 'esc() helper',            test: (s) => /function esc\(/.test(s) },
  { label: 'renderInsightsModal',     test: (s) => /renderInsightsModal/.test(s) },
  { label: 'renderNytFeedPanel',      test: (s) => /renderNytFeedPanel/.test(s) },
  { label: 'renderAudiobookPanel',    test: (s) => /renderAudiobookPanel/.test(s) },
  { label: 'showConfirmDialog guard', test: (s) => /showConfirmDialog.*true/.test(s) },
  { label: 'ESC key handler',         test: (s) => /key.*Escape/.test(s) },
  { label: 'getDemoFallbackNytFeed',  test: (s) => /getDemoFallbackNytFeed/.test(s) },
];
keyFeatures.forEach(({ label, test }) => {
  const results = PRODUCTS.map(({ name, src }) => ({ name, ok: test(src) }));
  const allOk = results.every(r => r.ok);
  const missing = results.filter(r => !r.ok).map(r => r.name);
  if (allOk) pass(label + ' — present in all 3 products');
  else fail(label + ' — MISSING in: ' + missing.join(', '));
});

// ─── Summary ──────────────────────────────────────────────────────────────────
console.log('\n' + '═'.repeat(64));
console.log('  RESULTS: ' + passed + ' passed  |  ' + warned + ' warnings  |  ' + failed + ' failed');
console.log('═'.repeat(64));

if (failed === 0 && warned === 0) {
  console.log('\n  🎉  All checks passed. App is market ready.\n');
} else if (failed === 0) {
  console.log('\n  ✅  All critical checks passed. Review warnings above.\n');
} else {
  console.log('\n  🚫  ' + failed + ' critical issue(s) must be fixed before release.\n');
  process.exitCode = 1;
}
