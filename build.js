/**
 * Build script: obfuscates the <script> block in index.html
 * and writes the result to index_dist.html
 *
 * Run: node build.js
 */
const fs   = require('fs');
const path = require('path');
const JavaScriptObfuscator = require('javascript-obfuscator');

const SRC  = path.join(__dirname, 'index.html');
const DIST = path.join(__dirname, 'index_dist.html');

const html = fs.readFileSync(SRC, 'utf8');

// Find the one inline <script> block (not the CDN src= ones)
const OPEN  = '  <script>\n';
const CLOSE = '  </script>';

const start = html.indexOf(OPEN);
const end   = html.lastIndexOf(CLOSE);

if (start === -1 || end === -1) {
  console.error('Could not locate <script> block. Check OPEN/CLOSE markers.');
  process.exit(1);
}

const before = html.slice(0, start + OPEN.length);
const jsCode = html.slice(start + OPEN.length, end);
const after  = html.slice(end);

console.log(`JS block: ${jsCode.length.toLocaleString()} characters`);
console.log('Obfuscating...');

const result = JavaScriptObfuscator.obfuscate(jsCode, {
  // ── Strength settings ─────────────────────────────────────────────────
  compact: true,                        // single line output
  controlFlowFlattening: false,         // keep off — your code is large, flattening explodes size
  deadCodeInjection: false,             // keep off — size explosion
  debugProtection: false,               // don't annoy yourself during testing
  disableConsoleOutput: false,          // keep console.error for your own debugging
  identifierNamesGenerator: 'hexadecimal', // vars become _0x1a2b3c style
  renameGlobals: false,                 // MUST be false — google.script.run must stay intact
  rotateStringArray: true,
  selfDefending: false,                 // makes code slow, not worth it
  splitStrings: false,                  // size explosion on 15k lines
  stringArray: true,                    // encode string literals
  stringArrayEncoding: ['base64'],      // harder to read encoded strings
  stringArrayThreshold: 0.75,          // encode 75% of strings
  transformObjectKeys: false,           // keep off — google.script run calls use object keys
  unicodeEscapeSequence: false          // keep off — size explosion
});

const obfuscated = result.getObfuscatedCode();
console.log(`Obfuscated: ${obfuscated.length.toLocaleString()} characters`);

const output = before + obfuscated + '\n' + after;
fs.writeFileSync(DIST, output, 'utf8');

console.log(`\nDone! → index_dist.html (${(output.length / 1024).toFixed(0)} KB)`);
console.log('Paste the contents of index_dist.html into Apps Script as your index.html');
