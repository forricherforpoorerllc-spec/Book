/**
 * Build script: obfuscates the inline <script> block in each product HTML
 * and writes the result to <name>_dist.html.
 *
 * Run: node build.js                  (build all 3 products)
 *      node build.js index            (build just Product 1)
 *      node build.js index2 index3    (build a subset)
 */
const fs   = require('fs');
const path = require('path');
const JavaScriptObfuscator = require('javascript-obfuscator');

const PRODUCTS = ['index', 'index2', 'index3'];

const OPEN  = '  <script>\n';
const CLOSE = '  </script>';

function buildOne(name) {
	const SRC  = path.join(__dirname, `${name}.html`);
	const DIST = path.join(__dirname, `${name}_dist.html`);

	if (!fs.existsSync(SRC)) {
		console.log(`\n[${name}] SKIP — ${name}.html not found`);
		return;
	}

	const html = fs.readFileSync(SRC, 'utf8');
	const start = html.indexOf(OPEN);
	const end   = html.lastIndexOf(CLOSE);

	if (start === -1 || end === -1) {
		console.error(`[${name}] ERROR — could not locate inline <script> block.`);
		process.exitCode = 1;
		return;
	}

	const before = html.slice(0, start + OPEN.length);
	const jsCode = html.slice(start + OPEN.length, end);
	const after  = html.slice(end);

	console.log(`\n[${name}] JS block: ${jsCode.length.toLocaleString()} chars — obfuscating…`);

	const result = JavaScriptObfuscator.obfuscate(jsCode, {
		compact: true,
		controlFlowFlattening: false,
		deadCodeInjection: false,
		debugProtection: false,
		disableConsoleOutput: false,
		identifierNamesGenerator: 'hexadecimal',
		renameGlobals: false,                 // google.script.run must stay intact
		rotateStringArray: true,
		selfDefending: false,
		splitStrings: false,
		stringArray: true,
		stringArrayEncoding: ['base64'],
		stringArrayThreshold: 0.75,
		transformObjectKeys: false,           // google.script.run object keys must stay intact
		unicodeEscapeSequence: false,
	});

	const obfuscated = result.getObfuscatedCode();
	const output = before + obfuscated + '\n' + after;
	fs.writeFileSync(DIST, output, 'utf8');

	console.log(`[${name}] Obfuscated: ${obfuscated.length.toLocaleString()} chars`);
	console.log(`[${name}] Done → ${path.basename(DIST)} (${(output.length / 1024).toFixed(0)} KB)`);
}

const targets = process.argv.slice(2).filter(Boolean);
const list = targets.length ? targets : PRODUCTS;
list.forEach(buildOne);

console.log('\nAll builds complete. Next: use a release script (or manual clasp push) per product.');
