/**
 * Release script: ships a specific product variant to its Apps Script project.
 *
 * Usage:
 *   node release.js 1            # ships Product 1 (Romantic / index.html)
 *   node release.js 2            # ships Product 2 (Horizon / index2.html)
 *   node release.js 3            # ships Product 3 (Blossom / index3.html)
 *   node release.js 2 --raw      # skip obfuscation (push readable source)
 *
 * What it does (atomic, with restore-on-error):
 *   1. node build.js indexN                  → produces indexN_dist.html
 *   2. swap .clasp.json with .clasp.productN.json
 *   3. set PRODUCT_VARIANT = 'indexN' in code1.gs
 *   4. rename index.html → __backup__/index.html
 *      copy indexN_dist.html → index.html
 *   5. npx clasp push -f
 *   6. restore everything in step 2-4 (always, even on failure)
 */
const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');

const ROOT = __dirname;
const args = process.argv.slice(2);
const productNum = args.find(a => /^[123]$/.test(a));
const useRaw     = args.includes('--raw');

if (!productNum) {
	console.error('Usage: node release.js <1|2|3> [--raw]');
	process.exit(1);
}

const productKey   = productNum === '1' ? 'index' : `index${productNum}`;
const productClasp = path.join(ROOT, `.clasp.product${productNum}.json`);
const activeClasp  = path.join(ROOT, '.clasp.json');
const indexHtml    = path.join(ROOT, 'index.html');
const variantHtml  = path.join(ROOT, `${productKey}.html`);
const distHtml     = path.join(ROOT, `${productKey}_dist.html`);
const codeFile     = path.join(ROOT, 'code1.gs');

if (!fs.existsSync(productClasp)) {
	console.error(`Missing ${path.basename(productClasp)} — create it with the product's scriptId.`);
	process.exit(1);
}
if (!fs.existsSync(variantHtml)) {
	console.error(`Missing ${productKey}.html`);
	process.exit(1);
}

// Save originals so we can restore on exit/error
const origClasp   = fs.readFileSync(activeClasp, 'utf8');
const origIndex   = fs.readFileSync(indexHtml,   'utf8');
const origCode    = fs.readFileSync(codeFile,    'utf8');
const origVariant = fs.readFileSync(variantHtml, 'utf8');

let restored = false;
function restore() {
	if (restored) return;
	restored = true;
	try { fs.writeFileSync(activeClasp, origClasp);   } catch (e) {}
	try { fs.writeFileSync(indexHtml,   origIndex);   } catch (e) {}
	try { fs.writeFileSync(codeFile,    origCode);    } catch (e) {}
	try { fs.writeFileSync(variantHtml, origVariant); } catch (e) {}
	console.log('\n✓ Restored original .clasp.json, index.html, code1.gs, ' + path.basename(variantHtml));
}
process.on('exit',  restore);
process.on('SIGINT', () => { restore(); process.exit(130); });

try {
	// 1. Build (skip if --raw)
	if (!useRaw) {
		console.log(`\n[1/4] Obfuscating ${productKey}.html…`);
		execSync(`node build.js ${productKey}`, { stdio: 'inherit', cwd: ROOT });
	} else {
		console.log(`\n[1/4] --raw flag set, skipping obfuscation`);
	}

	// 2. Swap clasp config
	console.log(`\n[2/4] Activating ${path.basename(productClasp)}…`);
	fs.copyFileSync(productClasp, activeClasp);

	// 3. Edit PRODUCT_VARIANT in code1.gs
	console.log(`[3/4] Setting PRODUCT_VARIANT = '${productKey}' in code1.gs…`);
	// Use a quote-agnostic pattern — matches straight or curly single-quotes
	const productVariantRe = /var PRODUCT_VARIANT\s*=\s*[''\u2018\u2019][^''\u2018\u2019]*[''\u2018\u2019];/;
	if (!productVariantRe.test(origCode)) {
		throw new Error('Could not find PRODUCT_VARIANT line in code1.gs');
	}
	const newCode = origCode.replace(
		productVariantRe,
		`var PRODUCT_VARIANT = '${productKey}';`
	);
	fs.writeFileSync(codeFile, newCode);

	// 4. Stage HTML under its real name (so doGet → createHtmlOutputFromFile(productKey) resolves)
	const sourceHtml = useRaw ? variantHtml : (fs.existsSync(distHtml) ? distHtml : variantHtml);
	console.log(`[4/4] Staging ${path.basename(sourceHtml)} as ${productKey}.html…`);
	if (sourceHtml !== variantHtml) {
		fs.copyFileSync(sourceHtml, variantHtml);
	}

	// 5. Push
	console.log(`\n→ Pushing to Apps Script project for Product ${productNum}…\n`);
	execSync('npx clasp push -f', { stdio: 'inherit', cwd: ROOT });

	console.log(`\n✓ Product ${productNum} pushed successfully.`);
	console.log(`Now run: npx clasp deploy   (or use a saved deployment ID)`);
} catch (err) {
	console.error(`\n✗ Release failed: ${err.message}`);
	process.exitCode = 1;
}
// restore() runs automatically on exit
