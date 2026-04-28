// Syntax validator for index.html / index2.html / index3.html
// Run:  node validate.js
// Reports the exact line/column of any JS syntax error in inline <script> blocks.

const fs = require('fs');
const path = require('path');

const FILES = ['index.html', 'index2.html', 'index3.html'];

function findScriptBlocks(html) {
  // Match <script> ... </script> where the open tag has no src= attribute.
  const re = /<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g;
  const blocks = [];
  let m;
  while ((m = re.exec(html))) {
    const body = m[1];
    if (!body.trim()) continue;
    const startOffset = m.index + m[0].indexOf('>') + 1;
    const startLine = html.slice(0, startOffset).split('\n').length;
    blocks.push({ body, startLine, startOffset });
  }
  return blocks;
}

function validateFile(file) {
  if (!fs.existsSync(file)) {
    console.log('SKIP (missing): ' + file);
    return;
  }
  const html = fs.readFileSync(file, 'utf8');
  const blocks = findScriptBlocks(html);
  console.log('\n=== ' + file + ' ===');
  console.log('Inline <script> blocks: ' + blocks.length);
  let errors = 0;
  blocks.forEach((blk, i) => {
    try {
      // eslint-disable-next-line no-new-func
      new Function(blk.body);
      console.log('  block ' + (i + 1) + ' @ html line ' + blk.startLine + ' : OK (' + blk.body.length + ' chars)');
    } catch (e) {
      errors++;
      console.log('  block ' + (i + 1) + ' @ html line ' + blk.startLine + ' : ERROR');
      console.log('    ' + e.message);
      // Report approximate html line from the error.
      // V8 errors include "<anonymous>:LINE" sometimes.
      const stack = String(e.stack || '');
      const m = stack.match(/<anonymous>:(\d+)(?::(\d+))?/);
      if (m) {
        const innerLine = parseInt(m[1], 10);
        const innerCol = m[2] ? parseInt(m[2], 10) : 0;
        const htmlLine = blk.startLine + innerLine - 1;
        console.log('    -> html line ' + htmlLine + (innerCol ? ', col ' + innerCol : ''));
        // Print 4 lines of context.
        const lines = html.split('\n');
        for (let L = Math.max(htmlLine - 2, 1); L <= Math.min(htmlLine + 2, lines.length); L++) {
          const marker = (L === htmlLine) ? '>>>' : '   ';
          console.log('    ' + marker + ' ' + L + ' | ' + lines[L - 1]);
        }
      }
    }
  });
  if (errors === 0) console.log('All blocks parsed cleanly.');
}

FILES.forEach(validateFile);
