const fs = require('fs');
const path = require('path');
const file = path.join(__dirname, 'index.html');
const text = fs.readFileSync(file, 'utf8');
const match = text.match(/<script[^>]*>([\s\S]*?)<\/script>/);
if (!match) {
  console.log('NO SCRIPT');
  process.exit(1);
}
const script = match[1];
try {
  new Function(script);
  console.log('PARSE OK');
} catch (e) {
  console.log('ERR', e.message);
  if (e.loc && e.loc.line) {
    const line = e.loc.line;
    const lines = script.split('\n');
    const start = Math.max(0, line - 5);
    console.log('LINE', line);
    console.log(lines.slice(start, line + 2).map((l, i) => `${start + i + 1}: ${l}`).join('\n'));
  }
}
