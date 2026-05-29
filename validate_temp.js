const fs = require("fs");
const path = "c:/Users/ESHA/Desktop/Codex/My Project/temp_script.js";
try {
  new Function(fs.readFileSync(path, 'utf8'));
  console.log('VALID_JS');
} catch (e) {
  console.error('JS_ERROR', e.message);
  process.exit(1);
}