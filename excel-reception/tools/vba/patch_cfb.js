// usage: node patch_cfb.js in.bin out.bin manifest.json   manifest: {"delete":[paths], "replace":{path: file}}
const CFB = require('cfb'), fs = require('fs');
const [,, inp, outp, manifestPath] = process.argv;
const m = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
const cfb = CFB.read(fs.readFileSync(inp), { type: 'buffer' });
for (const p of m.delete) { if (CFB.find(cfb, p)) CFB.utils.cfb_del(cfb, p); else console.error('missing', p); }
for (const [p, f] of Object.entries(m.replace)) {
  const buf = fs.readFileSync(f);
  if (CFB.find(cfb, p)) CFB.utils.cfb_del(cfb, p);
  CFB.utils.cfb_add(cfb, p, buf);
}
fs.writeFileSync(outp, CFB.write(cfb, { type: 'buffer' }));
console.log('entries:', cfb.FullPaths.length);
