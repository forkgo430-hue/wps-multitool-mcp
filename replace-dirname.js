/** 一次性脚本（用完删除）：把仓库内所有 wps-multitool-mcp 路径引用改为 wps-multitool-mcp。 */
const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const OLD = 'wps-multitool-mcp';
const NEW = 'wps-multitool-mcp';
const SKIP = new Set(['node_modules', 'dist', '.git', 'docs', 'logs']);

function walk(d, out = []) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    if (SKIP.has(e.name)) continue;
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p, out);
    else if (/\.(ts|js|json|md|yml|yaml|ps1|html|xml|toml)$/.test(e.name)) out.push(p);
  }
  return out;
}

let files = 0;
let hits = 0;
for (const f of walk(ROOT)) {
  const src = fs.readFileSync(f, 'utf8');
  if (!src.includes(OLD)) continue;
  const n = src.split(OLD).length - 1;
  fs.writeFileSync(f, src.split(OLD).join(NEW), 'utf8');
  console.log(`${String(n).padStart(3)}  ${path.relative(ROOT, f)}`);
  files++;
  hits += n;
}
console.log(`\n共 ${files} 个文件、${hits} 处替换`);
