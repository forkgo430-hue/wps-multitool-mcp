/**
 * 一次性诊断：找出"TS 传了参数、PS 却根本不读"的工具（静默丢弃 = 假成功）。
 *
 * 起因：验证者发现 wps_ppt_set_table_row_style 传嵌套 style，PS 读扁平 $p.backgroundColor，
 * 结果"返回成功但什么都没做"。这类 bug 最危险 —— 不报错，用户以为改完了。
 *
 * 做法（粗匹配，宁可多报不可漏报）：
 *   1. 从 wps-com.ps1 每个 case 块里收集 $p.<键>
 *   2. 从 dist/tools 下的编译产物里抓 executeMethod('<action>', {字面量})
 *   3. 字面量的键如果不在该 action 的 $p 集合里 → 列为候选
 *
 * 已知局限（报出来的不一定都是 bug，需人工过一遍）：
 *   - PS 可能用 $p 整体解构（$p.name 之外还有 $params["x"]）
 *   - TS 可能把参数放进嵌套对象由 PS 侧另一层解析
 *   - 动态拼参数的调用点抓不到（不报 = 漏报，不是安全）
 *
 * 用法：node scripts/manual/param-shape-audit.js
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const ps = fs.readFileSync(path.join(ROOT, 'scripts', 'wps-com.ps1'), 'utf8');

// 1. PS 侧每个 case 读了哪些 $p.<键>
const caseRe = /^\s{4}"([A-Za-z_][A-Za-z0-9_]*)"\s*\{/gm;
const marks = [];
let m;
while ((m = caseRe.exec(ps))) marks.push([m[1], m.index]);
const psKeys = {};
for (let i = 0; i < marks.length; i++) {
  const end = i + 1 < marks.length ? marks[i + 1][1] : ps.length;
  const body = ps.slice(marks[i][1], end);
  const set = new Set();
  const kr = /\$p\.([A-Za-z_][A-Za-z0-9_]*)/g;
  let k;
  while ((k = kr.exec(body))) set.add(k[1]);
  // PS 也可能读 $p["x"] / $p.$x
  const br = /\$p\[["']([A-Za-z0-9_]+)["']\]/g;
  while ((k = br.exec(body))) set.add(k[1]);
  psKeys[marks[i][0]] = set;
}

// 2. TS 侧（编译产物）每个 action 收到的字面量键
const files = [];
const walk = (d) => {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) {
      if (e.name !== 'node_modules') walk(p);
    } else if (e.name.endsWith('.js')) files.push(p);
  }
};
walk(path.join(ROOT, 'dist', 'tools'));

const rows = [];
for (const f of files) {
  const src = fs.readFileSync(f, 'utf8');
  const er = /executeMethod\(\s*'([A-Za-z0-9_]+)'\s*,\s*\{([^{}]*)\}/g;
  let x;
  while ((x = er.exec(src))) {
    const act = x[1];
    const obj = x[2];
    if (!psKeys[act]) continue;
    const keys = obj
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean)
      .map((s) => (s.indexOf(':') > 0 ? s.slice(0, s.indexOf(':')) : s).trim())
      .filter((s) => /^[A-Za-z_][A-Za-z0-9_]*$/.test(s));
    const missing = keys.filter((k) => !psKeys[act].has(k));
    if (missing.length) {
      rows.push({
        act,
        file: path.relative(ROOT, f).replace(/\\/g, '/'),
        missing: missing.join(','),
      });
    }
  }
}

rows.sort((a, b) => a.act.localeCompare(b.act));
console.log(`候选（TS 传了、PS 不读）共 ${rows.length} 个调用点：\n`);
for (const r of rows) console.log(`  ${r.act.padEnd(34)} ${r.missing.padEnd(28)} ${r.file}`);
console.log('\n下一步：逐个人工确认 —— 真丢参数的进 P1 修，PS 侧另有解析的标掉。');
