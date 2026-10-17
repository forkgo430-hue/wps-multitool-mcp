#!/usr/bin/env node
/**
 * Input: scripts/wps-com.ps1（读它的 switch ($Action) 分支）
 * Output: src/client/action-routes.generated.ts（action -> 应用 的路由表）
 * Pos: 路由表生成器。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要生成而不是手写
 * ----------------------
 * 桥是按应用拆子进程的，所以每个 action 必须知道该走哪一路。手写 256 条必错且会腐坏。
 * 这里直接从 wps-com.ps1 的 case 体里推断：该 case 里用到了哪个应用的对象
 *   Get-WpsExcel / $excel  -> excel
 *   Get-WpsWord  / $word   -> word
 *   Get-WpsPpt   / $ppt    -> ppt
 * 命中多个（或一个都没有，比如 ping / save / wpsStatus 这类跨应用或纯基础设施动作）-> common。
 *
 * 用法：node scripts/gen-action-routes.js
 * 改完 wps-com.ps1 就该重跑一次，然后 `npm run build`。
 */

const fs = require('fs');
const path = require('path');

const PS = path.join(__dirname, 'wps-com.ps1');
const OUT = path.join(__dirname, '..', 'src', 'client', 'action-routes.generated.ts');

const raw = fs.readFileSync(PS, 'utf8').replace(/^\uFEFF/, '');
const lines = raw.split(/\r?\n/);

const switchIdx = lines.findIndex((l) => l.trim().startsWith('switch ($Action)'));
if (switchIdx < 0) {
  console.error('未在 wps-com.ps1 里找到 switch ($Action)');
  process.exit(1);
}

const CASE_RE = /^\s*"([A-Za-z_][A-Za-z0-9_]*)"\s*\{\s*$/;

/**
 * 剥掉 PowerShell 注释（# 行注释 与 <# #> 块注释），保留字符串字面量里的 #。
 * 只用于路由推断，不求 100% 语法精确 —— 但必须避免"注释里的函数名被当成代码"。
 */
function stripComments(src) {
  const out = [];
  let i = 0;
  let inSingle = false;
  let inDouble = false;
  let inBlock = false;
  while (i < src.length) {
    const c = src[i];
    const next = src[i + 1];
    if (inBlock) {
      if (c === '#' && next === '>') {
        inBlock = false;
        i += 2;
        continue;
      }
      if (c === '\n') out.push('\n');
      i += 1;
      continue;
    }
    if (!inSingle && !inDouble && c === '<' && next === '#') {
      inBlock = true;
      i += 2;
      continue;
    }
    if (c === "'" && !inDouble) inSingle = !inSingle;
    else if (c === '"' && !inSingle) inDouble = !inDouble;
    else if (c === '#' && !inSingle && !inDouble) {
      // 行注释：跳到行尾
      while (i < src.length && src[i] !== '\n') i += 1;
      continue;
    }
    out.push(c);
    i += 1;
  }
  return out.join('');
}
const cases = [];
for (let i = switchIdx; i < lines.length; i++) {
  const m = lines[i].match(CASE_RE);
  if (m) cases.push({ name: m[1], line: i });
}
if (cases.length === 0) {
  console.error('未解析到任何 case');
  process.exit(1);
}

const table = {};
const duplicates = [];
for (let i = 0; i < cases.length; i++) {
  const start = cases[i].line;
  const end = i + 1 < cases.length ? cases[i + 1].line : lines.length;
  // 必须先剥掉注释再扫：注释里会提到 Get-WpsExcel 之类的函数名（例如 wpsStatus 的说明注释
  // 提到"对照 getAppInfo 那个走 Get-WpsExcel/Word/Ppt"），照抄会把 ping 误判成 excel。
  const body = stripComments(lines.slice(start, end).join('\n'));

  const hits = new Set();
  for (const m of body.matchAll(/Get-Wps(Excel|Word|Ppt)\b/g)) hits.add(m[1].toLowerCase());
  for (const m of body.matchAll(/\$(excel|word|ppt)\b/g)) hits.add(m[1].toLowerCase());
  for (const m of body.matchAll(/ComObject\s+'?(Ket|Kwps|Kwpp)\.Application/g)) {
    hits.add(m[1] === 'Ket' ? 'excel' : m[1] === 'Kwps' ? 'word' : 'ppt');
  }

  // PowerShell 的 switch 是【第一个匹配的分支生效】，同名分支后面的都是死代码。
  // 所以同名时保留第一次出现的判定，并把重复项报出来（本次实测有 4 个重复）。
  const name = cases[i].name;
  if (name in table) {
    duplicates.push(name);
    continue;
  }
  table[name] = hits.size === 1 ? [...hits][0] : 'common';
}

const counts = {};
for (const v of Object.values(table)) counts[v] = (counts[v] || 0) + 1;

const header = `/**
 * 本文件由 scripts/gen-action-routes.js 自动生成，请勿手改。
 * 源：scripts/wps-com.ps1 的 switch ($Action) 分支。
 * 重新生成：node scripts/gen-action-routes.js
 *
 * 统计：${Object.entries(counts).map(([k, v]) => `${k}=${v}`).join('  ')}  合计 ${cases.length}
 */

import type { AppKey } from './action-routes';

export const ACTION_ROUTES: Record<string, AppKey> = {
`;

const body = Object.keys(table)
  .sort()
  .map((k) => `  ${JSON.stringify(k)}: '${table[k]}',`)
  .join('\n');

fs.writeFileSync(OUT, header + body + '\n};\n', 'utf8');

console.log(`已生成 ${path.relative(path.join(__dirname, '..'), OUT)}`);
console.log(`case 总数 ${cases.length}，去重后 ${Object.keys(table).length}：${Object.entries(counts).map(([k, v]) => `${k}=${v}`).join('  ')}`);
if (duplicates.length) {
  console.warn(`注意：有 ${duplicates.length} 个重复 case 名（switch 里后者是死代码，已按第一个取值）：${[...new Set(duplicates)].join(', ')}`);
}
