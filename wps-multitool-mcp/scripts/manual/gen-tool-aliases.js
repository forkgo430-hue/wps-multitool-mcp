#!/usr/bin/env node
/**
 * Input: scripts/manual/tool-aliases.json（人维护的真相源：旧名 → 规范名）
 * Output: src/server/tool-aliases.generated.ts（运行时用的别名表 + 解析函数）
 * Pos: 别名表生成器。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么生成成 TS 常量，而不是运行期读 JSON
 * ----------------------------------------
 * package.json 的 files 只发 dist + scripts/wps-com.ps1 + skills，scripts/manual/*.json 不进包。
 * 运行期读它会在用户机器上 ENOENT —— 而这个表恰好是每次 callTool 都要查的，出错即全站不可用。
 *
 * 生成器只做【形状校验】，不查目标是否存在
 * --------------------------------------
 * "别名目标必须在注册表里"这条要在 alias-verify.js 里查，因为这里拿不到可信的注册表：
 * 生成器的输入就是 registry 要 import 的表，先 require(dist) 会把旧的编译产物当真相源（鸡生蛋）。
 * 所以：形状问题在这里非 0 退出；语义问题（目标不存在、链式别名）交给门禁。
 *
 * 幂等：输出按 key 排序，输入行序变动不会造成 diff（CI 有 git diff 门禁）。
 *
 * 用法：node scripts/manual/gen-tool-aliases.js
 */

'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const SRC = path.join(__dirname, 'tool-aliases.json');
const OUT = path.join(ROOT, 'src', 'server', 'tool-aliases.generated.ts');

const NAME_RE = /^wps_[a-z0-9]+(_[a-z0-9]+)*$/;

function die(msg) {
  console.error(`✗ ${msg}`);
  process.exit(1);
}

if (!fs.existsSync(SRC)) die(`找不到真相源 ${path.relative(ROOT, SRC)}`);

const raw = fs.readFileSync(SRC, 'utf8').replace(/^\uFEFF/, '');

// JSON.parse 会把重复 key 静默并成最后一个 —— 而重复 key 恰恰是最容易犯的编辑错误。
// 这里先扫一遍原始文本，把同一层里出现两次的 key 抓出来。
function findDuplicateKeys(text) {
  // 只扫 "aliases" 那一层的直接子键：缩进两格的 "key":
  const inAliases = text.indexOf('"aliases"');
  if (inAliases < 0) return [];
  const body = text.slice(inAliases);
  const seen = new Map();
  const dups = [];
  for (const m of body.matchAll(/^ {4}"([^"]+)"\s*:/gm)) {
    const k = m[1];
    if (seen.has(k)) dups.push(k);
    else seen.set(k, true);
  }
  return dups;
}
const dups = findDuplicateKeys(raw);
if (dups.length) die(`tool-aliases.json 里有重复的键：${dups.join(', ')}`);

let parsed;
try {
  parsed = JSON.parse(raw);
} catch (e) {
  die(`tool-aliases.json 不是合法 JSON：${e.message}`);
}

const aliases = parsed.aliases || {};
const entries = Object.entries(aliases);
if (!entries.length) die('aliases 是空的 —— 没有别名就别跑生成器（会把表清空）');

const problems = [];
const targets = new Map();
for (const [from, to] of entries) {
  if (typeof to !== 'string') {
    problems.push(`${from}: 目标必须是字符串，收到 ${JSON.stringify(to)}`);
    continue;
  }
  if (!NAME_RE.test(from)) problems.push(`${from}: 别名不合 wps_<组>_<动词>_<名词> 形状`);
  if (!NAME_RE.test(to)) problems.push(`${to}: 规范名不合 wps_<组>_<动词>_<名词> 形状（来自 ${from}）`);
  if (from === to) problems.push(`${from}: 自己别名自己`);
  if (targets.has(to)) problems.push(`${to}: 被 ${targets.get(to)} 和 ${from} 同时指向（一个规范名只该吸收一组重复）`);
  targets.set(to, from);
}
// 链式别名：a→b 且 b→c。解析只做一跳，链式会让 a 解析到不存在的 b。
for (const [from, to] of entries) {
  if (Object.prototype.hasOwnProperty.call(aliases, to)) {
    problems.push(`${from} → ${to}，但 ${to} 本身也是别名（解析只做一跳，链式会指向不存在的名字）`);
  }
}
if (problems.length) {
  console.error('✗ 别名表有问题：');
  for (const p of problems) console.error(`  - ${p}`);
  process.exit(1);
}

const sorted = entries.slice().sort((a, b) => a[0].localeCompare(b[0]));

const lines = [];
lines.push('/**');
lines.push(' * Input: 无（纯常量）');
lines.push(' * Output: 旧工具名 → 规范工具名的映射与解析函数');
lines.push(' * Pos: 别名表【生成物】。一旦我被修改，请更新我的头部注释，以及所属目录的md。');
lines.push(' *');
lines.push(' * 别手改这个文件：真相源是 scripts/manual/tool-aliases.json，');
lines.push(` * 生成器是 scripts/manual/gen-tool-aliases.js，共 ${sorted.length} 条。`);
lines.push(' *');
lines.push(' * 别名不注册成工具：它不出现在 tools/list、不计入工具总数，只在 registry 的入口被解析成规范名。');
lines.push(' * 为什么只做一跳解析：链式别名（a→b→c）会让"到底谁是真的"变成谜，一份表里就该一步到位。');
lines.push(' */');
lines.push('');
lines.push('/** 旧名 → 规范名。键是已下线的旧工具名，值是仍在注册表里的规范名 */');
lines.push('export const TOOL_ALIASES: Readonly<Record<string, string>> = {');
for (const [from, to] of sorted) {
  lines.push(`  ${JSON.stringify(from)}: ${JSON.stringify(to)},`);
}
lines.push('};');
lines.push('');
lines.push('/** 全部别名（旧名）。用于"这个请求是不是旧名"的判断与门禁遍历 */');
lines.push('export const ALIAS_NAMES: readonly string[] = Object.keys(TOOL_ALIASES);');
lines.push('');
lines.push('/**');
lines.push(' * 把（可能是旧名的）工具名解析成规范名。');
lines.push(' * 不是别名就原样返回 —— 调用方不用先判断再解析，少一个分支就少一处漏判。');
lines.push(' */');
lines.push('export function resolveToolAlias(name: string): string {');
lines.push('  return TOOL_ALIASES[name] ?? name;');
lines.push('}');
lines.push('');
lines.push('/** 这个名字是不是已下线的旧名（用于在结果里告诉调用方"你用的是旧名"） */');
lines.push('export function isToolAlias(name: string): boolean {');
lines.push('  return Object.prototype.hasOwnProperty.call(TOOL_ALIASES, name);');
lines.push('}');
lines.push('');

const content = lines.join('\n');
const eolContent = content.replace(/\n/g, detectEol());
function detectEol() {
  if (!fs.existsSync(OUT)) return '\n';
  return fs.readFileSync(OUT, 'utf8').includes('\r\n') ? '\r\n' : '\n';
}

const before = fs.existsSync(OUT) ? fs.readFileSync(OUT, 'utf8') : '';
if (before === eolContent) {
  console.log(`= ${path.relative(ROOT, OUT)} 无变化（${sorted.length} 条别名）`);
  process.exit(0);
}
fs.writeFileSync(OUT, eolContent, 'utf8');
console.log(`✓ ${path.relative(ROOT, OUT)}：${sorted.length} 条别名`);
for (const [from, to] of sorted) console.log(`    ${from} → ${to}`);
