/**
 * skill 包校验
 * -----------
 * 按 Agent Skills 规范（agentskills.io/specification）逐条查，不是随便看看：
 *   1) name 必须等于父目录名，且符合命名规则（规范硬要求，错���客户端不认）
 *   2) description ≤1024 字符 —— 上限是仓库里一条提交踩出来的（1176 → 992）
 *   3) SKILL.md <500 行 —— 超了模型一次读不完，渐进披露就失效了
 *   4) 引用文件必须存在，且【只一层深】（规范明确要求）
 *   5) 目录名用 references/（复数），别用 reference/
 *
 * 用法：node scripts/manual/skill-verify.js [skill目录]
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const skillDir = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(ROOT, 'skills', 'wps-multitool');

let pass = 0;
let fail = 0;
const ok = (label, cond, extra) => {
  if (cond) {
    pass++;
    console.log(`  ✓ ${label}${extra ? ' — ' + extra : ''}`);
  } else {
    fail++;
    console.log(`  ✗ ${label}${extra ? ' — ' + extra : ''}`);
  }
};

console.log(`\n校验 ${path.relative(ROOT, skillDir) || skillDir}`);

const skillPath = path.join(skillDir, 'SKILL.md');
if (!fs.existsSync(skillPath)) {
  console.log(`  ✗ 找不到 SKILL.md`);
  process.exit(1);
}
const raw = fs.readFileSync(skillPath, 'utf8');

// ---- frontmatter 解析（够用即可，不引 yaml 依赖） ----
const fmMatch = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n/);
ok('有 YAML frontmatter', !!fmMatch);
if (!fmMatch) process.exit(1);
const fm = fmMatch[1];

const getField = (key) => {
  const m = fm.match(new RegExp(`^${key}:\\s*([\\s\\S]*?)(?=\\n[a-z_]+:|$)`, 'm'));
  return m ? m[1].trim() : '';
};

const name = getField('name');
const description = getField('description');
const compatibility = getField('compatibility');

// ---- 1. name ----
const dirName = path.basename(skillDir);
ok('name 等于父目录名', name === dirName, `name=${name} dir=${dirName}`);
ok('name 长度 ≤64', name.length <= 64, `${name.length}`);
ok('name 只含小写字母/数字/连字符', /^[a-z0-9-]+$/.test(name));
ok('name 不以连字符开头结尾', !name.startsWith('-') && !name.endsWith('-'));
ok('name 无连续连字符', !name.includes('--'));
ok('name 不含保留词', !/claude|anthropic/i.test(name));

// ---- 2. description ----
ok('description 非空', description.length > 0);
ok('description ≤1024 字符', description.length <= 1024, `${description.length}`);
ok('description 写了"什么时候用"', /当|使用|适合|不要使用|需要/.test(description));

// ---- 3. compatibility ----
if (compatibility) {
  ok('compatibility ≤500 字符', compatibility.length <= 500, `${compatibility.length}`);
}

// ---- 4. 体量 ----
const lines = raw.split(/\r?\n/).length;
ok('SKILL.md <500 行', lines < 500, `${lines} 行`);
ok('SKILL.md 未超 5000 token 量级', raw.length < 20000, `${raw.length} 字符`);

// ---- 5. 目录约定 ----
ok('用 references/（复数）', fs.existsSync(path.join(skillDir, 'references')));
ok('没有 reference/（单数，规范名）', !fs.existsSync(path.join(skillDir, 'reference')));

// ---- 6. 引用完整性 + 只有一层深 ----
const body = raw.slice(fmMatch[0].length);
const refs = [...body.matchAll(/(?:references|scripts|assets)\/[\w.-]+/g)].map((m) => m[0]);
const uniqRefs = [...new Set(refs)];
ok('SKILL.md 里引用了参考文档', uniqRefs.length > 0, uniqRefs.join(', '));
for (const r of uniqRefs) {
  ok(`引用存在：${r}`, fs.existsSync(path.join(skillDir, r)));
}

// 参考文档内部不允许再互相引用（规范：只一层深）
let nested = 0;
const refDir = path.join(skillDir, 'references');
if (fs.existsSync(refDir)) {
  for (const f of fs.readdirSync(refDir)) {
    const txt = fs.readFileSync(path.join(refDir, f), 'utf8');
    const inner = [...txt.matchAll(/references\/[\w.-]+/g)];
    if (inner.length) {
      nested += inner.length;
      console.log(`  ! ${f} 里还有 ${inner.length} 处 references/ 引用`);
    }
  }
}
ok('参考文档没有嵌套引用（只一层深）', nested === 0);

const refFiles = fs.existsSync(refDir) ? fs.readdirSync(refDir) : [];
console.log(`\n  参考文档 ${refFiles.length} 份：${refFiles.join(', ')}`);

console.log(`\n汇总：${pass} 通过 / ${fail} 失败`);
process.exit(fail === 0 ? 0 : 1);