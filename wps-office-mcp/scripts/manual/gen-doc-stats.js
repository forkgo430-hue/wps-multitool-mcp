#!/usr/bin/env node
/**
 * Input: 真实注册路径（dist 里的注册表 + 注解计算）
 * Output: README.md 的两处 AUTO 区 + wps-claude-addon/panel.html 的数字/版本
 * Pos: 文档统计的单一真相源。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 要解决的问题
 * ------------
 * "工具总数/分布/注解统计/面板版本"这些数字曾经在 README 三处、panel.html、
 * SKILL.md 各抄一遍，漂了两轮没人发现 —— 根因是没有任何一处跟注册表挂钩，
 * 也没有任何一个地方会因为"数字对不上"而失败。
 *
 * 这个脚本让 README 与面板的数字【只能来自注册表】：
 *   - README `<!-- AUTO:tool-stats:start/end -->`：总数、分布、默认列出数
 *   - README `<!-- AUTO:tool-table:start/end -->`：工具清单表（计数列）
 *   - panel.html：工具数那行 + 版本号（版本以 package.json 为准）
 * SKILL.md 不归这个脚本管 —— 按约定它【不含数字】（"近 300 个，以 tools/list 为准"），
 * 背数字必腐化且没有门禁会红。
 *
 * 覆盖面必须走【真实的注册路径】（与 gen-tool-annotations.js 同一套教训）：
 *   registerBuiltinTools() 的 12 个内置 + registerAll(allTools) 的专业工具。
 * 少做第一步，总数就少 12 个 —— 这种"两处维护"的漂移已经吃过一次亏。
 *
 * 用法：node scripts/manual/gen-doc-stats.js
 * CI：跑完本脚本后 `git diff --exit-code` README.md 与 panel.html ——
 * 改了工具忘了重跑 → diff 非空 → 红。
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const REPO_ROOT = path.resolve(ROOT, '..');
const README = path.join(REPO_ROOT, 'README.md');
const PANEL = path.join(REPO_ROOT, 'wps-claude-addon', 'panel.html');

const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));
const { groupOfTool } = require(path.join(ROOT, 'dist', 'types', 'tools.js'));
const { annotationSummary } = require(path.join(ROOT, 'dist', 'server', 'tool-annotations.js'));

// ---- 1. 真实注册路径，量出数字 ----------------------------------------------
const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const all = toolRegistry.listAllTools();
const total = all.length;
const builtinCount = total - allTools.length; // registerBuiltinTools 注册的那批
const professionalCount = allTools.length;

// 按组统计（内置工具计入其 category 对应的组 —— 与 list_tool_groups 的口径一致）
const byGroup = {};
for (const t of all) {
  const g = groupOfTool(t);
  byGroup[g] = (byGroup[g] || 0) + 1;
}
const ann = annotationSummary(all.map((t) => t.name));
const listed = toolRegistry.listTools().tools.length;
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));

// ---- 2. 生成两块 AUTO 内容 ---------------------------------------------------

/** 各组在清单表里的展示名与能力描述（能力描述是给人看的，属于数据不是逻辑） */
const TABLE_ROWS = [
  {
    group: 'excel',
    label: 'Excel',
    caps: '公式/数据/图表/透视表/工作表/格式/工作簿/行列/批注保护/图片导出',
  },
  {
    group: 'word',
    label: 'Word',
    caps: '格式/内容/文档管理/页眉页脚/批注/书签/表格/模板填写/学术排版/修订/索引/交叉引用',
  },
  {
    group: 'ppt',
    label: 'PPT',
    caps: '幻灯片/形状/图片/表格/美化/动画/图表/3D/数据可视化/图片导出',
  },
  {
    // common + core：core 是工具组管理本身，与通用工具合一行展示
    group: 'common',
    label: '通用',
    caps: '保存/连接检测/文本选取/格式转换/工具组管理/意图路由/批量执行',
    includeCore: true,
  },
];

const tableLines = [
  '| 应用 | 工具数 | 主要能力 |',
  '|------|--------|---------|',
];
// 口径与上方统计区【完全一致】：内置计入对应组、core 计入通用。
// 曾经表格用"专业互斥 + 单列内置"口径（19/68/115/92 + 12），数字虽对但和统计区
// 长得不像，读者一律当成"旧的没同步"—— 口径不统一比数字错更费沟通成本。
for (const row of TABLE_ROWS) {
  let n = byGroup[row.group] || 0;
  if (row.includeCore) n += byGroup.core || 0;
  tableLines.push(`| ${row.label} | ${n} | ${row.caps} |`);
}
tableLines.push(`| 内置 | ${builtinCount}（已计入上方各组） | 连接检查/万能方法调用/数据缓存 |`);
const tableBlock = tableLines.join('\n');

const statsBlock = [
  `- **${total} 个工具**（${builtinCount} 个内置 + ${professionalCount} 个专业/组工具）；` +
    `**默认只列 ${listed} 个**（基础档，按需加载 —— 口径与校验见 \`scripts/manual/basic-tier-verify.js\`）`,
  `- 分布（按工具组，内置计入对应组）：` +
    `通用 ${(byGroup.common || 0) + (byGroup.core || 0)}（含 core）/ ` +
    `Word ${byGroup.word || 0} / Excel ${byGroup.excel || 0} / PPT ${byGroup.ppt || 0}`,
  `- 注解（实测）：只读 ${ann.readOnly} / 破坏性 ${ann.destructive} / 非幂等 ${ann.nonIdempotent}，` +
    `覆盖 ${ann.annotated}/${ann.total}`,
].join('\n');

// ---- 3. 替换 README 的 AUTO 区 -------------------------------------------------

/** 把 start/end 标记之间的内容换掉。标记缺失 = 调用方没按约定接好，直接失败。 */
function replaceAutoRegion(text, startMarker, endMarker, block, label) {
  const startIdx = text.indexOf(startMarker);
  const endIdx = text.indexOf(endMarker);
  if (startIdx < 0 || endIdx < 0 || endIdx < startIdx) {
    console.error(`✗ README 里找不到 ${label} 的 AUTO 标记（${startMarker} … ${endMarker}）`);
    process.exit(1);
  }
  return text.slice(0, startIdx + startMarker.length) + '\n' + block + '\n' + text.slice(endIdx);
}

let readme = fs.readFileSync(README, 'utf8');
const before = readme;
readme = replaceAutoRegion(
  readme,
  '<!-- AUTO:tool-stats:start -->',
  '<!-- AUTO:tool-stats:end -->',
  statsBlock,
  'tool-stats'
);
readme = replaceAutoRegion(
  readme,
  '<!-- AUTO:tool-table:start -->',
  '<!-- AUTO:tool-table:end -->',
  tableBlock,
  'tool-table'
);
if (readme !== before) {
  fs.writeFileSync(README, readme, 'utf8');
  console.log(`已更新 README.md（AUTO 区两块）`);
} else {
  console.log('README.md 已是最新（AUTO 区无需改动）');
}

// ---- 4. panel.html：工具数那行 + 版本号 ----------------------------------------
// 版本以 package.json 为唯一真相源（与 src/version.ts 同源），面板不许自己背一个。
let panel = fs.readFileSync(PANEL, 'utf8');
const panelBefore = panel;
const ver = pkg.version;
panel = panel.replace(
  /<li>\d+ 个 MCP 工具，默认只列 \d+ 个基础工具<\/li>/,
  `<li>${total} 个 MCP 工具，默认只列 ${listed} 个基础工具</li>`
);
panel = panel.replace(/(<span class="ver">v)\d+\.\d+\.\d+(<\/span>)/, `$1${ver}$2`);
panel = panel.replace(
  /(<span class="k">版本<\/span><span>)\d+\.\d+\.\d+(<\/span>)/,
  `$1${ver}$2`
);
if (panel !== panelBefore) {
  fs.writeFileSync(PANEL, panel, 'utf8');
  console.log(`已更新 panel.html（工具数 / 版本 → ${ver}）`);
} else {
  console.log('panel.html 已是最新');
}

// ---- 5. 报告 -------------------------------------------------------------------
console.log(`\n注册表实测（真实注册路径）:`);
console.log(`  总数 ${total} = 内置 ${builtinCount} + 专业 ${professionalCount}`);
console.log(
  `  按组: common ${byGroup.common || 0} + core ${byGroup.core || 0} / word ${byGroup.word || 0} / ` +
    `excel ${byGroup.excel || 0} / ppt ${byGroup.ppt || 0}`
);
console.log(`  默认列出 ${listed} 个（基础档）`);
console.log(
  `  注解: 只读 ${ann.readOnly} / 破坏性 ${ann.destructive} / 非幂等 ${ann.nonIdempotent} / 覆盖 ${ann.annotated}/${ann.total}`
);
process.exit(0);
