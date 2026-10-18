/**
 * Input: docs/agents/PPT_SMOKE_MATRIX.md（工具 + 安全参数 + 判据）
 * Output: 每个工具在【全新靶子】上单独跑一次的结果（PASSED=/FAILED= 清单）
 * Pos: 冒烟脚本的"隔离版"。一旦我被修改，请更新我的头部注释，以及所属目录的md。
 *
 * 为什么要有它
 * ----------
 * ppt-smoke-scan.js 是顺序跑的：前一条留下的状态会影响后一条（版式被换、形状被删、
 * 尺寸被改），于是"图没变化"既可能是工具坏，也可能是靶子已经不是当初那个靶子了。
 * 实测过：wps_ppt_set_background 在序列里判失败，单独跑三个（含两个别名）全部成功。
 * 所以这里每条用例都重建一次靶子 —— 慢，但结论可信。
 */
process.env.WPS_BRIDGE_PIPE = `\\\\.\\pipe\\wps-multitool-mcp-isolated-${Date.now()}`;
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_DAEMON_TIMEOUT = '60000';

const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const REPO = path.resolve(ROOT, '..');
const MATRIX = path.join(REPO, 'docs', 'agents', 'PPT_SMOKE_MATRIX.md');
const { WpsMcpServer } = require(path.join(ROOT, 'dist/server/mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist/server/tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist/tools/index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'ppt-iso-'));
const png = path.join(tmpDir, 'one.png');
fs.writeFileSync(
  png,
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )
);

const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => String(((r && r.content && r.content[0] && r.content[0].text) || (r && r.error) || '')).replace(/\s+/g, ' ');

function parseMatrix(file) {
  const rows = [];
  for (const line of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    if (!/^\|\s*`?wps_/.test(line)) continue;
    const cells = line.split('|').map((c) => c.trim());
    const tool = (cells[1] || '').replace(/`/g, '');
    const argRaw = (cells[2] || '').replace(/`/g, '');
    const judge = cells[3] || '';
    const note = cells[4] || '';
    if (!tool || judge === '跳过') continue;
    let args = {};
    try { args = JSON.parse(argRaw); } catch { continue; }
    rows.push({ tool, args, judge: judge.replace(/\s/g, ''), note });
  }
  return rows;
}

async function rebuild() {
  // 关掉上一轮的（只关我们自己建的无路径文稿）
  try {
    const list = await call('wps_ppt_get_open_presentations', {});
    const names = (txt(list).match(/演示文稿\d+/g) || []);
    for (const n of names) await call('wps_ppt_close_presentation', { name: n, save: false });
  } catch { /* ignore */ }
  await call('wps_ppt_create_presentation', {});
  await call('wps_ppt_add_slide', { layout: 'blank' });
  await call('wps_ppt_add_slide', { layout: 'blank' });
  for (const p of [1, 2]) {
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'rectangle', left: 60, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_shape', { slideIndex: p, type: 'oval', left: 240, top: 60, width: 120, height: 60 });
    await call('wps_ppt_add_textbox', { slideIndex: p, left: 60, top: 160, width: 240, height: 60, text: `文字${p}` });
    await call('wps_ppt_insert_table', { slideIndex: p, rows: 2, cols: 2 });
    await call('wps_ppt_insert_image', { slideIndex: p, path: png });
    await call('wps_ppt_set_shape_text', { slideIndex: p, shapeIndex: 1, text: `文字${p}` });
  }
}

async function fingerprint() {
  const c = await call('wps_ppt_get_slide_count', {});
  const parts = [txt(c).match(/(\d+)/)?.[1] || '0'];
  for (const i of [1, 2]) {
    const s = await call('wps_ppt_get_shapes', { slideIndex: i });
    parts.push(String((txt(s).match(/共 (\d+) 个/) || [])[1] || '0'));
    const a = await call('wps_ppt_get_animations', { slideIndex: i });
    parts.push(String((txt(a).match(/共 (\d+) 个|(\d+) 个动画/) || [, '0'])[1] || '0'));
  }
  return parts.join('/');
}

async function shot(tag) {
  const p = path.join(tmpDir, `s-${tag}.png`);
  await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: p });
  return fs.existsSync(p) ? fs.readFileSync(p) : null;
}

(async () => {
  const rows = parseMatrix(MATRIX);
  console.log(`=== 隔离验证：${rows.length} 条（每条重建靶子）===\n`);
  const passed = [];
  const failed = [];

  for (const row of rows) {
    await rebuild();
    const fpBefore = await fingerprint();
    const imgBefore = await shot('a');
    const imgCheck = await shot('b');
    const stable = imgBefore && imgCheck && Buffer.compare(imgBefore, imgCheck) === 0;

    let r = null;
    let err = '';
    try {
      r = await call(row.tool, row.args);
      if (!r || r.success === false) err = txt(r).slice(0, 90) || 'success=false';
    } catch (e) {
      err = `THROW ${e.message}`.slice(0, 90);
    }

    let ok = !err;
    let why = '';
    if (ok) {
      if (row.judge.includes('数量差')) {
        const fpAfter = await fingerprint();
        ok = fpAfter !== fpBefore;
        why = `计数 ${fpBefore} → ${fpAfter}`;
      } else if (row.judge.includes('读回')) {
        const t = txt(r);
        ok = t.length > 8 && !/undefined|null/.test(t.slice(0, 40));
        why = `读回 ${t.length} 字`;
      } else {
        const imgAfter = await shot('c');
        ok = !!stable && !!imgAfter && Buffer.compare(imgBefore, imgAfter) !== 0;
        why = stable ? (ok ? '导出图有变化' : '导出图无变化') : '导出不稳定';
      }
    }
    console.log(`${ok ? '✓' : '✗'} ${row.tool.padEnd(34)} ${ok ? why : err || why}`);
    (ok ? passed : failed).push(row.tool);
  }

  console.log(`\n通过 ${passed.length}/${rows.length}`);
  if (failed.length) console.log(`失败：${failed.join(', ')}`);
  console.log(`PASSED=${passed.join(',')}`);
  console.log(`FAILED=${failed.join(',')}`);

  try {
    const list = await call('wps_ppt_get_open_presentations', {});
    for (const n of txt(list).match(/演示文稿\d+/g) || []) {
      await call('wps_ppt_close_presentation', { name: n, save: false });
    }
  } catch { /* ignore */ }
  fs.rmSync(tmpDir, { recursive: true, force: true });
  process.exit(failed.length ? 1 : 0);
})();
