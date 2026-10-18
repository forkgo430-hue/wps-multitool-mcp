/**
 * PPT 基础档真机验证（需要 WPS 演示可访问）。
 *
 * 为什么整组 PPT 一直是 static
 * --------------------------
 * Word/Excel 都有 caps-probe，PPT 没有 —— 116 个工具从上线起就没在真机上跑过。
 * 默认档因此把 8 个基础 PPT 工具也挡在外面（默认只有 32 个）。
 * 这个脚本只回答一个问题：**这 8 个基础工具到底能不能用**。
 *
 * 探完就关、不保存、不留垃圾；只动临时目录里造的文件。
 * 无 WPS 时按既有约定"跳过不判失败"，但会明说未验证。
 *
 * 用法：node scripts/manual/ppt-caps-probe.js
 */
const fs = require('fs');
const os = require('os');
const path = require('path');

// 独立管道：桥是常驻的，共享管道会拿到旧脚本/旧状态（与 snapshot-verify 同一招）
process.env.WPS_BRIDGE_PIPE = '\\\\.\\pipe\\wps-multitool-mcp-ppt-probe';
process.env.WPS_BRIDGE_EPHEMERAL = '1';
process.env.WPS_BRIDGE_IDLE_MS = '30000';
// PPT 首次被 COM 拉起比 Word 慢，给足时间
process.env.WPS_DAEMON_TIMEOUT = process.env.WPS_DAEMON_TIMEOUT || '120000';

const ROOT = path.resolve(__dirname, '..', '..');
const { WpsMcpServer } = require(path.join(ROOT, 'dist', 'server', 'mcp-server.js'));
const { toolRegistry } = require(path.join(ROOT, 'dist', 'server', 'tool-registry.js'));
const { allTools } = require(path.join(ROOT, 'dist', 'tools', 'index.js'));

const inst = new WpsMcpServer();
inst.registerBuiltinTools();
toolRegistry.registerAll(allTools);

let skipped = false;
const results = [];
const check = (tool, ok, detail = '') => {
  results.push({ tool, ok });
  console.log(`  ${ok ? '✓' : '✗'} ${tool}${detail ? ' — ' + detail : ''}`);
};

const call = (name, args = {}) => toolRegistry.getTool(name).handler(args);
const txt = (r) => (r.content && r.content[0] && r.content[0].text) || r.error || '';

// 1x1 红点 PNG，够 WPS 认格式；不用真图片是为了不给仓库塞二进制
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

(async () => {
  const tmpDir = path.join(os.tmpdir(), `wps-ppt-probe-${Date.now()}`);
  fs.mkdirSync(tmpDir, { recursive: true });
  const imgPath = path.join(tmpDir, 'probe.png');
  fs.writeFileSync(imgPath, Buffer.from(PNG_1PX, 'base64'));
  let createdName = null;

  try {
    console.log('\n=== PPT 基础档真机探测 ===');

    // 探活：WPS 演示没装 / 起不来时不该判失败（CI runner 上就是这样）。
    // 与 word/excel-caps-probe 的约定一致：跳过，但要说清"未验证"。
    const hb = await call('wps_ppt_get_open_presentations', {});
    if (!hb.success) {
      skipped = true;
      console.log(`  ⚠ WPS 演示不可用（${txt(hb).slice(0, 80)}），本轮跳过，不判失败`);
      return;
    }

    // 先记下"探测前已开着哪些"：这样才能分清哪个文稿是我们建的，收尾只关自己的
    const before = txt(hb);
    const beforeNames = new Set((before.match(/[^\s\\/:*?"<>|]+\.(?:pptx?|dps)/gi) || []));

    const created = await call('wps_ppt_create_presentation', {});
    const m = txt(created).match(/「([^」]+)」|:\s*([^（\n]+)（/);
    createdName = m ? (m[1] || m[2] || '').trim() || null : null;

    if (created.success) {
      check('wps_ppt_create_presentation', true, txt(created).split('\n')[0].slice(0, 60));
    } else {
      // 冷启动病史：WPS 演示第一次被 COM 拉起要十几秒，调用方先超时了，
      // 但文稿其实已经建出来（实测窗口标题已变成"演示文稿N"）。
      // 所以用"有没有新出现的文稿"来判定，而不是只看这次调用的返回。
      const after = txt(await call('wps_ppt_get_open_presentations', {}));
      const afterNames = after.match(/[^\s\\/:*?"<>|]+\.(?:pptx?|dps)/gi) || [];
      const fresh = afterNames.find((n) => !beforeNames.has(n));
      if (fresh) {
        createdName = fresh;
        check(
          'wps_ppt_create_presentation',
          true,
          `首次调用超时，但文稿已建立（${fresh}）—— 冷启动历史问题，不是功能缺陷`
        );
      } else {
        check('wps_ppt_create_presentation', false, txt(created).split('\n')[0].slice(0, 60));
        throw new Error('演示文稿建不起来，后面的探测没有意义');
      }
    }

    const added = await call('wps_ppt_add_slide', { layout: 'title_content', title: 'PPT 探测标题' });
    check('wps_ppt_add_slide', added.success, txt(added).split('\n')[0].slice(0, 60));

    const count = await call('wps_ppt_get_slide_count', {});
    const n = Number((txt(count).match(/(\d+)/) || [])[1] || 0);
    check('wps_ppt_get_slide_count', count.success && n >= 1, `读到 ${n} 页`);

    const title = await call('wps_ppt_set_slide_title', { slideIndex: 1, title: 'PPT 探测标题（改写）' });
    check('wps_ppt_set_slide_title', title.success, txt(title).split('\n')[0].slice(0, 60));

    const notes = await call('wps_ppt_add_speaker_notes', { slideIndex: 1, notes: '真机探测备注' });
    check('wps_ppt_add_speaker_notes', notes.success, txt(notes).split('\n')[0].slice(0, 60));

    const img = await call('wps_ppt_insert_image', { slideIndex: 1, path: imgPath });
    check('wps_ppt_insert_image', img.success, txt(img).split('\n')[0].slice(0, 60));

    const outPath = path.join(tmpDir, 'slide1.png');
    const exp = await call('wps_ppt_export_slide_as_image', { slideIndex: 1, outputPath: outPath });
    const exported = fs.existsSync(outPath) && fs.statSync(outPath).size > 0;
    check(
      'wps_ppt_export_slide_as_image',
      exp.success && exported,
      exported ? `${Math.round(fs.statSync(outPath).size / 1024)} KB` : txt(exp).slice(0, 60)
    );

    // 判据不能看扩展名：未保存的文稿叫"演示文稿N"，没有 .pptx 后缀（第一版就写错了）
    const list = await call('wps_ppt_get_open_presentations', {});
    check(
      'wps_ppt_get_open_presentations',
      list.success && /演示文稿/.test(txt(list)) && /\d/.test(txt(list)),
      txt(list).split('\n')[0].slice(0, 60)
    );
  } catch (e) {
    console.log(`  ⚠ 探测中断：${e && e.message ? e.message : String(e)}`);
    console.log('    （需要 WPS 演示可访问。中断部分不判失败，但也不算数 —— 仍按 static 计）');
  } finally {
    try {
      if (createdName) {
        const closed = await call('wps_ppt_close_presentation', { name: createdName, save: false });
        console.log(`  ${closed.success ? '✓' : '·'} 收尾关闭 ${createdName}${closed.success ? '' : '（没关掉，可能名字不对）'}`);
      } else {
        const closed = await call('wps_ppt_close_presentation', { save: false });
        console.log(`  ${closed.success ? '✓' : '·'} 收尾关闭（按活动文稿）`);
      }
    } catch { /* 已经关了 */ }
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch { /* 占用 */ }

    if (skipped) {
      console.log('未探测到任何 PPT 能力（WPS 演示不可用），按约定跳过。');
      process.exit(0);
    }
    const passed = results.filter((r) => r.ok).map((r) => r.tool);
    const failed = results.filter((r) => !r.ok).map((r) => r.tool);
    console.log(`\n通过 ${passed.length}/${results.length}`);
    if (failed.length) console.log(`失败：${failed.join(', ')}`);
    console.log('\n把通过的这些逐条写进 scripts/manual/tool-verification.json 的 tools 里（值 live），');
    console.log('再跑 node scripts/manual/gen-tool-verification.js && npm run build。');
    console.log(`PASSED=${passed.join(',')}`);
    process.exit(failed.length ? 1 : 0);
  }
})();
