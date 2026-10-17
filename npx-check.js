/**
 * 一次性脚本（用完删除）：验证 npx 能真的拉起 MCP Server 并完成握手。
 * 不测别的，只回答一个问题：用户照 README 写 npx 配置，能不能通。
 */
const { spawn } = require('child_process');

const child = spawn('cmd', ['/c', 'npx', '-y', 'wps-multitool-mcp'], { stdio: ['pipe', 'pipe', 'pipe'] });
let buf = '';
const send = (o) => child.stdin.write(JSON.stringify(o) + '\n');
const done = (code) => { try { child.kill(); } catch {} process.exit(code); };

child.stdout.on('data', (d) => {
  buf += d.toString();
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i).trim();
    buf = buf.slice(i + 1);
    if (!line) continue;
    let m;
    try { m = JSON.parse(line); } catch { continue; }
    if (m.id === 1) {
      console.log('initialize OK:', m.result.serverInfo.name, m.result.serverInfo.version);
      send({ jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} });
    } else if (m.id === 2) {
      const n = m.result.tools.length;
      console.log('tools/list =', n);
      const hasStats = m.result.tools.some((t) => t.name === 'wps_common_status');
      console.log('内置状态探针在列 =', hasStats);
      done(n === 40 && hasStats ? 0 : 1);
    }
  }
});

child.stderr.on('data', (d) => {
  const s = d.toString();
  if (/error|Error/.test(s)) process.stderr.write('[stderr] ' + s);
});

send({
  jsonrpc: '2.0', id: 1, method: 'initialize',
  params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'npx-check', version: '0.1' } },
});
send({ jsonrpc: '2.0', method: 'notifications/initialized' });

setTimeout(() => { console.log('TIMEOUT（多次首次 npx 会先下载依赖，可重试一次）'); done(2); }, 90000);
