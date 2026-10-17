/** 一次性脚本（用完删除）：轮询最近一次 CI，直到结束并打印每步结论。 */
const https = require('https');
const fs = require('fs');
const os = require('os');
const path = require('path');

const cfg = JSON.parse(fs.readFileSync(path.join(os.homedir(), '.codebuddy', 'mcp.json'), 'utf8'));
const TOKEN = cfg.mcpServers.GitHub.env.GITHUB_PERSONAL_ACCESS_TOKEN;

function api(p) {
  return new Promise((resolve, reject) => {
    https.get(
      { hostname: 'api.github.com', path: p, headers: { Authorization: `Bearer ${TOKEN}`, 'User-Agent': 'ci-wait', 'X-GitHub-Api-Version': '2022-11-28' } },
      (res) => {
        let d = '';
        res.on('data', (c) => (d += c));
        res.on('end', () => {
          try { resolve(JSON.parse(d)); } catch (e) { reject(new Error(d.slice(0, 200))); }
        });
      }
    ).on('error', reject);
  });
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

(async () => {
  for (let i = 0; i < 40; i++) {
    const runs = await api('/repos/forkgo430/wps-multitool-mcp/actions/runs?per_page=1');
    const r = runs.workflow_runs && runs.workflow_runs[0];
    if (!r) { console.log('无 run'); return; }
    console.log(`[${i}] #${r.run_number} ${r.head_sha.slice(0, 7)} status=${r.status} conclusion=${r.conclusion}`);
    if (r.status === 'completed') {
      const jobs = await api(`/repos/forkgo430/wps-multitool-mcp/actions/runs/${r.id}/jobs`);
      for (const j of jobs.jobs || []) {
        console.log(`JOB ${j.name}: ${j.conclusion}`);
        for (const s of j.steps || []) console.log(`  ${s.conclusion === 'success' ? '✓' : s.conclusion === 'failure' ? '✗' : '·'} ${s.name}`);
      }
      return;
    }
    await sleep(10000);
  }
})();
