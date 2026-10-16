# WPS Skills 安装指南

> （只有代码、没有配套加载项，实际跑不起来），相关代码与加载项已删除。

> **本文件面向 AI 编程助手**（Claude Code、Cursor 等）。读取本文件后，你应能完全理解安装流程并自动执行所有步骤。按顺序执行，不要跳步。

---

## 安装前状态检查

在执行安装前，先判断是否已经安装过。依次执行以下检查：

```bash
# 1. 检查 MCP Server 是否已注册
claude mcp list 2>/dev/null | grep -i wps

# 2. 检查 Skills 是否已注册
ls ~/.claude/skills/wps-excel/SKILL.md 2>/dev/null

# 3. 检查编译产物是否存在
ls wps-office-mcp/dist/index.js 2>/dev/null
```

**判断逻辑**：
- 三项全部存在 → 已安装，无需重复执行。告知用户"WPS Skills 已安装，重启 Claude Code 和 WPS Office 即可使用"
- 部分存在 → 从缺失的步骤开始补装
- 全部不存在 → 执行完整安装流程

---

## Windows 安装

### Step 1: 检测前置条件

```powershell
# 检测 WPS Office 加载项目录是否存在
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons"

# 检测 Node.js 版本（必须 >= 18.0.0）
node -v
```

**如果前置条件不满足**：
- WPS 加载项目录不存在 → 告知用户安装 WPS Office：https://www.wps.cn/
- Node.js 不存在或版本 < 18 → 告知用户安装：https://nodejs.org/
- 前置条件不满足时，停止安装流程，不要继续

### Step 2: 运行一键安装脚本

```powershell
# 进入项目根目录
cd <项目根目录>

# 执行安装脚本
powershell -ExecutionPolicy Bypass -File scripts/auto-install.ps1
```

脚本自动完成以下全部操作：
- 检测 Node.js 18+ 版本
- 执行 `npm install` 安装依赖
- 执行 `npm run build` 编译 TypeScript
- 配置 Claude Code MCP（写入 `%USERPROFILE%\.claude\settings.json`）
- 复制 Skills 到 `%USERPROFILE%\.claude\skills\`
- 安装 WPS 加载项到 `%APPDATA%\kingsoft\wps\jsaddons\wps-claude-addon_\`（目录名尾部 `_` 是必须的）
- 更新 publish.xml 注册加载项

### Step 3: 安装后验证

```powershell
# 1. MCP Server 已注册
claude mcp list

# 2. Skills 已注册（应列出 wps-excel, wps-word, wps-ppt, wps-office）
Get-ChildItem "$env:USERPROFILE\.claude\skills"

# 3. 编译产物存在
Test-Path "wps-office-mcp\dist\index.js"

# 4. WPS 加载项已安装
Test-Path "$env:APPDATA\kingsoft\wps\jsaddons\wps-claude-addon_"

# 5. publish.xml 已注册加载项
Select-String -Path "$env:APPDATA\kingsoft\wps\jsaddons\publish.xml" -Pattern "wps-claude-addon"
```

### Step 4: 告知用户

安装完成后，必须告知用户：
1. **重启 Claude Code**（必须！）
2. 重启 WPS Office
3. 查看 "Claude助手" 选项卡

---

## 关键路径参考

### Windows

| 项目 | 路径 |
|------|------|
| WPS 加载项基础目录 | `%APPDATA%\kingsoft\wps\jsaddons\` |
| 加载项安装目录 | `<基础目录>\wps-claude-addon_\`（尾部 `_` 必须） |
| publish.xml | `<基础目录>\publish.xml` |
| Skills 注册目录 | `%USERPROFILE%\.claude\skills\`（复制方式，非软链接） |
| MCP Server 配置 | `%USERPROFILE%\.claude\settings.json` |

---

## 错误处理

遇到安装错误时，按以下对照表处理：

### npm install 失败

```bash
# 清除缓存后重试
cd wps-office-mcp
rm -rf node_modules package-lock.json
npm install
```

如果仍失败，检查 Node.js 版本：
```bash
node -v
# 必须 >= 18.0.0，否则升级 Node.js
```

### npm run build（TypeScript 编译）失败

```bash
cd wps-office-mcp
rm -rf dist node_modules
npm install
npm run build
```

如果报 `tsc: command not found`，说明 typescript 未安装为依赖，检查 package.json 的 devDependencies 中是否有 typescript。

### MCP Server 注册失败

手动注册：
```bash
claude mcp add wps-office node <项目根目录的绝对路径>/wps-office-mcp/dist/index.js
```

注意：`<项目根目录的绝对路径>` 必须替换为实际路径，不能使用相对路径或变量。

### Skills 软链接创建失败

```bash
PROJECT_DIR=<项目根目录的绝对路径>
mkdir -p ~/.claude/skills
ln -sf "$PROJECT_DIR/skills/wps-excel" ~/.claude/skills/wps-excel
ln -sf "$PROJECT_DIR/skills/wps-word" ~/.claude/skills/wps-word
ln -sf "$PROJECT_DIR/skills/wps-ppt" ~/.claude/skills/wps-ppt
ln -sf "$PROJECT_DIR/skills/wps-office" ~/.claude/skills/wps-office
```

验证：
```bash
ls -la ~/.claude/skills/
# 应看到 4 个软链接，指向项目中的 skills/ 子目录
```

### WPS 加载项未显示 "Claude助手" 选项卡

1. 确认加载项目录已正确复制且名称以 `_` 结尾：
```bash
# 应包含 main.js, manifest.xml, ribbon.xml 等文件
```

2. 确认 publish.xml 包含注册条目：
```bash
# 应包含 <jsplugin name="claude-assistant" .../> 条目
```

3. 强制退出并重启 WPS：
```bash
sleep 2
```

## 已知问题与故障排除（GitHub Issues）

### Issue #6: MCP 连接失败（"Failed to connect"）

**现象**：`claude mcp list` 显示 `wps-office: Failed to connect`。

**根因**：dist 目录中的编译产物过期，存在工具名称重复注册导致 MCP Server 启动即崩溃。

**解决方案**：
```bash
cd wps-office-mcp
# 清除旧编译产物后重新编译（关键步骤！）
rm -rf dist
npm run build
# 验证启动是否正常（应看到 "Server started successfully"）
node dist/index.js 2>&1 | head -5
# Ctrl+C 退出
```

如果仍然失败，完整重建：
```bash
cd wps-office-mcp
rm -rf dist node_modules
npm install
npm run build
```

然后重新注册 MCP：
```bash
claude mcp remove wps-office
claude mcp add wps-office node <项目绝对路径>/wps-office-mcp/dist/index.js
```

### Issue #4: WPS 加载项启动报 "arguments error"

**现象**：WPS 弹出 `ERROR: arguments error at <anonymous>:1:89`。

**根因**：manifest.xml 中缺少 `<ribbon>` 和 `<scripts>` 标签声明，导致 WPS 无法正确解析加载项入口。

**解决方案**：

```bash
# 1. 检查 manifest.xml 是否包含 ribbon 和 scripts 声明
grep -E "ribbon|scripts" <加载项安装目录>/manifest.xml
# 应看到类似：
#   <ribbon src="ribbon.xml"/>
#   <script src="main.js"/>

# 2. 确认 ribbon.xml 和 main.js 文件存在
ls <加载项安装目录>/ribbon.xml
ls <加载项安装目录>/main.js

# 3. 重新复制加载项（使用最新版本）

# 4. 重启 WPS
sleep 2
```
