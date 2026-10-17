# Input: stdin 每行一个 JSON 请求 {"id":"...","action":"...","params":{...}}
# Output: stdout 每行一个 JSON 响应（原动作输出 + id）
# Pos: WPS COM 常驻守护进程。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
#
# 为什么要常驻
# ------------
# 改造前 MCP 每调一次工具就 spawn 一个 powershell.exe，解析 5000 行脚本再连一次 COM，
# 单次白花 0.3~0.6 秒。守护进程只启动一次、只解析一次，之后的请求在同一个进程里派发，
# COM 附着关系也留在同一个 STA 线程上。实测冷启动 5.1 秒，之后每次 3 毫秒。
#
# 与 wps-com.ps1 的关系
# --------------------
# 本文件设 $env:WPS_COM_DAEMON=1 后 dot-source wps-com.ps1。
# wps-com.ps1 末尾的入口段见到该变量就【只定义函数、不执行动作】，
# 于是 Invoke-WpsAction 成了可反复调用的分发入口。
# 直跑用法（-File wps-com.ps1 -Action X）完全不受影响。
#
# 编码：本文件【必须存成 UTF-8 with BOM】。
#   PowerShell 5.1 对无 BOM 文件按 ANSI(GBK) 解码，头部中文注释会变乱码，
#   而且会静默吃掉紧随其后的第一条输出（已实测踩过，握手行就是这么丢的）。
#
# 协议：一行一个 JSON，请求/响应都带 id；除动作输出外本进程不往 stdout 写任何东西。
#       （已核验：wps-com.ps1 内没有 Write-Host / Write-Output / echo。）
#
# 生命周期（两条同时生效，缺一不可）
# --------------------------------
#   1) stdin 读到 EOF → 退出（父进程正常收管道时会走这条）
#   2) 父进程看门狗 → 退出
#   第 2 条是必须的：实测【父进程被强杀、甚至正常 exit 之后】stdin 都不一定立刻 EOF，
#   只靠第 1 条会留下一个抓着 COM 的孤儿 powershell。
#   看门狗必须跑在【独立 .NET 线程】上 —— 主线程阻塞在 ReadLine 里时，
#   PowerShell 的 Register-ObjectEvent / 定时器事件根本不会被泵到。
#   看门狗只做一件事：父进程没了就 Environment.Exit(0)。

$ErrorActionPreference = 'Continue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

$script:WpsReqId = $null
$parentPid = $null

# ==================== 父进程看门狗 ====================
# Add-Type 首次编译约 0.5~1 秒，相对 5 秒冷启动可以接受
try {
    if (-not ([System.Management.Automation.PSTypeName]'WpsParentWatchdog').Type) {
        Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Threading;

public static class WpsParentWatchdog
{
    public static void Start(int parentPid)
    {
        var t = new Thread(() =>
        {
            while (true)
            {
                try
                {
                    var p = Process.GetProcessById(parentPid);
                    if (p.HasExited) { break; }
                }
                catch
                {
                    // 父进程没了（或 PID 已失效）—— 我们也该走
                    Environment.Exit(0);
                }
                Thread.Sleep(2000);
            }
            Environment.Exit(0);
        });
        t.IsBackground = true;
        t.Name = "wps-parent-watchdog";
        t.Start();
    }
}
'@ -ErrorAction Stop
    }

    $parentPid = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID" -ErrorAction SilentlyContinue).ParentProcessId
    if ($parentPid) {
        [WpsParentWatchdog]::Start([int]$parentPid)
    }
} catch {
    # 看门狗起不来不能拖垮主流程：只是失去一层保护，照常工作
}

# ==================== 加载动作库 ====================
# 置位后再 dot-source，让 wps-com.ps1 只定义函数、不执行动作
$env:WPS_COM_DAEMON = '1'
. (Join-Path $PSScriptRoot 'wps-com.ps1')

# 就绪握手：Node 侧据此确认守护进程起来了（而不是白等到超时）
Output-Json @{ success = $true; data = @{ ready = $true; pid = $PID; daemon = 'wps-comd'; parent = $parentPid } }

while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }          # stdin 关闭 → 退出
    $line = $line.Trim()
    if (-not $line) { continue }

    $id = $null
    try {
        $req = $line | ConvertFrom-Json
        if ($req.id) { $id = [string]$req.id }
        $script:WpsReqId = $id

        $act = [string]$req.action
        if (-not $act) { throw '请求缺少 action' }

        # $req.params 为 $null 时给空表；动作内部统一按 $p.xxx 取值
        $par = if ($null -ne $req.params) { $req.params } else { @{} }

        Invoke-WpsAction -Action $act -p $par
    } catch {
        # 分发层异常兜底：绝不能把异常抛给宿主，否则整个守护进程会随之下线
        $script:WpsReqId = $id
        Output-Json @{ success = $false; error = "daemon: $($_.Exception.Message)" }
    } finally {
        $script:WpsReqId = $null
    }
}
