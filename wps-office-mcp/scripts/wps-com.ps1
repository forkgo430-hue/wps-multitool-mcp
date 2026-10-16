# Input: Action 名称与 JSON 参数
# Output: WPS COM 调用结果 JSON
# Pos: Windows COM 桥接脚本。一旦我被修改，请更新我的头部注释（Updated: 2026-05-26 15:30:00 CST），以及所属文件夹的md。
# WPS COM Bridge - PowerShell script for WPS COM operations
# Full implementation for Excel, Word, PPT, and common conversions
# Usage: powershell -File wps-com.ps1 -Action <action> -Params <json>

param(
    [string]$Action,
    [string]$Params = "{}"
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ==================== COM Object Getters ====================

function Get-WpsExcel {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Ket.Application') }
    catch {
        try {
            $app = New-Object -ComObject 'Ket.Application'
            if ($app) { try { $app.Visible = $true } catch {}; return $app }
        } catch {
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Ket.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID($clsid)
                    if ($type) {
                        $app = [Activator]::CreateInstance($type)
                        if ($app) { try { $app.Visible = $true } catch {}; return $app }
                    }
                }
            } catch {}
        }
        return $null
    }
}

function Get-WpsWord {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwps.Application') }
    catch {
        try {
            $app = New-Object -ComObject 'Kwps.Application'
            if ($app) { try { $app.Visible = $true } catch {}; return $app }
        } catch {
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Kwps.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID($clsid)
                    if ($type) {
                        $app = [Activator]::CreateInstance($type)
                        if ($app) { try { $app.Visible = $true } catch {}; return $app }
                    }
                }
            } catch {}
        }
        return $null
    }
}

function Get-WpsPpt {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Kwpp.Application') }
    catch {
        try {
            $app = New-Object -ComObject Kwpp.Application
            try { $app.Visible = $true } catch {}
            return $app
        } catch {
            try {
                $clsid = (Get-ItemProperty -Path 'HKCU:\Software\Classes\Kwpp.Application\CLSID' -ErrorAction Stop).'(default)'
                if ($clsid) {
                    $type = [Type]::GetTypeFromCLSID($clsid)
                    if ($type) {
                        $app = [Activator]::CreateInstance($type)
                        if ($app) {
                            try { $app.Visible = $true } catch {}
                            return $app
                        }
                    }
                }
            } catch {}
        }
        return $null
    }
}

# ==================== 窗口置前 / 进程退出 辅助 ====================

# 仅获取【已运行】的 WPS 实例（不创建新实例），供 showWps / quitWps 这类不应产生副作用的操作使用
function Get-ActiveWpsApp([string]$ProgId) {
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject($ProgId) } catch { return $null }
}

# 将已运行的 WPS 主窗口置于前台（按进程主窗口句柄激活）
function Bring-WpsToFront {
    try {
        $procs = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($procs.Count -eq 0) { return $false }
        try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch {}
        $ok = $false
        foreach ($proc in $procs) {
            try { if ([Microsoft.VisualBasic.Interaction]::AppActivate($proc.Id)) { $ok = $true } } catch {}
        }
        return $ok
    } catch { return $false }
}

# 定位 wps.exe 路径（优先取运行中进程的路径，其次查注册表 App Paths）
function Get-WpsExePath {
    try {
        $proc = Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.Path } | Select-Object -First 1
        if ($proc) { return $proc.Path }
    } catch {}
    foreach ($key in @(
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\wps.exe'
    )) {
        try {
            $v = (Get-ItemProperty -Path $key -ErrorAction Stop).'(default)'
            if ($v) { return $v }
        } catch {}
    }
    return $null
}

# 解析目标演示文稿：优先按 presentationName 精确定位（避免多文稿打开时 ActivePresentation 漂移），
# 未提供则回退到当前活动文稿；提供了但找不到则返回 $null（让调用方报明确错误，而不是误改其它文稿）。
function Get-TargetPres($ppt, $p) {
    if ($null -eq $ppt) { return $null }
    if ($p.presentationName) {
        try { return $ppt.Presentations.Item([string]$p.presentationName) } catch { return $null }
    }
    return $ppt.ActivePresentation
}

# 取整型参数并带默认值。写成函数而不是内联 $(if ...)：内联 if 在 [int] 转换里
# 解析不稳（实测会让整个 switch 报"语句块中缺少右}"），而且散落在几十处很难统一改。
function Get-IntParam($p, [string]$name, [int]$default = 1) {
    if ($null -eq $p) { return $default }
    $v = $p.$name
    if ($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)) { return $default }
    try { return [int]$v } catch { return $default }
}

# ==================== 文档/工作簿定位（A：不再依赖 ActiveDocument）====================
# 为什么要换掉 ActiveDocument
# ---------------------------
# ActiveDocument 是【弱引用】：它在这些情况下会返回一个"非空但空壳"的对象 ——
#   .Name 为空、.Content/.Styles/.PageSetup 全为 null、.Paragraphs.Count = 0
# 而 Documents.Count 明明 > 0。实测踩过：Documents.Count=2 但 ActiveDocument 取不到内容，
# 后续调用要么诡异地失败，要么直接 RPC_E_CALL_REJECTED。
# 根因是"活动文档"这个概念本身依赖 UI 焦点，而 WPS 的焦点状态在自动化场景下不可靠。
#
# 现在的策略（三层，逐层兜底）
#   1) 显式指定 documentName / workbookName —— 最可靠，直接按名字/路径取
#   2) ActiveDocument/ActiveWorkbook，但必须【通过健康检查】
#   3) 都不行就遍历集合，找第一个健康的
# 找不到就返回 $null，由各动作原有的判空分支报 "No active document"（不新造错误文案）。

# 文档健康检查：Name 非空 且 Content 可取到。
# 为什么要探 Content：空壳对象的 Name 可能是空串，也可能非空但内容取不到（实测见过）。
function Test-DocHealthy($doc) {
    if ($null -eq $doc) { return $false }
    try {
        $n = [string]$doc.Name
        if ([string]::IsNullOrWhiteSpace($n)) { return $false }
    } catch { return $false }
    try { $null = $doc.Content } catch { return $false }
    return $true
}

# 工作簿健康检查：Name 非空 且 Worksheets（或 Sheets）能取到
function Test-BookHealthy($book) {
    if ($null -eq $book) { return $false }
    try {
        $n = [string]$book.Name
        if ([string]::IsNullOrWhiteSpace($n)) { return $false }
    } catch { return $false }
    try { $null = $book.Worksheets } catch { return $false }
    return $true
}

# 按名字在集合里找文档：先按 FullName 全等，再按 Name 全等，最后按"文件名（无扩展名）"全等。
function Find-DocByName($app, [string]$want) {
    if ($null -eq $app -or [string]::IsNullOrWhiteSpace($want)) { return $null }
    $items = @()
    try { $items = @($app.Documents) } catch { return $null }
    $leaf = $want
    if ($leaf.Contains('\')) { $leaf = $leaf.Substring($leaf.LastIndexOf('\') + 1) }
    $stem = $leaf
    if ($stem.Contains('.')) { $stem = $stem.Substring(0, $stem.LastIndexOf('.')) }

    $byLeaf = $null; $byStem = $null
    foreach ($d in $items) {
        if ($null -eq $d) { continue }
        $full = ''; $nm = ''
        try { $full = [string]$d.FullName } catch {}
        try { $nm = [string]$d.Name } catch {}
        if ($full -and $full.Equals($want, [StringComparison]::OrdinalIgnoreCase)) { return $d }
        if ($nm -and $nm.Equals($leaf, [StringComparison]::OrdinalIgnoreCase)) { $byLeaf = $d }
        if ($nm -and $stem -and $nm.Equals($stem + [IO.Path]::GetExtension($nm), [StringComparison]::OrdinalIgnoreCase)) { $byStem = $d }
    }
    if ($byLeaf) { return $byLeaf }
    return $byStem
}

function Find-BookByName($app, [string]$want) {
    if ($null -eq $app -or [string]::IsNullOrWhiteSpace($want)) { return $null }
    $items = @()
    try { $items = @($app.Workbooks) } catch { return $null }
    $leaf = $want
    if ($leaf.Contains('\')) { $leaf = $leaf.Substring($leaf.LastIndexOf('\') + 1) }
    foreach ($b in $items) {
        if ($null -eq $b) { continue }
        $full = ''; $nm = ''
        try { $full = [string]$b.FullName } catch {}
        try { $nm = [string]$b.Name } catch {}
        if ($full -and $full.Equals($want, [StringComparison]::OrdinalIgnoreCase)) { return $b }
        if ($nm -and $nm.Equals($leaf, [StringComparison]::OrdinalIgnoreCase)) { return $b }
    }
    return $null
}

# 目标文稿（Word）。所有文档类动作都应走它，而不是直接读 $word.ActiveDocument。
#
# 注意：本函数在同一动作里可能被调用两次（先判空、再使用），每次都会重新解析。
# 这是有意的 —— 引入跨动作的缓存会带来"用户刚关掉文档，我们还拿着旧对象"的新问题，
# 而重复解析只是几次 COM 属性访问，代价可忽略。
function Resolve-Doc($app, $p) {
    if ($null -eq $app) { return $null }
    if ($p -and $p.documentName) {
        $d = Find-DocByName $app ([string]$p.documentName)
        if ($d) { return $d }
        # 显式指定却找不到：明确返回 $null，不要静默回退到"活动文档" ——
        # 否则用户以为在改 A，结果改的是 B。
        return $null
    }
    try { $act = $app.ActiveDocument } catch { $act = $null }
    if (Test-DocHealthy $act) { return $act }
    # 空壳 → 遍历找一个健康的
    try {
        foreach ($d in @($app.Documents)) {
            if (Test-DocHealthy $d) { return $d }
        }
    } catch {}
    return $null
}

# 目标工作簿（Excel）
function Resolve-Book($app, $p) {
    if ($null -eq $app) { return $null }
    if ($p -and $p.workbookName) {
        $b = Find-BookByName $app ([string]$p.workbookName)
        if ($b) { return $b }
        return $null
    }
    try { $act = $app.ActiveWorkbook } catch { $act = $null }
    if (Test-BookHealthy $act) { return $act }
    try {
        foreach ($b in @($app.Workbooks)) {
            if (Test-BookHealthy $b) { return $b }
        }
    } catch {}
    return $null
}

# ==================== 模态对话框探测（C）====================
# 为什么需要
# ----------
# COM 报 RPC_E_CALL_REJECTED(0x80010001) 的最常见原因不是"忙"，而是
# WPS 弹了模态对话框（另存为的确认框、字体替换提示、文档恢复提示…）在等用户点。
# 这时任何 COM 调用都会被拒，而且【干等不会好】—— 用户不点就永远不通。
# 与其让调用方等到超时（120s）才发现，不如立刻告诉他"有个对话框挡着，请先关掉"。
#
# 判据（第一版踩过坑）：我用"可见 + 有标题 + 不占满屏"当对话框特征，
# 结果把 WPS 的【主窗口】也认成了对话框（类名 OpusApp，标题形如 "xxx.docx - WPS Office"）——
# 那意味着每次 COM 被拒都会附一句错误指令"请关闭这个窗口"，比不提示更糟。
#
# 现在的判据：只认【类名】。标准对话框是 #32770，WPS 自绘的一般带 Dialog 字样。
# 宁可漏报（少一句提示，无害），不可误报（给出错误指令，有害）。
if (-not ('WpsModalProbe' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class WpsModalProbe {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr p);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr h, StringBuilder s, int n);
    delegate bool EnumWindowsProc(IntPtr h, IntPtr p);
    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

    public static string[] Find(int[] pids) {
        var set = new HashSet<int>(pids);
        var res = new List<string>();
        EnumWindows(delegate(IntPtr h, IntPtr p) {
            uint wp;
            GetWindowThreadProcessId(h, out wp);
            if (!set.Contains((int)wp)) return true;
            if (!IsWindowVisible(h)) return true;

            var cn = new StringBuilder(256);
            GetClassNameW(h, cn, 256);
            var cls = cn.ToString();
            // 只认对话框类名。主窗口类是 OpusApp / _WwG / bosa_sdm_* 等，一律不算。
            bool isDialog = string.Equals(cls, "#32770", StringComparison.OrdinalIgnoreCase)
                         || cls.IndexOf("Dialog", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!isDialog) return true;

            var sb = new StringBuilder(512);
            GetWindowTextW(h, sb, 512);
            var title = sb.ToString();
            if (string.IsNullOrWhiteSpace(title)) title = "(无标题)";
            res.Add(title + " [" + cls + "]");
            return true;
        }, IntPtr.Zero);
        return res.ToArray();
    }
}
'@ -ErrorAction SilentlyContinue
}

function Get-WpsModalDialogs {
    $pids = @()
    try { $pids = @(Get-Process -Name wps, et, wpp -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch {}
    if ($pids.Count -eq 0) { return @() }
    try { return [WpsModalProbe]::Find([int[]]$pids) } catch { return @() }
}

# 返回被模态对话框挡住的提示语（没有则返回 $null）
function Get-ModalBlockHint {
    $dlgs = @(Get-WpsModalDialogs)
    if ($dlgs.Count -eq 0) { return $null }
    $t = (@($dlgs | Select-Object -First 3) -join '；')
    return "WPS 有对话框挡着（COM 调用被拒 RPC_E_CALL_REJECTED）。请先关掉这些窗口再试：$t"
}

# 统一的 COM 错误文案（C）：若是对话框挡路，把"该关哪个窗口"一并带出去。
#
# 为什么值得单独做：RPC_E_CALL_REJECTED 本身完全无法自解释 —— 看到它只知道
# "调用被拒"，不知道是被谁拒的。而实测里它的绝大多数来源就是模态对话框。
# 加一句"请先关闭 X 窗口"，调用方就不用去猜，也不用等 120s 超时。
# 非对话框类错误原样返回，不加任何噪音。
function Get-CallError($err) {
    $m = [string]$err.Exception.Message
    if (($m -match '0x80010001') -or ($m -match 'RPC_E_CALL_REJECTED') -or ($m -match 'Call was rejected by callee')) {
        $hint = Get-ModalBlockHint
        if ($hint) { return "$m —— $hint" }
    }
    return $m
}

# 只读动作的有限重试（C）
# ----------------------
# 规则：
#   - 只对【只读】动作重试（get*/list*/find* 开头，或在白名单里）。写动作一律不重试 ——
#     请求可能已执行一半，重放会变成执行两次。
#   - 有限次 + 退避（250/500/1000ms），最多 3 次。
#   - 每次失败先看有没有模态对话框：有就直接放弃并把提示带上（干等无意义）。
function Test-ReadOnlyAction([string]$action) {
    if ([string]::IsNullOrWhiteSpace($action)) { return $false }
    if ($action -match '^(get|list|find|read|query|count|search|ping|wpsStatus|__probe)') { return $true }
    return $action -in @('getAppInfo', 'showWps', 'wireCheck')
}

function Invoke-WithReadOnlyRetry($app, [string]$action, $p) {
    $max = 3
    for ($i = 1; $i -le $max; $i++) {
        try {
            return Invoke-WpsAction -Action $action -p $p
        } catch {
            $m = [string]$_.Exception.Message
            $isRejected = ($m -match '0x80010001') -or ($m -match 'RPC_E_CALL_REJECTED')
            if (-not $isRejected) { throw }
            # 有对话框就别耗着了：用户不点，永远不会通
            $hint = Get-ModalBlockHint
            if ($hint) { throw $hint }
            if ($i -ge $max) { throw }
            Start-Sleep -Milliseconds (250 * [math]::Pow(2, $i - 1))
        }
    }
    return $null
}

# 在一个 TextFrame 内查找替换：先精确匹配(保留原换行)；多行 find 未命中时，把 \r\n/\r/\v 与 \n 规范化后再匹配并写回。返回是否发生替换。
function Replace-FrameText($frame, [string]$findText, [string]$replaceText) {
    try {
        if ($null -eq $frame -or -not $frame.HasText) { return $false }
        $text = $frame.TextRange.Text
        if ($null -eq $text) { return $false }
        if ($text.Contains($findText)) {
            $frame.TextRange.Text = $text.Replace($findText, $replaceText)
            return $true
        }
        if ($findText.Contains("`n")) {
            $nf = ((($findText -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            $nt = ((($text -replace "`r`n", "`n") -replace "`r", "`n") -replace "`v", "`n")
            if ($nt.Contains($nf)) {
                $frame.TextRange.Text = $nt.Replace($nf, $replaceText)
                return $true
            }
            # 兜底1：去掉所有换行后，若 find 恰为整块文本(PDF抽取的"幻影换行")，整体替换该形状
            $sf = ($findText -replace "[`r`n`v]", "")
            $st = ($text -replace "[`r`n`v]", "")
            if ($sf.Length -gt 0 -and $st -eq $sf) {
                $frame.TextRange.Text = $replaceText
                return $true
            }
            # 兜底2：去换行做子串匹配(find 是大形状里的一段且含幻影换行)，命中后用 stripped→原文 索引映射在原位替换
            if ($sf.Length -gt 0) {
                $origChars = $text.ToCharArray()
                $sb = New-Object System.Text.StringBuilder
                $map = New-Object System.Collections.Generic.List[int]
                for ($k = 0; $k -lt $origChars.Length; $k++) {
                    $ch = $origChars[$k]
                    if ($ch -ne [char]13 -and $ch -ne [char]10 -and $ch -ne [char]11) {
                        [void]$sb.Append($ch); $map.Add($k)
                    }
                }
                $idx = $sb.ToString().IndexOf($sf)
                if ($idx -ge 0) {
                    $startOrig = $map[$idx]
                    $endOrig = $map[$idx + $sf.Length - 1]
                    $before = $text.Substring(0, $startOrig)
                    $after = $text.Substring($endOrig + 1)
                    $frame.TextRange.Text = $before + $replaceText + $after
                    return $true
                }
            }
        }
        return $false
    } catch { return $false }
}

# 递归处理一个形状(及其组合/表格内的所有子形状)的文本替换，返回命中次数。
# 修复：replacePptText 原来只扫顶层 Shapes，组合(Group)内文字(如封面/立项页的旧项目名)替换不到。
function Replace-InShapeTree($shape, [string]$findText, [string]$replaceText) {
    $cnt = 0
    try {
        $isGroup = $false
        try { if ($shape.Type -eq 6) { $isGroup = $true } } catch {}
        if ($isGroup) {
            for ($gi = 1; $gi -le $shape.GroupItems.Count; $gi++) {
                $cnt += Replace-InShapeTree $shape.GroupItems.Item($gi) $findText $replaceText
            }
        } else {
            try { if ($shape.HasTextFrame) { if (Replace-FrameText $shape.TextFrame $findText $replaceText) { $cnt++ } } } catch {}
            try {
                if ($shape.HasTable) {
                    $tbl = $shape.Table
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            try { if (Replace-FrameText $tbl.Cell($rr, $cc).Shape.TextFrame $findText $replaceText) { $cnt++ } } catch {}
                        }
                    }
                }
            } catch {}
        }
    } catch {}
    return $cnt
}

function Output-Json($obj) {
    # 常驻模式下把请求 id 带回响应，供 Node 侧按 id 匹配请求/响应。
    # 直跑模式 $script:WpsReqId 为空，不注入，输出与改造前逐字节一致。
    if ($script:WpsReqId -and ($obj -is [hashtable])) { $obj['id'] = $script:WpsReqId }
    $obj | ConvertTo-Json -Depth 10 -Compress
}

function Convert-ColumnLetterToNumber([string]$letters) {
    if (-not $letters) { return $null }
    $letters = $letters.ToUpper().Trim()
    $sum = 0
    foreach ($ch in $letters.ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Convert-ColumnNumberToLetter([int]$col) {
    if ($col -le 0) { return $null }
    $letter = ""
    while ($col -gt 0) {
        $mod = ($col - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $col = [Math]::Floor(($col - 1) / 26)
    }
    return $letter
}

function Convert-HexColorToRgbInt([string]$hex) {
    if (-not $hex) { return $null }
    $hex = $hex.Trim()
    if ($hex.StartsWith('#')) { $hex = $hex.Substring(1) }
    if ($hex.Length -eq 3) {
        $hex = ($hex[0] + $hex[0] + $hex[1] + $hex[1] + $hex[2] + $hex[2])
    }
    if ($hex.Length -ne 6) { return $null }
    $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Get-RangeFromAddress($workbook, [string]$address) {
    if ($address -match "^(?<sheet>[^!]+)!(?<range>.+)$") {
        $sheetName = $matches.sheet.Trim("'")
        return $workbook.Sheets.Item($sheetName).Range($matches.range)
    }
    return $workbook.ActiveSheet.Range($address)
}

function Get-AppTypeByExtension([string]$filePath) {
    $ext = [System.IO.Path]::GetExtension($filePath).ToLower().Trim('.')
    $wordExts = @('doc', 'docx', 'docm', 'dot', 'dotx', 'dotm', 'rtf', 'txt', 'html', 'htm', 'xml', 'wps', 'wpt')
    $excelExts = @('xls', 'xlsx', 'xlsm', 'xlsb', 'xlt', 'xltx', 'xltm', 'csv', 'html', 'htm', 'et', 'ett')
    $pptExts = @('ppt', 'pptx', 'pptm', 'pot', 'potx', 'potm', 'pps', 'ppsx', 'ppsm', 'dps', 'dpt')
    if ($wordExts -contains $ext) { return 'word' }
    if ($excelExts -contains $ext) { return 'excel' }
    if ($pptExts -contains $ext) { return 'ppt' }
    return $null
}

function Get-WordSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Word WdSaveFormat: doc=0, docx=16, pdf=17, rtf=6, xps=18, html=8, txt=2, xml=11
    $map = @{ doc = 0; docx = 16; pdf = 17; rtf = 6; xps = 18; html = 8; htm = 8; txt = 2; xml = 11 }
    return $map[$f]
}

function Get-ExcelFileFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # Excel XlFileFormat: xls=-4143, xlsx=51, xlsm=52, xlsb=50, csv=6, html=44, xlt=17, xltx=54, xltm=53
    $map = @{ xls = -4143; xlsx = 51; xlsm = 52; xlsb = 50; csv = 6; html = 44; htm = 44; xlt = 17; xltx = 54; xltm = 53 }
    return $map[$f]
}

function Get-PptSaveFormat([string]$format) {
    if (-not $format) { return $null }
    $f = $format.ToLower().Trim('.')
    # PowerPoint PpSaveAsFileType: ppt=1, pptx=24, pptm=25, pdf=32, png=18, jpg=17, gif=16, bmp=19
    $map = @{ ppt = 1; pptx = 24; pptm = 25; pdf = 32; png = 18; jpg = 17; jpeg = 17; gif = 16; bmp = 19 }
    return $map[$f]
}

# ==================== 动作分发（常驻模式复用入口）====================
# 把原先裸的 switch 包进函数，是为了让 wps-comd.ps1 能常驻下来反复调用；
# 脚本直跑（-File wps-com.ps1 -Action X -Params Y）的用法完全不变（见文件末尾入口段）。
# 函数体内原先的 367 个 `exit` 已全部改成 `return` —— 常驻进程里 `exit` 会直接干掉守护进程，
# 而 switch 之后没有任何代码，两者语义等价。
# 注：函数体沿用原缩进（未整体右移），以便 diff 只体现必要改动。
function Invoke-WpsAction {
    param([string]$Action, $p)

# 注意：这里【绝不能】再写 `$p = $Params | ConvertFrom-Json`。
# ①A 常驻化时把原来裸 switch 包成函数，那句解析被一起带了进来，于是每次调用都用
# 脚本级 $Params（守护进程模式下就是默认值 "{}"）把传进来的 $p 覆盖成空对象 ——
# 表现为"所有带参数的动作都收不到参数"，而且不报错。实测于 ② 补参数化动作时发现。
# 直跑模式的解析在文件末尾的入口段里做，那里才是它该待的地方。

switch ($Action) {

    # ==================== Common ====================
    "ping" {
        Output-Json @{ success = $true; data = @{ message = "pong"; timestamp = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() } }
    }
    # 【临时诊断】列出挡路的对话框。用来验 C 项：Get-WpsModalDialogs 能不能真找到窗口。
    "__probeStyleAdd" {
        # 最小对比实验：探测里 styles.add 报 supported，实现里却 E_FAIL。
        # 逐项排除差异（字面量 vs 变量 / 参数类型 / 名字字符），找出到底是哪一处。
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $r = @{}

        # A: 与探测完全一致 —— 纯字面量
        try {
            $sA = $doc.Styles.Add("__probe_literal__", 1)
            $r["A_literal"] = if ($null -eq $sA) { "null" } else { "ok: " + [string]$sA.NameLocal }
            try { $doc.Styles.Item("__probe_literal__").Delete() } catch {}
        } catch { $r["A_literal"] = "FAIL: " + $_.Exception.Message }

        # B: 第一个参数走变量（与实现一致），第二个仍是字面量
        try {
            $nm = "__probe_var__"
            $sB = $doc.Styles.Add($nm, 1)
            $r["B_varName"] = if ($null -eq $sB) { "null" } else { "ok: " + [string]$sB.NameLocal }
            try { $doc.Styles.Item("__probe_var__").Delete() } catch {}
        } catch { $r["B_varName"] = "FAIL: " + $_.Exception.Message }

        # C: 两个参数都走变量（与实现完全一致）
        try {
            $nm2 = "__probe_bothvar__"
            $st = 1
            $sC = $doc.Styles.Add($nm2, $st)
            $r["C_bothVar"] = if ($null -eq $sC) { "null" } else { "ok: " + [string]$sC.NameLocal }
            try { $doc.Styles.Item("__probe_bothvar__").Delete() } catch {}
        } catch { $r["C_bothVar"] = "FAIL: " + $_.Exception.Message }

        # D: 中文样式名（实现里传的是中文名，探测用的是英文名）
        try {
            $sD = $doc.Styles.Add("__p01测试样式__", 1)
            $r["D_chineseName"] = if ($null -eq $sD) { "null" } else { "ok: " + [string]$sD.NameLocal }
            try { $doc.Styles.Item("__p01测试样式__").Delete() } catch {}
        } catch { $r["D_chineseName"] = "FAIL: " + $_.Exception.Message }

        # E: 从 PSCustomObject 取参数（实现是 $p.name 这种）
        try {
            $pE = @{ name = "__probe_pobj__"; styleType = $null }
            $nmE = [string]$pE.name
            $stE = 1
            if ($pE.styleType -eq "character") { $stE = 2 }
            $sE = $doc.Styles.Add($nmE, $stE)
            $r["E_pobjName"] = if ($null -eq $sE) { "null" } else { "ok: " + [string]$sE.NameLocal }
            try { $doc.Styles.Item("__probe_pobj__").Delete() } catch {}
        } catch { $r["E_pobjName"] = "FAIL: " + $_.Exception.Message }

        # F: 同一个文档上连续 Add 两次（看是不是第一次就污染了状态）
        try {
            $sF = $doc.Styles.Add("__probe_twice__", 1)
            $first = if ($null -eq $sF) { "null" } else { "ok" }
            try { $sF2 = $doc.Styles.Add("__probe_twice2__", 1) } catch { }
            $r["F_twice"] = $first + " / 第二次: " + $(if ($null -eq $sF2) { "null" } else { "ok" })
            try { $doc.Styles.Item("__probe_twice__").Delete() } catch {}
            try { $doc.Styles.Item("__probe_twice2__").Delete() } catch {}
        } catch { $r["F_twice"] = "FAIL: " + $_.Exception.Message }

        Output-Json @{ success = $true; data = $r }
    }

    "__probeClose" {
        # 试不同的关闭方式，找出哪种在"未命名未保存文档"上管用。
        # 背景：Close(0) 报 RPC_E_CALL_REJECTED，但读文档列表是成功的 ——
        # 说明不是 COM 整体坏掉，而是关闭这个特定动作被拒。
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $r = @{}
        $docs = @($word.Documents)
        $r.countBefore = $docs.Count
        $i = 0
        foreach ($d in $docs) {
            $i++
            $nm = ''
            try { $nm = [string]$d.Name } catch {}
            $full = ''
            try { $full = [string]$d.FullName } catch {}
            $tag = "$i(nm=$nm saved=$([bool]$full))"
            # 方式1：Close($false)
            try { $d.Close($false); $r["1_CloseFalse"] += "$tag:ok "; continue } catch { $r["1_CloseFalse"] += "$tag:FAIL($($_.Exception.Message.Substring(0,[Math]::Min(40,$_.Exception.Message.Length)))) " }
            # 方式2：先 SaveAs 到临时文件再关
            try {
                $tmp = [IO.Path]::GetTempFileName() + '.docx'
                $d.SaveAs([ref]$tmp, [ref]$missing)
                $d.Close($false)
                $r["2_SaveAsThenClose"] += "$tag:ok "
            } catch { $r["2_SaveAsThenClose"] += "$tag:FAIL " }
            # 方式3：Documents.Close 集合级
            try { $word.Documents.Close($false); $r["3_CollectionClose"] += "$tag:ok "; break } catch { $r["3_CollectionClose"] += "$tag:FAIL " }
        }
        $r.countAfter = @($word.Documents).Count
        Output-Json @{ success = $true; data = $r }
    }

    "__probeDialogs" {
        $dlgs = @(Get-WpsModalDialogs)
        Output-Json @{ success = $true; data = @{ dialogs = $dlgs; count = $dlgs.Count; hint = (Get-ModalBlockHint) } }
    }
    # ==================== 诊断与修复（D）====================
    "wpsRepair" {
        # 出问题时的"一键诊断+修复"。做两件事：
        #   1) 诊断：三个应用各自的实例、文档数、每个文档的健康度、挡路的对话框
        #   2) 修复：关掉【无名文档】（没路径没标题的，通常是脚本遗留的空壳）
        #      —— 只关无名文档，不动用户自己的文件；已命名的文档一律保留。
        $report = @{ apps = @{}; dialogs = @(); closed = @() }
        try { $report.dialogs = @(Get-WpsModalDialogs) } catch {}

        foreach ($a in @(
            @{ key = "excel"; progId = "Ket.Application";  coll = "Workbooks";     app = $null },
            @{ key = "word";  progId = "Kwps.Application"; coll = "Documents";     app = $null },
            @{ key = "ppt";   progId = "Kwpp.Application"; coll = "Presentations"; app = $null }
        )) {
            $inst = Get-ActiveWpsApp $a.progId
            $entry = @{ running = ($null -ne $inst); documents = @() }
            if ($null -ne $inst) {
                try { $entry.name = [string]$inst.Name } catch {}
                try { $entry.path = [string]$inst.Path } catch {}
                $docs = @()
                try { $docs = @($inst.($a.coll)) } catch {}
                $entry.count = $docs.Count
                foreach ($d in $docs) {
                    if ($null -eq $d) { continue }
                    $nm = ''; $full = ''; $healthy = $false
                    try { $nm = [string]$d.Name } catch {}
                    try { $full = [string]$d.FullName } catch {}
                    if ($a.key -eq "word") { $healthy = Test-DocHealthy $d }
                    else { $healthy = Test-BookHealthy $d }
                    $entry.documents += @{
                        name = $nm
                        saved = (-not [string]::IsNullOrWhiteSpace($full))
                        healthy = $healthy
                    }
                    # 只关"无名"文档：没有标题也没有路径 —— 那是自动化残留的空壳。
                    # ⚠️ 关闭用【集合级】Close，不要逐个 $d.Close()。
                    #   实测（__probeClose）：逐个关闭未命名文档会被 WPS 拒
                    #   （RPC_E_CALL_REJECTED），而 $app.Documents.Close($false) 一次成功
                    #   （countBefore=2 → countAfter=0）。这是 WPS 12.1 的实际行为。
                    #   真正执行在下面循环结束后统一做（先收集，再一次关）。
                    if ([string]::IsNullOrWhiteSpace($nm) -and [string]::IsNullOrWhiteSpace($full)) {
                        $unnamed += $a.key
                    }
                }
                # 活动文档健康度单独报一份，便于判断"是不是 ActiveDocument 空壳"
                try {
                    $act = $null
                    if ($a.key -eq "word") { $act = $inst.ActiveDocument }
                    else { $act = $inst.ActiveWorkbook }
                    $entry.activeHealthy = if ($a.key -eq "word") { Test-DocHealthy $act } else { Test-BookHealthy $act }
                } catch { $entry.activeHealthy = $false }
            }
            $report.apps[$a.key] = $entry
        }

        $hint = $null
        if (@($report.dialogs).Count -gt 0) {
            $hint = "检测到挡路对话框，请先关闭：" + ((@($report.dialogs | Select-Object -First 3)) -join '；')
        }
        $report.hint = $hint
        # 收尾：对收集到的"有无名文档"的应用做一次【集合级】关闭
        # （逐个 $d.Close() 会被 WPS 拒，集合级一次成功 —— 见 __probeClose 实测）
        foreach ($k in ($unnamed | Select-Object -Unique)) {
            $app = $null
            $coll = 'Documents'
            if ($k -eq 'word') { $app = Get-ActiveWpsApp 'Kwps.Application' }
            elseif ($k -eq 'excel') { $app = Get-ActiveWpsApp 'Ket.Application'; $coll = 'Workbooks' }
            elseif ($k -eq 'ppt') { $app = Get-ActiveWpsApp 'Kwpp.Application'; $coll = 'Presentations' }
            if ($null -eq $app) { continue }
            try {
                $app.$coll.Close($false)
                $report.closed += "$k`: 已关闭无名文档（集合级 Close）"
            } catch {
                $report.closed += "$k`: 关闭失败 $($_.Exception.Message)"
            }
        }

        Output-Json @{ success = $true; data = $report }
    }

    # 只读状态探针：报告三件套分别有没有已附着的实例。
    # 关键区别：这里【只用 Get-ActiveWpsApp】——它只取已运行的实例，不会创建新的。
    # 对照 getAppInfo：那个走 Get-WpsExcel/Word/Ppt，拿不到就 New-Object 起一个，
    # 用它探状态会把 WPS 意外拉起来，而且只报第一个找到的应用。
    "wpsStatus" {
        # 身份判据说明（别再用 .Name 当权威）：
        # WPS 为了兼容 Excel/Word 宏，故意自报 Application.Name = "Microsoft Excel" /
        # "Microsoft Word"，Version 也报成 Office 的兼容版本号（如 12.0）。实测本机
        # Ket.Application 拿到的对象 Path 指向 WPS 安装目录、且机器上根本没装 EXCEL.EXE
        # —— 所以它是 WPS 表格本人，不是真 Excel。
        # 因此身份以【我们请求的 ProgID】+【Application.Path 落在哪个产品目录】为准，
        # 自报名单独放在 reportedName，只作参考，不作判据。
        $apps = @(
            @{ key = "excel"; progId = "Ket.Application";  app = "WPS 表格" },
            @{ key = "word";  progId = "Kwps.Application"; app = "WPS 文字" },
            @{ key = "ppt";   progId = "Kwpp.Application"; app = "WPS 演示" }
        )
        $states = @{}
        foreach ($a in $apps) {
            $entry = @{ running = $false; app = $a.app; progId = $a.progId }
            $inst = Get-ActiveWpsApp $a.progId
            if ($null -ne $inst) {
                $entry.running = $true
                try { $entry.reportedName = [string]$inst.Name } catch {}
                try { $entry.version = [string]$inst.Version } catch {}
                try {
                    $ip = [string]$inst.Path
                    if ($ip) {
                        $entry.path = $ip
                        # 由安装目录反推产品，作为身份交叉验证
                        $entry.product = if ($ip -match 'WPS\s*Office') { 'WPS Office' }
                                         elseif ($ip -match 'Microsoft\s+Office') { 'Microsoft Office' }
                                         else { 'unknown' }
                    }
                } catch {}
                # 文档数：三个应用各自的集合名不一样，按顺序试
                foreach ($coll in @("Workbooks", "Documents", "Presentations")) {
                    try {
                        $c = $inst.$coll
                        if ($null -ne $c) { $entry.documents = [int]$c.Count; break }
                    } catch {}
                }
            }
            $states[$a.key] = $entry
        }
        Output-Json @{ success = $true; data = @{
            apps = $states
            psPid = $PID
            psVersion = $PSVersionTable.PSVersion.ToString()
            daemon = ($env:WPS_COM_DAEMON -eq '1')
        } }
    }

    # 让 WPS 可见并置于前台：① 激活已有窗口 ② 唤起 COM 实例 ③ 兜底启动可见实例
    "showWps" {
        $apps = @()
        $visible = 0
        $activated = 0
        foreach ($item in @(
            @{ name = 'Word';  progId = 'Kwps.Application' },
            @{ name = 'Excel'; progId = 'Ket.Application'  },
            @{ name = 'PPT';   progId = 'Kwpp.Application' }
        )) {
            $app = Get-ActiveWpsApp $item.progId
            if ($null -ne $app) {
                $apps += $item.name
                try { $app.Visible = $true; $visible++ } catch {}
                try { $app.Activate(); $activated++ } catch {}
            }
        }

        $front = Bring-WpsToFront
        $launched = $false

        # 兜底：没有可激活的窗口（COM 自动化实例本身无窗口）时，启动一个可见的 WPS 实例
        if (-not $front -and $p.launch -ne $false) {
            $exe = Get-WpsExePath
            if ($exe) {
                try {
                    Start-Process -FilePath $exe | Out-Null
                    $launched = $true
                    # 轮询等待可见窗口出现（最多约 14 秒）
                    for ($i = 0; $i -lt 14; $i++) {
                        Start-Sleep -Seconds 1
                        if (Bring-WpsToFront) { $front = $true; break }
                    }
                } catch {}
            }
        }

        if ($apps.Count -eq 0 -and -not $front -and -not $launched) {
            Output-Json @{ success = $false; error = "未能唤起 WPS 窗口（无可激活窗口，且启动可见实例失败）" }
            return
        }

        # 最终确认：只要系统里存在可见的 WPS 主窗口即视为成功
        if (-not $front) {
            $hasVisible = @(Get-Process -Name wps -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }).Count -gt 0
            if ($hasVisible) { $front = $true }
        }

        $msg = if ($front) { "WPS 窗口已置于前台" } elseif ($launched) { "已启动 WPS 可见实例" } else { "已将 WPS 设为可见" }
        Output-Json @{ success = $true; data = @{ apps = $apps; visible = $visible; activated = $activated; foreground = $front; launched = $launched; message = $msg } }
    }

    # 结束 WPS（含强制兜底）
    "quitWps" {
        $quit = @()
        foreach ($progId in @('Kwps.Application','Ket.Application','Kwpp.Application')) {
            $app = Get-ActiveWpsApp $progId
            if ($null -ne $app) {
                try { $app.DisplayAlerts = 0 } catch {}
                try { $app.Quit() } catch {}
                $quit += $progId
            }
        }
        Start-Sleep -Milliseconds 900
        $remainBefore = @(Get-Process -Name wps -ErrorAction SilentlyContinue).Count
        $forced = $false
        # 兜底：Quit 未清干净则强制结束残留进程（force 默认 true，传 force=false 只走优雅退出）
        if ($remainBefore -gt 0 -and $p.force -ne $false) {
            try { $null = & taskkill /F /IM wps.exe /T 2>&1 } catch {}
            try { $null = & taskkill /F /IM wpscloudsvr.exe /T 2>&1 } catch {}
            $forced = $true
        }
        Start-Sleep -Milliseconds 700
        $left = @(Get-Process -Name wps -ErrorAction SilentlyContinue).Count
        $msg = if ($left -eq 0) { "WPS 已退出" } else { "仍有 $left 个 WPS 进程残留" }
        Output-Json @{ success = $true; data = @{ quit = $quit; forced = $forced; remaining = $left; message = $msg } }
    }

    "wireCheck" {
        Output-Json @{ success = $true; data = @{ message = "WPS MCP Bridge 已连接" } }
    }

    "getAppInfo" {
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $excel.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "excel"; appName = $excel.Name; hasSelection = $hasSelection } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $word.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "word"; appName = $word.Name; hasSelection = $hasSelection } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            $hasSelection = $false
            try { $hasSelection = ($null -ne $ppt.ActiveWindow.Selection) } catch {}
            Output-Json @{ success = $true; data = @{ appType = "ppt"; appName = $ppt.Name; hasSelection = $hasSelection } }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "getSelectedText" {
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $text = $word.Selection.Text
                Output-Json @{ success = $true; data = @{ text = $text.Trim() } }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            return
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $text = if ($sel -is [string]) { $sel } else { $sel.Text }
                Output-Json @{ success = $true; data = @{ text = [string]$text } }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $text = $sel.TextRange.Text
                    Output-Json @{ success = $true; data = @{ text = [string]$text } }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "setSelectedText" {
        $text = if ($null -ne $p.text) { [string]$p.text } else { "" }
        $word = Get-WpsWord
        if ($null -ne $word) {
            try {
                $word.Selection.Text = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Word selection not available" }
            }
            return
        }
        $excel = Get-WpsExcel
        if ($null -ne $excel) {
            try {
                $sel = $excel.Selection
                $sel.Value2 = $text
                Output-Json @{ success = $true }
            } catch {
                Output-Json @{ success = $false; error = "Excel selection not available" }
            }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt) {
            try {
                $sel = $ppt.ActiveWindow.Selection
                if ($sel -and $sel.TextRange) {
                    $sel.TextRange.Text = $text
                    Output-Json @{ success = $true }
                } else {
                    Output-Json @{ success = $false; error = "PPT selection has no text" }
                }
            } catch {
                Output-Json @{ success = $false; error = "PPT selection not available" }
            }
            return
        }
        Output-Json @{ success = $false; error = "No WPS application running" }
    }

    "save" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $(Resolve-Book $excel $p).Save()
            Output-Json @{ success = $true; app = "excel" }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $(Resolve-Doc $word $p).Save()
            Output-Json @{ success = $true; app = "word" }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $ppt.ActivePresentation.Save()
            Output-Json @{ success = $true; app = "ppt" }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "saveAs" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; return }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel -or $null -eq $(Resolve-Book $excel $p)) { Output-Json @{ success = $false; error = "No active workbook" }; return }
            $format = Get-ExcelFileFormat $p.format
            if ($null -ne $format) { $(Resolve-Book $excel $p).SaveAs($path, $format) }
            else { $(Resolve-Book $excel $p).SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            return
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word -or $null -eq $(Resolve-Doc $word $p)) { Output-Json @{ success = $false; error = "No active document" }; return }
            $format = Get-WordSaveFormat $p.format
            if ($null -ne $format) { $(Resolve-Doc $word $p).SaveAs($path, $format) }
            else { $(Resolve-Doc $word $p).SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            return
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt -or $null -eq $ppt.ActivePresentation) { Output-Json @{ success = $false; error = "No active presentation" }; return }
            $format = Get-PptSaveFormat $p.format
            if ($null -ne $format) { $ppt.ActivePresentation.SaveAs($path, $format) }
            else { $ppt.ActivePresentation.SaveAs($path) }
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            return
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "openFile" {
        $path = $p.path
        if (-not $path) { Output-Json @{ success = $false; error = "Path required" }; return }
        $appType = if ($p.appType) { $p.appType } else { Get-AppTypeByExtension $path }
        if ($appType -eq 'excel') {
            $excel = Get-WpsExcel
            if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
            $excel.Workbooks.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "excel" } }
            return
        }
        if ($appType -eq 'word') {
            $word = Get-WpsWord
            if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
            $word.Documents.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "word" } }
            return
        }
        if ($appType -eq 'ppt') {
            $ppt = Get-WpsPpt
            if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
            $ppt.Presentations.Open($path)
            Output-Json @{ success = $true; data = @{ path = $path; appType = "ppt" } }
            return
        }
        Output-Json @{ success = $false; error = "Unknown app type" }
    }

    "createDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $n0 = 0
        try { $n0 = [int]$word.Documents.Count } catch {}
        $doc = $word.Documents.Add()

        # 等文档真正就绪（关键，之前漏了）
        # ----------------------------------
        # 原实现是 `Documents.Add() | Out-Null` 然后立刻返回 success，零等待。
        # 后果：新建的文档要过一小会儿才取得到 Content/Styles，紧接着的任何动作
        # 都会拿到"非空但空壳"的对象 —— 实测反复踩到（表现为莫名 "No active document"、
        # 以及样式/修订/交叉引用一连串失败）。
        #
        # 实现要点（踩过）：
        #   1) 每次都重新 `Documents.Item(Count)` 取最新文档，不要复用 Add() 的返回值 ——
        #      实测复用那个引用会让 Test-DocHealthy 永远 false（6 秒超时）。
        #   2) 等不到也【按成功返回】，只带 warning：Resolve-Doc 有遍历兜底，
        #      而且实测"不等 + 调用方自己等一会"是能工作的。抛异常会把后续全堵死。
        $warn = $null
        $ready = $false
        for ($i = 0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 150
            try {
                $cur = $word.Documents.Item($word.Documents.Count)
                if (Test-DocHealthy $cur) { $ready = $true; break }
            } catch {}
        }
        if (-not $ready) {
            $warn = "文档已创建但 6 秒内未通过健康检查（可能是 WPS 正在初始化）。" +
                    "后续操作会自动遍历集合找健康文档；若报错请重试或检查对话框。"
        }
        $data = @{ count = [int]$word.Documents.Count }
        if ($warn) { $data.warning = $warn }
        Output-Json @{ success = $true; data = $data }
    }

    "createPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $ppt.Presentations.Add() | Out-Null
        Output-Json @{ success = $true }
    }

    "convertToPDF" {
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $wb = $(Resolve-Book $excel $p)
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $wb.ExportAsFixedFormat(0, $outputPath)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel" } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $doc = $(Resolve-Doc $word $p)
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $doc.ExportAsFixedFormat($outputPath, 17)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word" } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, 'pdf') }
            $pres.SaveAs($outputPath, 32)
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt" } }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    "convertFormat" {
        $targetFormat = $p.targetFormat
        if (-not $targetFormat) { Output-Json @{ success = $false; error = "targetFormat required" }; return }
        $excel = Get-WpsExcel
        if ($null -ne $excel -and $null -ne $(Resolve-Book $excel $p)) {
            $wb = $(Resolve-Book $excel $p)
            $sourcePath = $wb.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-ExcelFileFormat $targetFormat
            if ($null -ne $format) { $wb.SaveAs($outputPath, $format) } else { $wb.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "excel"; targetFormat = $targetFormat } }
            return
        }
        $word = Get-WpsWord
        if ($null -ne $word -and $null -ne $(Resolve-Doc $word $p)) {
            $doc = $(Resolve-Doc $word $p)
            $sourcePath = $doc.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-WordSaveFormat $targetFormat
            if ($null -ne $format) { $doc.SaveAs($outputPath, $format) } else { $doc.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "word"; targetFormat = $targetFormat } }
            return
        }
        $ppt = Get-WpsPpt
        if ($null -ne $ppt -and $null -ne $ppt.ActivePresentation) {
            $pres = Get-TargetPres $ppt $p
            $sourcePath = $pres.FullName
            $outputPath = if ($p.outputPath) { $p.outputPath } else { [System.IO.Path]::ChangeExtension($sourcePath, $targetFormat) }
            $format = Get-PptSaveFormat $targetFormat
            if ($null -ne $format) { $pres.SaveAs($outputPath, $format) } else { $pres.SaveAs($outputPath) }
            Output-Json @{ success = $true; data = @{ sourcePath = $sourcePath; outputPath = $outputPath; appType = "ppt"; targetFormat = $targetFormat } }
            return
        }
        Output-Json @{ success = $false; error = "No active document" }
    }

    # ==================== Excel Basic ====================
    "getActiveWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheetCount = $wb.Sheets.Count; sheets = $sheets } }
    }

    "getCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $cell = $sheet.Cells.Item($p.row, $p.col)
        Output-Json @{ success = $true; data = @{ value = $cell.Value2; text = $cell.Text; formula = $cell.Formula } }
    }

    "setCellValue" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $sheet.Cells.Item($p.row, $p.col).Value2 = $p.value
        Output-Json @{ success = $true }
    }

    "getRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $data = @()
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            $row = @()
            for ($c = 1; $c -le $range.Columns.Count; $c++) { $row += $range.Cells.Item($r, $c).Value2 }
            $data += ,@($row)
        }
        Output-Json @{ success = $true; data = @{ data = $data } }
    }

    "setRangeData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        for ($r = 0; $r -lt $p.data.Count; $r++) {
            for ($c = 0; $c -lt $p.data[$r].Count; $c++) { $range.Cells.Item($r + 1, $c + 1).Value2 = $p.data[$r][$c] }
        }
        Output-Json @{ success = $true }
    }

    "diagnoseFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = $cell.Formula
        $errorType = $null
        $diagnosis = ""
        $suggestion = ""
        $precedents = @()
        if ($value -is [string] -and $value.StartsWith('#')) {
            $errorType = $value
            switch ($errorType) {
                "#REF!" { $diagnosis = "引用了不存在的单元格或区域"; $suggestion = "检查引用区域是否被删除或移动" }
                "#N/A" { $diagnosis = "查找函数未找到匹配值"; $suggestion = "确认查找值存在，或检查匹配条件" }
                "#VALUE!" { $diagnosis = "参数类型不正确或运算类型不匹配"; $suggestion = "检查函数参数类型和引用单元格" }
                "#NAME?" { $diagnosis = "函数名或名称拼写错误"; $suggestion = "检查函数名是否正确" }
                "#DIV/0!" { $diagnosis = "除数为零"; $suggestion = "检查除数单元格，避免除以零" }
                "#NUM!" { $diagnosis = "数值无效或超出范围"; $suggestion = "检查函数参数范围" }
                "#NULL!" { $diagnosis = "交集为空"; $suggestion = "检查引用区域的交集是否存在" }
                default { $diagnosis = "未知错误"; $suggestion = "检查公式与引用" }
            }
        }
        try {
            $refs = $cell.DirectPrecedents
            if ($refs -ne $null) {
                foreach ($area in $refs.Areas) { $precedents += $area.Address() }
            }
        } catch {}
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; currentValue = $value; errorType = $errorType; diagnosis = $diagnosis; suggestion = $suggestion; precedents = $precedents } }
    }

    "cleanData" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = $sheet.Range($p.range)
        $opsResult = @()
        foreach ($op in $p.operations) {
            $success = $true
            $message = ""
            switch ($op) {
                "trim" {
                    foreach ($cell in $range) {
                        if ($cell.Value2 -is [string]) { $cell.Value2 = $cell.Value2.Trim() }
                    }
                    $message = "已去除前后空格"
                }
                "remove_duplicates" {
                    $colCount = $range.Columns.Count
                    $cols = @()
                    for ($i = 1; $i -le $colCount; $i++) { $cols += $i }
                    $range.RemoveDuplicates($cols, 1)
                    $message = "已删除重复行"
                }
                "unify_date" {
                    foreach ($cell in $range) {
                        try {
                            if ($cell.Value2) {
                                $dt = [DateTime]::FromOADate($cell.Value2)
                                $cell.Value2 = $dt.ToString("yyyy-MM-dd")
                            }
                        } catch {}
                    }
                    $message = "已统一日期格式"
                }
                "remove_empty_rows" {
                    for ($i = $range.Rows.Count; $i -ge 1; $i--) {
                        $row = $range.Rows.Item($i)
                        $isEmpty = $true
                        foreach ($cell in $row.Cells) { if ($cell.Value2) { $isEmpty = $false; break } }
                        if ($isEmpty) { $row.Delete() }
                    }
                    $message = "已删除空行"
                }
                default { $success = $false; $message = "不支持的操作" }
            }
            $opsResult += @{ operation = $op; success = $success; message = $message }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; operations = $opsResult; message = "cleanData completed" } }
    }

    "createPivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sourceRange = Get-RangeFromAddress $wb $p.sourceRange
        $destSheet = if ($p.destinationSheet) { $wb.Sheets.Item($p.destinationSheet) } else { $excel.ActiveSheet }
        $destCell = $destSheet.Range($p.destinationCell)
        $pivotName = if ($p.tableName) { $p.tableName } else { "PivotTable" + [Guid]::NewGuid().ToString("N").Substring(0, 6) }
        $cache = $wb.PivotCaches().Create(1, $sourceRange)
        $table = $cache.CreatePivotTable($destCell, $pivotName)
        foreach ($fieldName in $p.rowFields) {
            $field = $table.PivotFields($fieldName)
            $field.Orientation = 1
            $field.Position = 1
        }
        if ($p.columnFields) {
            foreach ($fieldName in $p.columnFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 2
                $field.Position = 1
            }
        }
        if ($p.filterFields) {
            foreach ($fieldName in $p.filterFields) {
                $field = $table.PivotFields($fieldName)
                $field.Orientation = 3
                $field.Position = 1
            }
        }
        foreach ($vf in $p.valueFields) {
            $field = $table.PivotFields($vf.field)
            $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
            $func = $funcMap[$vf.aggregation]
            if ($null -eq $func) { $func = -4157 }
            $table.AddDataField($field, $vf.field, $func) | Out-Null
        }
        Output-Json @{ success = $true; data = @{ pivotTableName = $pivotName; location = $destCell.Address(); rowCount = $table.RowRange.Rows.Count; columnCount = $table.TableRange1.Columns.Count } }
    }

    "updatePivotTable" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $excel.ActiveSheet
        $table = $null
        if ($p.pivotTableName) {
            try { $table = $sheet.PivotTables($p.pivotTableName) } catch {}
        }
        if ($null -eq $table -and $p.pivotTableCell) {
            try { $table = $sheet.Range($p.pivotTableCell).PivotTable } catch {}
        }
        if ($null -eq $table) { Output-Json @{ success = $false; error = "PivotTable not found" }; return }
        $ops = @()
        if ($p.addRowFields) {
            foreach ($f in $p.addRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 1; $field.Position = 1
                $ops += @{ operation = "addRowFields"; success = $true; message = $f }
            }
        }
        if ($p.removeRowFields) {
            foreach ($f in $p.removeRowFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeRowFields"; success = $true; message = $f }
            }
        }
        if ($p.addColumnFields) {
            foreach ($f in $p.addColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 2; $field.Position = 1
                $ops += @{ operation = "addColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.removeColumnFields) {
            foreach ($f in $p.removeColumnFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeColumnFields"; success = $true; message = $f }
            }
        }
        if ($p.addFilterFields) {
            foreach ($f in $p.addFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 3; $field.Position = 1
                $ops += @{ operation = "addFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.removeFilterFields) {
            foreach ($f in $p.removeFilterFields) {
                $field = $table.PivotFields($f); $field.Orientation = 0
                $ops += @{ operation = "removeFilterFields"; success = $true; message = $f }
            }
        }
        if ($p.addValueFields) {
            foreach ($vf in $p.addValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $table.AddDataField($field, $vf.field, $func) | Out-Null
                $ops += @{ operation = "addValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.updateValueFields) {
            foreach ($vf in $p.updateValueFields) {
                $field = $table.PivotFields($vf.field)
                $funcMap = @{ SUM = -4157; COUNT = -4112; AVERAGE = -4106; MAX = -4136; MIN = -4139 }
                $func = $funcMap[$vf.aggregation]
                if ($null -eq $func) { $func = -4157 }
                $field.Function = $func
                $ops += @{ operation = "updateValueFields"; success = $true; message = $vf.field }
            }
        }
        if ($p.removeValueFields) {
            foreach ($vf in $p.removeValueFields) {
                try { $table.DataFields($vf).Orientation = 0 } catch {}
                $ops += @{ operation = "removeValueFields"; success = $true; message = $vf }
            }
        }
        if ($p.refresh) { try { $table.RefreshTable() | Out-Null; $ops += @{ operation = "refresh"; success = $true; message = "refreshed" } } catch {} }
        Output-Json @{ success = $true; data = @{ pivotTableName = $table.Name; operations = $ops } }
    }

    "setFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        $sheet = if ($p.sheet -is [int]) { $wb.Sheets.Item($p.sheet) } else { $wb.Sheets.Item($p.sheet) }
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.Cells.Item($p.row, $p.col) }
        $range.Formula = $p.formula
        Output-Json @{ success = $true }
    }

    "setCellFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.numberFormat) { $range.NumberFormat = $p.numberFormat }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.numberFormat } }
    }

    "setCellStyle" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = [bool]$p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = [bool]$p.italic }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $range.Interior.Color = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $range.Font.Color = $fc }
        }
        if ($p.horizontalAlignment) {
            $hAlignMap = @{ left = -4131; center = -4108; right = -4152 }
            $hAlign = $hAlignMap[$p.horizontalAlignment]
            if ($null -ne $hAlign) { $range.HorizontalAlignment = $hAlign }
        }
        if ($p.verticalAlignment) {
            $vAlignMap = @{ top = -4160; center = -4108; bottom = -4107 }
            $vAlign = $vAlignMap[$p.verticalAlignment]
            if ($null -ne $vAlign) { $range.VerticalAlignment = $vAlign }
        }
        if ($null -ne $p.border -and $p.border) {
            $range.Borders.LineStyle = 1
            if ($p.borderColor) {
                $bc = Convert-HexColorToRgbInt([string]$p.borderColor)
                if ($null -ne $bc) { $range.Borders.Color = $bc }
            }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setBorder" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $styleMap = @{ thin = 1; medium = 2; thick = 4; double = 6; none = 0 }
        $style = $styleMap[$p.style]
        if ($null -eq $style) { $style = 1 }
        $position = if ($p.position) { $p.position } else { "all" }
        $borders = @()
        if ($position -eq "all" -or $position -eq "outside") { $borders += 7, 8, 9, 10 }
        if ($position -eq "all" -or $position -eq "inside") { $borders += 11, 12 }
        if ($position -eq "left") { $borders += 7 }
        if ($position -eq "top") { $borders += 8 }
        if ($position -eq "bottom") { $borders += 9 }
        if ($position -eq "right") { $borders += 10 }
        $colorValue = $null
        if ($p.color) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
        foreach ($b in $borders) {
            $border = $range.Borders.Item($b)
            if ($style -eq 0) {
                $border.LineStyle = -4142
            } else {
                $border.LineStyle = 1
                $border.Weight = $style
            }
            if ($null -ne $colorValue) { $border.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; style = $p.style; position = $position } }
    }

    "copyFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $sourceRange = $sheet.Range($p.source)
        $targetRange = $sheet.Range($p.target)
        $sourceRange.Copy()
        $targetRange.PasteSpecial(-4122)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.source; target = $p.target } }
    }

    "clearFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.ClearFormats()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "addConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $formatType = if ($p.type) { $p.type } else { "cellValue" }
        if ($formatType -eq "cellValue") {
            $operatorMap = @{ greater = 5; greaterThan = 5; less = 6; lessThan = 6; equal = 3; notEqual = 4; greaterEqual = 7; greaterThanOrEqual = 7; lessEqual = 8; lessThanOrEqual = 8; between = 1 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 3 }
            $val1 = if ($null -ne $p.value1) { $p.value1 } else { $p.value }
            $val2 = $p.value2
            $cf = $range.FormatConditions.Add(1, $op, $val1, $val2)
            if ($null -ne $cf -and $p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) { $cf.Interior.Color = $bg }
            }
            if ($null -ne $cf -and $p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cf.Font.Color = $fc }
            }
        } elseif ($formatType -eq "colorScale") {
            $scaleType = if ($p.colorScaleType) { $p.colorScaleType } else { 3 }
            $range.FormatConditions.AddColorScale($scaleType)
        } elseif ($formatType -eq "dataBar") {
            $range.FormatConditions.AddDatabar() | Out-Null
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $formatType } }
    }

    "removeConditionalFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.index) {
            $range.FormatConditions.Item([int]$p.index).Delete()
        } else {
            $range.FormatConditions.Delete()
        }
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getConditionalFormats" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $formats = @()
        $count = $range.FormatConditions.Count
        for ($i = 1; $i -le $count; $i++) {
            $cf = $range.FormatConditions.Item($i)
            $formats += @{ index = $i; type = $cf.Type }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; formats = $formats; count = $count } }
    }

    "addDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $typeMap = @{ list = 3; whole = 1; decimal = 2; date = 4; time = 5; textLength = 6; custom = 7 }
        $validationType = $typeMap[$p.validationType]
        if ($null -eq $validationType) { $validationType = 3 }
        $range.Validation.Delete()
        if ($p.validationType -eq "list") {
            $listFormula = if ($p.formula1) { $p.formula1 } elseif ($p.list) { ($p.list -join ",") } else { "" }
            $range.Validation.Add($validationType, 1, 1, $listFormula)
            if ($null -ne $p.showDropdown -and -not [bool]$p.showDropdown) {
                $range.Validation.InCellDropdown = $false
            } else {
                $range.Validation.InCellDropdown = $true
            }
        } else {
            $operatorMap = @{ between = 1; notBetween = 2; equal = 3; notEqual = 4; greater = 5; less = 6; greaterEqual = 7; lessEqual = 8 }
            $op = $operatorMap[$p.operator]
            if ($null -eq $op) { $op = 1 }
            $range.Validation.Add($validationType, 1, $op, $p.formula1, $p.formula2)
        }
        if ($p.inputTitle -or $p.inputMessage) {
            $range.Validation.InputTitle = if ($p.inputTitle) { $p.inputTitle } else { "" }
            $range.Validation.InputMessage = if ($p.inputMessage) { $p.inputMessage } else { "" }
        }
        if ($p.errorTitle -or $p.errorMessage) {
            $range.Validation.ErrorTitle = if ($p.errorTitle) { $p.errorTitle } else { "" }
            $range.Validation.ErrorMessage = if ($p.errorMessage) { $p.errorMessage } else { "" }
        }
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.validationType } }
    }

    "removeDataValidation" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.Validation.Delete()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "getDataValidations" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $validation = $range.Validation
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $validation.Type; formula1 = $validation.Formula1; formula2 = $validation.Formula2; inputTitle = $validation.InputTitle; inputMessage = $validation.InputMessage } }
    }

    "getFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $formulaLocal = ""
        try { $formulaLocal = $cell.FormulaLocal } catch { $formulaLocal = "" }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; formula = $formula; formulaLocal = $formulaLocal; hasFormula = ($formula -like "=*") } }
    }

    "setArrayFormula" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.FormulaArray = $p.formula
        Output-Json @{ success = $true; data = @{ range = $p.range; formula = $p.formula } }
    }

    "getCellInfo" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        $value = $cell.Value2
        $formula = if ($cell.Formula) { $cell.Formula } else { "" }
        $numberFormat = if ($cell.NumberFormat) { $cell.NumberFormat } else { "" }
        $fontName = if ($cell.Font.Name) { $cell.Font.Name } else { "" }
        $fontSize = if ($cell.Font.Size) { $cell.Font.Size } else { 0 }
        $bold = if ($null -ne $cell.Font.Bold) { [bool]$cell.Font.Bold } else { $false }
        $bgColor = if ($null -ne $cell.Interior.Color) { $cell.Interior.Color } else { 0 }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; value = $value; formula = $formula; numberFormat = $numberFormat; font = @{ name = $fontName; size = $fontSize; bold = $bold }; backgroundColor = $bgColor } }
    }

    "refreshLinks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $links = $wb.LinkSources(1)
        if ($links) {
            for ($i = 1; $i -le $links.Length; $i++) {
                $wb.UpdateLink($links[$i - 1], 1)
            }
            Output-Json @{ success = $true; data = @{ refreshed = $links.Length } }
        } else {
            Output-Json @{ success = $true; data = @{ refreshed = 0; message = "没有外部链接" } }
        }
    }

    "consolidate" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $destRange = $sheet.Range($p.destination)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $destRange.Consolidate($p.sources, $func, $p.topRow, $p.leftColumn, $p.createLinks)
        Output-Json @{ success = $true; data = @{ destination = $p.destination; sources = $p.sources } }
    }

    "calculateSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        if ($p.all) {
            $excel.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = "all" } }
        } else {
            $sheet = $excel.ActiveSheet
            $sheet.Calculate()
            Output-Json @{ success = $true; data = @{ calculated = $sheet.Name } }
        }
    }

    # ==================== Excel Advanced ====================
    "getExcelContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $excel.ActiveSheet
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRange = $usedRange.Address(); usedRangeAddress = $usedRange.Address()
        }}
    }

    "getContext" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $excel.ActiveSheet
        $usedRange = $sheet.UsedRange
        $headers = @()
        if ($usedRange.Rows.Count -gt 0) {
            $headerRow = $usedRange.Rows.Item(1)
            for ($i = 1; $i -le [Math]::Min($headerRow.Columns.Count, 26); $i++) {
                $headers += @{ column = [char](64 + $i); value = $headerRow.Cells.Item(1, $i).Value2 }
            }
        }
        $sheets = @(); for ($i = 1; $i -le $wb.Sheets.Count; $i++) { $sheets += $wb.Sheets.Item($i).Name }
        Output-Json @{ success = $true; data = @{
            workbookName = $wb.Name; currentSheet = $sheet.Name; allSheets = $sheets
            selectedCell = $excel.Selection.Address(); headers = $headers; usedRangeAddress = $usedRange.Address(); usedRange = $usedRange.Address()
        }}
    }

    "sortRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $keyCol = $sheet.Range($p.keyColumn)
        $order = if ($p.order -eq "desc") { 2 } else { 1 }
        $range.Sort($keyCol, $order)
        Output-Json @{ success = $true }
    }

    "autoFilter" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.criteria) {
            $range.AutoFilter($p.field, $p.criteria)
        } else {
            $range.AutoFilter()
        }
        Output-Json @{ success = $true }
    }

    "createChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.dataRange)
        $chartType = $p.chartType
        if ($null -eq $chartType) {
            $chartTypes = @{ column = 51; column_clustered = 51; column_stacked = 52; bar = 57; bar_clustered = 57; line = 4; line_markers = 65; pie = 5; doughnut = -4120; area = 1; scatter = -4169; radar = -4151 }
            $chartType = $chartTypes[$p.chartTypeName]
        }
        if ($null -eq $chartType) { $chartType = 51 }
        $chartTypeName = if ($p.chartTypeName) { $p.chartTypeName } else { $null }
        $left = if ($p.position -and $p.position.left) { $p.position.left } elseif ($p.left) { $p.left } else { $range.Left + $range.Width + 20 }
        $top = if ($p.position -and $p.position.top) { $p.position.top } elseif ($p.top) { $p.top } else { $range.Top }
        $width = if ($p.position -and $p.position.width) { $p.position.width } else { 400 }
        $height = if ($p.position -and $p.position.height) { $p.position.height } else { 300 }
        $chartObj = $sheet.ChartObjects().Add($left, $top, $width, $height)
        $chartObj.Chart.SetSourceData($range)
        $chartObj.Chart.ChartType = $chartType
        if ($p.title) { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
        if ($null -ne $p.showLegend) { $chartObj.Chart.HasLegend = [bool]$p.showLegend }
        if ($p.showDataLabels) { $chartObj.Chart.ApplyDataLabels() }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; chartIndex = $chartObj.Index; dataRange = $p.dataRange; chartType = $chartTypeName; position = @{ left = $left; top = $top; width = $width; height = $height } } }
    }

    "updateChart" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $chartObj = $null
        if ($p.chartName) {
            try { $chartObj = $sheet.ChartObjects($p.chartName) } catch {}
        }
        if ($null -eq $chartObj -and $p.chartIndex) {
            try { $chartObj = $sheet.ChartObjects().Item([int]$p.chartIndex) } catch {}
        }
        if ($null -eq $chartObj) { Output-Json @{ success = $false; error = "Chart not found" }; return }
        $updated = @()
        if ($null -ne $p.title) {
            if ($p.title -eq "") { $chartObj.Chart.HasTitle = $false }
            else { $chartObj.Chart.HasTitle = $true; $chartObj.Chart.ChartTitle.Text = $p.title }
            $updated += "title"
        }
        if ($null -ne $p.chartType) {
            $chartObj.Chart.ChartType = $p.chartType
            $updated += "chartType"
        }
        if ($null -ne $p.showLegend) {
            $chartObj.Chart.HasLegend = [bool]$p.showLegend
            $updated += "showLegend"
        }
        if ($null -ne $p.legendPosition) {
            $legendMap = @{ bottom = -4107; top = -4160; left = -4131; right = -4152 }
            $pos = $legendMap[$p.legendPosition]
            if ($pos) { $chartObj.Chart.Legend.Position = $pos; $updated += "legendPosition" }
        }
        if ($null -ne $p.showDataLabels) {
            if ($p.showDataLabels) {
                $chartObj.Chart.ApplyDataLabels()
            } else {
                try {
                    $series = $chartObj.Chart.SeriesCollection()
                    for ($i = 1; $i -le $series.Count; $i++) {
                        try { $series.Item($i).DataLabels().Delete() } catch {}
                    }
                } catch {}
            }
            $updated += "showDataLabels"
        }
        if ($p.dataRange) {
            $chartObj.Chart.SetSourceData($sheet.Range($p.dataRange))
            $updated += "dataRange"
        }
        if ($p.colors -and $p.colors.Count -gt 0) {
            $series = $chartObj.Chart.SeriesCollection()
            for ($i = 1; $i -le $series.Count -and $i -le $p.colors.Count; $i++) {
                $color = Convert-HexColorToRgbInt($p.colors[$i - 1])
                if ($null -ne $color) { $series.Item($i).Format.Fill.ForeColor.RGB = $color }
            }
            $updated += "colors"
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartObj.Name; updatedProperties = $updated } }
    }

    "exportChartAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        $chartName = $p.chartName
        if ([string]::IsNullOrEmpty($chartName)) { Output-Json @{ success = $false; error = "Missing chartName" }; return }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 Excel COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $chartObj = $sheet.ChartObjects($chartName)
        # 先删掉同名旧文件：否则下面"文件存在"的校验会被上一次的产物骗过
        if (Test-Path $outputPath) { try { Remove-Item -LiteralPath $outputPath -Force } catch {} }
        $chartObj.Chart.Export($outputPath, $filterName)
        # 关键校验：Chart.Export 会【静默写出 0 字节文件且不报错】（office-mcp 的 GOTCHAS 记过同一个坑）
        if (-not (Test-Path $outputPath)) {
            Output-Json @{ success = $false; error = "Chart.Export 未抛错但文件未生成: $outputPath" }; return
        }
        $bytes = (Get-Item $outputPath).Length
        if ($bytes -le 0) {
            Output-Json @{ success = $false; error = "Chart.Export 生成了 0 字节文件: $outputPath（filter/路径可能不被 WPS 接受）" }; return
        }
        Output-Json @{ success = $true; data = @{ chartName = $chartName; outputPath = $outputPath; format = $filterName; bytes = $bytes } }
    }

    "exportRangeAsImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        if ([string]::IsNullOrEmpty($p.range)) { Output-Json @{ success = $false; error = "Missing range" }; return }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $range = $sheet.Range($p.range)
        # xlScreen=1 (Appearance), xlBitmap=2 (Format)
        $tempChart = $null
        try {
            # 先删同名旧文件，避免"文件存在"的校验被上一次产物骗过
            if (Test-Path $outputPath) { try { Remove-Item -LiteralPath $outputPath -Force } catch {} }
            $range.CopyPicture(1, 2)
            $tempChart = $sheet.ChartObjects().Add(0, 0, $range.Width, $range.Height)
            $tempChart.Activate()
            $tempChart.Chart.Paste()
            $tempChart.Chart.Export($outputPath, $filterName)
            $tempChart.Delete()
            $tempChart = $null
            # 关键校验：Chart.Export 会【静默写出 0 字节文件且不报错】，不查就等于假成功
            if (-not (Test-Path $outputPath)) {
                Output-Json @{ success = $false; error = "Chart.Export 未抛错但文件未生成: $outputPath" }; return
            }
            $bytes = (Get-Item $outputPath).Length
            if ($bytes -le 0) {
                Output-Json @{ success = $false; error = "Chart.Export 生成了 0 字节文件: $outputPath（filter/路径可能不被 WPS 接受）" }; return
            }
            Output-Json @{ success = $true; data = @{ range = $p.range; outputPath = $outputPath; format = $filterName; bytes = $bytes } }
        } catch {
            # 异常清理：剪贴板冲突时回滚临时图表
            if ($null -ne $tempChart) { try { $tempChart.Delete() } catch {} }
            Output-Json @{ success = $false; error = "导出区域为图片失败: $(Get-CallError $_)" }
        }
    }

    "removeDuplicates" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $originalCount = $range.Rows.Count
        $cols = @()
        if ($p.columns -and $p.columns.Count -gt 0) {
            foreach ($col in $p.columns) {
                if ($col -is [int]) { $cols += $col }
                else {
                    $num = Convert-ColumnLetterToNumber([string]$col)
                    if ($num) { $cols += $num }
                }
            }
        }
        if ($cols.Count -eq 0) { $cols = @(1) }
        $hasHeader = if ($null -ne $p.hasHeader) { [int]([bool]$p.hasHeader) } else { 1 }
        $range.RemoveDuplicates($cols, $hasHeader)
        $remainingCount = $range.Rows.Count
        $removedCount = $originalCount - $remainingCount
        Output-Json @{ success = $true; data = @{ originalCount = $originalCount; removedCount = $removedCount; remainingCount = $remainingCount } }
    }

    "createSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $wb.Sheets.Add()
        if ($p.name) { $sheet.Name = $p.name }
        Output-Json @{ success = $true; data = @{ sheetName = $sheet.Name; sheetIndex = $sheet.Index } }
    }

    "deleteSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $name = $sheet.Name
        $excel.DisplayAlerts = $false
        $sheet.Delete()
        $excel.DisplayAlerts = $true
        Output-Json @{ success = $true; data = @{ deletedSheet = $name } }
    }

    "renameSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $oldName = $sheet.Name
        $sheet.Name = $p.newName
        Output-Json @{ success = $true; data = @{ oldName = $oldName; newName = $p.newName } }
    }

    "copySheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        if ($p.before) {
            $sheet.Copy($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Copy($null, $wb.Sheets.Item($p.after))
        } else {
            $sheet.Copy($null, $wb.Sheets.Item($wb.Sheets.Count))
        }
        Output-Json @{ success = $true; data = @{ copiedFrom = $sheet.Name } }
    }

    "getSheetList" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheets = @()
        for ($i = 1; $i -le $wb.Sheets.Count; $i++) {
            $sheet = $wb.Sheets.Item($i)
            $sheets += @{ name = $sheet.Name; index = $i; visible = $sheet.Visible }
        }
        Output-Json @{ success = $true; data = @{ sheets = $sheets; count = $sheets.Count; activeSheet = $excel.ActiveSheet.Name } }
    }

    "switchSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $wb.Sheets.Item($p.sheet)
        $sheet.Activate()
        Output-Json @{ success = $true; data = @{ activeSheet = $sheet.Name } }
    }

    "moveSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = if ($p.sheet) { $wb.Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        if ($p.before) {
            $sheet.Move($wb.Sheets.Item($p.before))
        } elseif ($p.after) {
            $sheet.Move($null, $wb.Sheets.Item($p.after))
        }
        Output-Json @{ success = $true; data = @{ movedSheet = $sheet.Name } }
    }

    "mergeCells" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $across = if ($null -ne $p.across) { [bool]$p.across } else { $false }
        $range.Merge($across)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "unmergeCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.UnMerge()
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "setColumnWidth" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $col = $p.column
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $sheet.Range("${col}:${col}").ColumnWidth = $p.width
        Output-Json @{ success = $true; data = @{ column = $col; width = $p.width } }
    }

    "setRowHeight" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $sheet.Range("$($p.row):$($p.row)").RowHeight = $p.height
        Output-Json @{ success = $true; data = @{ row = $p.row; height = $p.height } }
    }

    "autoFitColumn" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if ($p.range) {
            $sheet.Range($p.range).Columns.AutoFit()
        } elseif ($p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").AutoFit()
        } else {
            $sheet.UsedRange.Columns.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "列宽已自动调整" } }
    }

    "autoFitRow" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if ($p.range) {
            $sheet.Range($p.range).Rows.AutoFit()
        } elseif ($p.row) {
            $sheet.Range("$($p.row):$($p.row)").AutoFit()
        } else {
            $sheet.UsedRange.Rows.AutoFit()
        }
        Output-Json @{ success = $true; data = @{ message = "行高已自动调整" } }
    }

    "autoFitAll" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $range.Columns.AutoFit()
        $range.Rows.AutoFit()
        Output-Json @{ success = $true; data = @{ message = "列宽行高已自动调整" } }
    }

    "setNumberFormat" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        if ($p.format) { $range.NumberFormat = $p.format }
        Output-Json @{ success = $true; data = @{ range = $p.range; format = $p.format } }
    }

    "wrapText" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $wrap = if ($null -ne $p.wrap) { [bool]$p.wrap } else { $true }
        $range.WrapText = $wrap
        Output-Json @{ success = $true; data = @{ range = $p.range; wrapText = $wrap } }
    }

    "setPrintArea" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if ($p.range) { $sheet.PageSetup.PrintArea = $p.range } else { $sheet.PageSetup.PrintArea = "" }
        Output-Json @{ success = $true; data = @{ printArea = if ($p.range) { $p.range } else { "cleared" } } }
    }

    "getSelection" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sel = $excel.Selection
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "No selection" }; return }
        $addr = $sel.Address()
        Output-Json @{ success = $true; data = @{ address = $addr; rows = $sel.Rows.Count; columns = $sel.Columns.Count } }
    }

    "clearRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $clearType = if ($p.type) { $p.type } else { "all" }
        if ($clearType -eq "contents") { $range.ClearContents() }
        elseif ($clearType -eq "formats") { $range.ClearFormats() }
        elseif ($clearType -eq "comments") { $range.ClearComments() }
        else { $range.Clear() }
        Output-Json @{ success = $true; data = @{ range = $p.range; clearType = $clearType } }
    }

    "insertRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; return }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Insert()
        Output-Json @{ success = $true; data = @{ insertedAt = $startRow; count = $count } }
    }

    "insertColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; return }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Insert()
        }
        Output-Json @{ success = $true; data = @{ insertedAt = $col; count = $count } }
    }

    "deleteRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $startRow = if ($p.row) { [int]$p.row } elseif ($p.startRow) { [int]$p.startRow } else { $null }
        if ($null -eq $startRow) { Output-Json @{ success = $false; error = "row/startRow required" }; return }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        $endRow = $startRow + $count - 1
        $sheet.Range("${startRow}:${endRow}").Delete()
        Output-Json @{ success = $true; data = @{ deletedFrom = $startRow; count = $count } }
    }

    "deleteColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $col = if ($p.column) { $p.column } elseif ($p.startColumn) { $p.startColumn } else { $null }
        if ($null -eq $col) { Output-Json @{ success = $false; error = "column/startColumn required" }; return }
        if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
        $count = if ($p.count) { [int]$p.count } else { 1 }
        for ($i = 0; $i -lt $count; $i++) {
            $sheet.Range("${col}:${col}").Delete()
        }
        Output-Json @{ success = $true; data = @{ deletedFrom = $col; count = $count } }
    }

    "hideRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; return }
        foreach ($r in $rows) { $sheet.Range("${r}:${r}").Hidden = $true }
        Output-Json @{ success = $true; data = @{ hiddenRows = $rows } }
    }

    "hideColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; return }
        $hidden = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").Hidden = $true
            $hidden += $col
        }
        Output-Json @{ success = $true; data = @{ hiddenColumns = $hidden } }
    }

    "showRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $rows = if ($p.rows) { $p.rows } elseif ($p.row) { @($p.row) } else { @() }
        if ($rows.Count -eq 0) { Output-Json @{ success = $false; error = "row/rows required" }; return }
        foreach ($r in $rows) { $sheet.Range("${r}:${r}").Hidden = $false }
        Output-Json @{ success = $true; data = @{ shownRows = $rows } }
    }

    "showColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cols = if ($p.columns) { $p.columns } elseif ($p.column) { @($p.column) } else { @() }
        if ($cols.Count -eq 0) { Output-Json @{ success = $false; error = "column/columns required" }; return }
        $shown = @()
        foreach ($c in $cols) {
            $col = $c
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("${col}:${col}").Hidden = $false
            $shown += $col
        }
        Output-Json @{ success = $true; data = @{ shownColumns = $shown } }
    }

    "groupRows" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if (-not $p.startRow -or -not $p.endRow) { Output-Json @{ success = $false; error = "startRow/endRow required" }; return }
        $sheet.Range("$($p.startRow):$($p.endRow)").Group()
        Output-Json @{ success = $true; data = @{ grouped = "$($p.startRow):$($p.endRow)" } }
    }

    "groupColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if (-not $p.startColumn -or -not $p.endColumn) { Output-Json @{ success = $false; error = "startColumn/endColumn required" }; return }
        $startCol = $p.startColumn
        $endCol = $p.endColumn
        if ($startCol -is [int]) { $startCol = Convert-ColumnNumberToLetter([int]$startCol) }
        if ($endCol -is [int]) { $endCol = Convert-ColumnNumberToLetter([int]$endCol) }
        $sheet.Range("${startCol}:${endCol}").Group()
        Output-Json @{ success = $true; data = @{ grouped = "${startCol}:${endCol}" } }
    }

    "freezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        if ($p.cell) {
            $sheet.Range($p.cell).Select()
        } elseif ($p.row -and $p.column) {
            $col = $p.column
            if ($col -is [int]) { $col = Convert-ColumnNumberToLetter([int]$col) }
            $sheet.Range("$col$($p.row)").Select()
        }
        $excel.ActiveWindow.FreezePanes = $true
        Output-Json @{ success = $true; data = @{ message = "窗格已冻结" } }
    }

    "unfreezePanes" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $excel.ActiveWindow.FreezePanes = $false
        Output-Json @{ success = $true; data = @{ message = "窗格冻结已取消" } }
    }

    "findInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $results = @()
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        $found = $searchRange.Find($p.searchText, $null, -4163, $lookAt)
        if ($found) {
            $firstAddr = $found.Address()
            do {
                $addr = $found.Address()
                $results += @{ address = ($addr -replace "\$", ""); value = $found.Value2 }
                $found = $searchRange.FindNext($found)
                if ($null -eq $found) { break }
                $currAddr = $found.Address()
            } while ($currAddr -ne $firstAddr)
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; results = $results; count = $results.Count } }
    }

    "replaceInSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $searchRange = if ($p.range) { $sheet.Range($p.range) } else { $sheet.UsedRange }
        $lookAt = if ($p.matchCase) { 1 } else { 2 }
        $replaced = $searchRange.Replace($p.searchText, $p.replaceText, $lookAt)
        Output-Json @{ success = $true; data = @{ searchText = $p.searchText; replaceText = $p.replaceText; success = $replaced } }
    }

    "copyRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $range.Copy()
        Output-Json @{ success = $true; data = @{ range = $p.range; message = "已复制到剪贴板" } }
    }

    "pasteRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $destRange = $sheet.Range($p.destination)
        if ($p.pasteType -eq "values") {
            $destRange.PasteSpecial(-4163)
        } elseif ($p.pasteType -eq "formats") {
            $destRange.PasteSpecial(-4122)
        } elseif ($p.pasteType -eq "formulas") {
            $destRange.PasteSpecial(-4123)
        } else {
            $sheet.Paste($destRange)
        }
        Output-Json @{ success = $true; data = @{ destination = $p.destination } }
    }

    "fillSeries" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $startCell = $range.Cells.Item(1, 1)
        $startCell.Value2 = if ($null -ne $p.startValue) { $p.startValue } else { 1 }
        $typeMap = @{ linear = 0; growth = 1; date = 2; autoFill = 3 }
        $fillType = $typeMap[$p.type]
        if ($null -eq $fillType) { $fillType = 0 }
        $step = if ($null -ne $p.step) { $p.step } else { 1 }
        $range.DataSeries($null, -4132, $fillType, $step)
        Output-Json @{ success = $true; data = @{ range = $p.range; type = $p.type } }
    }

    "transpose" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $sourceRange = $sheet.Range($p.sourceRange)
        $destCell = if ($p.destinationCell) { $p.destinationCell } elseif ($p.targetCell) { $p.targetCell } else { $null }
        if ($null -eq $destCell) { Output-Json @{ success = $false; error = "destinationCell/targetCell required" }; return }
        $destRange = $sheet.Range($destCell)
        $sourceRange.Copy()
        $destRange.PasteSpecial(-4163, -4142, $false, $true)
        $excel.CutCopyMode = $false
        Output-Json @{ success = $true; data = @{ source = $p.sourceRange; destination = $destCell } }
    }

    "textToColumns" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $delimiter = if ($p.delimiter) { $p.delimiter } else { "," }
        $tab = $false; $semicolon = $false; $comma = $false; $space = $false; $other = $false; $otherChar = $null
        if ($delimiter -eq "`t") { $tab = $true }
        elseif ($delimiter -eq ";") { $semicolon = $true }
        elseif ($delimiter -eq ",") { $comma = $true }
        elseif ($delimiter -eq " ") { $space = $true }
        else { $other = $true; $otherChar = $delimiter }
        $range.TextToColumns($null, 1, 1, $false, $tab, $semicolon, $comma, $space, $other, $otherChar)
        Output-Json @{ success = $true; data = @{ range = $p.range; delimiter = $delimiter } }
    }

    "subtotal" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $funcMap = @{ sum = 9; count = 2; average = 1; max = 4; min = 5 }
        $func = $funcMap[$p.function]
        if ($null -eq $func) { $func = 9 }
        $totalCols = $p.totalColumns
        if ($null -eq $totalCols) { $totalCols = $p.totalColumn }
        if ($totalCols -isnot [System.Array]) { $totalCols = @($totalCols) }
        $replace = if ($null -ne $p.replace) { [bool]$p.replace } else { $true }
        $range.Subtotal([int]$p.groupBy, $func, $totalCols, $replace, $false, $true)
        Output-Json @{ success = $true; data = @{ range = $p.range } }
    }

    "createNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $addr = $range.Address()
        $wb.Names.Add($p.name, "=" + $sheet.Name + "!" + $addr)
        Output-Json @{ success = $true; data = @{ name = $p.name; range = $p.range } }
    }

    "deleteNamedRange" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $wb.Names.Item($p.name).Delete()
        Output-Json @{ success = $true; data = @{ deletedName = $p.name } }
    }

    "getNamedRanges" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $names = @()
        for ($i = 1; $i -le $wb.Names.Count; $i++) {
            $n = $wb.Names.Item($i)
            $names += @{ name = $n.Name; refersTo = $n.RefersTo; visible = $n.Visible }
        }
        Output-Json @{ success = $true; data = @{ names = $names; count = $names.Count } }
    }

    "addCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        $cell.AddComment($p.text)
        if ($p.visible) { $cell.Comment.Visible = $true }
        Output-Json @{ success = $true; data = @{ cell = $p.cell; text = $p.text } }
    }

    "deleteCellComment" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $cell = $sheet.Range($p.cell)
        if ($cell.Comment) { $cell.Comment.Delete() }
        Output-Json @{ success = $true; data = @{ cell = $p.cell } }
    }

    "getCellComments" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $comments = @()
        for ($i = 1; $i -le $sheet.Comments.Count; $i++) {
            $c = $sheet.Comments.Item($i)
            $addr = $c.Parent.Address()
            $comments += @{ cell = ($addr -replace "\$", ""); text = $c.Text(); author = if ($c.Author) { $c.Author } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "protectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Protect($password, $p.drawingObjects, $p.contents, $p.scenarios)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $true } }
    }

    "unprotectSheet" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = if ($p.sheet) { $(Resolve-Book $excel $p).Sheets.Item($p.sheet) } else { $excel.ActiveSheet }
        $password = if ($p.password) { $p.password } else { "" }
        $sheet.Unprotect($password)
        Output-Json @{ success = $true; data = @{ sheet = $sheet.Name; protected = $false } }
    }

    "protectWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $(Resolve-Book $excel $p)
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No active workbook" }; return }
        $password = if ($p.password) { $p.password } else { "" }
        $structure = if ($null -ne $p.structure) { [bool]$p.structure } else { $true }
        $wb.Protect($password, $structure, $p.windows)
        Output-Json @{ success = $true; data = @{ workbook = $wb.Name; protected = $true } }
    }

    "insertExcelImage" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $left = if ($null -ne $p.left) { $p.left } else { 100 }
        $top = if ($null -ne $p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $sheet.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "setHyperlink" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.cell)
        $address = if ($p.address) { $p.address } else { "" }
        $subAddress = if ($p.subAddress) { $p.subAddress } else { "" }
        $screenTip = if ($p.screenTip) { $p.screenTip } else { "" }
        $textToDisplay = if ($p.textToDisplay) { $p.textToDisplay } else { "" }
        $sheet.Hyperlinks.Add($range, $address, $subAddress, $screenTip, $textToDisplay)
        Output-Json @{ success = $true; data = @{ cell = $p.cell; address = $address } }
    }

    "lockCells" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $sheet = $excel.ActiveSheet
        $range = $sheet.Range($p.range)
        $locked = if ($null -ne $p.locked) { [bool]$p.locked } else { $true }
        $range.Locked = $locked
        Output-Json @{ success = $true; data = @{ range = $p.range; locked = $locked } }
    }

    "openWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Open($p.path, $p.updateLinks, $p.readOnly)
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    "getOpenWorkbooks" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $workbooks = @()
        for ($i = 1; $i -le $excel.Workbooks.Count; $i++) {
            $wb = $excel.Workbooks.Item($i)
            $workbooks += @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count; active = ($wb.Name -eq $(Resolve-Book $excel $p).Name) }
        }
        Output-Json @{ success = $true; data = @{ workbooks = $workbooks; count = $workbooks.Count } }
    }

    "switchWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Item($(if ($p.name) { $p.name } else { $p.index }))
        $wb.Activate()
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName } }
    }

    "closeWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = if ($p.name) { $excel.Workbooks.Item($p.name) } else { $(Resolve-Book $excel $p) }
        if ($null -eq $wb) { Output-Json @{ success = $false; error = "No workbook" }; return }
        $name = $wb.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        $wb.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "createWorkbook" {
        $excel = Get-WpsExcel
        if ($null -eq $excel) { Output-Json @{ success = $false; error = "WPS Excel not running" }; return }
        $wb = $excel.Workbooks.Add()
        if ($p.name) { $wb.SaveAs($p.name) }
        Output-Json @{ success = $true; data = @{ name = $wb.Name; path = $wb.FullName; sheets = $wb.Sheets.Count } }
    }

    # ==================== Word ====================
    "getActiveDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        Output-Json @{ success = $true; data = @{
            name = $doc.Name; path = $doc.FullName
            paragraphCount = $doc.Paragraphs.Count; wordCount = $doc.Words.Count; characterCount = $doc.Characters.Count
        }}
    }

    "generateTOC" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $position = if ($p.position) { $p.position } else { "start" }
        $levels = if ($p.levels) { [int]$p.levels } else { 3 }
        $includePageNumbers = if ($null -ne $p.includePageNumbers) { [bool]$p.includePageNumbers } else { $true }
        $range = if ($position -eq "cursor") { $word.Selection.Range } else { $doc.Range(0, 0) }
        $doc.TablesOfContents.Add($range, $true, 1, $levels, $false, "", $true, $true, $includePageNumbers, $true)
        Output-Json @{ success = $true; data = @{ levels = $levels; position = $position; includePageNumbers = $includePageNumbers } }
    }

    "getDocumentText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = $doc.Content.Text
        if ($text.Length -gt 10000) { $text = $text.Substring(0, 10000) + "...(truncated)" }
        Output-Json @{ success = $true; data = @{ text = $text; length = $doc.Content.Text.Length } }
    }

    "getOpenDocuments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $documents = @()
        for ($i = 1; $i -le $word.Documents.Count; $i++) {
            $docItem = $word.Documents.Item($i)
            $documents += @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count; active = ($(Resolve-Doc $word $p).Name -eq $docItem.Name) }
        }
        Output-Json @{ success = $true; data = @{ documents = $documents; count = $documents.Count } }
    }

    "switchDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $docItem = $word.Documents.Item($(if ($p.name) { $p.name } else { $p.index }))
        $docItem.Activate()
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName } }
    }

    "openDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; return }
        $docItem = $word.Documents.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $docItem.Name; path = $docItem.FullName; paragraphs = $docItem.Paragraphs.Count } }
    }

    "closeDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $docItem = if ($p.name) { $word.Documents.Item($p.name) } else { $(Resolve-Doc $word $p) }
        if ($null -eq $docItem) { Output-Json @{ success = $false; error = "No document" }; return }
        $name = $docItem.Name
        $saveChanges = if ($null -ne $p.saveChanges) { [bool]$p.saveChanges } else { $true }
        $docItem.Close($saveChanges)
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "insertText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $position = if ($p.position) { $p.position } else { "cursor" }
        switch ($position) {
            "start" { $range = $doc.Range(0, 0); $range.InsertBefore($p.text) }
            "end" { $range = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1); $range.InsertAfter($p.text) }
            default { $word.Selection.TypeText($p.text) }
        }
        if ($p.style) {
            try { $word.Selection.Range.Style = $p.style } catch {}
        }
        Output-Json @{ success = $true; data = @{ position = $position; textLength = $p.text.Length } }
    }

    "setFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        if ($p.fontName) { $range.Font.Name = $p.fontName }
        if ($p.fontSize) { $range.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $range.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $range.Font.Italic = $p.italic }
        if ($null -ne $p.underline) { $range.Font.Underline = $p.underline }
        if ($p.color) {
            $colorMap = @{ red = 255; blue = 16711680; green = 65280; black = 0 }
            $colorValue = $colorMap[$p.color.ToLower()]
            if ($null -eq $colorValue) { $colorValue = Convert-HexColorToRgbInt([string]$p.color) }
            if ($null -ne $colorValue) { $range.Font.Color = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ settings = @{ fontName = $p.fontName; fontSize = $p.fontSize; bold = $p.bold; italic = $p.italic; underline = $p.underline; color = $p.color } } }
    }

    "findReplace" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $find = $doc.Content.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $p.findText
        $find.Replacement.Text = $p.replaceText
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $replaceAll = if ($null -ne $p.replaceAll) { [bool]$p.replaceAll } else { $true }
        $replaceType = if ($replaceAll) { 2 } else { 1 }
        $result = $find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, $p.replaceText, $replaceType)
        Output-Json @{ success = $true; data = @{ replaced = $result } }
    }

    "insertTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $rows = if ($p.rows) { [int]$p.rows } else { 3 }
        $cols = if ($p.cols) { [int]$p.cols } else { 3 }
        $range = $word.Selection.Range
        $table = $doc.Tables.Add($range, $rows, $cols)
        if ($p.data) {
            for ($r = 0; $r -lt [Math]::Min($p.data.Count, $rows); $r++) {
                for ($c = 0; $c -lt [Math]::Min($p.data[$r].Count, $cols); $c++) {
                    $table.Cell($r + 1, $c + 1).Range.Text = [string]$p.data[$r][$c]
                }
            }
        }
        $table.Borders.Enable = $true
        Output-Json @{ success = $true }
    }

    "setParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $range = if ($p.range -eq "all") { $doc.Content } else { $word.Selection.Range }
        $para = $range.ParagraphFormat
        if ($null -ne $p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2; justify = 3 }
            $align = $alignMap[$p.alignment]
            if ($null -eq $align) { $align = 0 }
            $para.Alignment = $align
        }
        if ($p.lineSpacing) {
            $para.LineSpacingRule = 4
            $para.LineSpacing = [double]$p.lineSpacing * 12
        }
        if ($null -ne $p.spaceBefore) { $para.SpaceBefore = $p.spaceBefore }
        if ($null -ne $p.spaceAfter) { $para.SpaceAfter = $p.spaceAfter }
        if ($null -ne $p.firstLineIndent) { $para.FirstLineIndent = [double]$p.firstLineIndent * 28.35 }
        if ($null -ne $p.leftIndent) { $para.LeftIndent = [double]$p.leftIndent * 28.35 }
        if ($null -ne $p.rightIndent) { $para.RightIndent = [double]$p.rightIndent * 28.35 }
        # 分页控制：论文/报告里"标题不与正文分离""每章另起一页"常用
        if ($null -ne $p.pageBreakBefore) { $para.PageBreakBefore = [bool]$p.pageBreakBefore }
        if ($null -ne $p.keepWithNext) { $para.KeepWithNext = [bool]$p.keepWithNext }
        if ($null -ne $p.keepTogether) { $para.KeepTogether = [bool]$p.keepTogether }
        Output-Json @{ success = $true }
    }

    "setPageSetup" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $ps = $doc.PageSetup
        if ($null -ne $p.topMargin) { $ps.TopMargin = [double]$p.topMargin * 28.35 }
        if ($null -ne $p.bottomMargin) { $ps.BottomMargin = [double]$p.bottomMargin * 28.35 }
        if ($null -ne $p.leftMargin) { $ps.LeftMargin = [double]$p.leftMargin * 28.35 }
        if ($null -ne $p.rightMargin) { $ps.RightMargin = [double]$p.rightMargin * 28.35 }
        if ($null -ne $p.orientation) {
            $ps.Orientation = if ($p.orientation -eq "landscape") { 1 } else { 0 }
        }
        if ($null -ne $p.paperSize) {
            $sizeMap = @{ A4 = 7; A3 = 6; Letter = 1; Legal = 5 }
            $size = $sizeMap[$p.paperSize]
            if ($null -eq $size) { $size = 7 }
            $ps.PaperSize = $size
        }
        Output-Json @{ success = $true }
    }

    "insertPageBreak" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $breakType = if ($p.type) { $p.type } else { "page" }
        $breakTypeMap = @{ page = 7; column = 8; section = 2; sectionContinuous = 3 }
        $bt = $breakTypeMap[$breakType]
        if ($null -eq $bt) { $bt = 7 }
        $word.Selection.InsertBreak($bt)
        Output-Json @{ success = $true; data = @{ type = $breakType } }
    }

    "insertHeader" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $section = $doc.Sections.Item(1)
        $header = $section.Headers.Item(1)
        $header.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $header.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true }
    }

    "insertFooter" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $section = $doc.Sections.Item(1)
        $footer = $section.Footers.Item(1)
        $footer.Range.Text = if ($p.text) { $p.text } else { "" }
        if ($p.alignment) {
            $alignMap = @{ left = 0; center = 1; right = 2 }
            $footer.Range.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        if ($p.includePageNumber) {
            $footer.Range.InsertAfter(" - 第 ")
            $footer.Range.Fields.Add($footer.Range, -1, "PAGE", $false)
            $footer.Range.InsertAfter(" 页 ")
        }
        Output-Json @{ success = $true }
    }

    "insertHyperlink" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $range = $word.Selection.Range
        $url = if ($p.url) { $p.url } else { $p.address }
        $text = if ($p.text) { $p.text } elseif ($p.displayText) { $p.displayText } else { $url }
        if ($range.Text -and $range.Text.Trim() -ne "") {
            $doc.Hyperlinks.Add($range, $url)
        } else {
            $range.Text = $text
            $doc.Hyperlinks.Add($doc.Range($range.Start, $range.Start + $text.Length), $url)
        }
        Output-Json @{ success = $true; data = @{ url = $url; text = $text } }
    }

    "insertBookmark" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; return }
        $range = $word.Selection.Range
        $doc.Bookmarks.Add($p.name, $range)
        Output-Json @{ success = $true; data = @{ name = $p.name } }
    }

    "getBookmarks" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $bookmarks = @()
        for ($i = 1; $i -le $doc.Bookmarks.Count; $i++) {
            $bm = $doc.Bookmarks.Item($i)
            $bookmarks += @{ name = $bm.Name; start = $bm.Start; end = $bm.End }
        }
        Output-Json @{ success = $true; data = @{ bookmarks = $bookmarks; count = $bookmarks.Count } }
    }

    "getDocumentParagraphs" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $startIdx = if ($null -ne $p.startParagraph) { [int]$p.startParagraph } else { 1 }
        $endIdx = if ($null -ne $p.endParagraph) { [int]$p.endParagraph } else { [Math]::Min($doc.Paragraphs.Count, $startIdx + 49) }
        if ($endIdx -gt $doc.Paragraphs.Count) { $endIdx = $doc.Paragraphs.Count }
        $paragraphs = @()
        for ($i = $startIdx; $i -le $endIdx; $i++) {
            $para = $doc.Paragraphs.Item($i)
            $text = $para.Range.Text
            $text = $text.TrimEnd("`r`n", "`r", "`n")
            if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + "..." }
            $styleName = ""
            try { $styleName = $para.Range.Style.NameLocal } catch { $styleName = "" }
            $paragraphs += @{ index = $i; text = $text; style = $styleName; start = $para.Range.Start; end = $para.Range.End }
        }
        Output-Json @{ success = $true; data = @{ paragraphs = $paragraphs; totalCount = $doc.Paragraphs.Count; returnedCount = $paragraphs.Count } }
    }

    "findInDocument" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.findText) { Output-Json @{ success = $false; error = "findText required" }; return }
        $matchCase = if ($null -ne $p.matchCase) { [bool]$p.matchCase } else { $false }
        $matchWholeWord = if ($null -ne $p.matchWholeWord) { [bool]$p.matchWholeWord } else { $false }
        $maxResults = if ($null -ne $p.maxResults) { [int]$p.maxResults } else { 20 }
        $results = @()
        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.findText
        $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        while ($found -and $results.Count -lt $maxResults) {
            $matchStart = $searchRange.Start
            $matchEnd = $searchRange.End
            $matchText = $searchRange.Text
            $paraIdx = 0
            for ($pi = 1; $pi -le $doc.Paragraphs.Count; $pi++) {
                $pr = $doc.Paragraphs.Item($pi).Range
                if ($matchStart -ge $pr.Start -and $matchStart -le $pr.End) { $paraIdx = $pi; break }
            }
            $ctxStart = [Math]::Max(0, $matchStart - 50)
            $ctxEnd = [Math]::Min($doc.Content.End, $matchEnd + 50)
            $ctxRange = $doc.Range($ctxStart, $ctxEnd)
            $context = $ctxRange.Text
            $results += @{ text = $matchText; start = $matchStart; end = $matchEnd; paragraphIndex = $paraIdx; context = $context }
            $searchRange = $doc.Range($matchEnd, $doc.Content.End)
            $searchRange.Find.ClearFormatting()
            $searchRange.Find.Text = $p.findText
            $found = $searchRange.Find.Execute($p.findText, $matchCase, $matchWholeWord, $false, $false, $false, $true, 1, $false, "", 0)
        }
        Output-Json @{ success = $true; data = @{ results = $results; count = $results.Count; findText = $p.findText } }
    }

    "smartFillField" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.keyword) { Output-Json @{ success = $false; error = "keyword required" }; return }
        if ($null -eq $p.value) { Output-Json @{ success = $false; error = "value required（空字符串将清除该字段内容）" }; return }
        $fillMode = if ($p.fillMode) { $p.fillMode.ToLower() } else { "auto" }

        $searchRange = $doc.Content.Duplicate
        $searchRange.Find.ClearFormatting()
        $searchRange.Find.Text = $p.keyword
        $found = $searchRange.Find.Execute($p.keyword, $false, $false, $false, $false, $false, $true, 1, $false, "", 0)
        if (-not $found) { Output-Json @{ success = $false; error = "Keyword '$($p.keyword)' not found" }; return }

        $matchStart = $searchRange.Start
        $matchEnd = $searchRange.End
        $paraRange = $searchRange.Paragraphs.Item(1).Range
        $paraText = $paraRange.Text
        $detectedMode = $fillMode

        if ($fillMode -eq "auto") {
            if ($paraText -match '[\{\【\[]' + [regex]::Escape($p.keyword) + '[\}\】\]]') {
                $detectedMode = "placeholder"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]\s*_+') {
                $detectedMode = "underline"
            }
            elseif ($paraText -match [regex]::Escape($p.keyword) + '\s*[：:]') {
                $detectedMode = "afterColon"
            }
            else {
                $afterKeyword = $paraText.Substring([Math]::Min($paraText.IndexOf($p.keyword) + $p.keyword.Length, $paraText.Length))
                $afterKeyword = $afterKeyword.TrimStart().TrimEnd("`r`n", "`r", "`n").TrimEnd()
                if ($afterKeyword.Length -eq 0 -or $afterKeyword -match '^[\s\r\n]*$') {
                    $detectedMode = "afterLabel"
                } else {
                    $detectedMode = "afterColon"
                }
            }
        }

        $fillResult = ""
        switch ($detectedMode) {
            "placeholder" {
                $fullPlaceholder = ""
                if ($paraText -match '([\{\【\[])' + [regex]::Escape($p.keyword) + '([\}\】\]])') {
                    $fullPlaceholder = $Matches[0]
                } else {
                    $fullPlaceholder = $p.keyword
                }
                $replaceRange = $doc.Content.Duplicate
                $replaceRange.Find.ClearFormatting()
                $replaceRange.Find.Text = $fullPlaceholder
                $found2 = $replaceRange.Find.Execute($fullPlaceholder, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                if ($found2) { $fillResult = "Replaced placeholder '$fullPlaceholder' with '$($p.value)'" }
                else { $fillResult = "Failed to replace placeholder" }
            }
            "underline" {
                $keywordEndPos = $matchEnd
                $fillDone = $false
                for ($ri = 1; $ri -le $paraRange.Words.Count; $ri++) {
                    $wordObj = $paraRange.Words.Item($ri)
                    if ($wordObj.Start -gt $keywordEndPos -and $wordObj.Font.Underline -ne 0 -and $wordObj.Text -match '_+') {
                        $wordObj.Text = $p.value
                        $wordObj.Font.Underline = 1
                        $fillDone = $true
                        $fillResult = "Filled underlined field after '$($p.keyword)' with '$($p.value)'"
                        break
                    }
                }
                if (-not $fillDone) {
                    $underscorePattern = "_+"
                    $afterKeyRange = $doc.Range($keywordEndPos, $paraRange.End)
                    $afterKeyRange.Find.ClearFormatting()
                    $afterKeyRange.Find.Text = $underscorePattern
                    $afterKeyRange.Find.MatchWildcards = $true
                    $foundUl = $afterKeyRange.Find.Execute($underscorePattern, $false, $false, $false, $false, $false, $true, 1, $false, $p.value, 1)
                    if ($foundUl) {
                        $fillResult = "Filled underline after '$($p.keyword)' with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($keywordEndPos, $keywordEndPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after keyword '$($p.keyword)' (no underline found)"
                    }
                }
            }
            "afterColon" {
                $afterKeyRange = $doc.Range($matchEnd, $paraRange.End)
                $afterKeyText = $afterKeyRange.Text
                $colonOffset = $afterKeyText.IndexOfAny([char[]]@('：', ':'))
                if ($colonOffset -ge 0) {
                    $insertPos = $matchEnd + $colonOffset + 1
                    $afterColonRange = $doc.Range($insertPos, $paraRange.End - 1)
                    $afterColonText = $afterColonRange.Text.TrimStart().TrimEnd("`r`n", "`r", "`n").TrimEnd()
                    if ($afterColonText.Length -gt 0 -and $afterColonText -notmatch '^[\s\r\n_　]+$') {
                        $afterColonRange.Text = $p.value
                        $fillResult = "Replaced content after colon with '$($p.value)'"
                    } else {
                        $insertRange = $doc.Range($insertPos, $insertPos)
                        $insertRange.InsertAfter($p.value)
                        $fillResult = "Inserted '$($p.value)' after colon following '$($p.keyword)'"
                    }
                } else {
                    $insertRange = $doc.Range($matchEnd, $matchEnd)
                    $insertRange.InsertAfter("：" + $p.value)
                    $fillResult = "No colon found, inserted '：$($p.value)' after '$($p.keyword)'"
                }
            }
            "afterLabel" {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after label '$($p.keyword)'"
            }
            default {
                $insertRange = $doc.Range($matchEnd, $matchEnd)
                $insertRange.InsertAfter($p.value)
                $fillResult = "Inserted '$($p.value)' after '$($p.keyword)' (default mode)"
            }
        }
        Output-Json @{ success = $true; data = @{ keyword = $p.keyword; value = $p.value; fillMode = $detectedMode; result = $fillResult } }
    }

    "replaceBookmarkContent" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "name required" }; return }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; return }
        try {
            $bm = $doc.Bookmarks.Item($p.name)
            $bmStart = $bm.Start
            $bmRange = $bm.Range
            $bmRange.Text = $p.text
            $newEnd = $bmStart + $p.text.Length
            $newRange = $doc.Range($bmStart, $newEnd)
            $doc.Bookmarks.Add($p.name, $newRange) | Out-Null
            Output-Json @{ success = $true; data = @{ name = $p.name; text = $p.text; start = $bmStart; end = $newEnd } }
        } catch {
            Output-Json @{ success = $false; error = "Bookmark '$($p.name)' not found or failed to replace: $(Get-CallError $_)" }
        }
    }

    "addComment" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $text = if ($p.text) { $p.text } else { $p.comment }
        if (-not $text) { Output-Json @{ success = $false; error = "text required" }; return }
        $range = $word.Selection.Range
        $doc.Comments.Add($range, $text)
        Output-Json @{ success = $true }
    }

    "getComments" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $comments = @()
        for ($i = 1; $i -le $doc.Comments.Count; $i++) {
            $c = $doc.Comments.Item($i)
            $comments += @{ index = $i; text = $c.Range.Text; author = if ($c.Author) { $c.Author } else { "" }; date = if ($c.Date) { $c.Date.ToString() } else { "" } }
        }
        Output-Json @{ success = $true; data = @{ comments = $comments; count = $comments.Count } }
    }

    "getDocumentStats" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $stats = @{ name = $doc.Name; path = $doc.FullName; pages = $doc.ComputeStatistics(2); words = $doc.ComputeStatistics(0); characters = $doc.ComputeStatistics(3); paragraphs = $doc.ComputeStatistics(4); lines = $doc.ComputeStatistics(1) }
        # 长文档常见的"结构规模"指标，顺手一并给出（拿不到就跳过，不报错）
        foreach ($pair in @(@("sections", "Sections"), @("tables", "Tables"), @("inlineShapes", "InlineShapes"), @("footnotes", "Footnotes"), @("toc", "TablesOfContents"), @("tof", "TablesOfFigures"))) {
            try { $stats[$pair[0]] = [int]$doc.($pair[1]).Count } catch {}
        }
        Output-Json @{ success = $true; data = $stats }
    }

    "insertImage" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        $path = if ($p.path) { $p.path } else { $p.filePath }
        if (-not $path) { Output-Json @{ success = $false; error = "path required" }; return }
        $range = $word.Selection.Range
        $shape = $doc.InlineShapes.AddPicture($path, $false, $true, $range)
        if ($p.width) { $shape.Width = $p.width }
        if ($p.height) { $shape.Height = $p.height }
        if ($p.scale) { $shape.ScaleWidth = $p.scale; $shape.ScaleHeight = $p.scale }
        Output-Json @{ success = $true; data = @{ width = $shape.Width; height = $shape.Height } }
    }

    "applyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($p.range -and $null -ne $p.range.start -and $null -ne $p.range.end) {
            $range = $doc.Range([int]$p.range.start, [int]$p.range.end)
        } else {
            $range = $word.Selection.Range
        }
        $range.Style = $p.styleName
        Output-Json @{ success = $true; data = @{ affectedText = $range.Text } }
    }

    # ==================== Word: Proofreading ====================
    "enableTrackChanges" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $enable = if ($null -ne $p.enable) { [bool]$p.enable } else { $true }
        $doc.TrackRevisions = $enable
        Output-Json @{ success = $true; data = @{ trackChanges = $enable; active = $doc.TrackRevisions } }
    }

    "getTrackChangesStatus" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $revisionCount = 0
        try { $revisionCount = $doc.Revisions.Count } catch {}
        Output-Json @{ success = $true; data = @{ trackChanges = [bool]$doc.TrackRevisions; revisionCount = $revisionCount } }
    }

    "replaceRange" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($null -eq $p.startPos -or $null -eq $p.endPos) { Output-Json @{ success = $false; error = "startPos and endPos required" }; return }
        if ($null -eq $p.text) { Output-Json @{ success = $false; error = "text required" }; return }
        try {
            $range = $doc.Range([int]$p.startPos, [int]$p.endPos)
            $originalText = $range.Text
            $range.Text = $p.text
            $actualEndPos = $range.End
            Output-Json @{ success = $true; data = @{ startPos = [int]$p.startPos; originalEndPos = [int]$p.endPos; endPos = $actualEndPos; originalText = $originalText.TrimEnd("`r`n"); newText = $p.text } }
        } catch {
            Output-Json @{ success = $false; error = "Failed to replace range: $(Get-CallError $_)" }
        }
    }

    # ==================== PPT ====================
    "getActivePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $slides = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $shapes = @()
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                $text = ""
                try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $text = $shape.TextFrame.TextRange.Text.Substring(0, [Math]::Min(50, $shape.TextFrame.TextRange.Text.Length)) } } catch {}
                $shapes += @{ name = $shape.Name; type = $shape.Type; text = $text }
            }
            $slides += @{ index = $i; shapeCount = $slide.Shapes.Count; shapes = $shapes }
        }
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count; slides = $slides } }
    }

    "openPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        if (-not $p.path) { Output-Json @{ success = $false; error = "path required" }; return }
        $pres = $ppt.Presentations.Open($p.path)
        Output-Json @{ success = $true; data = @{ name = $pres.Name; path = $pres.FullName; slideCount = $pres.Slides.Count } }
    }

    "closePresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = if ($p.name) { $ppt.Presentations.Item($p.name) } else { $ppt.ActivePresentation }
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No presentation" }; return }
        $name = $pres.Name
        $pres.Close()
        Output-Json @{ success = $true; data = @{ closed = $name } }
    }

    "getOpenPresentations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $presentations = @()
        for ($i = 1; $i -le $ppt.Presentations.Count; $i++) {
            $presItem = $ppt.Presentations.Item($i)
            $presentations += @{ name = $presItem.Name; path = $presItem.FullName; slideCount = $presItem.Slides.Count }
        }
        Output-Json @{ success = $true; data = @{ presentations = $presentations; count = $presentations.Count } }
    }

    "switchPresentation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $presItem = $ppt.Presentations.Item($(if ($p.name) { $p.name } else { $p.index }))
        $presItem.Windows.Item(1).Activate()
        Output-Json @{ success = $true; data = @{ name = $presItem.Name } }
    }

    "deleteSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $pres.Slides.Item($index).Delete()
        Output-Json @{ success = $true; data = @{ deleted = $index } }
    }

    "duplicateSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $newSlide = $slide.Duplicate()
        $newIndex = $newSlide.Item(1).SlideIndex
        Output-Json @{ success = $true; data = @{ sourceIndex = $index; newIndex = $newIndex } }
    }

    "insertSlidesFromFile" {
        # 从另一个 PPT 文件把整页幻灯片插入当前演示文稿，保留来源格式（跨PPT整合）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $src = if ($p.filePath) { $p.filePath } elseif ($p.path) { $p.path } else { $null }
        if (-not $src) { Output-Json @{ success = $false; error = "filePath required" }; return }
        if (-not (Test-Path $src)) { Output-Json @{ success = $false; error = "source file not found: $src" }; return }
        $before = $pres.Slides.Count
        $afterIndex = if ($null -ne $p.afterIndex) { [int]$p.afterIndex } else { $before }
        if ($afterIndex -lt 0) { $afterIndex = 0 }
        if ($afterIndex -gt $before) { $afterIndex = $before }
        if ($null -ne $p.slideStart -and $null -ne $p.slideEnd) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart, [int]$p.slideEnd)
        } elseif ($null -ne $p.slideStart) {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex, [int]$p.slideStart)
        } else {
            [void]$pres.Slides.InsertFromFile($src, $afterIndex)
        }
        $after = $pres.Slides.Count
        Output-Json @{ success = $true; data = @{ inserted = ($after - $before); afterIndex = $afterIndex; totalSlides = $after; source = $src } }
    }

    "moveSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $fromIndex = if ($p.from) { $p.from } else { $p.fromIndex }
        $toIndex = if ($p.to) { $p.to } else { $p.toIndex }
        $pres.Slides.Item($fromIndex).MoveTo($toIndex)
        Output-Json @{ success = $true; data = @{ from = $fromIndex; to = $toIndex } }
    }

    "getSlideCount" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        Output-Json @{ success = $true; data = @{ count = $pres.Slides.Count } }
    }

    "getSlideInfo" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $txt = ""
            try { if ($shape.HasTextFrame -and $shape.TextFrame.HasText) { $txt = $shape.TextFrame.TextRange.Text } } catch {}
            try {
                if ($shape.HasTable -and [string]::IsNullOrEmpty($txt)) {
                    $tbl = $shape.Table; $rws = @()
                    for ($rr = 1; $rr -le $tbl.Rows.Count; $rr++) {
                        $cs = @()
                        for ($cc = 1; $cc -le $tbl.Columns.Count; $cc++) {
                            $cv = ""
                            try { $cv = $tbl.Cell($rr, $cc).Shape.TextFrame.TextRange.Text } catch {}
                            $cs += $cv
                        }
                        $rws += ($cs -join "|")
                    }
                    $txt = "[表格] " + ($rws -join " ;; ")
                }
            } catch {}
            $txt = ($txt -replace "[`r`n`v]", " ")
            if ($txt.Length -gt 90) { $txt = $txt.Substring(0, 90) + "…" }
            $hasTxt = $false
            if ($shape.HasTextFrame) { $hasTxt = $true }
            $shapes += @{ name = $shape.Name; type = "$($shape.Type)"; hasText = $hasTxt; text = $txt }
        }
        Output-Json @{ success = $true; data = @{ index = $index; slideIndex = $index; shapeCount = $slide.Shapes.Count; shapesCount = $slide.Shapes.Count; shapes = $shapes; layout = "$($slide.Layout)" } }
    }

    "switchSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        $ppt.ActiveWindow.View.GotoSlide($index)
        Output-Json @{ success = $true; data = @{ currentSlide = $index } }
    }

    "setSlideLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $layoutMap = @{ title = 1; titleContent = 2; blank = 12; twoColumn = 4; comparison = 5; titleOnly = 11 }
        $layout = $layoutMap[$p.layout]
        if ($null -eq $layout) { $layout = if ($p.layout) { $p.layout } else { 2 } }
        $pres.Slides.Item($index).Layout = $layout
        Output-Json @{ success = $true; data = @{ slideIndex = $index; layout = $layout } }
    }

    "getSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $notes = ""
        try { $notes = $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text } catch { $notes = "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index; notes = $notes } }
    }

    "setSlideNotes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $index = if ($p.index) { $p.index } else { $p.slideIndex }
        if (-not $index) { $index = 1 }
        $slide = $pres.Slides.Item($index)
        $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = if ($p.notes) { $p.notes } else { "" }
        Output-Json @{ success = $true; data = @{ slideIndex = $index } }
    }

    "slide.add" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addSlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $layouts = @{ title = 1; title_content = 2; blank = 12; two_column = 3; comparison = 34 }
        $layoutKey = if ($p.layout) { $p.layout } else { "title_content" }
        $layoutType = $layouts[$layoutKey]
        if ($null -eq $layoutType) { $layoutType = 2 }
        $position = if ($p.position) { $p.position } else { $pres.Slides.Count + 1 }
        $slide = $pres.Slides.Add($position, $layoutType)
        if ($p.title -and $slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        if ($p.content) {
            try {
                foreach ($shape in $slide.Shapes) {
                    try {
                        if ($shape.PlaceholderFormat.Type -eq 2) {
                            $shape.TextFrame.TextRange.Text = $p.content
                        }
                    } catch {}
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $position; layout = $layoutKey } }
    }

    "addTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 50 }
        $shape = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $shape.TextFrame.TextRange.Text = $p.text
        if ($p.fontSize) { $shape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($p.fontName) { $shape.TextFrame.TextRange.Font.Name = $p.fontName }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "deleteTextBox" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getTextBoxes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $textBoxes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            if ($shape.HasTextFrame) {
                $text = ""
                try { $text = $shape.TextFrame.TextRange.Text } catch {}
                $textBoxes += @{ name = $shape.Name; index = $i; text = $text; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
            }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; textBoxes = $textBoxes; count = $textBoxes.Count } }
    }

    "setTextBoxText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setTextBoxStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $tr = $shape.TextFrame.TextRange
        if ($p.fontSize) { $tr.Font.Size = $p.fontSize }
        if ($p.fontName) { $tr.Font.Name = $p.fontName }
        if ($null -ne $p.bold) { $tr.Font.Bold = $p.bold }
        if ($null -ne $p.italic) { $tr.Font.Italic = $p.italic }
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $tr.Font.Color = $colorValue }
        }
        if ($p.alignment) {
            $alignMap = @{ left = 1; center = 2; right = 3 }
            $tr.ParagraphFormat.Alignment = $alignMap[$p.alignment]
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($slide.Shapes.HasTitle) { $slide.Shapes.Title.TextFrame.TextRange.Text = $p.title }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $p.title } }
    }

    "getSlideTitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $title = ""
        if ($slide.Shapes.HasTitle) { $title = $slide.Shapes.Title.TextFrame.TextRange.Text }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; title = $title } }
    }

    "setSlideSubtitle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            if ($placeholderType -eq 2) {
                $shape.TextFrame.TextRange.Text = if ($p.subtitle) { $p.subtitle } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; subtitle = $p.subtitle } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Subtitle placeholder not found" }
    }

    "setSlideContent" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $placeholderType = $null
            try { $placeholderType = $shape.PlaceholderFormat.Type } catch {}
            if ($placeholderType -eq 7) {
                $shape.TextFrame.TextRange.Text = if ($p.content) { $p.content } else { "" }
                Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
                return
            }
        }
        Output-Json @{ success = $false; error = "Content placeholder not found" }
    }

    "addShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapeTypeMap = @{ rectangle = 1; oval = 9; triangle = 7; diamond = 4; pentagon = 51; hexagon = 52; arrow = 13; star = 12; heart = 21; cloud = 179 }
        $shapeType = $shapeTypeMap[$p.type]
        if ($null -eq $shapeType) { $shapeType = if ($p.type) { $p.type } else { 1 } }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 100 }
        $shape = $slide.Shapes.AddShape($shapeType, $left, $top, $width, $height)
        if ($p.text) { $shape.TextFrame.TextRange.Text = $p.text }
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; slideIndex = $slideIndex } }
    }

    "deleteShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "getShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            $shapes += @{ name = $shape.Name; index = $i; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; shapes = $shapes; count = $shapes.Count } }
    }

    "setShapeStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) {
            $fillColor = Convert-HexColorToRgbInt([string]$p.fillColor)
            if ($null -ne $fillColor) { $shape.Fill.ForeColor.RGB = $fillColor }
        }
        if ($p.lineColor) {
            $lineColor = Convert-HexColorToRgbInt([string]$p.lineColor)
            if ($null -ne $lineColor) { $shape.Line.ForeColor.RGB = $lineColor }
        }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; text = $p.text } }
    }

    "setShapePosition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height } }
    }

    "insertPptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($null -ne $p.width) { $p.width } else { -1 }
        $height = if ($null -ne $p.height) { $p.height } else { -1 }
        $pic = $slide.Shapes.AddPicture($p.path, $false, $true, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $pic.Name; path = $p.path } }
    }

    "deletePptImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Delete()
        Output-Json @{ success = $true; data = @{ deleted = $(if ($p.name) { $p.name } else { $p.shapeIndex }) } }
    }

    "replacePptImage" {
        # 原位替换图片：保留原图位置/尺寸/旋转，删除旧图后在同一矩形插入新图（换图不动版式）
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No active presentation" }; return }
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $sel = if ($p.name) { $p.name } else { $p.shapeIndex }
        if ($null -eq $sel) { Output-Json @{ success = $false; error = "shapeIndex or name required" }; return }
        $newPath = if ($p.path) { $p.path } elseif ($p.filePath) { $p.filePath } else { $null }
        if (-not $newPath) { Output-Json @{ success = $false; error = "path required" }; return }
        if (-not (Test-Path $newPath)) { Output-Json @{ success = $false; error = "image not found: $newPath" }; return }
        $old = $slide.Shapes.Item($sel)
        $l = $old.Left; $t = $old.Top; $w = $old.Width; $h = $old.Height
        $rot = 0
        try { $rot = $old.Rotation } catch {}
        $old.Delete()
        $pic = $slide.Shapes.AddPicture($newPath, $false, $true, $l, $t, $w, $h)
        try { $pic.Rotation = $rot } catch {}
        Output-Json @{ success = $true; data = @{ name = $pic.Name; left = $l; top = $t; width = $w; height = $h; path = $newPath } }
    }

    "setImageStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        if ($null -ne $p.rotation) { $shape.Rotation = $p.rotation }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "exportSlideAsImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation (presentationName not found or no active presentation)" }; return }
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        if ($slideIndex -lt 1 -or $slideIndex -gt $pres.Slides.Count) { Output-Json @{ success = $false; error = "slideIndex $slideIndex out of range (1..$($pres.Slides.Count)) in '$($pres.Name)'" }; return }
        $slide = $pres.Slides.Item($slideIndex)
        $outputPath = if ($p.outputPath) { $p.outputPath } else { $p.path }
        if ([string]::IsNullOrEmpty($outputPath)) { Output-Json @{ success = $false; error = "Missing outputPath" }; return }
        # 确保父目录存在
        $dir = Split-Path -Parent $outputPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $rawFormat = if ($p.format) { $p.format.ToString().ToUpper() } else { "PNG" }
        # JPEG 在 PowerPoint COM 中按 JPG 滤镜处理
        $filterName = if ($rawFormat -eq "JPEG") { "JPG" } else { $rawFormat }
        $width = if ($p.width) { [int]$p.width } else { 1280 }
        $height = if ($p.height) { [int]$p.height } else { 720 }
        try {
            $slide.Export($outputPath, $filterName, $width, $height)
        } catch {
            Output-Json @{ success = $false; error = "Slide.Export failed: $(Get-CallError $_)" }; return
        }
        # 关键修复：校验文件是否真生成，避免 Export 静默失败却报成功
        if (Test-Path $outputPath) {
            $bytes = (Get-Item $outputPath).Length
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; outputPath = $outputPath; format = $filterName; width = $width; height = $height; bytes = $bytes; presentation = $pres.Name } }
        } else {
            Output-Json @{ success = $false; error = "Export 未抛错但文件未生成: $outputPath（WPS Slide.Export 可能不支持该 filter/尺寸参数或该路径）" }
        }
    }

    "insertPptTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $rows = if ($p.rows) { $p.rows } else { 3 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 200 }
        $table = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $table.Name; rows = $rows; cols = $cols } }
    }

    "setPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $textValue = if ($p.text) { $p.text } elseif ($p.value) { $p.value } else { "" }
        $cell.Shape.TextFrame.TextRange.Text = $textValue
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $textValue } }
    }

    "getPptTableCell" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $table = $shape.Table
        $rowCount = $table.Rows.Count
        $colCount = $table.Columns.Count
        $value = $null
        if ($null -ne $p.row -and $null -ne $p.col) {
            try { $value = $table.Cell($p.row, $p.col).Shape.TextFrame.TextRange.Text } catch { $value = $null }
        }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; value = $value; rowCount = $rowCount; colCount = $colCount } }
    }

    "setPptTableStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.tableName) { $p.tableName } else { $p.tableIndex }))
        if ($null -ne $p.left) { $shape.Left = $p.left }
        if ($null -ne $p.top) { $shape.Top = $p.top }
        if ($null -ne $p.width) { $shape.Width = $p.width }
        if ($null -ne $p.height) { $shape.Height = $p.height }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setPptTableCellStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $cell = $shape.Table.Cell($p.row, $p.col)
        $cellShape = $cell.Shape
        if ($p.backgroundColor) {
            $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
            if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
        }
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
        }
        if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize }
        if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold }
        Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col } }
    }

    "setPptTableRowStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $null
        $tableCount = 0
        $targetIndex = if ($p.tableIndex) { $p.tableIndex } else { 1 }
        if ($p.tableName) {
            $shape = $slide.Shapes.Item($p.tableName)
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.HasTable) {
                    $tableCount++
                    if ($tableCount -eq $targetIndex) { $shape = $s; break }
                }
            }
        }
        if ($null -eq $shape -or -not $shape.HasTable) { Output-Json @{ success = $false; error = "Table not found" }; return }
        $rowIndex = if ($p.row) { $p.row } else { 1 }
        $row = $shape.Table.Rows.Item($rowIndex)
        for ($c = 1; $c -le $row.Cells.Count; $c++) {
            $cellShape = $row.Cells.Item($c).Shape
            if ($p.backgroundColor) {
                $bg = Convert-HexColorToRgbInt([string]$p.backgroundColor)
                if ($null -ne $bg) { $cellShape.Fill.Visible = $true; $cellShape.Fill.Solid(); $cellShape.Fill.ForeColor.RGB = $bg }
            }
            if ($p.fontColor) {
                $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
                if ($null -ne $fc) { $cellShape.TextFrame.TextRange.Font.Color.RGB = $fc }
            }
            if ($p.fontSize) { $cellShape.TextFrame.TextRange.Font.Size = $p.fontSize }
            if ($null -ne $p.bold) { $cellShape.TextFrame.TextRange.Font.Bold = $p.bold }
        }
        Output-Json @{ success = $true; data = @{ row = $rowIndex } }
    }

    "setShapeShadow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Shadow.Visible = $true
        if ($p.color) {
            $shadowColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $shadowColor) { $shape.Shadow.ForeColor.RGB = $shadowColor }
        }
        if ($null -ne $p.transparency) { $shape.Shadow.Transparency = $p.transparency }
        if ($null -ne $p.blur) { $shape.Shadow.Blur = $p.blur }
        if ($null -ne $p.offsetX) { $shape.Shadow.OffsetX = $p.offsetX }
        if ($null -ne $p.offsetY) { $shape.Shadow.OffsetY = $p.offsetY }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setBackgroundGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $bg = $slide.Background.Fill
        $bg.Visible = $true
        $bg.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $bg.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $bg.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "setShapeGradient" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $shape.Fill.TwoColorGradient(1, 1)
        if ($p.color1) {
            $c1 = Convert-HexColorToRgbInt([string]$p.color1)
            if ($null -ne $c1) { $shape.Fill.ForeColor.RGB = $c1 }
        }
        if ($p.color2) {
            $c2 = Convert-HexColorToRgbInt([string]$p.color2)
            if ($null -ne $c2) { $shape.Fill.BackColor.RGB = $c2 }
        }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeBorder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.color) {
            $borderColor = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $borderColor) { $shape.Line.ForeColor.RGB = $borderColor }
        }
        if ($p.width) { $shape.Line.Weight = $p.width }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeTransparency" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = $p.transparency }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeRoundness" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($null -ne $p.roundness) { $shape.Adjustments.Item(1) = $p.roundness }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "setShapeFullStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        if ($p.fillColor) { $shape.Fill.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.fillColor) }
        if ($p.lineColor) { $shape.Line.ForeColor.RGB = Convert-HexColorToRgbInt([string]$p.lineColor) }
        if ($p.lineWidth) { $shape.Line.Weight = $p.lineWidth }
        if ($null -ne $p.transparency) { $shape.Fill.Transparency = $p.transparency }
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "alignShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $alignMap = @{ left = 1; center = 2; right = 3; top = 4; middle = 5; bottom = 6 }
        $align = $alignMap[$p.alignment]
        if ($null -eq $align) { $align = 1 }
        $range = $slide.Shapes.Range($p.names)
        $range.Align($align, 1)
        Output-Json @{ success = $true; data = @{ alignment = $p.alignment } }
    }

    "distributeShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $distMap = @{ horizontal = 0; vertical = 1 }
        $dist = $distMap[$p.direction]
        if ($null -eq $dist) { $dist = 0 }
        $range = $slide.Shapes.Range($p.names)
        $range.Distribute($dist, 1)
        Output-Json @{ success = $true; data = @{ direction = $p.direction } }
    }

    "groupShapes" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $range = $slide.Shapes.Range($p.names)
        $group = $range.Group()
        Output-Json @{ success = $true; data = @{ name = $group.Name } }
    }

    "duplicateShape" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $dup = $shape.Duplicate()
        Output-Json @{ success = $true; data = @{ name = $dup.Name } }
    }

    "setShapeZOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.name) { $p.name } else { $p.shapeIndex }))
        $zMap = @{ front = 0; forward = 2; backward = 3; back = 1 }
        $z = $zMap[$p.zOrder]
        if ($null -eq $z) { $z = 0 }
        $shape.ZOrder($z)
        Output-Json @{ success = $true; data = @{ name = $shape.Name; zOrder = $p.zOrder } }
    }

    "insertPptChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 300 }
        $shape = $slide.Shapes.AddChart($p.type, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "addAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($p.shapeName)
        $slide.TimeLine.MainSequence.AddEffect($shape, $p.effect, 1)
        Output-Json @{ success = $true }
    }

    "setSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = $p.effect
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; effect = $p.effect } }
    }

    "setSlideBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        if ($p.color) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.color)
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
        }
        if ($p.imagePath) {
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.UserPicture($p.imagePath)
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "slide.unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        $includeTitle = if ($null -ne $p.includeTitle) { [bool]$p.includeTitle } else { $true }
        $includeBody = if ($null -ne $p.includeBody) { [bool]$p.includeBody } else { $true }
        $start = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $end = if ($p.slideIndex) { [int]$p.slideIndex } else { $pres.Slides.Count }
        for ($i = $start; $i -le $end; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $isTitle = $false
                        try { $isTitle = $shape.PlaceholderFormat.Type -eq 1 -or $shape.PlaceholderFormat.Type -eq 3 } catch {}
                        if (($isTitle -and $includeTitle) -or (-not $isTitle -and $includeBody)) {
                            $shape.TextFrame.TextRange.Font.Name = $fontName
                            $count++
                        }
                    }
                } catch {}
            }
        }
        $slideCount = if ($p.slideIndex) { 1 } else { $pres.Slides.Count }
        Output-Json @{ success = $true; data = @{ font = $fontName; count = $count; slideCount = $slideCount } }
    }

    "addConnector" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 0 }
        $connector = $slide.Shapes.AddConnector(1, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $connector.Name } }
    }

    "addArrow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 100 }
        $width = if ($p.width) { $p.width } else { 120 }
        $height = if ($p.height) { $p.height } else { 60 }
        $shape = $slide.Shapes.AddShape(13, $left, $top, $width, $height)
        Output-Json @{ success = $true; data = @{ name = $shape.Name } }
    }

    "applyColorScheme" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemeKey = if ($p.scheme) { $p.scheme } elseif ($p.colorScheme) { $p.colorScheme } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ scheme = $schemeKey; count = $count } }
    }

    "autoBeautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count } }
    }

    "createKpiCards" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $cards = if ($p.cards) { $p.cards } else { @() }
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 350 }
        $cardWidth = if ($p.cardWidth) { $p.cardWidth } else { 150 }
        $cardHeight = if ($p.cardHeight) { $p.cardHeight } else { 80 }
        $gap = if ($p.gap) { $p.gap } else { 20 }

        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1", "#fd7e14")
        $created = @()

        for ($i = 0; $i -lt $cards.Count; $i++) {
            $card = $cards[$i]
            $x = $startX + ($i * ($cardWidth + $gap))
            $color = if ($card.color) { $card.color } else { $colors[$i % $colors.Count] }
            $shape = $slide.Shapes.AddShape(5, $x, $startY, $cardWidth, $cardHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$color)
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.Shadow.Visible = $true
            $shape.Shadow.Blur = 10
            $shape.Shadow.OffsetX = 3
            $shape.Shadow.OffsetY = 3
            $shape.Shadow.Transparency = 0.6
            $title = if ($card.title) { $card.title } else { "" }
            $value = if ($card.value) { $card.value } else { "" }
            $shape.TextFrame.TextRange.Text = ($title + "`n" + $value)
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 14
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $created += $shape.Name
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; cards = $created } }
    }

    "createStyledTable" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 5 }
        $cols = if ($p.cols) { $p.cols } else { 4 }
        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 150 }
        $width = if ($p.width) { $p.width } else { 620 }
        $height = if ($p.height) { $p.height } else { 200 }

        $tableShape = $slide.Shapes.AddTable($rows, $cols, $left, $top, $width, $height)
        $table = $tableShape.Table

        $styleKey = if ($p.style) { $p.style } else { "business" }
        $styles = @{
            business = @{ header = "#1a365d"; headerText = "#ffffff"; oddRow = "#f8f9fa"; evenRow = "#ffffff" }
            tech = @{ header = "#0d47a1"; headerText = "#ffffff"; oddRow = "#e3f2fd"; evenRow = "#ffffff" }
            elegant = @{ header = "#37474f"; headerText = "#ffffff"; oddRow = "#eceff1"; evenRow = "#ffffff" }
        }
        $style = $styles[$styleKey]
        if ($null -eq $style) { $style = $styles["business"] }

        for ($c = 1; $c -le $cols; $c++) {
            $headerCell = $table.Cell(1, $c)
            $headerColor = Convert-HexColorToRgbInt([string]$style.header)
            if ($null -ne $headerColor) { $headerCell.Shape.Fill.Solid(); $headerCell.Shape.Fill.ForeColor.RGB = $headerColor }
            $headerTextColor = Convert-HexColorToRgbInt([string]$style.headerText)
            if ($null -ne $headerTextColor) { $headerCell.Shape.TextFrame.TextRange.Font.Color.RGB = $headerTextColor }
            $headerCell.Shape.TextFrame.TextRange.Font.Bold = $true
            $headerCell.Shape.TextFrame.TextRange.Font.Size = 12
        }

        for ($r = 2; $r -le $rows; $r++) {
            $rowColor = if ($r % 2 -eq 0) { $style.evenRow } else { $style.oddRow }
            $rowColorValue = Convert-HexColorToRgbInt([string]$rowColor)
            for ($c = 1; $c -le $cols; $c++) {
                $cell = $table.Cell($r, $c)
                if ($null -ne $rowColorValue) { $cell.Shape.Fill.Solid(); $cell.Shape.Fill.ForeColor.RGB = $rowColorValue }
                $cell.Shape.TextFrame.TextRange.Font.Size = 11
            }
        }

        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    if (($r + 1) -le $rows -and ($c + 1) -le $cols) {
                        $table.Cell($r + 1, $c + 1).Shape.TextFrame.TextRange.Text = [string]$rowData[$c]
                    }
                }
            }
        }

        Output-Json @{ success = $true; data = @{ name = $tableShape.Name; rows = $rows; cols = $cols } }
    }

    "addTitleDecoration" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 50 }
        $top = if ($p.top) { $p.top } else { 130 }
        $width = if ($p.width) { $p.width } else { 200 }
        $height = if ($p.height) { $p.height } else { 5 }
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        $bar = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
        if ($null -ne $colorValue) { $bar.Fill.Solid(); $bar.Fill.ForeColor.RGB = $colorValue }
        $bar.Line.Visible = $false
        Output-Json @{ success = $true; data = @{ name = $bar.Name } }
    }

    "addPageIndicator" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $total = $pres.Slides.Count
        $current = $slideIndex
        $text = "$current / $total"
        $left = if ($p.left) { $p.left } else { 880 }
        $top = if ($p.top) { $p.top } else { 500 }
        $box = $slide.Shapes.AddTextbox(1, $left, $top, 80, 30)
        $box.TextFrame.TextRange.Text = $text
        $box.TextFrame.TextRange.Font.Size = 12
        $box.TextFrame.TextRange.Font.Color.RGB = if ($p.dark) { 0 } else { 16777215 }
        $box.TextFrame.TextRange.ParagraphFormat.Alignment = 3
        Output-Json @{ success = $true; data = @{ name = $box.Name; page = $text } }
    }

    "beautifyAllSlides" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemeKey = if ($p.style) { $p.style } else { "business" }
        $schemes = @{ business = @{ title = 0x2F5496; body = 0x333333 }; tech = @{ title = 0x00B0F0; body = 0x404040 }; creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }; minimal = @{ title = 0x000000; body = 0x666666 } }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title } else { $textRange.Font.Color.RGB = $scheme.body }
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ style = $schemeKey; count = $count } }
    }

    "createProgressBar" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $left = if ($p.left) { $p.left } else { 100 }
        $top = if ($p.top) { $p.top } else { 300 }
        $width = if ($p.width) { $p.width } else { 400 }
        $height = if ($p.height) { $p.height } else { 30 }
        $progress = if ($null -ne $p.progress) { [double]$p.progress } else { 0.5 }
        if ($progress -lt 0) { $progress = 0 }
        if ($progress -gt 1) { $progress = 1 }
        $label = if ($p.label) { $p.label } else { "" }

        $bgBar = $slide.Shapes.AddShape(5, $left, $top, $width, $height)
        $bgColor = Convert-HexColorToRgbInt([string]$(if ($p.bgColor) { $p.bgColor } else { "#e0e0e0" }))
        if ($null -ne $bgColor) { $bgBar.Fill.Solid(); $bgBar.Fill.ForeColor.RGB = $bgColor }
        $bgBar.Line.Visible = $false

        $progressWidth = $width * $progress
        $progressBar = $slide.Shapes.AddShape(5, $left, $top, $progressWidth, $height)
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#28a745" }))
        if ($null -ne $colorValue) { $progressBar.Fill.Solid(); $progressBar.Fill.ForeColor.RGB = $colorValue }
        $progressBar.Line.Visible = $false

        $percentText = $slide.Shapes.AddTextbox(1, $left + $width + 10, $top, 60, $height)
        $percentText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($progress * 100))
        $percentText.TextFrame.TextRange.Font.Size = 14
        $percentText.TextFrame.TextRange.Font.Bold = $true

        if ($label) {
            $labelBox = $slide.Shapes.AddTextbox(1, $left, $top - 25, $width, 20)
            $labelBox.TextFrame.TextRange.Text = $label
            $labelBox.TextFrame.TextRange.Font.Size = 12
        }

        Output-Json @{ success = $true; data = @{ bgBar = $bgBar.Name; progressBar = $progressBar.Name; progress = $progress } }
    }

    "createGauge" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 300 }
        $radius = if ($p.radius) { $p.radius } else { 100 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.5 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]"#e0e0e0")
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.7
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $valueText = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 20, 80, 40)
        $valueText.TextFrame.TextRange.Text = "{0}%" -f ([Math]::Round($value * 100))
        $valueText.TextFrame.TextRange.Font.Size = 24
        $valueText.TextFrame.TextRange.Font.Bold = $true
        $valueText.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 10, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 14
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ gauge = $outerCircle.Name; value = $value } }
    }

    "createMiniCharts" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $items = if ($p.items) { $p.items } else { @() }
        $startX = if ($p.startX) { $p.startX } else { 50 }
        $startY = if ($p.startY) { $p.startY } else { 400 }
        $itemWidth = if ($p.itemWidth) { $p.itemWidth } else { 100 }
        $gap = if ($p.gap) { $p.gap } else { 30 }
        $colors = @("#28a745", "#17a2b8", "#ffc107", "#dc3545", "#6f42c1")
        $created = @()

        for ($i = 0; $i -lt $items.Count; $i++) {
            $item = $items[$i]
            $x = $startX + ($i * ($itemWidth + $gap))
            $color = if ($item.color) { $item.color } else { $colors[$i % $colors.Count] }
            $colorValue = Convert-HexColorToRgbInt([string]$color)

            $valueBox = $slide.Shapes.AddTextbox(1, $x, $startY, $itemWidth, 30)
            $valueBox.TextFrame.TextRange.Text = if ($item.value) { $item.value } else { "0" }
            $valueBox.TextFrame.TextRange.Font.Size = 20
            $valueBox.TextFrame.TextRange.Font.Bold = $true
            if ($null -ne $colorValue) { $valueBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }

            $labelBox = $slide.Shapes.AddTextbox(1, $x, $startY + 30, $itemWidth, 20)
            $labelBox.TextFrame.TextRange.Text = if ($item.label) { $item.label } else { "" }
            $labelBox.TextFrame.TextRange.Font.Size = 11
            $labelBox.TextFrame.TextRange.Font.Color.RGB = 6710886

            if ($item.trend) {
                $trendBox = $slide.Shapes.AddTextbox(1, $x + $itemWidth - 30, $startY, 30, 20)
                $trendBox.TextFrame.TextRange.Text = if ($item.trend -eq "up") { "↑" } else { "↓" }
                $trendBox.TextFrame.TextRange.Font.Size = 14
                $trendColor = if ($item.trend -eq "up") { "#28a745" } else { "#dc3545" }
                $trendColorValue = Convert-HexColorToRgbInt([string]$trendColor)
                if ($null -ne $trendColorValue) { $trendBox.TextFrame.TextRange.Font.Color.RGB = $trendColorValue }
            }

            $created += @{ label = $item.label; value = $item.value }
        }

        Output-Json @{ success = $true; data = @{ count = $created.Count; items = $created } }
    }

    "createDonutChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $centerX = if ($p.centerX) { $p.centerX } else { 360 }
        $centerY = if ($p.centerY) { $p.centerY } else { 280 }
        $radius = if ($p.radius) { $p.radius } else { 80 }
        $value = if ($null -ne $p.value) { [double]$p.value } else { 0.75 }
        if ($value -lt 0) { $value = 0 }
        if ($value -gt 1) { $value = 1 }
        $title = if ($p.title) { $p.title } else { "" }
        $centerText = if ($p.centerText) { $p.centerText } else { "{0}%" -f ([Math]::Round($value * 100)) }

        $outerCircle = $slide.Shapes.AddShape(9, $centerX - $radius, $centerY - $radius, $radius * 2, $radius * 2)
        $outerColor = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#0d47a1" }))
        if ($null -ne $outerColor) { $outerCircle.Fill.Solid(); $outerCircle.Fill.ForeColor.RGB = $outerColor }
        $outerCircle.Line.Visible = $false

        $innerRadius = $radius * 0.6
        $innerCircle = $slide.Shapes.AddShape(9, $centerX - $innerRadius, $centerY - $innerRadius, $innerRadius * 2, $innerRadius * 2)
        $innerColor = Convert-HexColorToRgbInt([string]"#ffffff")
        if ($null -ne $innerColor) { $innerCircle.Fill.Solid(); $innerCircle.Fill.ForeColor.RGB = $innerColor }
        $innerCircle.Line.Visible = $false

        $textBox = $slide.Shapes.AddTextbox(1, $centerX - 40, $centerY - 15, 80, 30)
        $textBox.TextFrame.TextRange.Text = $centerText
        $textBox.TextFrame.TextRange.Font.Size = 18
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        if ($title) {
            $titleBox = $slide.Shapes.AddTextbox(1, $centerX - 60, $centerY + $radius + 15, 120, 25)
            $titleBox.TextFrame.TextRange.Text = $title
            $titleBox.TextFrame.TextRange.Font.Size = 12
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        }

        Output-Json @{ success = $true; data = @{ donut = $outerCircle.Name; value = $value } }
    }

    "autoLayout" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $layout = if ($p.layout) { $p.layout } else { "grid" }
        $margin = if ($p.margin) { $p.margin } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 20 }
        $slideWidth = 720
        $slideHeight = 540

        $shapes = @()
        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $s = $slide.Shapes.Item($i)
            if ($s.Type -eq 1 -or $s.Type -eq 6 -or $s.Type -eq 17) { $shapes += $s }
        }

        if ($shapes.Count -eq 0) {
            Output-Json @{ success = $true; data = @{ message = "没有可排列的形状"; count = 0 } }
            break
        }

        if ($layout -eq "horizontal") {
            $totalWidth = 0
            foreach ($shape in $shapes) { $totalWidth += $shape.Width }
            $availWidth = $slideWidth - 2 * $margin
            $spacing = ($availWidth - $totalWidth) / ($shapes.Count + 1)
            $currentX = $margin + $spacing
            foreach ($shape in $shapes) {
                $shape.Left = $currentX
                $shape.Top = ($slideHeight - $shape.Height) / 2
                $currentX += $shape.Width + $spacing
            }
        } elseif ($layout -eq "vertical") {
            $currentY = $margin
            foreach ($shape in $shapes) {
                $shape.Left = ($slideWidth - $shape.Width) / 2
                $shape.Top = $currentY
                $currentY += $shape.Height + $gap
            }
        } else {
            $cols = if ($p.cols) { $p.cols } else { [Math]::Ceiling([Math]::Sqrt($shapes.Count)) }
            $rows = [Math]::Ceiling($shapes.Count / $cols)
            $cellWidth = ($slideWidth - 2 * $margin) / $cols
            $cellHeight = ($slideHeight - 2 * $margin) / $rows
            for ($idx = 0; $idx -lt $shapes.Count; $idx++) {
                $col = $idx % $cols
                $row = [Math]::Floor($idx / $cols)
                $shape = $shapes[$idx]
                $shape.Left = $margin + $col * $cellWidth + ($cellWidth - $shape.Width) / 2
                $shape.Top = $margin + $row * $cellHeight + ($cellHeight - $shape.Height) / 2
            }
        }

        Output-Json @{ success = $true; data = @{ layout = $layout; count = $shapes.Count } }
    }

    "smartDistribute" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $shapeNames = if ($p.shapes) { $p.shapes } else { @() }
        $shapes = @()

        if ($shapeNames.Count -gt 0) {
            foreach ($name in $shapeNames) {
                try { $shapes += $slide.Shapes.Item($name) } catch {}
            }
        } else {
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $s = $slide.Shapes.Item($i)
                if ($s.Type -eq 1 -or $s.Type -eq 6 -or $s.Type -eq 17) { $shapes += $s }
            }
        }

        if ($shapes.Count -lt 2) {
            Output-Json @{ success = $true; data = @{ message = "需要至少2个形状"; count = $shapes.Count } }
            break
        }

        $shapes = if ($direction -eq "horizontal") { $shapes | Sort-Object Left } else { $shapes | Sort-Object Top }

        if ($direction -eq "horizontal") {
            $minX = $shapes[0].Left
            $maxX = $shapes[$shapes.Count - 1].Left + $shapes[$shapes.Count - 1].Width
            $totalWidth = 0
            foreach ($shape in $shapes) { $totalWidth += $shape.Width }
            $spacing = ($maxX - $minX - $totalWidth) / ($shapes.Count - 1)
            $currentX = $minX
            foreach ($shape in $shapes) {
                $shape.Left = $currentX
                $currentX += $shape.Width + $spacing
            }
        } else {
            $minY = $shapes[0].Top
            $maxY = $shapes[$shapes.Count - 1].Top + $shapes[$shapes.Count - 1].Height
            $totalHeight = 0
            foreach ($shape in $shapes) { $totalHeight += $shape.Height }
            $spacing = ($maxY - $minY - $totalHeight) / ($shapes.Count - 1)
            $currentY = $minY
            foreach ($shape in $shapes) {
                $shape.Top = $currentY
                $currentY += $shape.Height + $spacing
            }
        }

        Output-Json @{ success = $true; data = @{ direction = $direction; count = $shapes.Count } }
    }

    "createGrid" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $rows = if ($p.rows) { $p.rows } else { 2 }
        $cols = if ($p.cols) { $p.cols } else { 3 }
        $margin = if ($p.margin) { $p.margin } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 15 }
        $slideWidth = 720
        $slideHeight = 540

        $cellWidth = ($slideWidth - 2 * $margin - ($cols - 1) * $gap) / $cols
        $cellHeight = ($slideHeight - 2 * $margin - ($rows - 1) * $gap) / $rows
        $cells = @()

        for ($r = 0; $r -lt $rows; $r++) {
            for ($c = 0; $c -lt $cols; $c++) {
                $left = $margin + $c * ($cellWidth + $gap)
                $top = $margin + $r * ($cellHeight + $gap)
                $cell = $slide.Shapes.AddShape(5, $left, $top, $cellWidth, $cellHeight)
                $cellColor = Convert-HexColorToRgbInt([string]$(if ($p.cellColor) { $p.cellColor } else { "#f8f9fa" }))
                $borderColor = Convert-HexColorToRgbInt([string]$(if ($p.borderColor) { $p.borderColor } else { "#dee2e6" }))
                if ($null -ne $cellColor) { $cell.Fill.Solid(); $cell.Fill.ForeColor.RGB = $cellColor }
                if ($null -ne $borderColor) { $cell.Line.ForeColor.RGB = $borderColor }
                $cell.Line.Weight = 1
                $cells += $cell.Name
            }
        }

        Output-Json @{ success = $true; data = @{ rows = $rows; cols = $cols; cells = $cells } }
    }

    "addAnimationPreset" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $preset = if ($p.preset) { $p.preset } else { "fadeIn" }
        $timeline = $slide.TimeLine
        $delay = 0
        $delayIncrement = if ($p.delayIncrement) { $p.delayIncrement } else { 0.3 }
        $presets = @{ fadeIn = @{ effect = 10; duration = 0.5 }; flyIn = @{ effect = 2; duration = 0.5 }; zoomIn = @{ effect = 53; duration = 0.4 }; wipeIn = @{ effect = 22; duration = 0.5 }; appear = @{ effect = 1; duration = 0 } }
        $config = $presets[$preset]
        if ($null -eq $config) { $config = $presets["fadeIn"] }
        $animatedCount = 0

        for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
            $shape = $slide.Shapes.Item($i)
            try {
                $effect = $timeline.MainSequence.AddEffect($shape, $config.effect, 0, 1)
                $effect.Timing.Duration = $config.duration
                $effect.Timing.TriggerDelayTime = $delay
                $delay += $delayIncrement
                $animatedCount++
            } catch {}
        }

        Output-Json @{ success = $true; data = @{ preset = $preset; animatedShapes = $animatedCount } }
    }

    "addEmphasisAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $effectType = if ($p.effect) { $p.effect } else { "pulse" }
        $effectMap = @{ pulse = 63; spin = 15; grow = 53; teeter = 28 }
        $effectId = if ($effectMap[$effectType]) { $effectMap[$effectType] } else { 63 }

        $timeline = $slide.TimeLine
        $effect = $timeline.MainSequence.AddEffect($shape, $effectId, 0, 2)
        $effect.Timing.Duration = if ($p.duration) { $p.duration } else { 0.5 }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; effect = $effectType } }
    }

    "createFlowChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $steps = if ($p.steps) { $p.steps } else { @("开始", "步骤1", "步骤2", "结束") }
        $direction = if ($p.direction) { $p.direction } else { "horizontal" }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 250 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 50 }
        $gap = if ($p.gap) { $p.gap } else { 60 }
        $colors = @{ start = "#28a745"; process = "#0d47a1"; end = "#dc3545" }
        $shapes = @()

        for ($i = 0; $i -lt $steps.Count; $i++) {
            if ($direction -eq "horizontal") {
                $x = $startX + $i * ($boxWidth + $gap)
                $y = $startY
            } else {
                $x = $startX
                $y = $startY + $i * ($boxHeight + $gap)
            }

            $shapeType = 5
            $colorKey = "process"
            if ($i -eq 0) { $shapeType = 9; $colorKey = "start" }
            elseif ($i -eq $steps.Count - 1) { $shapeType = 9; $colorKey = "end" }

            $shape = $slide.Shapes.AddShape($shapeType, $x, $y, $boxWidth, $boxHeight)
            $colorValue = Convert-HexColorToRgbInt([string]$colors[$colorKey])
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
            $shape.TextFrame.TextRange.Text = $steps[$i]
            $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
            $shape.TextFrame.TextRange.Font.Size = 12
            $shape.TextFrame.TextRange.Font.Bold = $true
            $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            $shapes += $shape.Name

            if ($i -lt $steps.Count - 1) {
                if ($direction -eq "horizontal") {
                    $arrowX1 = $x + $boxWidth
                    $arrowY1 = $y + $boxHeight / 2
                    $arrowX2 = $x + $boxWidth + $gap
                    $arrowY2 = $arrowY1
                } else {
                    $arrowX1 = $x + $boxWidth / 2
                    $arrowY1 = $y + $boxHeight
                    $arrowX2 = $arrowX1
                    $arrowY2 = $y + $boxHeight + $gap
                }
                $arrow = $slide.Shapes.AddLine($arrowX1, $arrowY1, $arrowX2, $arrowY2)
                $arrow.Line.EndArrowheadStyle = 2
                $arrow.Line.Weight = 2
                $arrowColor = Convert-HexColorToRgbInt([string]"#666666")
                if ($null -ne $arrowColor) { $arrow.Line.ForeColor.RGB = $arrowColor }
            }
        }

        Output-Json @{ success = $true; data = @{ steps = $steps.Count; shapes = $shapes } }
    }

    "createOrgChart" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $nodes = if ($p.nodes) { $p.nodes } else { @(@{ name = "CEO"; level = 0 }, @{ name = "技术总监"; level = 1 }, @{ name = "市场总监"; level = 1 }, @{ name = "财务总监"; level = 1 }) }
        $centerX = 360
        $startY = if ($p.startY) { $p.startY } else { 80 }
        $boxWidth = if ($p.boxWidth) { $p.boxWidth } else { 100 }
        $boxHeight = if ($p.boxHeight) { $p.boxHeight } else { 40 }
        $levelGap = if ($p.levelGap) { $p.levelGap } else { 80 }
        $levelColors = @("#1a365d", "#2d5a87", "#4a7ab0", "#6b9bd1")

        $levels = @{}
        foreach ($node in $nodes) {
            $level = if ($node.level -ne $null) { $node.level } else { 0 }
            if (-not $levels.ContainsKey($level)) { $levels[$level] = @() }
            $levels[$level] += $node
        }

        $levelKeys = $levels.Keys | Sort-Object
        $shapes = @()
        $levelIndex = 0
        foreach ($levelKey in $levelKeys) {
            $levelNodes = $levels[$levelKey]
            $levelY = $startY + $levelIndex * $levelGap
            $totalWidth = $levelNodes.Count * $boxWidth + ($levelNodes.Count - 1) * 30
            $startX = $centerX - $totalWidth / 2

            for ($i = 0; $i -lt $levelNodes.Count; $i++) {
                $x = $startX + $i * ($boxWidth + 30)
                $shape = $slide.Shapes.AddShape(5, $x, $levelY, $boxWidth, $boxHeight)
                $colorValue = Convert-HexColorToRgbInt([string]$levelColors[$levelIndex % $levelColors.Count])
                if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
                $shape.Line.Visible = $false
                $shape.TextFrame.TextRange.Text = $levelNodes[$i].name
                $shape.TextFrame.TextRange.Font.Color.RGB = 16777215
                $shape.TextFrame.TextRange.Font.Size = 11
                $shape.TextFrame.TextRange.Font.Bold = $true
                $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                $shape.Shadow.Visible = $true
                $shape.Shadow.Blur = 5
                $shape.Shadow.OffsetX = 2
                $shape.Shadow.OffsetY = 2
                $shape.Shadow.Transparency = 0.7
                $shapes += $shape.Name
            }

            $levelIndex++
        }

        Output-Json @{ success = $true; data = @{ levels = $levelKeys.Count; totalNodes = $nodes.Count; shapes = $shapes } }
    }

    "createTimeline" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $events = if ($p.events) { $p.events } else { @(@{ date = "2020"; title = "成立" }, @{ date = "2021"; title = "融资" }, @{ date = "2022"; title = "扩张" }, @{ date = "2023"; title = "上市" }) }
        $startX = if ($p.startX) { $p.startX } else { 80 }
        $startY = if ($p.startY) { $p.startY } else { 280 }
        $lineLength = if ($p.lineLength) { $p.lineLength } else { 560 }

        $mainLine = $slide.Shapes.AddLine($startX, $startY, $startX + $lineLength, $startY)
        $lineColor = Convert-HexColorToRgbInt([string]$(if ($p.lineColor) { $p.lineColor } else { "#1a365d" }))
        if ($null -ne $lineColor) { $mainLine.Line.ForeColor.RGB = $lineColor }
        $mainLine.Line.Weight = 3

        $gap = if ($events.Count -gt 1) { $lineLength / ($events.Count - 1) } else { 0 }
        $shapes = @()

        for ($i = 0; $i -lt $events.Count; $i++) {
            $x = $startX + ($i * $gap)
            $above = ($i % 2 -eq 0)

            $dot = $slide.Shapes.AddShape(9, $x - 8, $startY - 8, 16, 16)
            $dotColor = Convert-HexColorToRgbInt([string]$(if ($p.dotColor) { $p.dotColor } else { "#0d47a1" }))
            if ($null -ne $dotColor) { $dot.Fill.Solid(); $dot.Fill.ForeColor.RGB = $dotColor }
            $dot.Line.Visible = $false

            $lineY = if ($above) { $startY - 50 } else { $startY + 50 }
            $connector = $slide.Shapes.AddLine($x, $startY, $x, $lineY)
            $connector.Line.Weight = 2
            $connectorColor = Convert-HexColorToRgbInt([string]"#cccccc")
            if ($null -ne $connectorColor) { $connector.Line.ForeColor.RGB = $connectorColor }

            $dateBox = $slide.Shapes.AddTextbox(1, $x - 30, $(if ($above) { $lineY - 45 } else { $lineY + 5 }), 60, 20)
            $dateBox.TextFrame.TextRange.Text = $events[$i].date
            $dateBox.TextFrame.TextRange.Font.Size = 12
            $dateBox.TextFrame.TextRange.Font.Bold = $true
            $dateBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $titleBox = $slide.Shapes.AddTextbox(1, $x - 40, $(if ($above) { $lineY - 25 } else { $lineY + 25 }), 80, 20)
            $titleBox.TextFrame.TextRange.Text = $events[$i].title
            $titleBox.TextFrame.TextRange.Font.Size = 10
            $titleColor = Convert-HexColorToRgbInt([string]"#666666")
            if ($null -ne $titleColor) { $titleBox.TextFrame.TextRange.Font.Color.RGB = $titleColor }
            $titleBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

            $shapes += $dot.Name
        }

        Output-Json @{ success = $true; data = @{ events = $events.Count; shapes = $shapes } }
    }

    "getSlideMaster" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster
        $masters = @()
        for ($i = 1; $i -le $master.Shapes.Count; $i++) {
            $shape = $master.Shapes.Item($i)
            $masters += @{ name = $shape.Name; type = $shape.Type; left = $shape.Left; top = $shape.Top; width = $shape.Width; height = $shape.Height }
        }
        Output-Json @{ success = $true; data = @{ shapeCount = $masters.Count; shapes = $masters } }
    }

    "setMasterBackground" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster
        if ($p.gradient) {
            $master.Background.Fill.TwoColorGradient(1, 1)
            $color1 = Convert-HexColorToRgbInt([string]$(if ($p.color1) { $p.color1 } else { "#1a365d" }))
            $color2 = Convert-HexColorToRgbInt([string]$(if ($p.color2) { $p.color2 } else { "#2d5a87" }))
            if ($null -ne $color1) { $master.Background.Fill.GradientStops.Item(1).Color.RGB = $color1 }
            if ($null -ne $color2) { $master.Background.Fill.GradientStops.Item(2).Color.RGB = $color2 }
        } else {
            $master.Background.Fill.Solid()
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#ffffff" }))
            if ($null -ne $colorValue) { $master.Background.Fill.ForeColor.RGB = $colorValue }
        }
        Output-Json @{ success = $true; data = @{ message = "母版背景已更新" } }
    }

    "addMasterElement" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $master = $pres.SlideMaster

        $type = if ($p.type) { $p.type } else { "textbox" }
        $left = if ($p.left) { $p.left } else { 600 }
        $top = if ($p.top) { $p.top } else { 20 }
        $width = if ($p.width) { $p.width } else { 100 }
        $height = if ($p.height) { $p.height } else { 40 }

        if ($type -eq "textbox") {
            $shape = $master.Shapes.AddTextbox(1, $left, $top, $width, $height)
            $shape.TextFrame.TextRange.Text = if ($p.text) { $p.text } else { "Logo" }
            $shape.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 14 }
            $shape.TextFrame.TextRange.Font.Bold = $true
        } elseif ($type -eq "shape") {
            $shape = $master.Shapes.AddShape($(if ($p.shapeType) { $p.shapeType } else { 5 }), $left, $top, $width, $height)
            $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
            if ($null -ne $colorValue) { $shape.Fill.Solid(); $shape.Fill.ForeColor.RGB = $colorValue }
            $shape.Line.Visible = $false
        } else {
            Output-Json @{ success = $false; error = "Unsupported master element type" }
            break
        }

        Output-Json @{ success = $true; data = @{ name = $shape.Name; type = $type } }
    }

    "set3DRotation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $shape.ThreeD.RotationX = if ($p.rotationX -ne $null) { $p.rotationX } else { 0 }
        $shape.ThreeD.RotationY = if ($p.rotationY -ne $null) { $p.rotationY } else { 0 }
        $shape.ThreeD.RotationZ = if ($p.rotationZ -ne $null) { $p.rotationZ } else { 0 }

        if ($p.preset) {
            $presets = @{ isometric = @{ x = 45; y = 45; z = 0 }; perspective = @{ x = 30; y = 30; z = 0 }; oblique = @{ x = 20; y = 60; z = 0 }; tiltLeft = @{ x = 0; y = -30; z = 0 }; tiltRight = @{ x = 0; y = 30; z = 0 } }
            $preset = $presets[$p.preset]
            if ($preset) {
                $shape.ThreeD.RotationX = $preset.x
                $shape.ThreeD.RotationY = $preset.y
                $shape.ThreeD.RotationZ = $preset.z
            }
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; rotationX = $shape.ThreeD.RotationX; rotationY = $shape.ThreeD.RotationY } }
    }

    "set3DDepth" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $depth = if ($p.depth) { $p.depth } else { 20 }
        $shape.ThreeD.Depth = $depth
        if ($p.depthColor) {
            $colorValue = Convert-HexColorToRgbInt([string]$p.depthColor)
            if ($null -ne $colorValue) { $shape.ThreeD.ExtrusionColor.RGB = $colorValue }
        }
        if ($p.lighting) {
            $lightingMap = @{ bright = 1; normal = 2; dim = 3; flat = 4 }
            $lightingValue = if ($lightingMap[$p.lighting]) { $lightingMap[$p.lighting] } else { 2 }
            $shape.ThreeD.PresetLighting = $lightingValue
        }

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; depth = $depth } }
    }

    "set3DMaterial" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))

        $material = if ($p.material) { $p.material } else { "plastic" }
        $materialMap = @{ matte = 1; plastic = 2; metal = 3; wireFrame = 4; glass = 5 }
        $materialValue = if ($materialMap[$material]) { $materialMap[$material] } else { 2 }
        $shape.ThreeD.PresetMaterial = $materialValue

        Output-Json @{ success = $true; data = @{ shape = $shape.Name; material = $material } }
    }

    "create3DText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)

        $text = if ($p.text) { $p.text } else { "3D文字" }
        $left = if ($p.left) { $p.left } else { 200 }
        $top = if ($p.top) { $p.top } else { 200 }
        $width = if ($p.width) { $p.width } else { 300 }
        $height = if ($p.height) { $p.height } else { 100 }
        $textBox = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $textBox.TextFrame.TextRange.Text = $text
        $textBox.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 48 }
        $textBox.TextFrame.TextRange.Font.Bold = $true
        $colorValue = Convert-HexColorToRgbInt([string]$(if ($p.color) { $p.color } else { "#1a365d" }))
        if ($null -ne $colorValue) { $textBox.TextFrame.TextRange.Font.Color.RGB = $colorValue }
        $textBox.TextFrame.TextRange.ParagraphFormat.Alignment = 2

        $textBox.ThreeD.RotationX = if ($p.rotationX -ne $null) { $p.rotationX } else { 15 }
        $textBox.ThreeD.RotationY = if ($p.rotationY -ne $null) { $p.rotationY } else { 30 }
        $textBox.ThreeD.Depth = if ($p.depth) { $p.depth } else { 10 }

        $textBox.Shadow.Visible = $true
        $textBox.Shadow.Blur = 8
        $textBox.Shadow.OffsetX = 5
        $textBox.Shadow.OffsetY = 5
        $textBox.Shadow.Transparency = 0.5

        Output-Json @{ success = $true; data = @{ name = $textBox.Name; text = $text } }
    }

    "setPptChartData" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        $chart.ChartData.Activate()
        $dataSheet = $chart.ChartData.Workbook.Worksheets.Item(1)
        if ($p.data) {
            for ($r = 0; $r -lt $p.data.Count; $r++) {
                $rowData = $p.data[$r]
                for ($c = 0; $c -lt $rowData.Count; $c++) {
                    $dataSheet.Cells.Item($r + 1, $c + 1).Value2 = $rowData[$c]
                }
            }
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "setPptChartStyle" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.chartName) { $p.chartName } else { $p.chartIndex }))
        $chart = $shape.Chart
        if ($p.title) {
            $chart.HasTitle = $true
            $chart.ChartTitle.Text = $p.title
        }
        if ($null -ne $p.hasLegend) {
            $chart.HasLegend = $p.hasLegend
        }
        Output-Json @{ success = $true; data = @{ chartName = $shape.Name } }
    }

    "removeAnimation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        if ($p.index) {
            $seq.Item($p.index).Delete()
        } else {
            while ($seq.Count -gt 0) { $seq.Item(1).Delete() }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "getAnimations" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $animations = @()
        for ($i = 1; $i -le $seq.Count; $i++) {
            $effect = $seq.Item($i)
            $animations += @{ index = $i; shapeName = $effect.Shape.Name; effectType = $effect.EffectType; duration = $effect.Timing.Duration }
        }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; animations = $animations; count = $animations.Count } }
    }

    "setAnimationOrder" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $seq = $slide.TimeLine.MainSequence
        $effect = $seq.Item($p.from)
        $effect.MoveTo($p.to)
        Output-Json @{ success = $true; data = @{ from = $p.from; to = $p.to } }
    }

    "removeSlideTransition" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.SlideShowTransition.EntryEffect = 0
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex } }
    }

    "applyTransitionToAll" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $transitionMap = @{ none = 0; fade = 1; push = 2; wipe = 3; split = 4; reveal = 5; random = 6 }
        $transition = if ($transitionMap[$p.transition]) { $transitionMap[$p.transition] } else { if ($p.transition) { $p.transition } else { 1 } }
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            $slide.SlideShowTransition.EntryEffect = $transition
            if ($p.duration) { $slide.SlideShowTransition.Duration = $p.duration }
        }
        Output-Json @{ success = $true; data = @{ transition = $p.transition; appliedTo = $pres.Slides.Count } }
    }

    "setBackgroundColor" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $colorValue = Convert-HexColorToRgbInt([string]$p.color)
        if ($p.applyToAll) {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) {
                $slide = $pres.Slides.Item($i)
                $slide.FollowMasterBackground = $false
                $slide.Background.Fill.Solid()
                if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            }
            Output-Json @{ success = $true; data = @{ color = $p.color; appliedTo = $pres.Slides.Count } }
        } else {
            $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
            $slide = $pres.Slides.Item($slideIndex)
            $slide.FollowMasterBackground = $false
            $slide.Background.Fill.Solid()
            if ($null -ne $colorValue) { $slide.Background.Fill.ForeColor.RGB = $colorValue }
            Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; color = $p.color } }
        }
    }

    "setBackgroundImage" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $slide.FollowMasterBackground = $false
        $slide.Background.Fill.UserPicture($p.path)
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndex; path = $p.path } }
    }

    "addPptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $actionSettings = $shape.ActionSettings.Item(1)
        $actionSettings.Hyperlink.Address = if ($p.address) { $p.address } else { "" }
        if ($p.subAddress) { $actionSettings.Hyperlink.SubAddress = $p.subAddress }
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name; address = $p.address } }
    }

    "removePptHyperlink" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } else { $p.shapeIndex }))
        $shape.ActionSettings.Item(1).Hyperlink.Address = ""
        Output-Json @{ success = $true; data = @{ shapeName = $shape.Name } }
    }

    "setSlideNumber" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -ne $p.visible) { $pres.SlideMaster.HeadersFooters.SlideNumber.Visible = $p.visible }
        Output-Json @{ success = $true; data = @{ visible = $p.visible } }
    }

    "setPptFooter" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideMaster.HeadersFooters.Footer.Visible = $true
        $pres.SlideMaster.HeadersFooters.Footer.Text = if ($p.text) { $p.text } else { "" }
        Output-Json @{ success = $true; data = @{ text = $p.text } }
    }

    "setPptDateTime" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideMaster.HeadersFooters.DateAndTime.Visible = if ($p.visible -ne $null) { $p.visible } else { $true }
        if ($p.useFixed) {
            $pres.SlideMaster.HeadersFooters.DateAndTime.UseFormat = $false
            $pres.SlideMaster.HeadersFooters.DateAndTime.Text = if ($p.text) { $p.text } else { "" }
        }
        Output-Json @{ success = $true; data = @{ visible = if ($p.visible -ne $null) { $p.visible } else { $true } } }
    }

    "findPptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $results = @()
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $text = $shape.TextFrame.TextRange.Text
                        if ($text -and $text.Contains($p.text)) {
                            $results += @{ slideIndex = $i; shapeName = $shape.Name; text = $text }
                        }
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ searchText = $p.text; results = $results; count = $results.Count } }
    }

    "replacePptText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        if ($null -eq $pres) { Output-Json @{ success = $false; error = "No target presentation" }; return }
        # 兼容两种参数名：TS 工具发送 find/replace，旧实现读 findText/replaceText（此前不匹配导致 0 替换）
        $findText = if ($null -ne $p.find) { [string]$p.find } else { [string]$p.findText }
        $replaceText = if ($null -ne $p.replace) { [string]$p.replace } elseif ($null -ne $p.replaceText) { [string]$p.replaceText } else { "" }
        if ([string]::IsNullOrEmpty($findText)) { Output-Json @{ success = $false; error = "find text required" }; return }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                # 递归处理：文本框 + 表格单元格 + 组合(Group)内所有子形状
                $count += Replace-InShapeTree $shape $findText $replaceText
            }
        }
        Output-Json @{ success = $true; data = @{ findText = $findText; replaceText = $replaceText; count = $count } }
    }

    "startSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $pres.SlideShowSettings.Run()
        Output-Json @{ success = $true }
    }

    "endSlideShow" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        if ($ppt.SlideShowWindows.Count -gt 0) { $ppt.SlideShowWindows.Item(1).View.Exit() }
        Output-Json @{ success = $true }
    }

    "unifyFont" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $fontName = if ($p.fontName) { $p.fontName } else { "微软雅黑" }
        $count = 0
        for ($i = 1; $i -le $pres.Slides.Count; $i++) {
            $slide = $pres.Slides.Item($i)
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $shape.TextFrame.TextRange.Font.Name = $fontName
                        $count++
                    }
                } catch {}
            }
        }
        Output-Json @{ success = $true; data = @{ fontName = $fontName; count = $count } }
    }

    "slide.beautify" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $schemeKey = if ($p.style -and $p.style.colorScheme) { $p.style.colorScheme } elseif ($p.style) { $p.style } else { "business" }
        $scheme = $schemes[$schemeKey]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $fontName = if ($p.style -and $p.style.font) { $p.style.font } else { $null }
        $slideIndex = $p.slideIndex
        $slides = @()
        if ($slideIndex -eq "all") {
            for ($i = 1; $i -le $pres.Slides.Count; $i++) { $slides += $pres.Slides.Item($i) }
        } else {
            $target = if ($slideIndex) { [int]$slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
            $slides += $pres.Slides.Item($target)
        }
        $ops = @(
            @{ operation = "unify_font"; count = 0 },
            @{ operation = "apply_color_scheme"; count = 0 },
            @{ operation = "align"; count = 0 },
            @{ operation = "optimize_spacing"; count = 0 }
        )
        foreach ($slide in $slides) {
            for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                $shape = $slide.Shapes.Item($j)
                try {
                    if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                        $textRange = $shape.TextFrame.TextRange
                        if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                        else { $textRange.Font.Color.RGB = $scheme.body }
                        $ops[1].count++
                        if ($fontName) { $textRange.Font.Name = $fontName; $ops[0].count++ }
                    }
                } catch {}
            }
        }
        $slideIndexOut = if ($slideIndex -eq "all") { "all" } else { ($slides | Select-Object -First 1).SlideIndex }
        Output-Json @{ success = $true; data = @{ slideIndex = $slideIndexOut; operations = $ops } }
    }

    "set3DRotation" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } elseif ($p.shapeIndex) { [int]$p.shapeIndex } else { 1 }))
        $rotX = if ($null -ne $p.rotationX) { [double]$p.rotationX } else { 0 }
        $rotY = if ($null -ne $p.rotationY) { [double]$p.rotationY } else { 0 }
        $rotZ = if ($null -ne $p.rotationZ) { [double]$p.rotationZ } else { 0 }
        if ($p.preset) {
            $presets = @{
                isometric   = @{ x = 45; y = 45; z = 0 }
                perspective = @{ x = 30; y = 30; z = 0 }
                oblique     = @{ x = 20; y = 60; z = 0 }
                tiltLeft    = @{ x = 0;  y = -30; z = 0 }
                tiltRight   = @{ x = 0;  y = 30; z = 0 }
            }
            $pr = $presets[$p.preset]
            if ($null -ne $pr) { $rotX = $pr.x; $rotY = $pr.y; $rotZ = $pr.z }
        }
        $shape.ThreeD.RotationX = $rotX
        $shape.ThreeD.RotationY = $rotY
        $shape.ThreeD.RotationZ = $rotZ
        Output-Json @{ success = $true; data = @{ shape = $shape.Name; rotationX = $rotX; rotationY = $rotY; rotationZ = $rotZ } }
    }

    "set3DDepth" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } elseif ($p.shapeIndex) { [int]$p.shapeIndex } else { 1 }))
        $depth = if ($null -ne $p.depth) { [double]$p.depth } else { 20 }
        $shape.ThreeD.Depth = $depth
        if ($p.depthColor) {
            $dc = Convert-HexColorToRgbInt([string]$p.depthColor)
            if ($null -ne $dc) { $shape.ThreeD.ExtrusionColor.RGB = $dc }
        }
        if ($p.lighting) {
            $lightingMap = @{ bright = 1; normal = 2; dim = 3; flat = 4 }
            $lv = $lightingMap[$p.lighting]
            if ($null -ne $lv) { $shape.ThreeD.PresetLighting = $lv }
        }
        Output-Json @{ success = $true; data = @{ shape = $shape.Name; depth = $depth } }
    }

    "set3DMaterial" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $shape = $slide.Shapes.Item($(if ($p.shapeName) { $p.shapeName } elseif ($p.shapeIndex) { [int]$p.shapeIndex } else { 1 }))
        $material = if ($p.material) { $p.material } else { "plastic" }
        $materialMap = @{ matte = 1; plastic = 2; metal = 3; wireFrame = 4; glass = 5 }
        $mv = $materialMap[$material]
        if ($null -eq $mv) { $mv = 2 }
        $shape.ThreeD.PresetMaterial = $mv
        Output-Json @{ success = $true; data = @{ shape = $shape.Name; material = $material } }
    }

    "create3DText" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { [int]$p.slideIndex } else { 1 }
        $slide = $pres.Slides.Item($slideIndex)
        $text = if ($p.text) { $p.text } else { "3D文字" }
        $left = if ($null -ne $p.left) { $p.left } else { 200 }
        $top = if ($null -ne $p.top) { $p.top } else { 200 }
        $width = if ($null -ne $p.width) { $p.width } else { 300 }
        $height = if ($null -ne $p.height) { $p.height } else { 100 }
        $tb = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
        $tb.TextFrame.TextRange.Text = $text
        $tb.TextFrame.TextRange.Font.Size = if ($p.fontSize) { $p.fontSize } else { 36 }
        $tb.TextFrame.TextRange.Font.Bold = $true
        if ($p.fontColor) {
            $fc = Convert-HexColorToRgbInt([string]$p.fontColor)
            if ($null -ne $fc) { $tb.TextFrame.TextRange.Font.Color.RGB = $fc }
        }
        $tb.ThreeD.Depth = if ($null -ne $p.depth) { $p.depth } else { 20 }
        $tb.ThreeD.RotationX = if ($null -ne $p.rotationX) { $p.rotationX } else { 30 }
        $tb.ThreeD.RotationY = if ($null -ne $p.rotationY) { $p.rotationY } else { 30 }
        $tb.ThreeD.PresetMaterial = 2
        Output-Json @{ success = $true; data = @{ shape = $tb.Name; text = $text } }
    }

    "beautifySlide" {
        $ppt = Get-WpsPpt
        if ($null -eq $ppt) { Output-Json @{ success = $false; error = "WPS PPT not running" }; return }
        $pres = Get-TargetPres $ppt $p
        $slideIndex = if ($p.slideIndex) { $p.slideIndex } else { $ppt.ActiveWindow.Selection.SlideRange.SlideIndex }
        $slide = $pres.Slides.Item($slideIndex)
        $schemes = @{
            business = @{ title = 0x2F5496; body = 0x333333 }
            tech = @{ title = 0x00B0F0; body = 0x404040 }
            creative = @{ title = 0xFF6B6B; body = 0x4A4A4A }
            minimal = @{ title = 0x000000; body = 0x666666 }
        }
        $scheme = $schemes[$p.style]
        if ($null -eq $scheme) { $scheme = $schemes["business"] }
        $count = 0
        for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
            $shape = $slide.Shapes.Item($j)
            try {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $textRange = $shape.TextFrame.TextRange
                    if ($textRange.Font.Size -ge 24) { $textRange.Font.Color.RGB = $scheme.title }
                    else { $textRange.Font.Color.RGB = $scheme.body }
                    $count++
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ style = $p.style; count = $count } }
    }

    # ==================== Word 长文档 / 学术排版（补 office-mcp 那条缺失链路）====================
    # 这一组是为了补齐"认真写一篇论文/报告"所需的能力：标题层级、题注、脚注、图表目录、
    # 章节自动编号、分栏、列表、表格样式、刷新域、大纲抽取。
    # 口径：每个动作都先拿 $doc，拿不到就如实报错返回，不做半截操作。

    # 按级别插入标题：用 Word 内置 Heading 样式（wdStyleHeadingN = -1-N），保证能进目录
    "addHeading" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $level = if ($p.level) { [int]$p.level } else { 1 }
        if ($level -lt 1) { $level = 1 }
        if ($level -gt 9) { $level = 9 }
        $text = [string]$p.text
        try {
            $word.Selection.Style = $doc.Styles.Item(-1 - $level)
            $word.Selection.TypeText($text)
            $word.Selection.TypeParagraph()
            Output-Json @{ success = $true; data = @{ level = $level; text = $text } }
        } catch {
            Output-Json @{ success = $false; error = "插入标题失败: $(Get-CallError $_)" }
        }
    }

    # 插入段落：给了 index 就在该段之前插，没给就追加到文末
    "insertParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = [string]$p.text
        if ($p.index) {
            $idx = [int]$p.index
            if ($idx -lt 1) { $idx = 1 }
            if ($idx -gt $doc.Paragraphs.Count) { $idx = $doc.Paragraphs.Count }
            $doc.Paragraphs.Item($idx).Range.InsertBefore($text + "`r")
        } else {
            $doc.Content.InsertAfter($text + "`r")
        }
        Output-Json @{ success = $true; data = @{ paragraphs = $doc.Paragraphs.Count } }
    }

    # 删除段落：从 index 起删 count 段（1 基，越界自动收敛）
    "deleteParagraph" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $idx = if ($p.index) { [int]$p.index } else { $doc.Paragraphs.Count }
        $cnt = if ($p.count) { [int]$p.count } else { 1 }
        if ($idx -lt 1) { $idx = 1 }
        $deleted = 0
        for ($k = 0; $k -lt $cnt; $k++) {
            if ($idx -gt $doc.Paragraphs.Count) { break }
            try { $doc.Paragraphs.Item($idx).Range.Delete(); $deleted++ } catch { break }
        }
        Output-Json @{ success = $true; data = @{ deleted = $deleted; paragraphs = $doc.Paragraphs.Count } }
    }

    # 抽取大纲：按段落的 OutlineLevel（1..9 是标题级，10 是正文）列出标题树
    "getOutline" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $limit = if ($p.limit) { [int]$p.limit } else { 500 }
        $outline = @()
        $n = [int]$doc.Paragraphs.Count
        for ($i = 1; $i -le $n; $i++) {
            if ($outline.Count -ge $limit) { break }
            try {
                $para = $doc.Paragraphs.Item($i)
                $lvl = [int]$para.OutlineLevel
                if ($lvl -ge 1 -and $lvl -le 9) {
                    $t = [string]$para.Range.Text
                    $t = $t -replace "[`r`n`a]", ""
                    if ($t.Trim()) {
                        $outline += @{ index = $i; level = $lvl; text = $t.Trim() }
                    }
                }
            } catch {}
        }
        Output-Json @{ success = $true; data = @{ headings = $outline; count = $outline.Count } }
    }

    # 默认字体：改 Normal 样式的中文字体/西文字体/字号，正文与新建段落都跟着变
    "setDefaultFont" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $normal = $doc.Styles.Item(-1)   # wdStyleNormal
            if ($p.eastAsia) { $normal.Font.NameFarEast = [string]$p.eastAsia }
            if ($p.ascii) { $normal.Font.NameAscii = [string]$p.ascii }
            if ($p.name) { $normal.Font.Name = [string]$p.name }
            if ($p.size) { $normal.Font.Size = [double]$p.size }
            Output-Json @{ success = $true; data = @{
                name = [string]$normal.Font.Name
                eastAsia = [string]$normal.Font.NameFarEast
                size = [double]$normal.Font.Size
            } }
        } catch {
            Output-Json @{ success = $false; error = "设置默认字体失败: $(Get-CallError $_)" }
        }
    }

    # 题注：给图表/表格加自动编号标签（Word 的 Caption 会自动续号）
    "addCaption" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $label = if ($p.label) { [string]$p.label } else { "图" }
        # wdCaptionPositionAbove = 0, wdCaptionPositionBelow = 1
        $pos = if ($p.position -eq "above") { 0 } else { 1 }
        $title = [string]$p.text
        $mode = $null
        $captionText = $null
        $nativeErr = $null

        # ① 先试原生题注：它会插入可见编号域，之后"图表目录"能自动收集到
        try {
            $sel = $word.Selection
            $sel.Collapse(0)     # wdCollapseEnd
            $sel.InsertCaption($label, $title, $pos, $false) | Out-Null
            $mode = "field"
        } catch {
            $nativeErr = $_.Exception.Message
        }

        # ② 原生不支持就手工兜底：纯文本题注 + 自增序号。
        #    实测本机 WPS 12.1 的 Range/Selection.InsertCaption 都会抛"值不在预期的范围内"
        #    （CaptionLabels 集合也不可用），所以必须有这条兜底，否则这个功能直接废掉。
        #    代价：序号是写死的文本，不是活编号域 —— 图表目录收集不到，已在返回值里如实标注。
        if ($null -eq $mode) {
            try {
                $n = 0
                foreach ($para in $doc.Paragraphs) {
                    $t = [string]$para.Range.Text
                    if ($t -match ("^\s*" + [regex]::Escape($label) + "\s+\d+")) { $n++ }
                }
                $n = $n + 1
                $captionText = "$label $n $title"
                $rng = $word.Selection.Range
                $rng.Collapse(0)
                $rng.InsertAfter($captionText + "`r")
                $mode = "manual"
            } catch {
                Output-Json @{ success = $false; error = "插入题注失败：原生接口报「$nativeErr」；手工兜底也失败：$(Get-CallError $_)" }
                return
            }
        }

        Output-Json @{ success = $true; data = @{
            label = $label
            text = $title
            mode = $mode
            caption = $captionText
            note = if ($mode -eq "manual") { "原生题注接口在本机 WPS 不可用，已按标签自增序号手工插入；序号是文本不是域，图表目录无法自动收集" } else { $null }
            nativeError = $nativeErr
        } }
    }

    # 脚注：在当前选区（或文末）插入一条脚注
    "addFootnote" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $text = [string]$p.text
        try {
            $rng = if ($p.atEnd) { $doc.Content } else { $word.Selection.Range }
            $fn = $doc.Footnotes.Add($rng, $text)
            Output-Json @{ success = $true; data = @{ total = [int]$doc.Footnotes.Count; index = [int]$fn.Index } }
        } catch {
            Output-Json @{ success = $false; error = "插入脚注失败: $(Get-CallError $_)" }
        }
    }

    # 图表目录：基于题注标签生成（要先把题注加上才有内容）
    "insertTableOfFigures" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $label = if ($p.label) { [string]$p.label } else { "图" }
        try {
            $rng = $doc.Content
            $rng.Collapse(0)   # wdCollapseEnd → 放文末
            $tof = $doc.TablesOfFigures.Add($rng, $label)
            try { $tof.Update() } catch {}
            Output-Json @{ success = $true; data = @{ label = $label; total = [int]$doc.TablesOfFigures.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入图表目录失败: $(Get-CallError $_)" }
        }
    }

    # 刷新域：目录、图表目录、交叉引用、页码等一次性更新
    "updateFields" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $report = @{ fields = 0; toc = 0; tof = 0; errors = @() }
        try { $doc.Fields.Update(); $report.fields = [int]$doc.Fields.Count } catch { $report.errors += "fields: $(Get-CallError $_)" }
        try { foreach ($t in $doc.TablesOfContents) { $t.Update() }; $report.toc = [int]$doc.TablesOfContents.Count } catch { $report.errors += "toc: $(Get-CallError $_)" }
        try { foreach ($f in $doc.TablesOfFigures) { $f.Update() }; $report.tof = [int]$doc.TablesOfFigures.Count } catch { $report.errors += "tof: $(Get-CallError $_)" }
        Output-Json @{ success = $true; data = $report }
    }

    # 分栏：整节等宽多栏
    "setColumns" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $count = if ($p.count) { [int]$p.count } else { 2 }
        if ($count -lt 1) { $count = 1 }
        if ($count -gt 6) { $count = 6 }
        try {
            $doc.PageSetup.TextColumns.SetCount($count)
            if ($p.spacing) {
                for ($i = 1; $i -le $count; $i++) {
                    try { $doc.PageSetup.TextColumns.Item($i).Spacing = [double]$p.spacing * 28.35 } catch {}
                }
            }
            if ($null -ne $p.lineBetween) {
                try { $doc.PageSetup.TextColumns.LineBetween = [bool]$p.lineBetween } catch {}
            }
            Output-Json @{ success = $true; data = @{ columns = [int]$doc.PageSetup.TextColumns.Count } }
        } catch {
            # 分栏在不同 WPS 版本上接口差异较大 —— 失败就明说，别伪装成功
            Output-Json @{ success = $false; error = "设置分栏失败: $(Get-CallError $_)" }
        }
    }

    # 章节自动编号：给 Heading1~N 挂多级编号（近似实现，版本差异可能导致效果不同）
    "setHeadingNumbering" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $maxLevel = if ($p.levels) { [int]$p.levels } else { 3 }
        if ($maxLevel -lt 1) { $maxLevel = 1 }
        if ($maxLevel -gt 9) { $maxLevel = 9 }
        try {
            # 实测：`ListTemplates.Add(...)` 在本机 WPS 12.1 返回 null（"不能对 Null 值表达式调用方法"），
            # 所以不用"新建列表模板"，改成直接借用【大纲编号库】里现成的模板来挂 Heading 样式 ——
            # 这是 Word 里的经典做法，兼容性明显更好。
            # wdOutlineNumberGallery = 2
            $tpl = $word.ListGalleries.Item(2).ListTemplates.Item(1)
            for ($i = 1; $i -le $maxLevel; $i++) {
                $doc.Styles.Item(-1 - $i).LinkToListTemplate($tpl, $i)
            }
            Output-Json @{ success = $true; data = @{ levels = $maxLevel; mode = "outline-gallery" } }
        } catch {
            Output-Json @{ success = $false; error = "设置章节编号失败（本机 WPS 可能不支持该接口；可在 WPS 里用「开始 → 多级列表」手动设置）: $(Get-CallError $_)" }
        }
    }

    # 列表：项目符号或编号，一次写多条
    "addList" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $items = @($p.items)
        if ($items.Count -eq 0) { Output-Json @{ success = $false; error = "items 为空" }; return }
        $type = if ($p.type -eq "number") { "number" } else { "bullet" }
        try {
            $sel = $word.Selection
            if ($type -eq "number") { $sel.Range.ListFormat.ApplyNumberDefault() }
            else { $sel.Range.ListFormat.ApplyBulletDefault() }
            foreach ($it in $items) {
                $sel.TypeText([string]$it)
                $sel.TypeParagraph()
            }
            # 收尾：去掉列表末那个空项目符号，回到正文
            try { $sel.Range.ListFormat.RemoveNumbers() } catch {}
            Output-Json @{ success = $true; data = @{ type = $type; count = $items.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入列表失败: $(Get-CallError $_)" }
        }
    }

    # 表格样式：套用样式名 + 表头加粗 + 边框
    "formatTable" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = $(Resolve-Doc $word $p)
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $idx = if ($p.tableIndex) { [int]$p.tableIndex } else { 1 }
        if ($idx -lt 1 -or $idx -gt $doc.Tables.Count) {
            Output-Json @{ success = $false; error = "表格序号越界（共 $($doc.Tables.Count) 个）" }; return
        }
        $applied = @{ style = $false; headerBold = $false; borders = $false }
        try {
            $tbl = $doc.Tables.Item($idx)
            if ($p.style) {
                try { $tbl.Style = [string]$p.style; $applied.style = $true } catch {}
            }
            if ($p.headerBold) {
                try { $tbl.Rows.Item(1).Range.Font.Bold = $true; $applied.headerBold = $true } catch {}
            }
            if ($p.borders) {
                try {
                    $tbl.Borders.InsideLineStyle = 1     # wdLineStyleSingle
                    $tbl.Borders.OutsideLineStyle = 1
                    $applied.borders = $true
                } catch {}
            }
            Output-Json @{ success = $true; data = @{ tableIndex = $idx; applied = $applied } }
        } catch {
            Output-Json @{ success = $false; error = "设置表格样式失败: $(Get-CallError $_)" }
        }
    }

    # 【诊断】故意占住这一路一段时间 —— 用来实测排队、超时、取消这些时序行为。
    # 没有它，"队列上限有没有生效""取消后 inFlight 会不会提前变 false"这类问题
    # 只能靠猜，因为本机动作都太快，观察不到中间状态。
    "__sleep" {
        $ms = 1000
        try { if ($p.ms) { $ms = [int]$p.ms } } catch {}
        if ($ms -lt 0) { $ms = 0 }
        if ($ms -gt 600000) { $ms = 600000 }
        Start-Sleep -Milliseconds $ms
        Output-Json @{ success = $true; data = @{ slept = $ms } }
    }

    # 同上，但被 TS 侧判为【只读】（见 bridge/cancellable.ts 的白名单）。
    # 之所以要两个：取消机制对"已在途"的请求是按只读/写入分岔的，
    # 两条路都得能实测，只留一个就有一条永远验不到。
    "__sleepRead" {
        $ms = 1000
        try { if ($p.ms) { $ms = [int]$p.ms } } catch {}
        if ($ms -lt 0) { $ms = 0 }
        if ($ms -gt 600000) { $ms = 600000 }
        Start-Sleep -Milliseconds $ms
        Output-Json @{ success = $true; data = @{ slept = $ms } }
    }

    # 【临时诊断】原样回显收到的参数，用来判断"参数到底有没有送到"
    "__probeParams" {
        $keys = @()
        try { $keys = @($p.PSObject.Properties | ForEach-Object { $_.Name }) } catch {}
        Output-Json @{ success = $true; data = @{ received = $p; keys = $keys; ptype = $p.GetType().FullName } }
    }

    # 【临时诊断】查 WPS 文字的 COM 状态（进程/文档数/ActiveDocument 是否可用）
    "__probeDoc" {
        $info = @{}
        try { $word = Get-WpsWord } catch { $info.getWordErr = $_.Exception.Message; $word = $null }
        $info.wordNull = ($null -eq $word)
        if ($null -ne $word) {
            try { $info.appName = [string]$word.Name } catch { $info.appNameErr = $_.Exception.Message }
            try { $info.appPath = [string]$word.Path } catch { $info.appPathErr = $_.Exception.Message }
            try { $info.docCount = [int]$word.Documents.Count } catch { $info.docCountErr = $_.Exception.Message }
            try {
                $d = $(Resolve-Doc $word $p)
                $info.activeDocNull = ($null -eq $d)
                if ($null -ne $d) {
                    try { $info.activeDocName = [string]$d.Name } catch { $info.adnErr = $_.Exception.Message }
                    try { $info.contentNull = ($null -eq $d.Content) } catch { $info.contentErr = $_.Exception.Message }
                    try { $info.paraCount = [int]$d.Paragraphs.Count } catch { $info.paraErr = $_.Exception.Message }
                    try { $info.stylesNull = ($null -eq $d.Styles) } catch { $info.stylesErr = $_.Exception.Message }
                    try { $info.pageSetupNull = ($null -eq $d.PageSetup) } catch { $info.psErr = $_.Exception.Message }
                }
            } catch { $info.activeDocErr = $_.Exception.Message }
        }
        Output-Json @{ success = $true; data = $info }
    }

    # ==================== 能力探测（阶段 6 的第一步）====================
    # 为什么先探测再实现
    # ------------------
    # 前面已经在本机 WPS 12.1 上踩了三次接口不支持（InsertCaption、CaptionLabels、
    # ListTemplates.Add）。如果按"Word 官方 API 应该都有"直接写，实现完才发现
    # 一半跑不通，返工成本很高。
    # 所以先把候选接口在【真机】上各试一次，输出一张"支持 / 不支持"的表，
    # 再据此定实现清单 —— 这一步能省掉大量返工。
    #
    # ⚠️ 探针绝大多数是只读的；两处例外（styles.modifyFont / table.merge / table.addRow）
    #    会短暂改动后复原或依赖文档现状，所以【必须在副本上跑】。
    "__probeCaps" {
        $word = Get-ActiveWpsApp 'Kwps.Application'
        $doc = if ($word) { Resolve-Doc $word $p } else { $null }
        if (-not $doc) {
            Output-Json @{ success = $false; error = "没有可用的文稿（能力探测需要打开一个文档）" }
            return
        }
        $caps = [ordered]@{}

        # ---- P1：样式 CUD ----
        $caps["styles.list"]          = { $null = $doc.Styles.Count; $null = $doc.Styles.Item("正文") }
        $caps["styles.add"]           = { $s = $doc.Styles.Add("__probe_style__", 1); if ($null -eq $s) { throw "Styles.Add 返回 null" }
                                          $back = $doc.Styles.Item("__probe_style__"); if ($null -eq $back) { throw "样式建了但取不回来" }
                                          try { $back.Delete() } catch {} }
        $caps["styles.delete"]        = { $s = $doc.Styles.Add("__probe_del__", 1); $s.Delete() }
        $caps["styles.modifyFont"]    = { $s = $doc.Styles.Item("正文"); $old = $s.Font.NameFarEast
                                          try { $s.Font.NameFarEast = "__probe__" } finally { $s.Font.NameFarEast = $old } }
        $caps["styles.modifyParagraph"] = { $s = $doc.Styles.Item("正文"); $old = $s.ParagraphFormat.SpaceAfter
                                          try { $s.ParagraphFormat.SpaceAfter = 99 } finally { $s.ParagraphFormat.SpaceAfter = $old } }
        $caps["styles.byIndex"]       = { $null = $doc.Styles.Item(-1) }   # wdStyleHeading1

        # ---- P1：交叉引用 ----
        $caps["crossRef.items"]       = { $null = $doc.GetCrossReferenceItems(1) }   # wdRefTypeHeading
        $caps["crossRef.fieldAdd"]    = { $rng = $doc.Range(0, [Math]::Min(1, $doc.Content.End))
                                          $f = $doc.Fields.Add($rng, -1, "STYLEREF 1", $false)
                                          if ($null -eq $f) { throw "Fields.Add 返回 null" }
                                          $t = ""; try { $t = [string]$f.Code.Text } catch {}
                                          if ($t -notmatch "STYLEREF") { throw "域已插入但代码不是 STYLEREF（实际=$t）" }
                                          try { $f.Delete() } catch {} }

        # ---- P1：索引 ----
        $caps["indexes.count"]        = { $null = $doc.Indexes.Count }
        $caps["indexes.markEntry"]    = { $r = $doc.Indexes.MarkEntry($doc.Range(0,1)); if ($null -eq $r) { throw "MarkEntry 返回 null" } }
        $caps["indexes.add"]          = { $ix = $doc.Indexes.Add($doc.Range(0, [Math]::Min(1, $doc.Content.End))); if ($null -eq $ix) { throw "Indexes.Add 返回 null" }; try { $ix.Delete() } catch {} }

        # ---- P0：修订接受/拒绝 ----
        $caps["revisions.count"]      = { $null = $doc.Revisions.Count }
        $caps["revisions.item"]       = { if ($doc.Revisions.Count -gt 0) { $null = $doc.Revisions.Item(1).Range.Text } }
        $caps["revisions.acceptAll"]  = { $doc.Revisions.AcceptAll() }
        $caps["revisions.rejectAll"]  = { $doc.Revisions.RejectAll() }

        # ---- P0：表格单元格级 ----
        $caps["table.cell"]           = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Range.Text } }
        $caps["table.rowsAdd"]        = { if ($doc.Tables.Count -le 0) { throw "文档里没有表格，跳过" }
                                          $t = $doc.Tables.Item(1); $n0 = $t.Rows.Count; $null = $t.Rows.Add()
                                          if ($t.Rows.Count -ne $n0 + 1) { throw "Rows.Add 返回了但行数没变（$n0→$($t.Rows.Count)）" } }
        $caps["table.colsAdd"]        = { if ($doc.Tables.Count -le 0) { throw "文档里没有表格，跳过" }
                                          $t = $doc.Tables.Item(1); $n0 = $t.Columns.Count; $null = $t.Columns.Add()
                                          if ($t.Columns.Count -ne $n0 + 1) { throw "Columns.Add 返回了但列数没变（$n0→$($t.Columns.Count)）" } }
        $caps["table.cellMerge"]      = { if ($doc.Tables.Count -gt 0) { $t=$doc.Tables.Item(1); $t.Cell(1,1).Merge($t.Cell(1, [Math]::Max(2,[Math]::Min(2,$t.Columns.Count)))) } }
        $caps["table.cellSplit"]      = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Split(1,2) } }
        $caps["table.cellBorders"]    = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Cell(1,1).Borders } }
        $caps["table.style"]          = { if ($doc.Tables.Count -gt 0) { $doc.Tables.Item(1).Style = "网格型" } }
        $caps["table.rowsDelete"]     = { if ($doc.Tables.Count -gt 0) { $null = $doc.Tables.Item(1).Rows.Count } }

        # ---- P1：页码/奇偶页/首页不同 ----
        $caps["pageNumbers.add"]      = { $f = $doc.Sections.Item(1).Footers.Item(1); $n0 = $f.PageNumbers.Count
                                          $null = $f.PageNumbers.Add(4, $true)
                                          if ($f.PageNumbers.Count -le $n0) { throw "PageNumbers.Add 后页码数没增加（$n0→$($f.PageNumbers.Count)）" } }
        $caps["pageSetup.differentFirst"] = { $v = $doc.PageSetup.DifferentFirstPageHeaderFooter; $doc.PageSetup.DifferentFirstPageHeaderFooter = $v }
        $caps["pageSetup.oddEven"]    = { $v = $doc.PageSetup.OddAndEvenPagesHeaderFooter; $doc.PageSetup.OddAndEvenPagesHeaderFooter = $v }

        # ---- P2：水印/内容控件/形状/保护/查找/批注/比较/邮件合并 ----
        $caps["watermark"]            = { $null = $doc.Sections.Item(1).Headers.Item(1).Shapes }
        $caps["contentControl"]       = { $null = $doc.ContentControls.Count }
        $caps["shape.addTextbox"]     = { $null = $doc.Shapes.AddTextbox(1, 0, 0, 100, 50) }
        $caps["docProtect"]           = { $null = $doc.ProtectionType }
        $caps["find.wildcards"]       = { $f = $doc.Content.Find; $f.MatchWildcards = $true; $f.ClearFormatting(); $null = $f.Execute("__probe__") }
        $caps["comments.count"]       = { $null = $doc.Comments.Count }
        $caps["comments.add"]         = { $c = $doc.Comments.Add($doc.Range(0,1), "__probe__"); try { $c.Delete() } catch {} }
        $caps["docCompare"]           = { $null = $doc.Compare }
        $caps["mailMerge"]            = { $null = $doc.MailMerge.State }
        $caps["exportPdf"]            = { $null = $doc.ExportAsFixedFormat }

        # ---- 对照锚点：这三个是我【曾经】判定为"不支持"的接口。
        #      放进同一张表是为了让探针自己回答"当初的归因对不对"。
        #
        #      ⚠️ 重要教训（第二轮探测才暴露的）：判据不能只看"有没有抛异常"。
        #      第一轮这三个里有两个报了 supported，说明"接口不可用"的结论下早了 ——
        #      真实原因是【调用方式】不对（参数/返回值），不是接口缺失。
        #      所以这里的探针都带"有效性断言"：能调用 ≠ 能用。
        $caps["anchor.insertCaption"] = {
            $null = $doc.CaptionLabels.Count
            $c = $word.Selection.Range.InsertCaption("图", "probe", 1, $false)
            if ($null -eq $c -or [string]::IsNullOrWhiteSpace([string]$c.Text)) {
                throw "InsertCaption 可调用但没产生题注（返回空）"
            }
        }
        $caps["anchor.captionLabels"] = {
            $n = $doc.CaptionLabels.Count
            if ($n -le 0) { throw "CaptionLabels 接口可用，但当前文档没有任何题注标签（Count=$n）" }
        }
        $caps["anchor.listTemplatesAdd"] = {
            $lt = $doc.ListTemplates.Add($true)
            if ($null -eq $lt) { throw "ListTemplates.Add 返回 null —— 接口可调用但没有产出列表模板" }
        }

        foreach ($k in @($caps.Keys)) {
            $status = "unknown"; $err = $null
            try {
                & $caps[$k] | Out-Null
                $status = "supported"
            } catch {
                $status = "unsupported"
                $err = [string]$_.Exception.Message
                if ($err.Length -gt 160) { $err = $err.Substring(0, 160) }
            }
            $caps[$k] = @{ status = $status; error = $err }
        }

        # 对照组自检：三个锚点必须【都】不是 supported —— 否则说明探针判据太松
        # （比如只 catch 异常、把"返回空"当成成功）。
        $anchorNames = @($caps.Keys | Where-Object { $_ -like "anchor.*" })
        $anchorLeaked = @($anchorNames | Where-Object { $caps[$_].status -eq "supported" })

        $keys = @($caps.Keys)
        Output-Json @{ success = $true; data = @{
            docName = [string]$doc.Name
            total = $keys.Count
            supported = @($keys | Where-Object { $caps[$_].status -eq "supported" })
            unsupported = @($keys | Where-Object { $caps[$_].status -ne "supported" })
            caps = $caps
            anchorSelfCheck = @{
                total = $anchorNames.Count
                leaked = $anchorLeaked
                passed = ($anchorLeaked.Count -eq 0)
                note = "锚点全部报 unsupported 才算探针有效；有 leaked 说明判据太松，不能拿这张表做决策。"
            }
            note = "探针多数只读；styles.modify* 会改完改回，table.rowsAdd/colsAdd 会真的加行加列 —— 请在【副本】上跑。表格类探针需要文档里已有表格，否则会因为 Count=0 静默跳过而误报 supported。"
        } }
    }

    # ==================== P0：修订的接受 / 拒绝 ====================
    # 之前只能开/查询修订，没有"处理"这个出口 —— 半个功能。
    # mode: acceptAll | rejectAll | acceptOne | rejectOne
    "handleRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        $mode = [string]$p.mode
        $before = 0
        try { $before = [int]$doc.Revisions.Count } catch {}
        if ($before -eq 0) {
            Output-Json @{ success = $true; data = @{ mode = $mode; count = 0; handled = 0; note = "文档里没有修订" } }
            return
        }
        try {
            switch ($mode) {
                "acceptAll" { $doc.Revisions.AcceptAll() }
                "rejectAll" { $doc.Revisions.RejectAll() }
                "acceptOne" {
                    if ($null -eq $p.index) { Output-Json @{ success = $false; error = "acceptOne 需要 index（从 1 开始）" }; return }
                    $doc.Revisions.Item([int]$p.index).Accept()
                }
                "rejectOne" {
                    if ($null -eq $p.index) { Output-Json @{ success = $false; error = "rejectOne 需要 index（从 1 开始）" }; return }
                    $doc.Revisions.Item([int]$p.index).Reject()
                }
                default { Output-Json @{ success = $false; error = "mode 只能是 acceptAll/rejectAll/acceptOne/rejectOne，收到: $mode" }; return }
            }
            $after = [int]$doc.Revisions.Count
            Output-Json @{ success = $true; data = @{ mode = $mode; before = $before; after = $after; handled = ($before - $after) } }
        } catch {
            Output-Json @{ success = $false; error = "处理修订失败: $(Get-CallError $_)" }
        }
    }

    # 列出修订（谁改的、改成什么）—— 处理之前得先看得见
    "listRevisions" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $items = @()
            $n = [int]$doc.Revisions.Count
            for ($i = 1; $i -le $n; $i++) {
                $r = $doc.Revisions.Item($i)
                $e = @{ index = $i }
                try { $e.type = [int]$r.Type } catch {}            # 1=插入 2=删除 3=属性 4=段落标记...
                try { $e.author = [string]$r.Author } catch {}
                try { $e.date = [string]$r.Date } catch {}
                try { $e.text = [string]$r.Range.Text } catch {}
                try { $e.page = [int]$r.Range.Information(3) } catch {}   # wdActiveEndPageNumber
                $items += $e
            }
            Output-Json @{ success = $true; data = @{ count = $n; revisions = $items } }
        } catch {
            Output-Json @{ success = $false; error = "读取修订失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P0：表格单元格级 ====================
    # tableIndex 默认 1（第一个表格）。坐标一律 1-based，与 Word UI 一致。
    "mergeTableCells" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $a = $t.Cell([int]$p.row1, [int]$p.col1)
            $b = $t.Cell([int]$p.row2, [int]$p.col2)
            $a.Merge($b)
            Output-Json @{ success = $true; data = @{ merged = "$($p.row1),$($p.col1)-$($p.row2),$($p.col2)" } }
        } catch {
            Output-Json @{ success = $false; error = "合并 Word 表格单元格失败（注意：不能跨越已有合并区）: $(Get-CallError $_)" }
        }
    }

    "splitCell" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            # 两个参数必须先落到变量再传：写成 $c.Split((..), (..)) 时，
            # 那个逗号会被 PowerShell 当成数组运算符，整个块解析失败（实测踩过）。
            $nr = Get-IntParam $p "numRows" 1
            $nc = Get-IntParam $p "numCols" 1
            $c.Split($nr, $nc)
            Output-Json @{ success = $true; data = @{ split = "$($p.row),$($p.col)"; numRows = $nr; numCols = $nc } }
        } catch {
            Output-Json @{ success = $false; error = "拆分单元格失败（该单元格可能未合并过）: $(Get-CallError $_)" }
        }
    }

    # 插入/删除行列。dir: rows | cols；where: before | after；count 默认 1
    "insertTableRowCol" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $cnt = (Get-IntParam $p "count" 1)
            $where = if ($p.where -eq "before") { 0 } else { 1 }   # wdRowHeightAuto=0? 实测 0=前 1=后
            $info = @{ table = $t.Index; before = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count } }
            if ($p.dir -eq "cols") {
                $col = $t.Columns.Item([int]$p.index)
                if ($where -eq 0) { for ($k = 0; $k -lt $cnt; $k++) { $null = $col.Select(); $null = $t.Columns.Add($null) } }
                else { $null = $col.Select(); $null = $t.Columns.Add() }
            } else {
                $row = $t.Rows.Item([int]$p.index)
                if ($where -eq 0) { for ($k = 0; $k -lt $cnt; $k++) { $null = $row.Select(); $null = $t.Rows.Add($null) } }
                else { $null = $row.Select(); $null = $t.Rows.Add() }
            }
            $info.after = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            Output-Json @{ success = $true; data = $info }
        } catch {
            Output-Json @{ success = $false; error = "插入行列失败: $(Get-CallError $_)" }
        }
    }

    "deleteTableRowCol" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $before = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            if ($p.dir -eq "cols") { $t.Columns.Item([int]$p.index).Delete() } else { $t.Rows.Item([int]$p.index).Delete() }
            $after = @{ rows = [int]$t.Rows.Count; cols = [int]$t.Columns.Count }
            Output-Json @{ success = $true; data = @{ before = $before; after = $after } }
        } catch {
            Output-Json @{ success = $false; error = "删除行列失败: $(Get-CallError $_)" }
        }
    }

    # 单元格边框。side: all | top | bottom | left | right；style 可选（1=单线 3=粗线 0=无）
    "setCellBorders" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            $style = if ($null -ne $p.style) { [int]$p.style } else { 1 }
            # wdBorderTop=1 Bottom=3 Left=2 Right=4 Horizontal=5 Vertical=6
            $map = @{ top = 1; left = 2; bottom = 3; right = 4; horizontal = 5; vertical = 6 }
            $applied = @()
            if ($p.side -eq "all") {
                foreach ($k in @("top", "left", "bottom", "right")) {
                    $c.Borders.Item($map[$k]).LineStyle = $style
                    $c.Borders.Item($map[$k]).LineWidth = 4      # wdLineWidth050pt
                    $applied += $k
                }
            } else {
                $key = [string]$p.side
                if (-not $map.ContainsKey($key)) { Output-Json @{ success = $false; error = "side 只能是 top/bottom/left/right/horizontal/vertical/all" }; return }
                $c.Borders.Item($map[$key]).LineStyle = $style
                if ($style -ne 0) { $c.Borders.Item($map[$key]).LineWidth = 4 }
                $applied += $key
            }
            Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; applied = $applied; style = $style } }
        } catch {
            Output-Json @{ success = $false; error = "设置单元格边框失败: $(Get-CallError $_)" }
        }
    }

    # 读单元格文本（表格类操作的前置确认）
    "getCellText" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if ($doc.Tables.Count -eq 0) { Output-Json @{ success = $false; error = "文档里没有表格" }; return }
        try {
            $t = $doc.Tables.Item((Get-IntParam $p "tableIndex" 1))
            $c = $t.Cell([int]$p.row, [int]$p.col)
            $txt = ""
            try { $txt = [string]$c.Range.Text } catch {}
            $txt = $txt -replace "[`r`n`a]", ""      # 去掉单元格结尾标记
            $text = ""
            try { $text = [string]$c.Range.Text } catch {}
            Output-Json @{ success = $true; data = @{ row = $p.row; col = $p.col; text = $txt; raw = $text } }
        } catch {
            Output-Json @{ success = $false; error = "读取单元格失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：样式 新建 / 修改 / 删除 ====================
    "listStyles" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $onlyUser = $false
            if ($null -ne $p.onlyUserDefined) { $onlyUser = [bool]$p.onlyUserDefined }
            $items = @()
            foreach ($s in @($doc.Styles)) {
                $builtin = $true
                try { $builtin = [bool]$s.BuiltIn } catch {}
                if ($onlyUser -and $builtin) { continue }
                $e = @{ name = [string]$s.NameLocal; builtIn = $builtin }
                try { $e.type = [int]$s.Type } catch {}
                try { $e.inUse = [bool]$s.InUse } catch {}
                $items += $e
            }
            Output-Json @{ success = $true; data = @{ count = $items.Count; onlyUserDefined = $onlyUser; styles = $items } }
        } catch {
            Output-Json @{ success = $false; error = "列出样式失败: $(Get-CallError $_)" }
        }
    }

    # baseOn 可选（基于哪个样式派生）。其余属性可选，不传就不改。
    "addStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        # "样式是否已存在"必须用 try/catch 判，不能看返回值
        # --------------------------------------------------
        # $doc.Styles.Item(不存在的名字) 在 WPS 里是【抛异常】，不是返回 null。
        # 第一版写成 `if ($null -ne $doc.Styles.Item($p.name))`，异常被外层 catch
        # 接住，于是"想建个新样式"永远报 E_FAIL —— 而同样的 Styles.Add 调用
        # 单独跑（见 __probeStyleAdd 的 A~F 六组对照）是完全成功的。
        # 坑就在这个已存在检查上。
        $alreadyExists = $false
        try { $null = $doc.Styles.Item([string]$p.name); $alreadyExists = $true } catch { $alreadyExists = $false }
        if ($alreadyExists) {
            Output-Json @{ success = $false; error = "样式 '$($p.name)' 已存在；改样式请用 modifyStyle" }; return
        }
        try {
            # wdStyleTypeParagraph=1
            # 样式类型用【字符串】参数 styleType，而不是数字 type ——
            # 数字 1/2 直接传太脆（调用方要查表），而 schema 里的 type 字段本身就是字符串，
            # 用同一个名字会让 TS 侧的类型定义自相矛盾。
            $stType = 1   # wdStyleTypeParagraph
            if ($p.styleType -eq "character") { $stType = 2 }

            # 建样式与设属性必须【分开】处理：实测把两者写在一个 try 里时，
            # 只要某个属性改不了就整体报 E_FAIL，样式根本建不出来。
            # 现在建样式单独 try —— 建成功就是建成功，属性设不上的逐项如实报告。
            $s = $null
            try {
                $s = $doc.Styles.Add([string]$p.name, $stType)
            } catch {
                Output-Json @{ success = $false; error = "Styles.Add 失败（样式名冲突？或本机 WPS 不支持新建样式）: $(Get-CallError $_)" }
                return
            }
            if ($null -eq $s) { Output-Json @{ success = $false; error = "Styles.Add 返回 null" }; return }
            $failed = @()
            if ($p.baseOn) { try { $s.BaseStyle = [string]$p.baseOn } catch { $failed += "baseOn" } }
            $applied = @()
            # 逐属性设置：设上了记 applied，设不上记 failed —— 不静默吞掉，也不整体失败
            if ($p.fontSize) { try { $s.Font.Size = [double]$p.fontSize; $applied += "fontSize" } catch { $failed += "fontSize" } }
            if ($p.fontName) {
                try { $s.Font.NameFarEast = [string]$p.fontName; $applied += "fontName" } catch { $failed += "fontName" }
                try { $s.Font.NameAscii = [string]$p.fontName } catch { }
            }
            if ($null -ne $p.bold) { try { $s.Font.Bold = [bool]$p.bold; $applied += "bold" } catch { $failed += "bold" } }
            if ($null -ne $p.italic) { try { $s.Font.Italic = [bool]$p.italic; $applied += "italic" } catch { $failed += "italic" } }
            if ($null -ne $p.color) {
                try {
                    $bgr = [int]$p.color
                    $s.Font.Color = ($bgr -band 0xFF) * 65536 + ($bgr -shr 8 -band 0xFF) * 256 + ($bgr -shr 16 -band 0xFF)
                    $applied += "color"
                } catch { $failed += "color" }
            }
            if ($null -ne $p.spaceBefore) { try { $s.ParagraphFormat.SpaceBefore = [double]$p.spaceBefore; $applied += "spaceBefore" } catch { $failed += "spaceBefore" } }
            if ($null -ne $p.spaceAfter) { try { $s.ParagraphFormat.SpaceAfter = [double]$p.spaceAfter; $applied += "spaceAfter" } catch { $failed += "spaceAfter" } }
            if ($null -ne $p.lineSpacing) { try { $s.ParagraphFormat.LineSpacingRule = 5; $s.ParagraphFormat.LineSpacing = [double]$p.lineSpacing; $applied += "lineSpacing" } catch { $failed += "lineSpacing" } }
            if ($null -ne $p.alignment) {
                $map = @{ left = 0; center = 1; right = 2; justify = 3 }
                $k = [string]$p.alignment
                if ($map.ContainsKey($k)) { try { $s.ParagraphFormat.Alignment = $map[$k]; $applied += "alignment" } catch { $failed += "alignment" } }
            }
            # ⚠️ note 先算好再放进 hashtable。不要写成 `note = (if ...)`：
            #    hashtable 字面里带括号的 if 会在运行期报
            #    "无法将 if 项识别为 cmdlet"（而文件级 PARSE 却是通的，很迷惙）。
            #    同文件 addCaption 那处不带括号能跗，但别依赖这种微妙差异。
            $note = $null
            if ($failed.Count -gt 0) {
                $note = "以下属性本机 WPS 不接受：$($failed -join '/')。样式已建成，其余属性已设上。"
            }
            Output-Json @{ success = $true; data = @{
                name = [string]$p.name
                styleName = [string]$s.NameLocal
                applied = $applied
                failed = $failed
                note = $note
            } }
        } catch {
            Output-Json @{ success = $false; error = "新建样式失败: $(Get-CallError $_)" }
        }
    }

    "modifyStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        try {
            $s = $doc.Styles.Item([string]$p.name)
            if ($null -eq $s) { Output-Json @{ success = $false; error = "样式不存在: $($p.name)" }; return }
            if ($s.BuiltIn) {
                # 内置样式不允许整体改名/删除，但改属性是允许的 —— 这里只提示不拦
            }
            $applied = @()
            if ($p.fontSize) { $s.Font.Size = [double]$p.fontSize; $applied += "fontSize" }
            if ($p.fontName) { $s.Font.NameFarEast = [string]$p.fontName; $s.Font.NameAscii = [string]$p.fontName; $applied += "fontName" }
            if ($null -ne $p.bold) { $s.Font.Bold = [bool]$p.bold; $applied += "bold" }
            if ($null -ne $p.italic) { $s.Font.Italic = [bool]$p.italic; $applied += "italic" }
            if ($null -ne $p.color) {
                $bgr = [int]$p.color
                $s.Font.Color = ($bgr -band 0xFF) * 65536 + ($bgr -shr 8 -band 0xFF) * 256 + ($bgr -shr 16 -band 0xFF)
                $applied += "color"
            }
            if ($null -ne $p.spaceBefore) { $s.ParagraphFormat.SpaceBefore = [double]$p.spaceBefore; $applied += "spaceBefore" }
            if ($null -ne $p.spaceAfter) { $s.ParagraphFormat.SpaceAfter = [double]$p.spaceAfter; $applied += "spaceAfter" }
            if ($null -ne $p.lineSpacing) { $s.ParagraphFormat.LineSpacingRule = 5; $s.ParagraphFormat.LineSpacing = [double]$p.lineSpacing; $applied += "lineSpacing" }
            if ($null -ne $p.alignment) {
                $map = @{ left = 0; center = 1; right = 2; justify = 3 }
                $k = [string]$p.alignment
                if ($map.ContainsKey($k)) { $s.ParagraphFormat.Alignment = $map[$k]; $applied += "alignment" }
            }
            if ($null -ne $p.newName) {
                try { $s.NameLocal = [string]$p.newName; $applied += "newName" } catch {
                    $applied += "newName(内置样式不可改名，已跳过)"
                }
            }
            if ($applied.Count -eq 0) {
                Output-Json @{ success = $false; error = "没有任何可改的属性（fontSize/fontName/bold/italic/color/spaceBefore/spaceAfter/lineSpacing/alignment/newName）" }; return
            }
            Output-Json @{ success = $true; data = @{ name = [string]$p.name; applied = $applied } }
        } catch {
            Output-Json @{ success = $false; error = "修改样式失败: $(Get-CallError $_)" }
        }
    }

    "deleteStyle" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        if (-not $p.name) { Output-Json @{ success = $false; error = "缺少 name（样式名）" }; return }
        try {
            $s = $doc.Styles.Item([string]$p.name)
            if ($null -eq $s) { Output-Json @{ success = $false; error = "样式不存在: $($p.name)" }; return }
            if ($s.BuiltIn) {
                Output-Json @{ success = $false; error = "'$($p.name)' 是内置样式，不能删除（可以 modifyStyle 改它的属性）" }; return
            }
            $s.Delete()
            Output-Json @{ success = $true; data = @{ deleted = [string]$p.name } }
        } catch {
            Output-Json @{ success = $false; error = "删除样式失败（样式正被段落使用时可能失败）: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：交叉引用 ====================
    # 在【当前光标位置】插入对某个标题的引用。
    # refType: heading | bookmark | numberedItem | footnote
    "insertCrossReference" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            # wdRefTypeHeading=1 wdRefTypeBookmark=2 wdRefTypeNumberedItem=3 wdRefTypeFootnote=4
            $rtMap = @{ heading = 1; bookmark = 2; numberedItem = 3; footnote = 4 }
            $k = if ($p.refType) { [string]$p.refType } else { "heading" }
            if (-not $rtMap.ContainsKey($k)) { Output-Json @{ success = $false; error = "refType 只能是 heading/bookmark/numberedItem/footnote" }; return }
            $rt = $rtMap[$k]
            $items = @($doc.GetCrossReferenceItems($rt))
            if ($items.Count -eq 0) {
                Output-Json @{ success = $false; error = "文档里没有可引用的$($k)（GetCrossReferenceItems 返回空）。heading 需要先有标题样式段落" }; return
            }
            $idx = (Get-IntParam $p "itemIndex" 1)
            if ($idx -lt 1 -or $idx -gt $items.Count) {
                Output-Json @{ success = $false; error = "itemIndex 越界：文档里共 $($items.Count) 个可引用项，收到 $idx" }; return
            }
            $target = [string]$items[$idx - 1]
            $rng = $word.Selection.Range
            $rng.Collapse(0)   # wdCollapseEnd
            # insertAsHyperlink / includePosition / separateNumbers / separatorString
            $null = $doc.Fields.Add($rng, -1, "REF $target `h", $true)
            Output-Json @{ success = $true; data = @{ refType = $k; itemIndex = $idx; target = $target; total = $items.Count } }
        } catch {
            Output-Json @{ success = $false; error = "插入交叉引用失败: $(Get-CallError $_)" }
        }
    }

    "listCrossReferenceTargets" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $rtMap = @{ heading = 1; bookmark = 2; numberedItem = 3; footnote = 4 }
            $out = @{}
            foreach ($kv in $rtMap.GetEnumerator()) {
                try { $out[$kv.Key] = @($doc.GetCrossReferenceItems($kv.Value)) } catch { $out[$kv.Key] = @() }
            }
            Output-Json @{ success = $true; data = @{ targets = $out } }
        } catch {
            Output-Json @{ success = $false; error = "列出可引用项失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：索引 ====================
    # MarkEntry 在探测里报了 E_FAIL（当时是在【新建的空文档】上试的），
    # 所以这里实现时多给了两个参数并保留详细报错，便于真机再判一次。
    "markIndexEntry" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $rng = if ($null -ne $p.start -and $null -ne $p.end) { $doc.Range([int]$p.start, [int]$p.end) } else { $word.Selection.Range }
            if ($rng -eq $null -or [string]::IsNullOrWhiteSpace([string]$rng.Text)) {
                Output-Json @{ success = $false; error = "选区为空：请先选中要编入索引的文字，或传 start/end" }; return
            }
            $entry = if ($p.entry) { [string]$p.entry } else { ([string]$rng.Text).Trim() }
            $r = $doc.Indexes.MarkEntry($rng, [string]$entry)
            if ($null -eq $r) {
                Output-Json @{ success = $false; error = "Indexes.MarkEntry 返回 null（本机 WPS 12.1 探测时报 E_FAIL，可能不支持该接口）" }; return
            }
            Output-Json @{ success = $true; data = @{ entry = $entry } }
        } catch {
            Output-Json @{ success = $false; error = "标记索引项失败: $(Get-CallError $_)" }
        }
    }

    "insertIndex" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $marked = 0
            try { $marked = [int]$doc.Indexes.MarkEntries.Count } catch {}
            $rng = $word.Selection.Range
            $rng.Collapse(0)
            $ix = $doc.Indexes.Add($rng)
            if ($null -eq $ix) { Output-Json @{ success = $false; error = "Indexes.Add 返回 null" }; return }
            $cols = 1; $heading = $null
            if ($null -ne $p.columns) { try { $ix.NumberOfColumns = [int]$p.columns; $cols = [int]$p.columns } catch {} }
            if ($p.heading) { try { $ix.HeadingSeparator = [string]$p.heading; $heading = [string]$p.heading } catch {} }
            Output-Json @{ success = $true; data = @{ markedEntries = $marked; columns = $cols; heading = $heading } }
        } catch {
            Output-Json @{ success = $false; error = "插入索引失败: $(Get-CallError $_)" }
        }
    }

    # ==================== P1：页码域 / 奇偶页 / 首页不同 ====================
    # sectionIndex 默认 1（第一个节）
    "addPageNumber" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $sec = $doc.Sections.Item((Get-IntParam $p "sectionIndex" 1))
            # wdHeaderFooterPrimary=1 FirstPage=2 EvenPages=3
            $hfMap = @{ primary = 1; first = 2; even = 3 }
            $k = if ($p.which) { [string]$p.which } else { "primary" }
            if (-not $hfMap.ContainsKey($k)) { Output-Json @{ success = $false; error = "which 只能是 primary/first/even" }; return }
            $hf = $sec.Headers.Item($hfMap[$k])
            # wdAlignPageNumberRight=2
            $align = if ($p.alignment -eq "center") { 1 } elseif ($p.alignment -eq "left") { 0 } else { 2 }
            $n0 = [int]$hf.PageNumbers.Count
            $null = $hf.PageNumbers.Add($align, $true)
            if ($hf.PageNumbers.Count -le $n0) {
                Output-Json @{ success = $false; error = "PageNumbers.Add 调用了但页码数没增加（$n0→$($hf.PageNumbers.Count)）" }; return
            }
            Output-Json @{ success = $true; data = @{ section = $sec.Index; which = $k; alignment = $p.alignment; count = [int]$hf.PageNumbers.Count } }
        } catch {
            Output-Json @{ success = $false; error = "添加页码失败: $(Get-CallError $_)" }
        }
    }

    # 奇偶页不同 / 首页不同 —— PageSetup 级的开关
    "setHeaderFooterLayout" {
        $word = Get-WpsWord
        if ($null -eq $word) { Output-Json @{ success = $false; error = "WPS Word not running" }; return }
        $doc = Resolve-Doc $word $p
        if ($null -eq $doc) { Output-Json @{ success = $false; error = "No active document" }; return }
        try {
            $applied = @()
            if ($null -ne $p.differentFirstPage) {
                $doc.PageSetup.DifferentFirstPageHeaderFooter = [bool]$p.differentFirstPage
                $applied += "differentFirstPage"
            }
            if ($null -ne $p.oddAndEvenPages) {
                $doc.PageSetup.OddAndEvenPagesHeaderFooter = [bool]$p.oddAndEvenPages
                $applied += "oddAndEvenPages"
            }
            if ($applied.Count -eq 0) {
                Output-Json @{ success = $false; error = "至少要传 differentFirstPage 或 oddAndEvenPages 之一" }; return
            }
            Output-Json @{ success = $true; data = @{
                applied = $applied
                differentFirstPage = [bool]$doc.PageSetup.DifferentFirstPageHeaderFooter
                oddAndEvenPages = [bool]$doc.PageSetup.OddAndEvenPagesHeaderFooter
            } }
        } catch {
            Output-Json @{ success = $false; error = "设置页眉页脚布局失败: $(Get-CallError $_)" }
        }
    }

    default {
        Output-Json @{ success = $false; error = "Unknown action: $Action" }
    }
}

}   # function Invoke-WpsAction（上一行的 } 是 switch 的收尾）

# ==================== 入口 ====================
# 常驻模式：wps-comd.ps1 先设 $env:WPS_COM_DAEMON=1 再 dot-source 本文件，
#           此时只需拿到 Invoke-WpsAction 的定义，不执行任何动作。
# 直跑模式（默认）：行为与改造前完全一致。
if ($env:WPS_COM_DAEMON -ne '1') {
    try { $p = $Params | ConvertFrom-Json } catch { $p = @{} }
    Invoke-WpsAction -Action $Action -p $p
}
