/**
 * Input: WPS 加载项回调与用户交互请求
 * Output: 状态展示与日志输出
 * Pos: Windows WPS 加载项脚本入口。一旦我被修改，请更新我的头部注释，以及所属文件夹的md。
 * MCP 工具 - COM 桥接版
 * MCP Server 通过 PowerShell COM 接口直接操作 WPS
 * @author 老王
 */

// ==================== 全局状态 ====================
var addonStatus = {
    initialized: false
};

var LINE_TOP = '┌─ MCP 工具 ─────────────────';
var LINE_BOT = '└────────────────────────────';

// ==================== 工具函数 ====================

// 汇总当前活动文档信息（Word / Excel / PPT 通用）
function getActiveDocInfo() {
    if (typeof Application === 'undefined') { return null; }
    try {
        if (Application.ActiveWorkbook) {
            var wb = Application.ActiveWorkbook;
            return { name: wb.Name, type: 'WPS 表格', extra: '工作表数: ' + wb.Sheets.Count };
        }
    } catch (e) {}
    try {
        if (Application.ActivePresentation) {
            var pres = Application.ActivePresentation;
            return { name: pres.Name, type: 'WPS 演示', extra: '幻灯片数: ' + pres.Slides.Count };
        }
    } catch (e) {}
    try {
        if (Application.ActiveDocument) {
            var d = Application.ActiveDocument;
            var extra = '';
            try { extra = '字数: ' + d.Words.Count; } catch (e2) {}
            return { name: d.Name, type: 'WPS 文字', extra: extra };
        }
    } catch (e) {}
    return null;
}

// ==================== ribbon.xml 回调函数 ====================

function OnAddinLoad(ribbonUI) {
    console.log('[MCP工具] 加载项已加载 - COM 桥接模式');
    addonStatus.initialized = true;
}

// 连接状态：一屏展示桥接与当前文档信息
function ribbonCheckStatus(control) {
    var lines = [];
    lines.push(LINE_TOP);
    lines.push('  连接状态 : ' + (addonStatus.initialized ? '● 已就绪' : '○ 未就绪'));
    lines.push('  通信模式 : PowerShell COM 桥接');
    lines.push('  桥接说明 : MCP Server 通过 COM 接口直接操作 WPS');
    lines.push(LINE_BOT);

    try {
        if (typeof Application !== 'undefined') {
            lines.push('');
            lines.push('当前应用 : ' + (Application.Name || 'WPS Office'));
            var doc = getActiveDocInfo();
            if (doc) {
                lines.push('当前文档 : ' + doc.name + '（' + doc.type + '）');
                if (doc.extra) { lines.push('           ' + doc.extra); }
            } else {
                lines.push('当前文档 : 无');
            }
        }
    } catch (e) {
        lines.push('');
        lines.push('WPS 信息获取失败: ' + e.message);
    }
    alert(lines.join('\n'));
}

// 显示窗口：将 WPS 窗口置于前台
function ribbonShowWindow(control) {
    var lines = [];
    lines.push(LINE_TOP);
    try {
        if (typeof Application === 'undefined') {
            lines.push('  ⚠ 无法获取 WPS 应用对象');
        } else {
            try { Application.Visible = true; } catch (e1) {}
            try {
                if (Application.WindowState === 2) { Application.WindowState = 1; }
            } catch (e3) {}
            var activated = false;
            try {
                if (Application.Activate) { Application.Activate(); activated = true; }
                else if (Application.ActiveWindow && Application.ActiveWindow.Activate) { Application.ActiveWindow.Activate(); activated = true; }
            } catch (e2) {}
            lines.push(activated ? '  ● 已尝试将 WPS 窗口置于前台' : '  ⚠ 窗口激活接口不可用，已恢复可见');
        }
    } catch (e) {
        lines.push('  ⚠ 出错: ' + e.message);
    }
    lines.push(LINE_BOT);
    alert(lines.join('\n'));
}

// 文档概览：汇总当前文档信息
function ribbonDocInfo(control) {
    var lines = [];
    lines.push(LINE_TOP);
    try {
        var doc = getActiveDocInfo();
        if (doc) {
            lines.push('  名称 : ' + doc.name);
            lines.push('  类型 : ' + doc.type);
            if (doc.extra) { lines.push('  ' + doc.extra); }
        } else {
            lines.push('  当前没有打开的文档');
        }
    } catch (e) {
        lines.push('  ⚠ 读取失败: ' + e.message);
    }
    lines.push(LINE_BOT);
    alert(lines.join('\n'));
}

// 使用说明
function ribbonHelp(control) {
    var lines = [];
    lines.push(LINE_TOP);
    lines.push('  本加载项为 MCP 工具集提供功能区入口。');
    lines.push('  AI 助手通过 MCP 协议调用 PowerShell');
    lines.push('  COM 桥接，直接操控 WPS 文档。');
    lines.push('');
    lines.push('  · 连接状态 —— 查看桥接与当前文档');
    lines.push('  · 显示窗口 —— 把 WPS 置于前台');
    lines.push('  · 文档概览 —— 汇总当前文档信息');
    lines.push('');
    lines.push('  提示：MCP 操作后若窗口看不见，');
    lines.push('  点「显示窗口」即可唤起。');
    lines.push(LINE_BOT);
    alert(lines.join('\n'));
}

// ==================== WPS 生命周期函数 ====================

function OnAddinNotify(params) {
    console.log('[MCP工具] 收到通知:', JSON.stringify(params));
}

// ==================== 初始化 ====================
console.log('[MCP工具] main.js 已加载 - COM 桥接模式');
