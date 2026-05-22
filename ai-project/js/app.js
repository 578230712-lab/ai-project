// ==================== TOAST ====================
function showToast(msg, type) {
    var t = document.getElementById('toast');
    if (!t) return;
    var colors = { success: 'bg-emerald-500/90', error: 'bg-red-500/90', info: 'bg-accent-500/90' };
    t.className = 'fixed top-4 right-4 z-[100] px-5 py-3 rounded-xl text-white text-sm ' + (colors[type] || colors.info) + ' toast';
    t.textContent = msg;
    t.style.display = '';
    setTimeout(function () { t.style.display = 'none'; }, 3000);
}

// ==================== REVEAL ON SCROLL (首页) ====================
function initReveal() {
    var reveals = document.querySelectorAll('.reveal');
    if (!reveals.length) return;
    function check() {
        reveals.forEach(function (el) {
            var top = el.getBoundingClientRect().top;
            if (top < window.innerHeight - 100) el.classList.add('active');
        });
    }
    window.addEventListener('scroll', check);
    check();
}

// ==================== CONFIG PERSISTENCE ====================
function restoreConfig() {
    var wh = localStorage.getItem('yzj_webhook') || '';
    var doc = localStorage.getItem('tdoc_url') || '';
    var webhookInput = document.getElementById('webhookInput');
    var docUrlInput = document.getElementById('docUrlInput');
    if (webhookInput) webhookInput.value = wh;
    if (docUrlInput) docUrlInput.value = doc;
    if (wh) {
        var testBtn = document.getElementById('testWebhookBtn');
        var statusEl = document.getElementById('webhookStatus');
        if (testBtn) testBtn.disabled = false;
        if (statusEl) statusEl.innerHTML = '<span class="text-emerald-400">✅ 已保存</span>';
        updateStep(2, 'done');
    }
    if (doc) {
        embedDoc();
        updateStep(1, 'done');
    }
}

// ==================== STEP INDICATOR ====================
function updateStep(num, state) {
    var el = document.getElementById('step-indicator-' + num);
    if (!el) return;
    el.className = 'flex items-center gap-2 px-4 py-2 rounded-full border text-sm';
    if (state === 'active') el.classList.add('step-active');
    else if (state === 'done') el.classList.add('step-done');
    var dot = el.querySelector('span:first-child');
    if (dot && state === 'done') {
        dot.className = 'w-6 h-6 rounded-full bg-emerald-500 text-white text-xs flex items-center justify-center font-bold';
        dot.textContent = '✓';
    }
}

// ==================== STEP 1: EMBED DOCUMENT ====================
function embedDoc() {
    var docUrlInput = document.getElementById('docUrlInput');
    if (!docUrlInput) return;
    var url = docUrlInput.value.trim();
    if (!url) { showToast('请先粘贴腾讯文档链接', 'error'); return; }
    if (url.indexOf('docs.qq.com') === -1) { showToast('请输入有效的腾讯文档链接', 'error'); return; }

    var embedUrl = url;
    if (embedUrl.indexOf('?') > -1) embedUrl += '&nlc=1';
    else embedUrl += '?nlc=1';

    var iframe = document.getElementById('docIframe');
    var preview = document.getElementById('docPreview');
    var statusEl = document.getElementById('embedStatus');
    if (iframe) iframe.src = embedUrl;
    if (preview) preview.style.display = '';
    if (statusEl) {
        statusEl.className = 'text-xs text-slate-500 mb-4';
        statusEl.textContent = '✅ 文档已嵌入。如果显示登录页，请确保已将文档设为「获得链接的任何人可查看」。';
        statusEl.style.display = '';
    }

    localStorage.setItem('tdoc_url', url);
    updateStep(1, 'done');
    showToast('文档嵌入成功', 'success');
}

// ==================== STEP 2: WEBHOOK ====================
function saveWebhook() {
    var input = document.getElementById('webhookInput');
    if (!input) return;
    var url = input.value.trim();
    if (!url) { showToast('请粘贴 Webhook 地址', 'error'); return; }
    if (url.indexOf('yunzhijia.com') === -1) { showToast('请输入有效的云之家 Webhook 地址', 'error'); return; }
    localStorage.setItem('yzj_webhook', url);
    var testBtn = document.getElementById('testWebhookBtn');
    var statusEl = document.getElementById('webhookStatus');
    if (testBtn) testBtn.disabled = false;
    if (statusEl) statusEl.innerHTML = '<span class="text-emerald-400">✅ 已保存</span>';
    updateStep(2, 'done');
    input.classList.add('config-saved');
    setTimeout(function () { input.classList.remove('config-saved'); }, 600);
    showToast('Webhook 已保存', 'success');
}

async function testWebhook() {
    var wh = localStorage.getItem('yzj_webhook');
    if (!wh) { showToast('请先保存 Webhook 地址', 'error'); return; }
    showToast('正在发送测试消息...', 'info');
    try {
        var resp = await fetch(wh, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                msgtype: 'text',
                content: '✅ AI 项目监理系统测试消息\n\n云之家机器人配置成功！系统将自动推送逾期任务提醒。'
            })
        });
        var data = await resp.json();
        if (data.success || data.errcode === 0) {
            showToast('测试消息发送成功！请查看群聊', 'success');
        } else {
            showToast('发送失败: ' + (data.errmsg || JSON.stringify(data)), 'error');
        }
    } catch (e) {
        var msg = e.message;
        if (msg.indexOf('Failed to fetch') > -1 || msg.indexOf('NetworkError') > -1) {
            showToast('CORS 跨域拦截！请通过本地服务器打开页面（双击 start-server.bat）', 'error');
        } else {
            showToast('网络错误: ' + msg, 'error');
        }
    }
}

// ==================== STEP 3: EXCEL PARSING ====================
var parsedTasks = [];
var overdueTasks = [];

function handleFile(event) {
    var file = event.target.files[0];
    if (!file) return;
    var statusEl = document.getElementById('uploadStatus');
    if (statusEl) statusEl.innerHTML = '<span class="text-slate-400">⏳ 正在解析文件...</span>';

    var reader = new FileReader();
    reader.onload = function (e) {
        try {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array' });
            var sheetName = workbook.SheetNames[0];
            var sheet = workbook.Sheets[sheetName];
            var json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

            if (json.length < 2) {
                if (statusEl) statusEl.innerHTML = '<span class="text-red-400">❌ 文件为空或只有表头</span>';
                return;
            }
            parseTasks(json);
        } catch (err) {
            if (statusEl) statusEl.innerHTML = '<span class="text-red-400">❌ 解析失败: ' + err.message + '</span>';
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseTasks(rows) {
    var headers = rows[0];
    var colMap = {};
    var keywords = {
        name: ['任务名称', '任务名', '名称', '标题', '需求描述', '问题/需求描述', '问题', 'name', 'task'],
        person: ['负责人', '责任人', '处理人', '提出人', 'owner', 'assignee'],
        status: ['状态', '进度', 'status', 'state'],
        deadline: ['截止日期', '截止时间', '预计解决时间', '解决时间', 'deadline', 'due', '日期'],
        priority: ['优先级', '重要程度', 'priority']
    };

    for (var i = 0; i < headers.length; i++) {
        var h = String(headers[i]).trim().toLowerCase();
        for (var key in keywords) {
            if (!colMap[key] && keywords[key].some(function (kw) { return h.indexOf(kw.toLowerCase()) > -1; })) {
                colMap[key] = i;
            }
        }
    }

    var statusEl = document.getElementById('uploadStatus');

    if (colMap['name'] === undefined) {
        if (statusEl) statusEl.innerHTML = '<span class="text-red-400">❌ 未找到「任务名称」列。请确保表格包含标准表头。</span>';
        return;
    }
    if (colMap['deadline'] === undefined) {
        if (statusEl) statusEl.innerHTML = '<span class="text-red-400">❌ 未找到「截止日期」列。</span>';
        return;
    }

    var today = new Date(); today.setHours(0, 0, 0, 0);
    parsedTasks = [];
    var statusDoneWords = ['已完成', '完成', 'done', 'closed', '已关闭', '结束'];

    for (var r = 1; r < rows.length; r++) {
        var row = rows[r];
        var name = String(row[colMap['name']] || '').trim();
        if (!name) continue;

        var person = colMap['person'] !== undefined ? String(row[colMap['person']] || '').trim() : '';
        var status = colMap['status'] !== undefined ? String(row[colMap['status']] || '').trim() : '';
        var priority = colMap['priority'] !== undefined ? String(row[colMap['priority']] || '').trim() : '';
        var deadlineRaw = colMap['deadline'] !== undefined ? row[colMap['deadline']] : '';

        var isDone = statusDoneWords.some(function (w) { return status.toLowerCase().indexOf(w.toLowerCase()) > -1; });

        var deadline = null;
        var deadlineStr = String(deadlineRaw).trim();

        if (deadlineStr) {
            if (/^\d{5}$/.test(deadlineStr)) {
                var excelEpoch = new Date(1899, 11, 30);
                deadline = new Date(excelEpoch.getTime() + parseInt(deadlineStr) * 86400000);
            } else {
                var d = new Date(deadlineStr);
                if (!isNaN(d.getTime())) deadline = d;
                else {
                    var m = deadlineStr.match(/(\d{4})[\/\-. ](\d{1,2})[\/\-. ](\d{1,2})/);
                    if (m) deadline = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
                }
            }
        }

        var isOverdue = false;
        if (deadline && !isDone) {
            var dl = new Date(deadline); dl.setHours(0, 0, 0, 0);
            if (dl < today) isOverdue = true;
        }

        parsedTasks.push({ name: name, person: person, status: status, priority: priority, deadline: deadline, deadlineStr: deadlineStr, isDone: isDone, isOverdue: isOverdue });
    }

    overdueTasks = parsedTasks.filter(function (t) { return t.isOverdue; });
    var totalTasks = parsedTasks.length;
    var doneTasks = parsedTasks.filter(function (t) { return t.isDone; }).length;
    var pendingTasks = totalTasks - doneTasks;

    if (statusEl) statusEl.innerHTML = '<span class="text-emerald-400">✅ 成功解析 ' + totalTasks + ' 条任务</span>';

    // Summary cards
    var summaryCards = document.getElementById('summaryCards');
    if (summaryCards) {
        summaryCards.innerHTML =
            '<div class="glass-card p-4 text-center"><div class="text-2xl font-bold text-white">' + totalTasks + '</div><div class="text-xs text-slate-500 mt-1">总任务数</div></div>' +
            '<div class="glass-card p-4 text-center"><div class="text-2xl font-bold text-emerald-400">' + doneTasks + '</div><div class="text-xs text-slate-500 mt-1">已完成</div></div>' +
            '<div class="glass-card p-4 text-center"><div class="text-2xl font-bold text-amber-400">' + pendingTasks + '</div><div class="text-xs text-slate-500 mt-1">进行中/未开始</div></div>' +
            '<div class="glass-card p-4 text-center"><div class="text-2xl font-bold text-red-400">' + overdueTasks.length + '</div><div class="text-xs text-slate-500 mt-1">⚠️ 已逾期</div></div>';
    }

    // Overdue table
    var overdueTable = document.getElementById('overdueTable');
    var pushBtn = document.getElementById('pushBtn');
    var resultsArea = document.getElementById('resultsArea');

    if (overdueTasks.length > 0) {
        var formatDate = function (d) {
            if (!d) return '-';
            return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
        };
        var tableRows = overdueTasks.map(function (t) {
            var priClass = (t.priority.toLowerCase().indexOf('重要') > -1 || t.priority.toLowerCase().indexOf('高') > -1) ? 'text-orange-400' : 'text-slate-400';
            return '<tr class="border-b border-white/5 hover:bg-white/5">' +
                '<td class="py-3 px-4 text-sm text-white">' + escHtml(t.name) + '</td>' +
                '<td class="py-3 px-4 text-sm text-slate-300">' + escHtml(t.person) || '-' + '</td>' +
                '<td class="py-3 px-4 text-sm text-slate-400">' + escHtml(t.status) || '-' + '</td>' +
                '<td class="py-3 px-4 text-sm text-red-400 font-medium">' + formatDate(t.deadline) + '</td>' +
                '<td class="py-3 px-4 text-sm ' + priClass + '">' + escHtml(t.priority) || '-' + '</td>' +
                '</tr>';
        }).join('');
        if (overdueTable) {
            overdueTable.innerHTML =
                '<h4 class="text-white font-semibold mb-3">⚠️ 逾期任务列表（' + overdueTasks.length + '条）</h4>' +
                '<table class="w-full text-left"><thead><tr class="border-b border-white/10 text-xs text-slate-500">' +
                '<th class="py-2 px-4 font-medium">任务名称</th><th class="py-2 px-4 font-medium">负责人</th><th class="py-2 px-4 font-medium">状态</th><th class="py-2 px-4 font-medium">截止日期</th><th class="py-2 px-4 font-medium">优先级</th>' +
                '</tr></thead><tbody>' + tableRows + '</tbody></table>';
        }
        if (pushBtn) pushBtn.disabled = false;
    } else {
        if (overdueTable) overdueTable.innerHTML = '<div class="text-emerald-400 text-sm py-4">🎉 没有逾期任务，一切正常！</div>';
        if (pushBtn) pushBtn.disabled = true;
    }

    if (resultsArea) resultsArea.style.display = '';
    updateStep(3, 'done');
}

function escHtml(s) {
    var d = document.createElement('div');
    d.textContent = s;
    return d.innerHTML;
}

// ==================== PUSH TO YUNZHIJIA ====================
async function pushOverdue() {
    var wh = localStorage.getItem('yzj_webhook');
    if (!wh) { showToast('请先在第二步配置云之家 Webhook', 'error'); return; }
    if (overdueTasks.length === 0) { showToast('没有逾期任务需要推送', 'info'); return; }

    var pushStatus = document.getElementById('pushStatus');
    if (pushStatus) pushStatus.innerHTML = '<span class="text-slate-400">⏳ 正在推送...</span>';

    var today = new Date();
    var dateStr = today.getFullYear() + '-' + String(today.getMonth() + 1).padStart(2, '0') + '-' + String(today.getDate()).padStart(2, '0');
    var formatDate = function (d) { return d ? d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') : '-'; };

    var atMentions = '';
    var mentioned = {};
    overdueTasks.forEach(function (t) { if (t.person) mentioned[t.person] = true; });
    var names = Object.keys(mentioned);
    if (names.length > 0) atMentions = '\n' + names.map(function (p) { return '@' + p; }).join(' ');

    var content = '📋 **项目逾期任务提醒** (' + dateStr + ')\n\n';
    content += '共 **' + overdueTasks.length + '** 条任务已逾期，请及时处理：\n\n';
    overdueTasks.forEach(function (t, i) {
        content += (i + 1) + '. ' + t.name + '\n';
        content += '   负责人: ' + (t.person || '未指定') + ' | 截止: ' + formatDate(t.deadline) + ' | 状态: ' + (t.status || '-') + '\n';
    });
    content += '\n请相关责任人尽快更新任务状态。' + atMentions;

    try {
        var resp = await fetch(wh, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ msgtype: 'text', content: content })
        });
        var data = await resp.json();
        if (data.success || data.errcode === 0) {
            if (pushStatus) pushStatus.innerHTML = '<span class="text-emerald-400">✅ 推送成功！请查看云之家群聊。共推送 ' + overdueTasks.length + ' 条逾期任务。</span>';
            showToast('推送成功！', 'success');
        } else {
            if (pushStatus) pushStatus.innerHTML = '<span class="text-red-400">❌ 推送失败: ' + (data.errmsg || JSON.stringify(data)) + '</span>';
        }
    } catch (e) {
        var msg = e.message;
        if (msg.indexOf('Failed to fetch') > -1 || msg.indexOf('NetworkError') > -1) {
            if (pushStatus) pushStatus.innerHTML = '<span class="text-orange-400">⚠️ CORS 跨域拦截！请通过本地服务器打开页面（双击 start-server.bat）</span>';
        } else {
            if (pushStatus) pushStatus.innerHTML = '<span class="text-red-400">❌ 网络错误: ' + msg + '</span>';
        }
    }
}

// ==================== DOWNLOAD SCRIPT ====================
function downloadScript() {
    var wh = localStorage.getItem('yzj_webhook') || 'YOUR_WEBHOOK_URL_HERE';
    var script = '#!/usr/bin/env python3\n# -*- coding: utf-8 -*-\n"""\nAI 项目监理系统 - 自动监控脚本\n用法: python monitor.py\n建议通过 Windows 计划任务 / Linux cron 每天定时运行\n"""\nimport json, urllib.request, datetime, sys\n\n# ========== 配置区 ==========\nWEBHOOK_URL = "' + wh + '"\nEXCEL_PATH = r"C:\\Users\\你的用户名\\Desktop\\空白表格.xlsx"  # 修改为你的 Excel 路径\n\n# ========== 读取 Excel ==========\ntry:\n    import openpyxl\nexcept ImportError:\n    print("请先安装 openpyxl: pip install openpyxl")\n    sys.exit(1)\n\nwb = openpyxl.load_workbook(EXCEL_PATH)\nws = wb.active\n\nheaders = [cell.value for cell in ws[1]]\ncol_map = {}\nfor i, h in enumerate(headers):\n    h_lower = str(h).lower() if h else \'\'\n    if \'任务\' in h_lower or \'名称\' in h_lower: col_map[\'name\'] = i\n    elif \'负责人\' in h_lower or \'处理人\' in h_lower: col_map[\'person\'] = i\n    elif \'状态\' in h_lower: col_map[\'status\'] = i\n    elif \'截止\' in h_lower or \'deadline\' in h_lower: col_map[\'deadline\'] = i\n    elif \'优先\' in h_lower: col_map[\'priority\'] = i\n\ntoday = datetime.date.today()\noverdue = []\n\nfor row in ws.iter_rows(min_row=2, values_only=True):\n    name = str(row[col_map[\'name\']]).strip() if \'name\' in col_map and row[col_map[\'name\']] else \'\'\n    if not name: continue\n    person = str(row[col_map[\'person\']]).strip() if \'person\' in col_map and row[col_map[\'person\']] else \'\'\n    status = str(row[col_map[\'status\']]).strip() if \'status\' in col_map and row[col_map[\'status\']] else \'\'\n    is_done = any(w in status for w in [\'已完成\', \'完成\', \'done\', \'已关闭\'])\n\n    deadline = None\n    if \'deadline\' in col_map and row[col_map[\'deadline\']]:\n        d = row[col_map[\'deadline\']]\n        if hasattr(d, \'date\'): deadline = d.date()\n        elif isinstance(d, str):\n            for fmt in [\'%Y-%m-%d\', \'%Y/%m/%d\', \'%Y.%m.%d\']:\n                try:\n                    deadline = datetime.datetime.strptime(d.strip(), fmt).date()\n                    break\n                except: pass\n\n    if deadline and not is_done and deadline < today:\n        overdue.append({\'name\': name, \'person\': person, \'status\': status, \'deadline\': str(deadline)})\n\nif not overdue:\n    print(f"[{today}] 没有逾期任务")\n    sys.exit(0)\n\ntoday_str = str(today)\ncontent = f"📋 **项目逾期任务提醒** ({today_str})\\n\\n共 **{len(overdue)}** 条任务已逾期，请及时处理：\\n\\n"\nmentioned = set(t.get(\'person\') for t in overdue if t.get(\'person\'))\nfor i, t in enumerate(overdue):\n    content += f"{i+1}. {t[\'name\']}\\n   负责人: {t[\'person\'] or \'未指定\'} | 截止: {t[\'deadline\']} | 状态: {t[\'status\']}\\n"\nif mentioned:\n    content += "\\n" + " ".join(f"@{p}" for p in mentioned)\n\ndata = json.dumps({"msgtype": "text", "text": {"content": content}}).encode()\nreq = urllib.request.Request(WEBHOOK_URL, data=data, headers={\'Content-Type\': \'application/json\'})\ntry:\n    resp = json.loads(urllib.request.urlopen(req).read())\n    if resp.get(\'success\') or resp.get(\'errcode\') == 0:\n        print(f"[{today}] 推送成功: {len(overdue)} 条逾期任务")\n    else:\n        print(f"推送失败: {resp}")\nexcept Exception as e:\n    print(f"网络错误: {e}")\n';

    var blob = new Blob([script], { type: 'text/plain' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'ai-monitor.py';
    a.click();
    URL.revokeObjectURL(url);
    showToast('脚本已下载！按注释修改 Excel 路径后即可使用', 'success');
}

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', function () {
    initReveal();
    restoreConfig();
});
