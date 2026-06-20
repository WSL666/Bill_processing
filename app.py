"""
账单处理 Web 应用
功能：上传 Excel → 删除"异常"列非空行 → 客户名校验 → 表格展示
      清洗结果 → 按店铺分类导出（今日账目处理）
      支持每日账目文件（含日期工作表）→ 重新计算费用 → 分类导出
"""
import atexit
import html
import io
import json
import math
import os
import re
import sys
import tempfile
import zipfile
from decimal import Decimal, ROUND_DOWN, ROUND_UP
import pandas as pd
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

HEAD_FREIGHT_RATE = 69.0
PACKING_FEE_PER = 2.0
LABEL_CHANGE_FEE_PER = 5.0

# ── 客户名 + 费率配置（JSON持久化）───────────────────

def _data_dir():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    d = os.path.join(base, '发货明细')
    os.makedirs(d, exist_ok=True)
    return d

def _customers_path():
    return os.path.join(_data_dir(), 'customers.json')

CUSTOMER_SET = set()
CONFIG = {"头程运费单价": 69.0, "打包费": 2.0, "换面单费": 5.0}

def load_customers():
    global CUSTOMER_SET, CONFIG, HEAD_FREIGHT_RATE, PACKING_FEE_PER, LABEL_CHANGE_FEE_PER
    path = _customers_path()
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict) and 'customers' in data:
                CUSTOMER_SET = set(data['customers'])
                if 'config' in data:
                    CONFIG.update(data['config'])
            else:
                # 兼容旧格式（纯数组）
                CUSTOMER_SET = set(data)
            print(f"[客户名] 加载 {len(CUSTOMER_SET)} 个")
        except Exception:
            CUSTOMER_SET = set()
    else:
        CUSTOMER_SET = set()
        save_customers()
        print("[客户名] 空名单，请通过前端添加")
    # 同步到全局变量
    HEAD_FREIGHT_RATE = float(CONFIG.get("头程运费单价", 69))
    PACKING_FEE_PER = float(CONFIG.get("打包费", 2))
    LABEL_CHANGE_FEE_PER = float(CONFIG.get("换面单费", 5))

def save_customers():
    _data_dir()
    with open(_customers_path(), 'w', encoding='utf-8') as f:
        json.dump({"customers": sorted(CUSTOMER_SET), "config": CONFIG}, f, ensure_ascii=False, indent=2)

def get_customer_list():
    return sorted(CUSTOMER_SET)

load_customers()

# ── CSS + JS ────────────────────────────────────────────

STYLE = """
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Microsoft YaHei',sans-serif;background:#f0f2f5;color:#333}
    .container{max-width:98%;margin:0 auto;padding:10px 20px}
    h1{text-align:center;margin-bottom:4px;color:#1a1a2e;font-size:24px}
    .subtitle{text-align:center;color:#999;margin-bottom:20px;font-size:13px}
    .card{background:#fff;border-radius:12px;padding:24px 28px;margin-bottom:20px;box-shadow:0 2px 12px rgba(0,0,0,0.05)}
    .row{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
    .file-label{display:inline-flex;align-items:center;gap:8px;padding:10px 24px;background:#4a90d9;color:#fff;border-radius:6px;cursor:pointer;font-size:15px;white-space:nowrap}
    .file-label:hover{background:#357abd}
    .file-name{color:#888;font-size:14px;max-width:600px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .file-name.ok{color:#333;font-weight:500}
    .btn{padding:10px 24px;border:none;border-radius:6px;font-size:15px;cursor:pointer;text-decoration:none;display:inline-block;white-space:nowrap;transition:background .2s}
    .btn-green{background:#52c41a;color:#fff}.btn-green:hover{background:#45a716}
    .btn:disabled{background:#ccc;cursor:not-allowed}
    .tip{color:#aaa;font-size:12px;margin-top:8px}
    .stats{display:flex;gap:14px;margin-bottom:20px;flex-wrap:wrap}
    .stat-card{flex:1;min-width:100px;background:#fff;border-radius:10px;padding:16px 18px;box-shadow:0 2px 8px rgba(0,0,0,0.04);text-align:center}
    .stat-num{font-size:28px;font-weight:700;color:#4a90d9}
    .stat-label{font-size:13px;color:#999;margin-top:2px}
    .table-wrap{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,0.05);overflow:hidden;margin-bottom:20px}
    .table-head{padding:14px 20px;border-bottom:1px solid #f0f0f0;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
    .table-head h2{font-size:16px;color:#1a1a2e}
    .table-tools{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
    .search-input{padding:7px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:13px;width:200px;outline:none}
    .search-input:focus{border-color:#4a90d9}
    .btn-sm{padding:7px 16px;border:none;border-radius:5px;font-size:13px;cursor:pointer;text-decoration:none;display:inline-block;transition:background .2s}
    .btn-sm.blue{background:#4a90d9;color:#fff}.btn-sm.blue:hover{background:#357abd}
    .btn-sm.purple{background:#722ed1;color:#fff}.btn-sm.purple:hover{background:#531dab}
    .btn-sm.gray{background:#999;color:#fff}.btn-sm.gray:hover{background:#777}
    .scroll{overflow-x:auto;max-height:55vh;overflow-y:auto}
    table{width:100%;border-collapse:collapse;font-size:13px}
    thead th{background:#fafafa;padding:9px 10px;text-align:left;font-weight:600;color:#555;border-bottom:2px solid #e8e8e8;white-space:nowrap;position:sticky;top:0;z-index:2}
    tbody td{padding:7px 10px;border-bottom:1px solid #f5f5f5;white-space:nowrap}
    tbody tr:hover{background:#f8fafc}
    .red{background:#ffe0e0!important}.red:hover{background:#ffcccc!important}
    .orange{background:#fff3cd!important}.orange:hover{background:#ffe69c!important}
    .red.orange{background:#ffe0e0!important}.red.orange:hover{background:#ffcccc!important}
    .search-match,.search-match td{background:#f6ffed!important}.search-match:hover,.search-match:hover td{background:#d9f7be!important}
    .badge{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px}
    .badge-ok{background:#f6ffed;color:#52c41a;border:1px solid #b7eb8f}
    .badge-del{background:#fff2f0;color:#e74c3c;border:1px solid #ffccc7}
    .badge-warn{background:#fff7e6;color:#fa8c16;border:1px solid #ffd591}
    .legend{display:flex;gap:16px;margin-bottom:12px;font-size:13px;color:#888;flex-wrap:wrap}
    .legend-item{display:flex;align-items:center;gap:4px}
    .dot{width:12px;height:12px;border-radius:3px}
    .dot.r{background:#ffccc7}.dot.o{background:#ffd591}.dot.n{background:#fff;border:1px solid #d9d9d9}
    .divider{border:none;border-top:2px dashed #e8e8e8;margin:28px 0 20px}
    .error{background:#fff2f0;border:1px solid #ffccc7;border-radius:8px;padding:12px 20px;color:#e74c3c;margin-bottom:20px;font-size:14px}
    .footer{text-align:center;color:#ddd;font-size:12px;padding:20px}
    .loading{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(255,255,255,0.7);z-index:999;justify-content:center;align-items:center;flex-direction:column}
    .loading.on{display:flex}
    .spinner{width:40px;height:40px;border:3px solid #e8e8e8;border-top-color:#4a90d9;border-radius:50%;animation:spin .8s linear infinite;margin-bottom:12px}
    @keyframes spin{to{transform:rotate(360deg)}}
    .loading-text{color:#666;font-size:15px}
    .back-top{position:fixed;bottom:80px;right:30px;width:42px;height:42px;background:#4a90d9;color:#fff;border:none;border-radius:50%;font-size:20px;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,0.15);display:none;z-index:100}
    .back-top.on{display:block}
    @media(max-width:768px){.stats{flex-direction:column}.row{flex-direction:column;align-items:stretch}}
    /* 客户名单弹窗 */
    .modal-bg{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:1000;justify-content:center;align-items:center}
    .modal-bg.on{display:flex}
    .modal{background:#fff;border-radius:12px;width:520px;max-height:75vh;overflow:hidden;box-shadow:0 8px 32px rgba(0,0,0,0.2)}
    .modal-head{padding:18px 24px;border-bottom:1px solid #f0f0f0;display:flex;justify-content:space-between;align-items:center}
    .modal-head h3{font-size:18px;color:#1a1a2e}
    .modal-close{background:none;border:none;font-size:22px;cursor:pointer;color:#999;padding:4px 8px}
    .modal-close:hover{color:#333}
    .modal-body{padding:16px 24px;max-height:50vh;overflow-y:auto}
    .modal-foot{padding:14px 24px;border-top:1px solid #f0f0f0;display:flex;gap:10px}
    .modal-foot input{flex:1;padding:8px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;outline:none}
    .modal-foot input:focus{border-color:#4a90d9}
    .cust-row{display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #f5f5f5}
    .cust-row .name{font-size:14px;flex:1}
    .cust-row .acts{display:flex;gap:6px}
    .cust-row .acts button{border:none;background:none;cursor:pointer;font-size:13px;padding:2px 8px;border-radius:4px}
    .cust-row .acts .edit{color:#4a90d9}.cust-row .acts .edit:hover{background:#e6f7ff}
    .cust-row .acts .del{color:#e74c3c}.cust-row .acts .del:hover{background:#fff2f0}
    .cust-count{font-size:13px;color:#999;margin-bottom:8px}
    td.editable{background:#fff;cursor:text}td.editable:focus{background:#fffde0;outline:2px solid #4a90d9}
"""

JS = """
function escHtml(s) { var d=document.createElement('div'); d.textContent=s; return d.innerHTML; }
function onFileSelected(input) {
    var n = input.files.length ? input.files[0].name : '';
    var s = document.getElementById('fileName');
    var b = document.getElementById('submitBtn');
    if (n) { s.textContent = n; s.className = 'file-name ok'; b.disabled = false; }
    else { s.textContent = '未选择文件'; s.className = 'file-name'; b.disabled = true; }
}
function showLoading() {
    if (!document.getElementById('fileInput').files.length) { alert('请先选择文件'); return false; }
    var y=document.querySelector('[name=year]').value;
    var m=document.querySelector('[name=month]').value;
    var d=document.querySelector('[name=day]').value;
    if(!y||!m||!d){ alert('请填写完整的年月日'); return false; }
    document.getElementById('loading').classList.add('on'); return true;
}
function initSearch(tableId, searchId, btnId) {
    function doFilter() {
        var q = document.getElementById(searchId).value.toLowerCase().trim();
        document.querySelectorAll('#'+tableId+' tbody tr').forEach(function(tr) {
            if (!q) { tr.style.display = ''; tr.classList.remove('search-match'); return; }
            var found = false;
            // 含中文 → 词边界匹配（避免"周银燕"误匹配"周银燕2店"）
            // 不含中文 → 普通子串匹配（方便搜订单号、日期等）
            var hasChinese = /[\\u4e00-\\u9fff]/.test(q);
            tr.querySelectorAll('td').forEach(function(td) {
                if (found) return;
                var text = td.textContent.trim().toLowerCase();
                if (hasChinese) {
                    var idx = text.indexOf(q);
                    while (idx >= 0) {
                        var before = idx > 0 ? text[idx - 1] : '';
                        var after = idx + q.length < text.length ? text[idx + q.length] : '';
                        if (!/[a-z0-9\\u4e00-\\u9fff]/.test(before) && !/[a-z0-9\\u4e00-\\u9fff]/.test(after)) {
                            found = true; return;
                        }
                        idx = text.indexOf(q, idx + 1);
                    }
                } else {
                    if (text.indexOf(q) >= 0) { found = true; return; }
                }
            });
            if (found) {
                tr.style.display = '';
                tr.classList.add('search-match');
            } else {
                tr.style.display = 'none';
                tr.classList.remove('search-match');
            }
        });
    }
    document.getElementById(searchId).addEventListener('input', doFilter);
    document.getElementById(searchId).addEventListener('keydown', function(e) {
        if(e.key === 'Enter') { doFilter(); e.preventDefault(); }
    });
    if(btnId) document.getElementById(btnId).addEventListener('click', doFilter);
}
window.addEventListener('scroll', function() {
    var bt = document.getElementById('backTop');
    if (bt) bt.classList.toggle('on', window.scrollY > 300);
});
function doExport(url) {
    fetch('/api/export/check').then(r=>r.json()).then(d=>{
        var msg = d.exists ? '文件已存在，是否替换？' : '确定导出？';
        if(confirm(msg)){
            fetch(url).then(r2=>r2.json()).then(d2=>{
                if(d2.ok){alert(d2.msg);}
                else{alert('导出失败');}
            });
        }
    });
}
function delRow(btn) {
    var idx = parseInt(btn.getAttribute('data-idx'));
    if(isNaN(idx)) return;
    if(!confirm('确定删除该行吗？')) return;
    // 禁用所有删除按钮，防止并发删除
    var allBtns = document.querySelectorAll('#table1 tbody button');
    allBtns.forEach(function(b){ b.disabled = true; });
    fetch('/api/row/delete', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({idx:idx})})
    .then(r=>r.json()).then(d=>{
        if(d.ok){
            btn.closest('tr').remove();
            // 更新行号和 data-row / data-idx（只对有删除按钮的客户行重新编号）
            var custCounter = 0;
            document.querySelectorAll('#table1 tbody tr').forEach(function(tr,i){
                tr.querySelector('td').textContent = i + 1;
                var delBtn = tr.querySelector('button');
                tr.querySelectorAll('.editable').forEach(function(td){
                    td.setAttribute('data-row', custCounter);
                });
                if(delBtn){
                    delBtn.setAttribute('data-idx', custCounter);
                    delBtn.setAttribute('onclick', 'delRow(this)');
                    custCounter++;
                }
            });
            // 更新统计
            var stats=document.querySelectorAll('.stat-num');
            if(stats.length>=4){ stats[3].textContent = d.rows; }
        }else{alert('删除失败:'+(d.msg||''));}
        allBtns.forEach(function(b){ b.disabled = false; });
    }).catch(function(e){
        alert('网络错误，删除失败');
        allBtns.forEach(function(b){ b.disabled = false; });
    });
}
document.addEventListener('blur', function(e) {
    if(e.target.classList.contains('editable')) {
        var td = e.target;
        var oldVal = td.getAttribute('data-old') || '';
        var newVal = td.textContent.trim();
        if(newVal === oldVal) return;
        var idx = parseInt(td.getAttribute('data-row'));
        var col = td.getAttribute('data-col');
        fetch('/api/row/update', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({idx:idx, col:col, val:newVal})})
        .then(r=>r.json()).then(d=>{
            if(d.ok){ td.setAttribute('data-old', newVal); }
            else{ td.textContent = oldVal; alert('保存失败'); }
        });
    }
}, true);
document.addEventListener('focus', function(e) {
    if(e.target.classList.contains('editable')) {
        e.target.setAttribute('data-old', e.target.textContent.trim());
    }
}, true);
"""

# ── 页面模板 ─────────────────────────────────────

def page(title, body):
    return f"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0"><title>{title}</title>
<style>{STYLE}</style></head><body>
<div class="loading" id="loading"><div class="spinner"></div><div class="loading-text">正在处理文件...</div></div>
<div class="container"><h1>{title}</h1><p class="subtitle">上传 Excel &#8594; 清洗数据 + 分类导出</p><script>{JS}</script>{body}
<div class="footer">Bill Processing &middot; Flask + Pandas + openpyxl</div></div>
<button class="back-top" id="backTop" onclick="window.scrollTo({{top:0,behavior:'smooth'}})">&#9650;</button>
<div class="modal-bg" id="custModal">
<div class="modal">
<div class="modal-head"><h3>&#128100; 客户名单管理</h3><button class="modal-close" onclick="closeCustomerModal()">&times;</button></div>
<div class="modal-body" id="custList"><div style="text-align:center;color:#999;padding:20px">加载中...</div></div>
<div class="modal-foot" style="flex-direction:column;align-items:stretch;gap:8px">
<div style="display:flex;gap:8px;align-items:center;padding-bottom:8px;border-bottom:1px dashed #e8e8e8">
<span style="font-size:13px;color:#888;white-space:nowrap">&#128193; 批量导入：</span>
<input type="file" id="custFileInput" accept=".xlsx,.xls" onchange="uploadCustomerFile()" style="display:none">
<button class="btn-sm blue" onclick="document.getElementById('custFileInput').click()">选择Excel文件</button>
<span id="uploadStatus" style="font-size:12px;color:#999"></span>
</div>
<div id="addRows"><div style="display:flex;gap:8px;align-items:center"><input type="text" class="cust-new-input" placeholder="输入客户名" style="flex:1;padding:8px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;outline:none" onkeydown="if(event.key==='Enter')batchAdd()"><button class="btn-sm gray" onclick="delAddRow(this)" style="flex-shrink:0">&#10005;</button></div></div>
<div style="display:flex;gap:8px">
<button class="btn-sm blue" onclick="addRow()">&#10133; 添加一行</button>
<button class="btn btn-green" onclick="batchAdd()" style="padding:8px 20px;font-size:14px">&#128640; 全部添加</button>
</div>
</div>
</div></div>
<div class="modal-bg" id="feeModal">
<div class="modal" style="width:420px">
<div class="modal-head"><h3>&#9881; 费率设置</h3><button class="modal-close" onclick="closeFeeModal()">&times;</button></div>
<div class="modal-body">
<div style="display:flex;flex-direction:column;gap:12px;font-size:14px">
<div style="display:flex;align-items:center;gap:8px"><span style="width:110px">头程运费单价：</span><input type="number" id="cfg1" step="0.1" style="flex:1;padding:6px 10px;border:1px solid #d9d9d9;border-radius:4px;text-align:center"> <span>元/公斤</span></div>
<div style="display:flex;align-items:center;gap:8px"><span style="width:110px">打包费：</span><input type="number" id="cfg2" step="0.1" style="flex:1;padding:6px 10px;border:1px solid #d9d9d9;border-radius:4px;text-align:center"> <span>元/单</span></div>
<div style="display:flex;align-items:center;gap:8px"><span style="width:110px">换面单费：</span><input type="number" id="cfg3" step="0.1" style="flex:1;padding:6px 10px;border:1px solid #d9d9d9;border-radius:4px;text-align:center"> <span>元/单</span></div>
</div>
</div>
<div class="modal-foot"><button class="btn btn-green" onclick="saveConfig();closeFeeModal()" style="width:100%">&#128190; 保存并关闭</button></div>
</div></div>
<script>
var custModal=document.getElementById('custModal');
function openCustomerModal(){{custModal.classList.add('on');loadCustomers();}}
function closeCustomerModal(){{custModal.classList.remove('on');}}
custModal.addEventListener('click',function(e){{if(e.target===custModal)closeCustomerModal();}});
var feeModal=document.getElementById('feeModal');
function openFeeModal(){{feeModal.classList.add('on');loadConfig();}}
function closeFeeModal(){{saveConfig();feeModal.classList.remove('on');}}
feeModal.addEventListener('click',function(e){{if(e.target===feeModal)closeFeeModal();}});
function loadConfig(){{fetch('/api/config').then(r=>r.json()).then(d=>{{document.getElementById('cfg1').value=d['头程运费单价'];document.getElementById('cfg2').value=d['打包费'];document.getElementById('cfg3').value=d['换面单费'];}});}}
function saveConfig(){{var r1=parseFloat(document.getElementById('cfg1').value);var r2=parseFloat(document.getElementById('cfg2').value);var r3=parseFloat(document.getElementById('cfg3').value);if(isNaN(r1)||isNaN(r2)||isNaN(r3))return;fetch('/api/config/update',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{'头程运费单价':r1,'打包费':r2,'换面单费':r3}})}}).then(r=>r.json()).then(d=>{{if(d.ok){{var rd=document.getElementById('rateDisplay');if(rd)rd.innerHTML='头程费：<b style=\"color:#e74c3c;font-size:17px\">'+r1+'</b> 元/kg · 打包费：<b style=\"color:#e74c3c;font-size:17px\">'+r2+'</b> 元/单 · 换面单费：<b style=\"color:#e74c3c;font-size:17px\">'+r3+'</b> 元/单 <a href=\"#\" onclick=\"openFeeModal()\" style=\"color:#4a90d9;margin-left:6px\">&#9881;修改</a>';}}else{{alert(d.msg||'保存失败');}}}});}}
function loadCustomers(){{fetch('/api/customers').then(r=>r.json()).then(data=>{{var h='<div class=\"cust-count\">共 '+data.length+' 个客户</div>';data.forEach(function(n){{var safe=escHtml(n);var safeJs=safe.replace(/'/g,'&#39;');h+='<div class=\"cust-row\"><span class=\"name\">'+safe+'</span><span class=\"acts\"><button class=\"edit\" onclick=\"editCust(\\''+safeJs+'\\')\">&#9998; 编辑</button><button class=\"del\" onclick=\"delCust(\\''+safeJs+'\\')\">&#10005; 删除</button></span></div>';}});document.getElementById('custList').innerHTML=h;}});}}
function addRow(){{var d=document.getElementById('addRows');var r=document.createElement('div');r.style.cssText='display:flex;gap:8px;align-items:center;margin-top:6px';r.innerHTML='<input type=\"text\" class=\"cust-new-input\" placeholder=\"输入客户名\" style=\"flex:1;padding:8px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;outline:none\" onkeydown=\"if(event.key===\\'Enter\\')batchAdd()\"><button class=\"btn-sm gray\" onclick=\"delAddRow(this)\" style=\"flex-shrink:0\">&#10005;</button>';d.appendChild(r);r.querySelector('input').focus();}}
function delAddRow(btn){{var rows=document.querySelectorAll('#addRows>div');if(rows.length<=1){{rows[0].querySelector('input').value='';return;}}btn.parentElement.remove();}}
function batchAdd(){{var names=[];document.querySelectorAll('.cust-new-input').forEach(function(inp){{var n=inp.value.trim();if(n)names.push(n);}});if(!names.length)return;fetch('/api/customers/batch_add',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{names:names}})}}).then(r=>r.json()).then(d=>{{if(d.ok){{document.getElementById('addRows').innerHTML='<div style=\"display:flex;gap:8px;align-items:center\"><input type=\"text\" class=\"cust-new-input\" placeholder=\"输入客户名\" style=\"flex:1;padding:8px 12px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;outline:none\" onkeydown=\"if(event.key===\\'Enter\\')batchAdd()\"><button class=\"btn-sm gray\" onclick=\"delAddRow(this)\" style=\"flex-shrink:0\">&#10005;</button></div>';loadCustomers();location.reload();}}else{{alert(d.msg);}}}});}}
function delCust(n){{if(!confirm('确定删除客户 \"'+n+'\" 吗？'))return;fetch('/api/customers/delete',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{name:n}})}}).then(r=>r.json()).then(d=>{{if(d.ok){{loadCustomers();location.reload();}}}});}}
function editCust(old){{var nw=prompt('修改客户名：',old);if(!nw||nw.trim()===old)return;fetch('/api/customers/update',{{method:'POST',headers:{{'Content-Type':'application/json'}},body:JSON.stringify({{old:old,'new':nw.trim()}})}}).then(r=>r.json()).then(d=>{{if(d.ok){{loadCustomers();location.reload();}}else{{alert(d.msg);}}}});}}
function uploadCustomerFile(){{var fi=document.getElementById('custFileInput');if(!fi.files.length)return;var fd=new FormData();fd.append('file',fi.files[0]);var st=document.getElementById('uploadStatus');st.textContent='导入中...';fetch('/api/customers/upload',{{method:'POST',body:fd}}).then(r=>r.json()).then(d=>{{if(d.ok){{st.textContent='已导入'+d.added+'个，共'+d.total+'个';fi.value='';loadCustomers();location.reload();}}else{{st.textContent='失败:'+(d.msg||'');}}}});}}
</script></body></html>"""

def render_upload():
    return f"""
<div class="card">
    <form method="POST" enctype="multipart/form-data" onsubmit="return showLoading()">
        <div class="row">
            <label class="file-label" for="fileInput"><svg viewBox="0 0 24 24" width="20" height="20"><path d="M9 16h6v-6h4l-7-7-7 7h4zm-4 2h14v2H5z" fill="#fff"/></svg>选择 Excel 文件</label>
            <input type="file" name="file" id="fileInput" accept=".xlsx,.xls" onchange="onFileSelected(this)" style="position:absolute;opacity:0;width:0;height:0">
            <span class="file-name" id="fileName">未选择文件</span>
            <span style="display:inline-flex;align-items:center;gap:4px;font-size:14px;color:#555">
                <input type="number" name="year" placeholder="xxx" style="width:60px;padding:8px 6px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;text-align:center;outline:none" min="2020" max="2099"> <b>年</b>
                <input type="number" name="month" placeholder="xx" style="width:45px;padding:8px 4px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;text-align:center;outline:none" min="1" max="12"> <b>月</b>
                <input type="number" name="day" placeholder="xx" style="width:45px;padding:8px 4px;border:1px solid #d9d9d9;border-radius:6px;font-size:14px;text-align:center;outline:none" min="1" max="31"> <b>日</b>
            </span>
            <button type="submit" class="btn btn-green" id="submitBtn" disabled>&#128640; 开始处理</button>
            <button type="button" class="btn btn-blue" onclick="openCustomerModal()" style="margin-left:8px">&#128100; 客户名单（{len(CUSTOMER_SET)}）</button>
            <span id="rateDisplay" style="font-size:15px;color:#888;white-space:nowrap;border:1px solid #d9d9d9;border-radius:6px;padding:8px 20px;background:#fafafa">头程费：<b style="color:#e74c3c;font-size:17px">{HEAD_FREIGHT_RATE}</b> 元/kg · 打包费：<b style="color:#e74c3c;font-size:17px">{PACKING_FEE_PER}</b> 元/单 · 换面单费：<b style="color:#e74c3c;font-size:17px">{LABEL_CHANGE_FEE_PER}</b> 元/单 <a href="#" onclick="openFeeModal()" style="color:#4a90d9;margin-left:6px">&#9881;修改</a></span>
        </div>
    </form>
</div>
"""

EMPTY = """
<div style="text-align:center;padding:60px 20px;color:#bbb">
    <div style="font-size:56px;margin-bottom:12px">&#128228;</div>
    <div style="font-size:15px">请上传 Excel 文件开始处理</div>
</div>
"""

# ── 辅助函数 ─────────────────────────────────────────────

def is_empty(v):
    if v is None: return True
    if pd.isna(v): return True
    return str(v).strip() == ""

def normalize_name(name):
    """去空格、去通快-前缀、去尾部"店铺"字样"""
    s = str(name).strip()
    if s.startswith("通快-"): s = s[3:]
    s = re.sub(r'店铺$', '', s)
    return s

def is_customer(name):
    """精确匹配：去通快-前缀后，必须在客户名单中完全一致"""
    return normalize_name(name) in CUSTOMER_SET

def ceil2(v):
    """截断到3位小数，第三位>0则向上进一。先格式化消浮点噪声，再用Decimal精确运算。"""
    # 格式化为6位小数消除浮点误差，再转Decimal
    d = Decimal(f'{float(v):.6f}')
    d3 = d.quantize(Decimal('0.001'), rounding=ROUND_DOWN)
    third = int(d3 * 1000) % 10
    if third > 0:
        return float(d.quantize(Decimal('0.01'), rounding=ROUND_UP))
    else:
        return float(d.quantize(Decimal('0.01')))

def clean_store_name(name):
    s = str(name).strip()
    if s.startswith("通快-"): s = s[3:]
    return s

def is_daily_sheet(name):
    return bool(re.match(r'^\d+\.\d+$', str(name).strip()))

def parse_daily_sheet(fp, sheet_name):
    df = pd.read_excel(fp, sheet_name=sheet_name, header=None)
    has_wh = (df.shape[1] == 10)
    if has_wh:
        df.columns = ['订单号','店铺名','仓库','包裹数','机器称重','头程运费单价','头程费','打包费','换面单费','小计']
        data = df.iloc[1:].drop(columns=['仓库']).copy()
    else:
        df.columns = ['订单号','店铺名','包裹数','机器称重','头程运费单价','头程费','打包费','换面单费','小计']
        data = df.iloc[1:].copy()
    data = data[~(data['订单号'].isna() & data['店铺名'].astype(str).str.contains('小计'))]
    data = data[data['订单号'].notna()]
    data['日期'] = sheet_name
    return data.reset_index(drop=True)

def recalc(df):
    df = df.copy()
    df['头程运费单价'] = HEAD_FREIGHT_RATE
    df['头程费'] = df['机器称重'] * HEAD_FREIGHT_RATE
    df['打包费'] = df['包裹数'] * PACKING_FEE_PER
    df['换面单费'] = df['包裹数'] * LABEL_CHANGE_FEE_PER
    df['小计'] = df['头程费'] + df['打包费'] + df['换面单费']
    return df


# ── 主路由 ────────────────────────────────────────────

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return page("&#128230; 发货明细总表", render_upload() + '<div class="error">&#9888; 未收到文件</div>')

        file = request.files["file"]
        if file.filename == "":
            return page("&#128230; 发货明细总表", render_upload() + '<div class="error">&#9888; 请选择文件</div>')


        # 保存临时文件
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        tmp_path = tmp.name
        atexit.register(lambda p=tmp_path: os.path.exists(p) and os.unlink(p))
        try: file.save(tmp_path)
        except Exception as e:
            tmp.close()
            try: os.unlink(tmp_path)
            except: pass
            return page("&#128230; 发货明细总表", render_upload() + f'<div class="error">&#9888; 保存失败: {e}</div>')
        tmp.close()

        # 读取填写的日期
        filter_date = None
        try:
            y = int(request.form.get('year', 0))
            m = int(request.form.get('month', 0))
            d = int(request.form.get('day', 0))
            if y and m and d:
                filter_date = f"{y}-{m:02d}-{d:02d}"
        except: pass

        bill_html = ""
        ship_html = ""

        # ===== A. 账单清洗（店铺名+上传时间）=====
        try:
            df = pd.read_excel(tmp_path)
            has_store = "店铺名" in df.columns
            has_time = False
            time_col = None
            for c in df.columns:
                if '上传时间' in str(c) or '时间' in str(c):
                    has_time = True
                    time_col = c
                    break

            if has_store and has_time:
                # 日期过滤：上传时间不匹配 → _del
                df["_del"] = False
                if filter_date and time_col:
                    try:
                        dates = pd.to_datetime(df[time_col]).dt.strftime('%Y-%m-%d')
                        df["_del"] = (dates != filter_date)
                    except: pass

                # 客户匹配
                df["_cust"] = df["店铺名"].apply(is_customer)
                # 前端展示：统一去掉"店铺"，客户加"通快-"前缀
                df["店铺名"] = df.apply(
                    lambda r: ("通快-"+normalize_name(r["店铺名"])) if r["_cust"] else normalize_name(r["店铺名"]), axis=1)

                orig = len(df); dc = int(df["_del"].sum()); kc = orig - dc
                nc = int((~df["_del"] & ~df["_cust"]).sum()); cc = kc - nc

                df["_sort"] = df.apply(
                    lambda r: 0 if (not r["_del"] and r["_cust"]) else (1 if not r["_del"] else 2), axis=1)
                df.sort_values("_sort", inplace=True)

                df_clean = df[~df["_del"] & df["_cust"]].drop(columns=["_del","_cust","_sort"])
                app.config["CLEAN_DF"] = df_clean
                app.config["EXPORT_DATE"] = filter_date  # 供导出建目录用

                cols = [x for x in df.columns if x not in ("_del","_cust","_sort")]
                tdata = df.to_dict("records")

                th = '<th>#</th>'+''.join(f'<th>{html.escape(str(x))}</th>' for x in cols)+'<th>标记</th><th>操作</th>'
                trs = []
                cust_idx = 0
                for i, row in enumerate(tdata):
                    cl = ''
                    editable = False
                    if row.get('_del'):
                        cl = 'red'
                        if not row.get('_cust'): cl += ' orange'
                    elif not row.get('_cust'):
                        cl = 'orange'
                    else:
                        editable = True  # 正常客户行可编辑
                    td = f'<td>{i+1}</td>'
                    for x in cols:
                        v = row.get(x)
                        if v is None: v = ''
                        if editable and x not in ('_del','_cust','_sort'):
                            td += f'<td contenteditable="true" data-row="{cust_idx}" data-col="{html.escape(str(x))}" class="editable">{html.escape(str(v))}</td>'
                        else:
                            td += f'<td>{html.escape(str(v))}</td>'
                    if row.get('_del'): b = '<span class="badge badge-del">日期不对</span>'
                    elif not row.get('_cust'): b = '<span class="badge badge-warn">非客户</span>'
                    else: b = '<span class="badge badge-ok">客户</span>'
                    td += f'<td>{b}</td>'
                    if editable:
                        td += f'<td><button class="btn-sm gray" data-idx="{cust_idx}" onclick="delRow(this)" style="padding:2px 8px;font-size:12px">&#10005; 删除</button></td>'
                        cust_idx += 1
                    else:
                        td += '<td></td>'
                    trs.append(f'<tr class="{cl}">{td}</tr>')

                bill_html = f"""
    <div class="stats">
        <div class="stat-card"><div class="stat-num">{orig}</div><div class="stat-label">&#128203; 原始行数</div></div>
        <div class="stat-card"><div class="stat-num" style="color:#e74c3c">{dc}</div><div class="stat-label">&#128465; 日期不对</div></div>
        <div class="stat-card"><div class="stat-num" style="color:#fa8c16">{nc}</div><div class="stat-label">&#9888; 非客户</div></div>
        <div class="stat-card"><div class="stat-num">{cc}</div><div class="stat-label">&#127978; 已匹配客户</div></div>
    </div>
    <div class="legend">
        <div class="legend-item"><span class="dot r"></span>日期不对（已删除）</div>
        <div class="legend-item"><span class="dot o"></span>非客户店铺</div>
        <div class="legend-item"><span class="dot n"></span>正常数据</div>
    </div>
    <div class="table-wrap">
        <div class="table-head"><h2>&#128202; 账单清洗结果</h2>
            <div class="table-tools">
                <input type="text" class="search-input" id="searchBox1" placeholder="&#128269; 输入关键词筛选...">
                <button class="btn-sm blue" id="searchBtn1">&#128269; 搜索</button>
                <a href="#" class="btn-sm purple" onclick="doExport('/download/export');return false">&#128202; 今日账目导出</a>
                <a href="/" class="btn-sm gray">&#8634; 重新上传</a>
            </div>
        </div>
        <div class="scroll"><table id="table1"><thead><tr>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>
    </div>
    <script>initSearch('table1','searchBox1','searchBtn1');</script>"""
        except Exception as e:
            bill_html = f'<div class="error">账单清洗出错: {e}</div>'

        # ===== B. 每日账目处理（有日期工作表）=====
        try:
            xl = pd.ExcelFile(tmp_path)
            sheets = [s for s in xl.sheet_names if is_daily_sheet(s)]
            if sheets:
                frames = []
                for s in sheets:
                    df_s = parse_daily_sheet(tmp_path, s)
                    if len(df_s) > 0:
                        df_s['包裹数'] = pd.to_numeric(df_s['包裹数'], errors='coerce').fillna(0).astype(int)
                        df_s['机器称重'] = pd.to_numeric(df_s['机器称重'], errors='coerce').fillna(0.0)
                        df_s = recalc(df_s)
                        frames.append(df_s)
                if frames:
                    all_data = pd.concat(frames, ignore_index=True)
                    co = ['订单号','店铺名','日期','包裹数','机器称重','头程运费单价','头程费','打包费','换面单费','小计']
                    all_data = all_data[[c for c in co if c in all_data.columns]]
                    for c in ['机器称重','打包费','换面单费','小计']:
                        all_data[c] = all_data[c].round(2)
                    all_data['头程费'] = all_data['头程费'].apply(ceil2)
                    sm = build_store_summary(all_data)
                    for c in ['机器称重合计','头程费合计','打包费合计','换面单费合计','小计合计']:
                        sm[c] = sm[c].round(2)

                    app.config["SHIP_ALL"] = all_data
                    app.config["SHIP_SM"] = sm

                    to = len(all_data); sc = len(sheets); stc = all_data['店铺名'].nunique()
                    thd = round(float(all_data['头程费'].sum()),2)
                    tpk = round(float(all_data['打包费'].sum()),2)
                    tlb = round(float(all_data['换面单费'].sum()),2)
                    tsub = round(float(all_data['小计'].sum()),2)
                    pc = [c for c in co if c in all_data.columns]
                    td2 = all_data.head(500).to_dict("records")

                    th2 = '<th>#</th>'+''.join(f'<th>{html.escape(str(x))}</th>' for x in pc)
                    trs2 = []
                    for i, row in enumerate(td2):
                        td = f'<td>{i+1}</td>'
                        for col in pc:
                            v = row.get(col)
                            if v is None or (isinstance(v,float) and pd.isna(v)): td += '<td></td>'
                            elif isinstance(v,float): td += f'<td>{v:.2f}</td>'
                            else: td += f'<td>{html.escape(str(v))}</td>'
                        trs2.append(f'<tr>{td}</tr>')

                    ship_html = f"""
    <hr class="divider">
    <div class="stats">
        <div class="stat-card"><div class="stat-num">{to}</div><div class="stat-label">&#128230; 总订单数</div></div>
        <div class="stat-card"><div class="stat-num">{sc}</div><div class="stat-label">&#128197; 日期数</div></div>
        <div class="stat-card"><div class="stat-num">{stc}</div><div class="stat-label">&#127978; 店铺数</div></div>
        <div class="stat-card"><div class="stat-num">{thd}</div><div class="stat-label">&#128176; 头程费合计</div></div>
        <div class="stat-card"><div class="stat-num" style="color:#e74c3c">{tpk}</div><div class="stat-label">&#128230; 打包费合计</div></div>
        <div class="stat-card"><div class="stat-num" style="color:#e74c3c">{tlb}</div><div class="stat-label">&#128221; 换面单费合计</div></div>
        <div class="stat-card"><div class="stat-num" style="color:#52c41a">{tsub}</div><div class="stat-label">&#128178; 总费用合计</div></div>
    </div>
    <div class="table-wrap">
        <div class="table-head"><h2>&#128203; 每日账目明细（前500行）</h2>
            <div class="table-tools">
                <input type="text" class="search-input" id="searchBox2" placeholder="&#128269; 输入关键词筛选...">
                <button class="btn-sm blue" id="searchBtn2">&#128269; 搜索</button>
                <a href="#" class="btn-sm purple" onclick="doExport('/download/ship/export');return false">&#128202; 今日账目导出</a>
                <a href="/" class="btn-sm gray">&#8634; 重新上传</a>
            </div>
        </div>
        <div class="scroll"><table id="table2"><thead><tr>{th2}</tr></thead><tbody>{"".join(trs2)}</tbody></table></div>
    </div>
    <script>initSearch('table2','searchBox2','searchBtn2');</script>"""
        except Exception as e:
            ship_html = f'<div class="error">每日账目处理出错: {e}</div>'

        try: os.unlink(tmp_path)
        except: pass

        if not bill_html and not ship_html:
            body = render_upload() + '<div class="error">&#9888; 文件中未找到"店铺名"和"上传时间"列，无法处理。</div>'
        else:
            body = render_upload() + bill_html + ship_html

        return page("&#128230; 发货明细总表", body)

    return page("&#128230; 发货明细总表", render_upload() + EMPTY)


# ── 下载路由 ─────────────────────────────────────────

@app.route("/download/export")
def dl_export():
    """今日账目导出：两个Excel（清洗数据 + 各店铺分类），打包成zip"""
    df = app.config.get("CLEAN_DF")
    if df is None or df.empty: return "暂无数据", 400

    df = df.copy()
    store_col = order_col = weight_col = None
    for c in df.columns:
        cs = str(c)
        if '店铺' in cs or '店名' in cs: store_col = c
        if '订单' in cs or '单号' in cs: order_col = c
        if '称重' in cs or '重量' in cs or 'weight' in cs.lower(): weight_col = c

    out = pd.DataFrame()
    out['订单号'] = df[order_col].astype(str) if order_col else ['ORD-'+str(i) for i in range(len(df))]
    out['店铺名'] = df[store_col] if store_col else ''
    out['包裹数'] = 1
    out['机器称重'] = pd.to_numeric(df[weight_col], errors='coerce').fillna(0) if weight_col else 0
    rate = float(CONFIG.get("头程运费单价", 69))
    pack = float(CONFIG.get("打包费", 2))
    label = float(CONFIG.get("换面单费", 5))
    out[f'头程运费单价（{rate}元/公斤）'] = rate
    out['头程费'] = (out['机器称重'] * rate).apply(ceil2)
    out[f'打包费（{pack}元/单）'] = pack
    out[f'换面单费（{label}元/单）'] = label
    out['小计'] = out['头程费'] + pack + label

    # 目录：发货明细/2026年/2026年6月每日账目/6月17日/
    filter_date = app.config.get("EXPORT_DATE", "")
    if filter_date:
        parts = filter_date.split('-')
        year_str = f"{parts[0]}年"
        month_str = f"{parts[0]}年{int(parts[1])}月"
        day_str = f"{int(parts[1])}月{int(parts[2])}日"
    else:
        year_str = month_str = day_str = "未知"

    base = _data_dir()
    dir_day = os.path.join(base, year_str, f"{month_str}每日账目", day_str)
    # 清空旧文件再写入
    if os.path.exists(dir_day):
        for f in os.listdir(dir_day):
            try: os.remove(os.path.join(dir_day, f))
            except: pass
    os.makedirs(dir_day, exist_ok=True)

    # 两个文件都存到日期文件夹
    df.to_excel(os.path.join(dir_day, f"{day_str}总发货明细表.xlsx"), index=False, sheet_name="清洗数据")
    if store_col:
        for store in sorted(out['店铺名'].dropna().unique()):
            sd = out[out['店铺名'] == store].copy().reset_index(drop=True)
            srow = {'订单号':'','店铺名':'小计','包裹数':int(sd['包裹数'].sum()),
                '机器称重':round(sd['机器称重'].sum(),3),
                f'头程运费单价（{rate}元/公斤）':'','头程费':round(sd['头程费'].sum(),2),
                f'打包费（{pack}元/单）':int(sd[f'打包费（{pack}元/单）'].sum()),
                f'换面单费（{label}元/单）':int(sd[f'换面单费（{label}元/单）'].sum()),
                '小计':round(sd['小计'].sum(),2)}
            sd = pd.concat([sd, pd.DataFrame([srow])], ignore_index=True)
            cname = clean_store_name(store)
            fname = f"通快-{cname}（{day_str}）-个人的每日账目.xlsx"
            sd.to_excel(os.path.join(dir_day, fname), index=False, sheet_name=cname[:31])
    return jsonify({'ok': True, 'msg': f'已导出到：{dir_day}', 'dir': str(base)})


@app.route("/download/ship/export")
def dl_ship_export():
    """每日账目导出：两个Excel（全部订单 + 各店铺汇总），打包成zip"""
    df = app.config.get("SHIP_ALL")
    if df is None or df.empty: return "暂无数据", 400
    savedir = _data_dir()
    # 文件1：全部订单
    df.to_excel(os.path.join(savedir, "全部订单.xlsx"), index=False, sheet_name="全部订单")
    # 文件2：各店铺每日汇总
    buf2 = io.BytesIO()
    with pd.ExcelWriter(buf2, engine="openpyxl") as w:
        for store in sorted(df['店铺名'].unique()):
            sd = df[df['店铺名'] == store].sort_values('日期')
            ds = sd.groupby('日期').agg(
                包裹数=('包裹数','sum'), 机器称重=('机器称重','sum'),
                头程费=('头程费','sum'), 打包费=('打包费','sum'),
                换面单费=('换面单费','sum'), 小计=('小计','sum'),
                订单数=('订单号','count')).reset_index()
            for c in ['机器称重','头程费','打包费','换面单费','小计']:
                ds[c] = ds[c].round(2)
            ds.to_excel(w, index=False, sheet_name=clean_store_name(store)[:31])
    buf2.seek(0)
    with open(os.path.join(savedir, "各店铺每日汇总.xlsx"), 'wb') as f:
        f.write(buf2.read())
    return jsonify({'ok': True, 'msg': f'已导出到 {savedir}', 'dir': savedir})


def build_store_summary(all_data):
    df = all_data.copy()
    df['店铺名_clean'] = df['店铺名'].apply(clean_store_name)
    s = df.groupby(['店铺名_clean','日期']).agg(
        包裹数合计=('包裹数','sum'), 机器称重合计=('机器称重','sum'),
        头程费合计=('头程费','sum'), 打包费合计=('打包费','sum'),
        换面单费合计=('换面单费','sum'), 小计合计=('小计','sum'),
        订单数=('订单号','count')).reset_index()
    return s.sort_values(['店铺名_clean','日期'])


# ── 客户名单 API ──────────────────────────────────

@app.route("/api/customers")
def api_customers():
    return jsonify(get_customer_list())

@app.route("/api/customers/add", methods=["POST"])
def api_customers_add():
    data = request.get_json()
    name = normalize_name(data.get('name', ''))
    if not name: return jsonify({'ok': False, 'msg': '名称不能为空'})
    if name in CUSTOMER_SET: return jsonify({'ok': False, 'msg': '已存在'})
    CUSTOMER_SET.add(name)
    save_customers()
    return jsonify({'ok': True})

@app.route("/api/config", methods=["GET"])
def api_config():
    return jsonify(CONFIG)

@app.route("/api/config/update", methods=["POST"])
def api_config_update():
    global HEAD_FREIGHT_RATE, PACKING_FEE_PER, LABEL_CHANGE_FEE_PER
    data = request.get_json()
    allowed_keys = {"头程运费单价", "打包费", "换面单费"}
    for k in data:
        if k not in allowed_keys:
            return jsonify({'ok': False, 'msg': f'未知配置项: {k}'})
        try:
            v = float(data[k])
        except (ValueError, TypeError):
            return jsonify({'ok': False, 'msg': f'{k} 必须是数字'})
        if v < 0 or v > 9999:
            return jsonify({'ok': False, 'msg': f'{k} 取值范围: 0 ~ 9999'})
        CONFIG[k] = v
    HEAD_FREIGHT_RATE = float(CONFIG.get("头程运费单价", 69))
    PACKING_FEE_PER = float(CONFIG.get("打包费", 2))
    LABEL_CHANGE_FEE_PER = float(CONFIG.get("换面单费", 5))
    save_customers()
    return jsonify({'ok': True})

@app.route("/api/customers/upload", methods=["POST"])
def api_customers_upload():
    if "file" not in request.files: return jsonify({'ok': False, 'msg': '未收到文件'})
    file = request.files["file"]
    try:
        df = pd.read_excel(file)
        col = df.columns[0]
        added = 0
        for raw in df[col].dropna():
            s = normalize_name(str(raw))
            if s and s not in CUSTOMER_SET:
                CUSTOMER_SET.add(s); added += 1
        if added > 0: save_customers()
        return jsonify({'ok': True, 'added': added, 'total': len(CUSTOMER_SET)})
    except Exception as e:
        return jsonify({'ok': False, 'msg': str(e)})

@app.route("/api/customers/batch_add", methods=["POST"])
def api_customers_batch_add():
    data = request.get_json()
    names = data.get('names', [])
    added = 0; skipped = []
    for n in names:
        n = normalize_name(n)
        if not n: continue
        if n in CUSTOMER_SET: skipped.append(n); continue
        CUSTOMER_SET.add(n); added += 1
    if added > 0:
        save_customers()
    return jsonify({'ok': True, 'added': added, 'skipped': skipped})

@app.route("/api/customers/delete", methods=["POST"])
def api_customers_delete():
    data = request.get_json()
    name = normalize_name(data.get('name', ''))
    if name in CUSTOMER_SET:
        CUSTOMER_SET.remove(name)
        save_customers()
    return jsonify({'ok': True})

@app.route("/api/customers/update", methods=["POST"])
def api_customers_update():
    data = request.get_json()
    old = normalize_name(data.get('old', ''))
    new = normalize_name(data.get('new', ''))
    if not new: return jsonify({'ok': False, 'msg': '名称不能为空'})
    if old in CUSTOMER_SET:
        CUSTOMER_SET.remove(old)
        CUSTOMER_SET.add(new)
        save_customers()
    return jsonify({'ok': True})


@app.route("/api/export/check")
def api_export_check():
    """检查导出文件是否已存在"""
    filter_date = app.config.get("EXPORT_DATE", "")
    if not filter_date: return jsonify({'exists': False})
    parts = filter_date.split('-')
    day_str = f"{int(parts[1])}月{int(parts[2])}日"
    dir_day = os.path.join(_data_dir(), f"{parts[0]}年", f"{parts[0]}年{int(parts[1])}月每日账目", day_str)
    exists = os.path.isdir(dir_day) and os.listdir(dir_day)
    return jsonify({'exists': exists})

@app.route("/api/row/delete", methods=["POST"])
def api_row_delete():
    data = request.get_json()
    idx = data.get('idx')
    df = app.config.get("CLEAN_DF")
    if df is None or df.empty: return jsonify({'ok': False, 'msg': '无数据'})
    if idx is not None and idx < len(df):
        df.drop(df.index[idx], inplace=True)
        df.reset_index(drop=True, inplace=True)
        app.config["CLEAN_DF"] = df
        return jsonify({'ok': True, 'rows': len(df)})
    return jsonify({'ok': False, 'msg': '无效索引'})

@app.route("/api/row/update", methods=["POST"])
def api_row_update():
    data = request.get_json()
    idx = data.get('idx')
    col = data.get('col')
    val = data.get('val')
    df = app.config.get("CLEAN_DF")
    if df is None or df.empty: return jsonify({'ok': False, 'msg': '无数据'})
    if idx is not None and col is not None and idx < len(df):
        if col in df.columns:
            df.at[df.index[idx], col] = val
            app.config["CLEAN_DF"] = df
            return jsonify({'ok': True})
    return jsonify({'ok': False, 'msg': '更新失败'})


if __name__ == "__main__":
    print("=" * 50)
    print("  发货明细总表  http://127.0.0.1:5000")
    print(f"  已加载客户: {len(CUSTOMER_SET)} 个")
    print("=" * 50)
    app.run(debug=False, host="127.0.0.1", port=5000)
