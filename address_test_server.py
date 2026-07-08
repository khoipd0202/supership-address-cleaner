# -*- coding: utf-8 -*-
"""
address_test_server.py — Server + UI test realtime cho parse_address.py

Chạy:
    python3 address_test_server.py [port]   (mặc định port 8765)

Rồi mở trình duyệt: http://localhost:8765

Chức năng:
- "Test nhanh": gõ/dán địa chỉ (mỗi dòng 1 địa chỉ) -> tách realtime bằng
  tầng luật, hiển thị từng trường + tô màu highlight phần nhận diện được
  ngay trên văn bản gốc, kèm thống kê tổng hợp.
- "Test theo file Excel": chọn file .xlsx trên máy (đọc bằng SheetJS ngay
  trên trình duyệt, không upload lên đâu cả) -> tách hàng loạt, ra thống
  kê % tách được từng cấp, xem trước bảng kết quả.
- Nút "Thử AI (Ollama)" trên từng dòng: so sánh kết quả TRƯỚC/SAU khi
  chạy thêm tầng AI local cho đúng địa chỉ đó.

Không cần cài thêm gì ngoài Python chuẩn (dùng lại toàn bộ logic trong
parse_address.py cùng thư mục).
"""
import json
import os
import re
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from parse_address import (  # noqa: E402
    Parser, extract_phone, confidence, mask_phone, base_name,
    strip_diacritics, ai_call_ollama, ai_merge, vtitle,
)

DATA_PATH = os.path.join(HERE, "vn_units_data.json")
DIR_PATH = os.path.join(HERE, "danh_ba_khach.json")

if not os.path.exists(DATA_PATH):
    sys.exit("Thiếu file vn_units_data.json cạnh address_test_server.py")

print("Đang nạp danh mục ĐVHC...")
_t0 = time.time()
PARSER = Parser(DATA_PATH)
print(f"  xong ({time.time()-_t0:.1f}s)")

DIRECTORY = {}
if os.path.exists(DIR_PATH):
    with open(DIR_PATH, encoding="utf-8") as f:
        DIRECTORY = json.load(f)
    print(f"  danh bạ SĐT khách cũ: {len(DIRECTORY)} số")

_lock = threading.Lock()   # Parser.parse() dùng chung state tạm, an toàn hơn khi khoá

ENTRY_FIELDS = ("street", "hamlet", "ward", "district", "province",
                "ward_new", "province_new")


def parse_one(raw, use_directory=True):
    with _lock:
        res = PARSER.parse(raw)
    phone = extract_phone(raw) if raw else None
    phone_hit = conflict = False
    if use_directory and phone and phone in DIRECTORY:
        e = DIRECTORY[phone]
        if res["province"] and e.get("province") and \
                base_name(res["province"]) != base_name(e["province"]):
            conflict = True
        else:
            phone_hit = True
            for f_ in ENTRY_FIELDS:
                if not res.get(f_) and e.get(f_):
                    res[f_] = e[f_]
    conf = ("Cần kiểm tra" if conflict else
            "Cao (SĐT khách cũ)" if phone_hit else
            confidence(res, False, False))
    res["phone"] = phone
    res["confidence"] = conf
    res["phone_hit"] = phone_hit
    res["conflict"] = conflict
    return res


def find_span(raw, value):
    """Tìm vị trí (start,end) của value trong raw (không phân biệt hoa/thường,
    không phân biệt dấu) để tô màu highlight. strip_diacritics giữ nguyên độ
    dài chuỗi (thay ký tự có dấu bằng ký tự gốc 1-1), nên vị trí tìm được
    trên chuỗi đã bỏ dấu áp dụng thẳng được lên chuỗi gốc.

    Có fallback cho trường hợp giá trị được suy ra từ chữ viết tắt liền số
    (vd res["hamlet"]="Khu phố 22" nhưng raw chỉ viết "Kp22"): tìm phần
    đuôi (thường là số) như 1 token độc lập, rồi mở rộng ngược ra các ký tự
    chữ liền kề phía trước để bắt trọn cụm viết tắt."""
    if not value:
        return None
    nraw = strip_diacritics(str(raw)).lower()
    nval = strip_diacritics(str(value)).lower()
    idx = nraw.find(nval)
    if idx >= 0:
        return [idx, idx + len(nval)]
    parts = nval.split()
    if len(parts) >= 2:
        tail = parts[-1]
        # chỉ chặn lặp số phía trước (vd "22" trong "122"), CHO PHÉP chữ cái
        # đứng ngay trước (vd "kp" trong "kp22") vì đó chính là viết tắt
        m = re.search(r"(?<![0-9])" + re.escape(tail) + r"(?![a-z0-9])", nraw)
        if m:
            j = m.start()
            while j > 0 and nraw[j - 1].isalpha() and (m.start() - j) < 4:
                j -= 1
            return [j, m.end()]
    return None


def build_stats(results):
    total = len(results)
    n_detail = sum(1 for r in results if r.get("street") or r.get("hamlet"))
    n_full = sum(1 for r in results if r.get("ward"))
    n_partial = sum(1 for r in results
                     if not r.get("ward") and r.get("province"))
    n_none = total - n_full - n_partial
    conf_count = {}
    for r in results:
        c = r.get("confidence", "?")
        conf_count[c] = conf_count.get(c, 0) + 1
    return {
        "total": total, "n_detail": n_detail, "n_full": n_full,
        "n_partial": n_partial, "n_none": n_none, "conf_count": conf_count,
    }


# ---------------------------------------------------------------- HTML/JS

PAGE = r"""<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<title>Test tách địa chỉ - Realtime</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --bg: #0f1115; --panel: #171a21; --border: #2a2f3a; --text: #e8e8ec;
    --muted: #9aa0ac; --accent: #4f9dff; --good: #37b26c; --mid: #e0a92f;
    --bad: #e35b5b; --c-street: #4f9dff55; --c-hamlet: #b06fe655;
    --c-ward: #37b26c55; --c-dist: #e0a92f55; --c-prov: #e35b5b55;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    background: var(--bg); color: var(--text);
  }
  header {
    padding: 14px 20px; border-bottom: 1px solid var(--border);
    display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
  }
  header h1 { font-size: 16px; margin: 0; font-weight: 600; }
  .tabs { display: flex; gap: 4px; }
  .tab {
    padding: 6px 14px; border-radius: 6px; cursor: pointer; font-size: 13px;
    color: var(--muted); border: 1px solid transparent;
  }
  .tab.active { background: var(--panel); color: var(--text); border-color: var(--border); }
  main { padding: 20px; max-width: 1300px; margin: 0 auto; }
  .view { display: none; }
  .view.active { display: block; }
  textarea {
    width: 100%; min-height: 130px; background: var(--panel); color: var(--text);
    border: 1px solid var(--border); border-radius: 8px; padding: 10px 12px;
    font-family: ui-monospace, monospace; font-size: 13px; resize: vertical;
  }
  .stats-bar {
    display: flex; gap: 10px; flex-wrap: wrap; margin: 14px 0;
  }
  .stat {
    background: var(--panel); border: 1px solid var(--border); border-radius: 8px;
    padding: 8px 14px; font-size: 12px; color: var(--muted); min-width: 110px;
  }
  .stat b { display: block; font-size: 18px; color: var(--text); margin-top: 2px; }
  .conf-row { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }
  .badge {
    font-size: 11px; padding: 3px 9px; border-radius: 20px; border: 1px solid var(--border);
    color: var(--muted);
  }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
  th, td {
    text-align: left; padding: 7px 9px; border-bottom: 1px solid var(--border);
    vertical-align: top;
  }
  th { color: var(--muted); font-weight: 600; position: sticky; top: 0; background: var(--bg); }
  .raw-cell { max-width: 320px; }
  mark.street { background: var(--c-street); color: inherit; border-radius: 3px; padding: 0 2px; }
  mark.hamlet { background: var(--c-hamlet); color: inherit; border-radius: 3px; padding: 0 2px; }
  .pill {
    display: inline-block; font-size: 11px; padding: 2px 8px; border-radius: 20px;
    white-space: nowrap;
  }
  .pill.cao { background: #37b26c33; color: var(--good); }
  .pill.tb { background: #e0a92f33; color: var(--mid); }
  .pill.thap { background: #e35b5b33; color: var(--bad); }
  .pill.ck { background: #a259ff33; color: #c79bff; }
  .legend { display: flex; gap: 14px; font-size: 11px; color: var(--muted); margin: 8px 0 14px; flex-wrap: wrap; }
  .legend span { display: inline-flex; align-items: center; gap: 5px; }
  .sw { width: 10px; height: 10px; border-radius: 3px; display: inline-block; }
  button {
    background: var(--accent); color: white; border: none; border-radius: 6px;
    padding: 6px 12px; font-size: 12px; cursor: pointer;
  }
  button.ghost {
    background: transparent; border: 1px solid var(--border); color: var(--text);
  }
  button:disabled { opacity: .5; cursor: default; }
  .ai-btn { padding: 3px 8px; font-size: 11px; }
  .field { font-size: 12px; }
  .field .lbl { color: var(--muted); font-size: 10px; }
  .diffwrap { display: flex; gap: 6px; flex-direction: column; }
  .diff-old { color: var(--muted); text-decoration: line-through; font-size: 11px; }
  .diff-new { color: var(--good); font-size: 12px; }
  #dropzone {
    border: 1.5px dashed var(--border); border-radius: 10px; padding: 30px;
    text-align: center; color: var(--muted); font-size: 13px;
  }
  #dropzone.drag { border-color: var(--accent); color: var(--text); }
  .row-actions { display: flex; gap: 6px; align-items: center; }
  small.hint { color: var(--muted); }
  .toolbar { display: flex; gap: 10px; align-items: center; margin: 8px 0 4px; flex-wrap: wrap; }
  label.chk { font-size: 12px; color: var(--muted); display: flex; gap: 6px; align-items: center; }
  #status { font-size: 12px; color: var(--muted); }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  @media (max-width: 900px) { .grid2 { grid-template-columns: 1fr; } }
</style>
</head>
<body>

<header>
  <h1>Test tách địa chỉ &mdash; Realtime</h1>
  <div class="tabs">
    <div class="tab active" data-view="quick">Test nhanh</div>
    <div class="tab" data-view="bulk">Test theo file Excel</div>
  </div>
  <span id="status"></span>
</header>

<main>
  <!-- ========================= TEST NHANH ========================= -->
  <div class="view active" id="view-quick">
    <p><small class="hint">Mỗi dòng 1 địa chỉ. Kết quả tự cập nhật khi bạn gõ (realtime).</small></p>
    <textarea id="quick-input" placeholder="Đường Luu Văn Việt Kp22 Phường Tam Hiệp, Biên Hòa, Đồng Nai
123 Nguyễn Trãi To5 Phường 5 Quận 5 TPHCM
Thôn Đông xã Yên Sở, Hoài Đức, Hà Nội"></textarea>

    <div class="toolbar">
      <label class="chk"><input type="checkbox" id="use-dir" checked> Dùng danh bạ SĐT khách cũ</label>
      <label class="chk">Ollama host <input type="text" id="ollama-host" value="http://127.0.0.1:11434" style="width:170px;background:var(--panel);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:2px 6px;"></label>
      <label class="chk">Model <input type="text" id="ollama-model" value="qwen2.5:3b" style="width:110px;background:var(--panel);color:var(--text);border:1px solid var(--border);border-radius:5px;padding:2px 6px;"></label>
    </div>

    <div class="legend">
      <span><span class="sw" style="background:var(--c-street)"></span> Tên đường</span>
      <span><span class="sw" style="background:var(--c-hamlet)"></span> Thôn/Xóm/Ấp...</span>
    </div>

    <div class="stats-bar" id="quick-stats"></div>
    <div style="overflow-x:auto;">
    <table id="quick-table">
      <thead>
        <tr>
          <th>#</th><th class="raw-cell">Địa chỉ gốc (highlight)</th>
          <th>SĐT</th><th>Tên đường</th><th>Thôn/Xóm</th><th>Phường/Xã</th>
          <th>Quận/Huyện</th><th>Tỉnh/TP</th><th>Độ tin cậy</th><th>Ghi chú</th><th>AI</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
    </div>
  </div>

  <!-- ========================= TEST BULK FILE ========================= -->
  <div class="view" id="view-bulk">
    <p><small class="hint">Chọn file Excel trên máy &mdash; xử lý ngay trong trình duyệt (không upload lên server nào), chỉ gửi phần văn bản địa chỉ về server local của bạn để tách.</small></p>
    <div id="dropzone">
      <input type="file" id="file-input" accept=".xlsx,.xls,.csv" style="display:none;">
      <div>Kéo thả file .xlsx vào đây hoặc <button class="ghost" id="pick-file">chọn file</button></div>
      <div id="file-name" style="margin-top:8px;font-size:12px;"></div>
    </div>

    <div class="stats-bar" id="bulk-stats"></div>
    <div class="conf-row" id="bulk-conf"></div>
    <div style="overflow-x:auto; max-height: 520px; overflow-y:auto;">
    <table id="bulk-table">
      <thead>
        <tr>
          <th>#</th><th class="raw-cell">Địa chỉ gốc</th>
          <th>Tên đường</th><th>Thôn/Xóm</th><th>Phường/Xã</th>
          <th>Quận/Huyện</th><th>Tỉnh/TP</th><th>Độ tin cậy</th><th>Ghi chú</th>
        </tr>
      </thead>
      <tbody></tbody>
    </table>
    </div>
  </div>
</main>

<script>
const $ = (s, el) => (el||document).querySelector(s);
const $$ = (s, el) => Array.from((el||document).querySelectorAll(s));

// ---------- tab switching
$$('.tab').forEach(t => t.addEventListener('click', () => {
  $$('.tab').forEach(x => x.classList.remove('active'));
  $$('.view').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  $('#view-' + t.dataset.view).classList.add('active');
}));

function setStatus(s) { $('#status').textContent = s; }

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function highlightRaw(raw, spans) {
  // spans: [{start,end,cls}], không chồng lấn
  const ok = spans.filter(s => s && s.start != null).sort((a,b)=>a.start-b.start);
  let out = '', pos = 0;
  for (const s of ok) {
    if (s.start < pos) continue;
    out += escapeHtml(raw.slice(pos, s.start));
    out += `<mark class="${s.cls}">${escapeHtml(raw.slice(s.start, s.end))}</mark>`;
    pos = s.end;
  }
  out += escapeHtml(raw.slice(pos));
  return out;
}

function confPill(c) {
  let cls = 'thap';
  if (c.startsWith('Cao')) cls = 'cao';
  else if (c.startsWith('Trung')) cls = 'tb';
  else if (c.startsWith('Cần')) cls = 'ck';
  return `<span class="pill ${cls}">${escapeHtml(c)}</span>`;
}

function renderStats(target, stats) {
  const pct = (n) => stats.total ? (100*n/stats.total).toFixed(1)+'%' : '0%';
  target.innerHTML = `
    <div class="stat">Tổng địa chỉ<b>${stats.total}</b></div>
    <div class="stat">Có tên đường/thôn xóm<b>${stats.n_detail} (${pct(stats.n_detail)})</b></div>
    <div class="stat">Tới phường/xã<b>${stats.n_full} (${pct(stats.n_full)})</b></div>
    <div class="stat">Chỉ tới tỉnh/quận<b>${stats.n_partial} (${pct(stats.n_partial)})</b></div>
    <div class="stat">Không tách được<b>${stats.n_none} (${pct(stats.n_none)})</b></div>
  `;
}

function renderConfRow(target, conf_count) {
  target.innerHTML = Object.entries(conf_count)
    .sort((a,b)=>b[1]-a[1])
    .map(([k,v]) => `<span class="badge">${escapeHtml(k)}: ${v}</span>`).join('');
}

// ---------- TEST NHANH (realtime) ----------
let quickTimer = null;
$('#quick-input').addEventListener('input', () => {
  clearTimeout(quickTimer);
  quickTimer = setTimeout(runQuick, 350);
});
window.addEventListener('DOMContentLoaded', runQuick);

async function runQuick() {
  const lines = $('#quick-input').value.split('\n').map(s=>s.trim()).filter(Boolean);
  if (!lines.length) {
    $('#quick-table tbody').innerHTML = '';
    $('#quick-stats').innerHTML = '';
    return;
  }
  setStatus('Đang tách...');
  const useDir = $('#use-dir').checked;
  const res = await fetch('/api/parse', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({lines, use_directory: useDir})
  }).then(r => r.json());
  setStatus('');
  renderStats($('#quick-stats'), res.stats);
  const tbody = $('#quick-table tbody');
  tbody.innerHTML = '';
  res.results.forEach((r, i) => {
    const tr = document.createElement('tr');
    const spans = [];
    if (r.street_span) spans.push({start:r.street_span[0], end:r.street_span[1], cls:'street'});
    if (r.hamlet_span) spans.push({start:r.hamlet_span[0], end:r.hamlet_span[1], cls:'hamlet'});
    tr.innerHTML = `
      <td>${i+1}</td>
      <td class="raw-cell">${highlightRaw(r.raw, spans)}</td>
      <td>${r.phone||''}</td>
      <td>${escapeHtml(r.street||'')}</td>
      <td>${escapeHtml(r.hamlet||'')}</td>
      <td>${escapeHtml(r.ward||'')}</td>
      <td>${escapeHtml(r.district||'')}</td>
      <td>${escapeHtml(r.province||'')}</td>
      <td>${confPill(r.confidence)}</td>
      <td style="max-width:220px;color:var(--muted);font-size:11px;">${escapeHtml(r.note||'')}</td>
      <td><button class="ai-btn ghost" data-idx="${i}">Thử AI</button><div class="ai-result" id="ai-${i}"></div></td>
    `;
    tbody.appendChild(tr);
  });
  $$('.ai-btn', tbody).forEach(btn => btn.addEventListener('click', () => testAi(btn, res.results)));
}

async function testAi(btn, results) {
  const idx = +btn.dataset.idx;
  const raw = results[idx].raw;
  btn.disabled = true; btn.textContent = 'Đang chạy...';
  const host = $('#ollama-host').value;
  const model = $('#ollama-model').value;
  try {
    const r = await fetch('/api/parse_ai', {
      method: 'POST', headers: {'Content-Type':'application/json'},
      body: JSON.stringify({text: raw, host, model})
    }).then(r => r.json());
    const box = document.getElementById('ai-' + idx);
    if (r.error) {
      box.innerHTML = `<div class="diff-old">Lỗi: ${escapeHtml(r.error)}</div>`;
    } else {
      const fields = ['street','hamlet','ward','district','province'];
      let html = '<div class="diffwrap">';
      let changed = false;
      fields.forEach(f => {
        const before = r.rule[f] || '(trống)';
        const after = r.ai[f] || '(trống)';
        if (before !== after) {
          changed = true;
          html += `<div><span class="lbl">${f}</span><div class="diff-old">${escapeHtml(before)}</div><div class="diff-new">${escapeHtml(after)}</div></div>`;
        }
      });
      html += changed ? '' : '<div class="diff-old">(AI không đổi gì)</div>';
      html += `<div style="color:var(--muted);font-size:10px;">${(r.elapsed||0).toFixed(1)}s</div></div>`;
      box.innerHTML = html;
    }
  } catch (e) {
    document.getElementById('ai-' + idx).innerHTML = `<div class="diff-old">Lỗi mạng: ${e}</div>`;
  }
  btn.disabled = false; btn.textContent = 'Thử AI';
}

// ---------- TEST FILE EXCEL ----------
$('#pick-file').addEventListener('click', () => $('#file-input').click());
$('#file-input').addEventListener('change', e => handleFile(e.target.files[0]));
const dz = $('#dropzone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
dz.addEventListener('drop', e => {
  e.preventDefault(); dz.classList.remove('drag');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});

function normKey(s) {
  return String(s||'').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g,'')
    .replace(/đ/g,'d').replace(/[^a-z0-9]+/g,' ');
}

async function handleFile(file) {
  if (!file) return;
  $('#file-name').textContent = 'Đang đọc ' + file.name + ' ...';
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  if (!rows.length) { $('#file-name').textContent = 'File rỗng.'; return; }
  const header = rows[0];
  let colIdx = header.findIndex(h => normKey(h).includes('dia chi'));
  if (colIdx < 0) colIdx = 0;
  $('#file-name').textContent = `${file.name} — cột "${header[colIdx]}", ${rows.length-1} dòng. Đang tách...`;
  const lines = rows.slice(1).map(r => (r[colIdx]||'').toString()).filter(Boolean);
  setStatus('Đang tách ' + lines.length + ' địa chỉ...');
  const res = await fetch('/api/parse', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({lines, use_directory: true})
  }).then(r => r.json());
  setStatus('');
  $('#file-name').textContent = `${file.name} — cột "${header[colIdx]}", ${lines.length} địa chỉ đã tách.`;
  renderStats($('#bulk-stats'), res.stats);
  renderConfRow($('#bulk-conf'), res.stats.conf_count);
  const tbody = $('#bulk-table tbody');
  tbody.innerHTML = '';
  const frag = document.createDocumentFragment();
  res.results.slice(0, 2000).forEach((r, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i+1}</td>
      <td class="raw-cell">${escapeHtml(r.raw)}</td>
      <td>${escapeHtml(r.street||'')}</td>
      <td>${escapeHtml(r.hamlet||'')}</td>
      <td>${escapeHtml(r.ward||'')}</td>
      <td>${escapeHtml(r.district||'')}</td>
      <td>${escapeHtml(r.province||'')}</td>
      <td>${confPill(r.confidence)}</td>
      <td style="max-width:220px;color:var(--muted);font-size:11px;">${escapeHtml(r.note||'')}</td>
    `;
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
}
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass    # im lặng, khỏi rác terminal

    def _send_json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            data = PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path == "/health":
            self._send_json({"ok": True})
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            body = json.loads(self.rfile.read(length) or b"{}")
        except json.JSONDecodeError:
            self._send_json({"error": "JSON không hợp lệ"}, 400)
            return

        if self.path == "/api/parse":
            lines = body.get("lines") or []
            use_dir = body.get("use_directory", True)
            results = []
            for raw in lines:
                res = parse_one(raw, use_dir)
                res["raw"] = raw
                res["street_span"] = find_span(raw, res.get("street"))
                res["hamlet_span"] = find_span(raw, res.get("hamlet"))
                results.append(res)
            self._send_json({"results": results, "stats": build_stats(results)})

        elif self.path == "/api/parse_ai":
            text = body.get("text", "")
            model = body.get("model", "qwen2.5:3b")
            host = body.get("host", "http://127.0.0.1:11434")
            rule_res = parse_one(text, True)
            t0 = time.time()
            try:
                items = ai_call_ollama([(0, text)], model, host)
                ai_res = dict(rule_res)
                if 0 in items:
                    ai_merge(PARSER, ai_res, items[0])
                self._send_json({
                    "rule": rule_res, "ai": ai_res,
                    "elapsed": time.time() - t0, "error": None,
                })
            except Exception as exc:
                self._send_json({
                    "rule": rule_res, "ai": rule_res,
                    "elapsed": time.time() - t0, "error": str(exc),
                })
        else:
            self.send_response(404)
            self.end_headers()


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"Server đang chạy: http://localhost:{port}  (Ctrl+C để dừng)")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
