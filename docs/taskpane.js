/* Tiện ích văn phòng – taskpane dùng chung cho Excel và Word (Office.js). */
const { numberToWords, TEXT_OPS, formatDate, textStats, columnLetter, findFormulaOutliers } = TextUtils;
let HOST = null;
let busy = false;
const MAX_CELLS = 100000;
function guard(r) { if (r.rowCount * r.columnCount > MAX_CELLS) throw new Error("Vùng quá lớn; hãy chọn tối đa 100.000 ô mỗi lần."); }

const $ = (sel) => document.querySelector(sel);

function setStatus(msg, isError) {
  const el = $('#status');
  el.textContent = msg || '';
  el.classList.toggle('error', !!isError);
}

/* ---------- Excel ---------- */

// Áp dụng hàm xử lý chuỗi lên các ô chứa văn bản (bỏ qua công thức).
function excelTransform(fn) {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    range.load(['formulas', 'address', 'rowCount', 'columnCount', 'valueTypes']);
    await ctx.sync();
    if (range.isNullObject) return 'Vùng chọn đang trống.';
    guard(range);
    let changed = 0;
    const edits = [];
    range.formulas.forEach((row,r) => row.forEach((v,c) => {
      if (typeof v !== 'string' || v.startsWith('=') || range.valueTypes[r][c] !== 'String') return;
      const n = fn(v);
      if (n !== v) edits.push([r,c,n]);
    }));
    if (edits.length > 5000) throw new Error('Chia nhỏ vùng chọn: tối đa 5.000 ô sửa mỗi lần.');
    if (edits.length && !confirm(`Sửa ${edits.length} ô văn bản? Hãy sao lưu trước khi thực hiện.`)) return 'Đã huỷ.';
    edits.forEach(([r,c,n]) => { range.getCell(r,c).values = [[/^[=+\-@]/.test(n) ? "'"+n : n]]; changed++; });
    await ctx.sync();
    return changed ? `Đã sửa ${changed} ô (${range.address}).` : 'Không có ô nào cần sửa.';
  });
}

function excelNumbersToWords(withCurrency) {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    range.load(['values', 'columnCount']);
    await ctx.sync();
    if (range.isNullObject) return 'Vùng chọn đang trống.';
    if (range.columnCount > 1) throw new Error('Chỉ chọn một cột số để tránh ghi đè dữ liệu bên cạnh.');
    const dest = range.getOffsetRange(0,1);
    dest.load(['values','formulas']);
    await ctx.sync();
    const rows = range.values.map((row,r) => [row[0],r]).filter(([v]) => typeof v === 'number');
    if (rows.some(([,r]) => dest.values[r][0] !== '' || dest.formulas[r][0] !== '')) throw new Error('Cột kết quả có dữ liệu hoặc công thức. Hãy để trống các ô đích.');
    if (rows.length > 5000) throw new Error('Tối đa 5.000 số mỗi lần.');
    if (rows.length && !confirm(`Ghi ${rows.length} kết quả vào cột bên phải? Số lẻ được làm tròn.`)) return 'Đã huỷ.';
    rows.forEach(([v,r]) => dest.getCell(r,0).values = [[numberToWords(v,withCurrency)]]);
    const count = rows.length;
    await ctx.sync();
    return count ? `Đã đọc ${count} số thành chữ.` : 'Không tìm thấy ô số nào trong vùng chọn.';
  });
}

function excelDuplicates(mode) {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    range.load('values');
    await ctx.sync();
    if (range.isNullObject) return 'Vùng chọn đang trống.';
    const seen = new Set();
    const dups = [];
    range.values.forEach((row, i) => {
      if (row.every(v => v === '')) return;
      const key = JSON.stringify(row.map((v) => (typeof v === 'string' ? v.trim().toLowerCase() : v)));
      if (seen.has(key)) dups.push(i);
      else seen.add(key);
    });
    if (!dups.length) return 'Không có dòng trùng.';
    if (mode === 'mark') {
      dups.forEach((i) => { range.getRow(i).format.fill.color = '#FFF2A8'; });
    } else {
      return 'Đã vô hiệu hoá xoá cả hàng để bảo vệ dữ liệu ngoài vùng chọn. Dùng báo cáo dòng trùng.';
    }
    await ctx.sync();
    return mode === 'mark' ? `Đã tô vàng ${dups.length} dòng trùng.` : `Đã xoá ${dups.length} dòng trùng.`;
  });
}

/* ---------- Excel: kiểm tra công thức bất thường trong cột ---------- */

const FX_MAX_SHOWN = 500;
let lastFlagged = null; // { sheet, addrs } của lần tô màu gần nhất

function excelCheckFormulas(highlight, manual) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    const used = ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    used.load(['formulas', 'formulasR1C1', 'rowIndex', 'columnIndex', 'columnCount']);
    await ctx.sync();
    if (used.isNullObject) return 'Vùng chọn đang trống.';

    const reports = [];
    const flagged = [];
    for (let c = 0; c < used.columnCount; c++) {
      const letter = columnLetter(used.columnIndex + c);
      const res = findFormulaOutliers(used.formulasR1C1.map((r) => r[c]), { manual, blank: true });
      if (res.total === 0) continue; // cột không có công thức thì bỏ qua
      const addrOf = (i) => letter + (used.rowIndex + i + 1);
      const rep = { letter, status: res.status, total: res.total };
      if (res.status === 'ok') {
        rep.matching = res.domCount;
        rep.sampleAddr = addrOf(res.sample);
        rep.sampleFormula = String(used.formulas[res.sample][c]);
        rep.items = res.outliers.map((o) => ({
          addr: addrOf(o.index),
          kind: o.kind,
          text: String(used.formulas[o.index][c]),
        }));
        rep.items.forEach((it) => flagged.push(it.addr));
      }
      reports.push(rep);
    }
    if (!reports.length) {
      renderFormulaReport([], sheet.name);
      return 'Vùng chọn không có công thức nào.';
    }
    if (highlight && flagged.length) {
      if (lastFlagged) throw new Error('Hãy hoàn nguyên màu lần kiểm tra trước.');
      if (flagged.length > 2000) throw new Error('Quá 2.000 ô lỗi; hãy tắt tô màu.');
      const fills = flagged.map(a => sheet.getRange(a).format.fill);
      fills.forEach(f => f.load('color'));
      await ctx.sync();
      const colors = fills.map(f => f.color);
      flagged.forEach((a) => { sheet.getRange(a).format.fill.color = '#FFC7CE'; });
      lastFlagged = { sheet: sheet.name, addrs: flagged, colors };
    }
    await ctx.sync();
    renderFormulaReport(reports, sheet.name);
    return flagged.length
      ? `Đã kiểm tra ${reports.length} cột có công thức, phát hiện ${flagged.length} ô bất thường.`
      : `Đã kiểm tra ${reports.length} cột có công thức, không thấy ô bất thường.`;
  });
}

function renderFormulaReport(reports, sheetName) {
  const box = $('#fx-results');
  box.replaceChildren();
  reports.forEach((rep) => {
    const wrap = document.createElement('div');
    wrap.className = 'fx-col';
    const head = document.createElement('p');
    head.className = 'fx-head';
    if (rep.status === 'few') {
      head.textContent = `Cột ${rep.letter}: chưa đủ 3 công thức để so sánh.`;
    } else if (rep.status === 'mixed') {
      head.textContent = `Cột ${rep.letter}: có nhiều nhóm công thức khác nhau, không nhóm nào chiếm đa số nên bỏ qua.`;
    } else {
      head.textContent =
        `Cột ${rep.letter}: ${rep.matching}/${rep.total} công thức giống ô ${rep.sampleAddr} (${rep.sampleFormula}). ` +
        (rep.items.length ? `${rep.items.length} ô bất thường:` : 'Không có ô bất thường.');
    }
    wrap.append(head);
    (rep.items || []).slice(0, FX_MAX_SHOWN).forEach((it) => {
      const b = document.createElement('button');
      b.className = 'fx-item';
      b.dataset.addr = it.addr;
      b.dataset.sheet = sheetName;
      const a = document.createElement('strong');
      a.textContent = it.addr;
      const k = document.createElement('span');
      k.textContent = it.kind === 'formula' ? 'Công thức khác' : it.kind === 'blank' ? 'Ô trống' : 'Giá trị nhập tay';
      const f = document.createElement('code');
      f.textContent = it.text;
      b.title = it.text;
      b.append(a, k, f);
      wrap.append(b);
    });
    if (rep.items && rep.items.length > FX_MAX_SHOWN) {
      const more = document.createElement('p');
      more.className = 'hint';
      more.textContent = `… và ${rep.items.length - FX_MAX_SHOWN} ô khác (vẫn được tô màu nếu bạn bật tuỳ chọn tô).`;
      wrap.append(more);
    }
    box.append(wrap);
  });
}

function excelClearFlags() {
  if (!lastFlagged) return Promise.resolve('Chưa có ô nào được tô màu trong phiên này.');
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(lastFlagged.sheet);
    const fills = lastFlagged.addrs.map(a => ws.getRange(a).format.fill);
    fills.forEach(f => f.load('color')); await ctx.sync();
    fills.forEach((f,i) => { if (String(f.color).toUpperCase() !== '#FFC7CE') return; const old = lastFlagged.colors[i]; if (old) f.color = old; else f.clear(); });
    await ctx.sync();
    const n = lastFlagged.addrs.length;
    lastFlagged = null;
    return `Đã bỏ tô màu ${n} ô.`;
  });
}

function excelGoto(sheetName, addr) {
  return Excel.run(async (ctx) => {
    const ws = ctx.workbook.worksheets.getItem(sheetName);
    ws.activate();
    ws.getRange(addr).select();
    await ctx.sync();
  });
}

/* ---------- Word ---------- */

function wordSelectionTransform(fn) {
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    sel.load('text');
    await ctx.sync();
    if (!sel.text) return 'Hãy bôi đen đoạn văn bản cần xử lý.';
    const next = fn(sel.text);
    if (next === sel.text) return 'Không có gì cần sửa.';
    if (!confirm('Thay văn bản đã chọn có thể làm mất định dạng từng phần. Tiếp tục?')) return 'Đã huỷ.';
    sel.insertText(next, 'Replace');
    await ctx.sync();
    return 'Đã cập nhật đoạn văn bản đang chọn.';
  });
}

function wordClean() {
  return Word.run(async (ctx) => {
    if (!confirm('Dọn khoảng trắng và đoạn trống toàn tài liệu? Có thể ảnh hưởng bố cục.')) return 'Đã huỷ.';
    const body = ctx.document.body;
    const spaces = body.search('[ ]{2,}', { matchWildcards: true });
    spaces.load('items');
    await ctx.sync();
    spaces.items.forEach((r) => r.insertText(' ', 'Replace'));
    await ctx.sync();

    const paras = body.paragraphs;
    paras.load('items/text,items/tableNestingLevel');
    await ctx.sync();
    let removed = 0;
    for (let i = paras.items.length - 1; i > 0; i--) {
      const p = paras.items[i];
      const prev = paras.items[i - 1];
      if (!p.text.trim() && !prev.text.trim() && p.tableNestingLevel === 0 && prev.tableNestingLevel === 0) {
        p.delete();
        removed++;
      }
    }
    await ctx.sync();
    return `Đã gộp ${spaces.items.length} chỗ trống thừa và xoá ${removed} đoạn trống.`;
  });
}

function wordReplace(pairs, matchCase) {
  return Word.run(async (ctx) => {
    if (!confirm('Thay thế lần lượt từng cặp trên toàn bộ nội dung chính? Kết quả cặp trước có thể khớp cặp sau.')) return 'Đã huỷ.';
    let total = 0;
    for (const [find,replacement] of pairs) {
      const res = ctx.document.body.search(find,{matchCase}); res.load('items'); await ctx.sync();
      res.items.forEach(r => r.insertText(replacement,'Replace')); total += res.items.length;
      await ctx.sync();
    }
    return `Đã thay ${total} vị trí (${pairs.length} cặp).`;
  });
}

function wordStats() {
  return Word.run(async (ctx) => {
    const body = ctx.document.body;
    const sel = ctx.document.getSelection();
    body.load('text');
    sel.load('text');
    await ctx.sync();
    const all = textStats(body.text);
    const s = textStats(sel.text);
    const dl = $('#stats');
    const rows = [
      ['Số từ', all.words],
      ['Số ký tự (cả khoảng trắng)', all.chars],
      ['Số ký tự (không khoảng trắng)', all.charsNoSpaces],
      ['Số đoạn', all.paragraphs],
    ];
    if (sel.text) rows.push(['Số từ trong vùng chọn', s.words]);
    dl.replaceChildren();
    rows.forEach(([k, v]) => {
      const dt = document.createElement('dt');
      const dd = document.createElement('dd');
      dt.textContent = k;
      dd.textContent = v.toLocaleString('vi-VN');
      dl.append(dt, dd);
    });
    dl.hidden = false;
    return 'Đã cập nhật thống kê.';
  });
}

function wordInsertDate(long) {
  return Word.run(async (ctx) => {
    ctx.document.getSelection().insertText(formatDate(new Date(), long), 'Replace');
    await ctx.sync();
    return 'Đã chèn ngày.';
  });
}

function parsePairs(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.split('=>'))
    .filter((p) => p.length === 2 && p[0].trim())
    .map(([a, b]) => [a.trim(), b.trim()]);
}

/* ---------- Điều phối ---------- */

const ACTIONS = {
  ...Extra.actions,
  ...General.actions,
  trim: () => (HOST === 'Excel' ? excelTransform(TEXT_OPS.trim) : wordSelectionTransform(TEXT_OPS.trim)),
  upper: () => (HOST === 'Excel' ? excelTransform(TEXT_OPS.upper) : wordSelectionTransform(TEXT_OPS.upper)),
  lower: () => (HOST === 'Excel' ? excelTransform(TEXT_OPS.lower) : wordSelectionTransform(TEXT_OPS.lower)),
  proper: () => (HOST === 'Excel' ? excelTransform(TEXT_OPS.proper) : wordSelectionTransform(TEXT_OPS.proper)),
  noaccent: () => (HOST === 'Excel' ? excelTransform(TEXT_OPS.noaccent) : wordSelectionTransform(TEXT_OPS.noaccent)),
  'xl-words': () => excelNumbersToWords($('#xl-currency').checked),
  'xl-dup-mark': () => excelDuplicates('mark'),
  'xl-dup-delete': () => excelDuplicates('delete'),
  'xl-fx-check': () => excelCheckFormulas($('#fx-highlight').checked, $('#fx-manual').checked),
  'xl-fx-clear': () => excelClearFlags(),
  'wd-clean': () => wordClean(),
  'wd-replace': async () => {
    const pairs = parsePairs($('#wd-pairs').value);
    if (!pairs.length) return 'Nhập ít nhất một dòng dạng: cũ => mới';
    return wordReplace(pairs, $('#wd-case').checked);
  },
  'wd-stats': () => wordStats(),
  'wd-date-short': () => wordInsertDate(false),
  'wd-date-long': () => wordInsertDate(true),
};

async function run(btn) {
  const action = ACTIONS[btn.dataset.action];
  if (!action || busy) return;
  busy = true;
  document.querySelectorAll('[data-action]').forEach(b => b.disabled = true);
  setStatus('Đang xử lý…');
  try {
    setStatus(await action());
  } catch (e) {
    console.error(e);
    setStatus('Không thực hiện được: ' + (e.message || e), true);
  } finally {
    busy = false;
    document.querySelectorAll('[data-action]').forEach(b => b.disabled = false);
  }
}

Office.onReady((info) => {
  HOST = info.host === Office.HostType.Excel ? 'Excel' : info.host === Office.HostType.Word ? 'Word' : null;
  document.querySelectorAll('[data-host]').forEach((el) => { el.hidden = el.dataset.host !== HOST; });
  if (!HOST) {
    $('#host-label').textContent = 'Hãy mở tiện ích này bên trong Excel hoặc Word.';
    document.querySelectorAll('button').forEach((b) => { b.disabled = true; });
    return;
  }
  document.documentElement.dataset.host = HOST;
  Extra.init(HOST);
  $('#host-label').textContent = HOST === 'Excel' ? 'Đang chạy trong Excel' : 'Đang chạy trong Word';
  document.querySelectorAll('button[data-action]').forEach((b) => b.addEventListener('click', () => run(b)));
  $('#fx-results').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-addr]');
    if (!b) return;
    excelGoto(b.dataset.sheet, b.dataset.addr).catch((err) => setStatus('Không chọn được ô: ' + (err.message || err), true));
  });
});
