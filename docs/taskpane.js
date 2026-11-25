
/* In-pane confirmation: native browser dialogs are unavailable in Office. */
function officeConfirm(message) {
  return new Promise(resolve => {
    const overlay=document.createElement('div');
    overlay.style.cssText='position:fixed;inset:0;z-index:2147483647;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;padding:12px;';
    const panel=document.createElement('div');
    panel.style.cssText='background:#fff;color:#222;padding:16px;border-radius:8px;width:100%;max-width:340px;box-shadow:0 4px 24px #0004;';
    panel.setAttribute('role','alertdialog');panel.setAttribute('aria-modal','true');panel.setAttribute('aria-label','Xác nhận thao tác');
    const title=document.createElement('strong');title.textContent='Xác nhận thao tác';
    const p=document.createElement('p');p.textContent=message;
    const buttons=document.createElement('div');buttons.style.cssText='display:flex;gap:8px;justify-content:flex-end;';
    const cancel=document.createElement('button'),ok=document.createElement('button');
    cancel.type=ok.type='button';cancel.textContent='Huỷ';ok.textContent='Tiếp tục';ok.style.cssText='background:#107c41;color:white;';
    const previous=document.activeElement;
    function finish(value){document.removeEventListener('keydown',key,true);overlay.remove();if(previous&&previous.focus&&!previous.disabled)previous.focus();resolve(value);}
    function key(e){if(e.key==='Escape'){e.preventDefault();e.stopPropagation();finish(false);}else if(e.key==='Tab'){e.preventDefault();(document.activeElement===cancel?ok:cancel).focus();}}
    cancel.addEventListener('click',()=>finish(false));ok.addEventListener('click',()=>finish(true));
    buttons.append(cancel,ok);panel.append(title,p,buttons);overlay.append(panel);document.body.append(overlay);document.addEventListener('keydown',key,true);cancel.focus();
  });
}
/* Tiện ích văn phòng – taskpane dùng chung cho Excel và Word (Office.js). */
const { numberToWords, TEXT_OPS, columnLetter, findFormulaOutliers } = TextUtils;
let HOST = null;
let busy = false;
const MAX_CELLS = 100000;
function guard(r) { if (r.rowCount * r.columnCount > MAX_CELLS) throw new Error("Vùng quá lớn; hãy chọn tối đa 100.000 ô mỗi lần."); }

const $ = (sel) => document.querySelector(sel);
async function loadBounded(ctx,range,props){
  range.load('rowCount,columnCount');await ctx.sync();
  if(range.isNullObject)return false;
  guard(range);range.load(props);await ctx.sync();return true;
}
function excelSupports(version){return Office.context.requirements.isSetSupported('ExcelApi',version);}
async function getFormulaMask(ctx,range){
  if(!excelSupports('1.9'))return null;
  const matches=range.getSpecialCellsOrNullObject('Formulas');
  matches.load('areaCount');
  range.load('rowIndex,columnIndex,rowCount,columnCount');
  await ctx.sync();
  if(matches.isNullObject)return TextUtils.formulaMask(range.rowCount,range.columnCount,range.rowIndex,range.columnIndex,[]);
  matches.areas.load('items/rowIndex,items/columnIndex,items/rowCount,items/columnCount');
  await ctx.sync();
  return TextUtils.formulaMask(range.rowCount,range.columnCount,range.rowIndex,range.columnIndex,matches.areas.items);
}
// Refuse writes if the selected/target range intersects a spill or saved array.
async function protectArrays(ctx,range,required=false){
  if(!excelSupports('1.12')){
    if(required)throw Error('Công cụ này cần ExcelApi 1.12 để bảo vệ công thức mảng. Không thay đổi dữ liệu.');
    return;
  }
  range.load('hasSpill,savedAsArray');await ctx.sync();
  if(range.hasSpill!==false||range.savedAsArray!==false)throw Error('Vùng có hoặc giao với công thức mảng/tràn. Hãy chọn vùng khác; không thay đổi dữ liệu.');
}
async function assertUnchanged(ctx,range){
  const snapshot=JSON.stringify([range.values,range.formulas]);
  range.load('values,formulas');await ctx.sync();
  if(snapshot!==JSON.stringify([range.values,range.formulas]))throw Error('Dữ liệu đã thay đổi. Hãy chạy lại tiện ích.');
}
function writeEdits(range,edits,textFormat=false){
  const sorted=[...edits].sort((a,b)=>a[1]-b[1]||a[0]-b[0]);
  for(let i=0;i<sorted.length;){
    const [row,col]=sorted[i];let end=i+1;
    while(end<sorted.length&&sorted[end][1]===col&&sorted[end][0]===sorted[end-1][0]+1)end++;
    const target=range.getCell(row,col).getResizedRange(end-i-1,0);
    if(textFormat)target.numberFormat=Array.from({length:end-i},()=>['@']);
    target.values=sorted.slice(i,end).map(x=>[TextUtils.safeSpreadsheetText(x[2])]);
    i=end;
  }
}

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
    if(!await loadBounded(ctx,range,['values','formulas','address','valueTypes']))return 'Vùng chọn đang trống.';
    const mask=await getFormulaMask(ctx,range);
    let changed = 0;
    const edits = [];
    range.formulas.forEach((row,r) => row.forEach((v,c) => {
      if (typeof v !== 'string' || (mask?mask[r][c]:v.startsWith('=')) || range.valueTypes[r][c] !== 'String') return;
      const n = fn(v);
      if (n !== v) edits.push([r,c,n]);
    }));
    if (edits.length > 5000) throw new Error('Chia nhỏ vùng chọn: tối đa 5.000 ô sửa mỗi lần.');
    if (edits.length && !await officeConfirm(`Sửa ${edits.length} ô văn bản? Hãy sao lưu trước khi thực hiện.`)) return 'Đã huỷ.';
    if(edits.length){await protectArrays(ctx,range);await assertUnchanged(ctx,range);writeEdits(range,edits);changed=edits.length;}
    await ctx.sync();
    return changed ? `Đã sửa ${changed} ô (${range.address}).` : 'Không có ô nào cần sửa.';
  });
}

function excelNumbersToWords(withCurrency) {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    if(!await loadBounded(ctx,range,['values','formulas','columnIndex']))return 'Vùng chọn đang trống.';
    if (range.columnCount > 1) throw new Error('Chỉ chọn một cột số để tránh ghi đè dữ liệu bên cạnh.');
    if(range.columnIndex>=16383)throw Error('Không còn cột bên phải.');
    const dest = range.getOffsetRange(0,1);
    dest.load(['values','formulas']);
    await ctx.sync();
    const rows = range.values.map((row,r) => [row[0],r]).filter(([v]) => typeof v === 'number');
    if (rows.some(([,r]) => dest.values[r][0] !== '' || dest.formulas[r][0] !== '')) throw new Error('Cột kết quả có dữ liệu hoặc công thức. Hãy để trống các ô đích.');
    if (rows.length > 5000) throw new Error('Tối đa 5.000 số mỗi lần.');
    if(rows.some(([v])=>!Number.isFinite(v)||Math.round(Math.abs(v))>999999999999999))throw Error('Có số vượt giới hạn 15 chữ số sau làm tròn.');
    if (rows.length && !await officeConfirm(`Ghi ${rows.length} kết quả vào cột bên phải? Số lẻ được làm tròn.`)) return 'Đã huỷ.';
    if(rows.length){await protectArrays(ctx,dest);await assertUnchanged(ctx,range);await assertUnchanged(ctx,dest);writeEdits(dest,rows.map(([v,r])=>[r,0,numberToWords(v,withCurrency)]));}
    const count = rows.length;
    await ctx.sync();
    return count ? `Đã đọc ${count} số thành chữ.` : 'Không tìm thấy ô số nào trong vùng chọn.';
  });
}

function renderDuplicateRows(items){
  const box=$('#duplicate-results');box.replaceChildren();
  const note=document.createElement('p');note.className='hint';
  note.textContent=items.length?'Đã tô '+items.length+' dòng trùng. Bấm vào dòng để tới ô đã tô.':'Không có dòng trùng.';
  box.append(note);
  for(const it of items.slice(0,200)){
    const btn=document.createElement('button');btn.type='button';btn.className='fx-item';
    btn.textContent=it.text;btn.title=it.sheet+'!'+it.addr;
    btn.onclick=()=>{if(!busy)excelGoto(it.sheet,it.addr).catch(e=>setStatus(e.message,true));};
    box.append(btn);
  }
  if(items.length>200){
    const more=document.createElement('p');more.className='hint';
    more.textContent='Hiện 200/'+items.length+' dòng. Xem đầy đủ bằng Xuất CSV hoặc Tạo sheet báo cáo ngay sau khi chạy.';
    box.append(more);
  }
  Extra.show(items);
}
function excelDuplicates(mode) {
  return Excel.run(async (ctx) => {
    if(mode!=='mark')return 'Đã vô hiệu hoá xoá cả hàng để bảo vệ dữ liệu ngoài vùng chọn.';
    const range=ctx.workbook.getSelectedRange().getUsedRangeOrNullObject(true);
    const sheet=ctx.workbook.worksheets.getActiveWorksheet();
    sheet.load('name');
    $('#duplicate-results').replaceChildren();
    Extra.show([]);
    if(!await loadBounded(ctx,range,['values','formulas','rowIndex','columnIndex']))return 'Vùng chọn đang trống.';
    const seen=new Map(),dups=[];
    range.values.forEach((row,i)=>{
      if(row.every(v=>v===''))return;
      const key=TextUtils.duplicateKey(row);
      if(seen.has(key))dups.push({i,first:seen.get(key)});
      else seen.set(key,i);
    });
    if(!dups.length){renderDuplicateRows([]);return 'Không có dòng trùng.';}
    if(dups.length>2000)throw Error('Quá 2.000 dòng trùng; dùng báo cáo thay cho tô màu.');
    if(!await officeConfirm('Tô màu '+dups.length+' dòng trùng? Màu nền cũ sẽ bị thay đổi.'))return 'Đã huỷ.';
    await assertUnchanged(ctx,range);
    for(const {i} of dups)range.getRow(i).format.fill.color='#FFF2A8';
    await ctx.sync();
    const col=columnLetter(range.columnIndex);
    const items=dups.map(({i,first})=>{
      const row=range.rowIndex+i+1,original=range.rowIndex+first+1;
      return {sheet:sheet.name,addr:col+row,kind:'Dòng trùng đã tô',text:'Dòng '+row+' trùng dòng '+original};
    });
    renderDuplicateRows(items);
    return 'Đã tô vàng '+items.length+' dòng trùng. Xem danh sách ngay dưới nút.';
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
    if(!await loadBounded(ctx,used,['values','formulas','formulasR1C1','rowIndex','columnIndex'])){renderFormulaReport([],sheet.name);Extra.show([]);return 'Vùng chọn đang trống.';}

    const mask=await getFormulaMask(ctx,used);
    const reports = [];
    const flagged = [];
    for (let c = 0; c < used.columnCount; c++) {
      const letter = columnLetter(used.columnIndex + c);
      const res = findFormulaOutliers(used.formulasR1C1.map((r) => r[c]), { manual, blank: true, values:used.values.map(r=>r[c]), formulaMask:mask?mask.map(r=>r[c]):null });
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
      renderFormulaReport([], sheet.name);Extra.show([]);
      return 'Vùng chọn không có công thức nào.';
    }
    if (highlight && flagged.length) {
      if (lastFlagged) throw new Error('Hãy hoàn nguyên màu lần kiểm tra trước.');
      if (flagged.length > 2000) throw new Error('Quá 2.000 ô lỗi; hãy tắt tô màu.');
      if(!await officeConfirm('Tô màu '+flagged.length+' ô lỗi? Chỉ phục hồi gần đúng màu nền trong phiên này.'))return 'Đã huỷ.';
      await assertUnchanged(ctx,used);
      const fills = flagged.map(a => sheet.getRange(a).format.fill);
      fills.forEach(f => f.load('color'));
      await ctx.sync();
      const colors = fills.map(f => f.color);
      flagged.forEach((a) => { sheet.getRange(a).format.fill.color = '#FFC7CE'; });
      await ctx.sync();
      lastFlagged = { sheet: sheet.name, addrs: flagged, colors };
    }
    await ctx.sync();
    renderFormulaReport(reports, sheet.name);
    const combined=reports.flatMap(rep=>(rep.items||[]).map(it=>({sheet:sheet.name,addr:it.addr,kind:it.kind==='blank'?'Thiếu công thức':it.kind==='manual'?'Nhập tay':'Khác công thức',text:it.text})));
    Extra.show(combined.slice(0,10000),combined.length);
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
    const original=sel.text;
    const nativeCase=fn===TEXT_OPS.upper?'Upper':fn===TEXT_OPS.lower?'Lower':null;
    if(nativeCase&&Office.context.requirements.isSetSupported('WordApiDesktop','1.4')){
      if(!await officeConfirm('Chuyển kiểu chữ bằng chức năng gốc của Word?'))return 'Đã huỷ.';
      sel.load('text');await ctx.sync();if(sel.text!==original)throw Error('Vùng chọn đã thay đổi. Hãy chạy lại.');
      sel.case=nativeCase;await ctx.sync();return 'Đã chuyển kiểu chữ bằng Word, không thay toàn bộ đoạn.';
    }
    if (!await officeConfirm('Thay văn bản đã chọn có thể làm mất định dạng từng phần. Tiếp tục?')) return 'Đã huỷ.';
    sel.load('text');await ctx.sync();if(sel.text!==original)throw Error('Vùng chọn đã thay đổi. Hãy chạy lại.');
    sel.insertText(next, 'Replace');
    await ctx.sync();
    return 'Đã cập nhật đoạn văn bản đang chọn.';
  });
}

function wordClean() {
  return Word.run(async (ctx) => {
    if(!Office.context.requirements.isSetSupported('WordApi','1.3'))throw Error('Cần Word hỗ trợ WordApi 1.3 để dọn tài liệu an toàn.');
    if (!await officeConfirm('Dọn khoảng trắng và đoạn trống toàn tài liệu? Có thể ảnh hưởng bố cục.')) return 'Đã huỷ.';
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
    if (!await officeConfirm('Thay thế lần lượt từng cặp trên toàn bộ nội dung chính? Kết quả cặp trước có thể khớp cặp sau.')) return 'Đã huỷ.';
    let total = 0;
    for (const [find,replacement] of pairs) {
      const res = ctx.document.body.search(find,{matchCase}); res.load('items'); await ctx.sync();
      res.items.forEach(r => r.insertText(replacement,'Replace')); total += res.items.length;
      await ctx.sync();
    }
    return `Đã thay ${total} vị trí (${pairs.length} cặp).`;
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
  'xl-fx-check': () => excelCheckFormulas($('#fx-highlight').checked, $('#fx-manual').checked),
  'xl-fx-clear': () => excelClearFlags(),
  'wd-clean': () => wordClean(),
  'wd-replace': async () => {
    const pairs = parsePairs($('#wd-pairs').value);
    if (!pairs.length) return 'Nhập ít nhất một dòng dạng: cũ => mới';
    return wordReplace(pairs, $('#wd-case').checked);
  },
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
    if (!b || busy) return;
    excelGoto(b.dataset.sheet, b.dataset.addr).catch((err) => setStatus('Không chọn được ô: ' + (err.message || err), true));
  });
});
