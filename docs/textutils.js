/* Hàm xử lý văn bản thuần, dùng chung cho Excel và Word (có thể test bằng Node). */
(function (root) {
  const DIGITS = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
  const UNITS = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ'];
  const LOCALE = 'vi-VN';

  function readGroup(n, full) {
    const tram = Math.floor(n / 100);
    const chuc = Math.floor((n % 100) / 10);
    const dv = n % 10;
    const parts = [];
    if (full || tram > 0) parts.push(DIGITS[tram], 'trăm');
    if (chuc > 1) {
      parts.push(DIGITS[chuc], 'mươi');
      if (dv === 1) parts.push('mốt');
      else if (dv === 5) parts.push('lăm');
      else if (dv > 0) parts.push(DIGITS[dv]);
    } else if (chuc === 1) {
      parts.push('mười');
      if (dv === 5) parts.push('lăm');
      else if (dv > 0) parts.push(DIGITS[dv]);
    } else if (dv > 0) {
      if (full || tram > 0) parts.push('lẻ');
      parts.push(DIGITS[dv]);
    }
    return parts.join(' ');
  }

  /** Đọc số nguyên (làm tròn) thành chữ tiếng Việt, tối đa 15 chữ số. */
  function numberToWords(value, withCurrency) {
    const raw=Number(value);
    let num=Math.sign(raw)*Math.round(Math.abs(raw));
    if (!Number.isFinite(num)) return '';
    if (Math.abs(num) > 999999999999999) return 'Số quá lớn';
    const neg = num < 0;
    num = Math.abs(num);
    let text;
    if (num === 0) {
      text = 'không';
    } else {
      const groups = [];
      while (num > 0) { groups.push(num % 1000); num = Math.floor(num / 1000); }
      const out = [];
      for (let i = groups.length - 1; i >= 0; i--) {
        if (groups[i] === 0) continue;
        const s = readGroup(groups[i], out.length > 0);
        out.push(UNITS[i] ? s + ' ' + UNITS[i] : s);
      }
      text = out.join(' ');
    }
    if (neg) text = 'âm ' + text;
    if (withCurrency) text += ' đồng';
    return text.charAt(0).toLocaleUpperCase(LOCALE) + text.slice(1);
  }

  function removeAccents(s) {
    return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').normalize('NFC');
  }

  const TEXT_OPS = {
    trim: (s) => s.replace(/[^\S\r\n]+/g, ' ').replace(/^ +| +$/gm, ''),
    upper: (s) => s.toLocaleUpperCase(LOCALE),
    lower: (s) => s.toLocaleLowerCase(LOCALE),
    proper: (s) => s.toLocaleLowerCase(LOCALE).replace(/(^|\s)(\S)/g, (m, a, b) => a + b.toLocaleUpperCase(LOCALE)),
    noaccent: removeAccents,
  };

  function formatDate(d, long) {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    return long ? `Ngày ${d.getDate()} tháng ${d.getMonth() + 1} năm ${d.getFullYear()}` : `${dd}/${mm}/${d.getFullYear()}`;
  }

  function textStats(text) {
    const t = text || '';
    const words = t.trim() ? t.trim().split(/\s+/).length : 0;
    return {
      words,
      chars: t.replace(/[\r\n]/g, '').length,
      charsNoSpaces: t.replace(/\s/g, '').length,
      paragraphs: t.trim() ? t.split(/\r|\n/).filter((p) => p.trim()).length : 0,
    };
  }

  /** Chỉ số cột (bắt đầu từ 0) -> chữ cái cột: 0 -> A, 27 -> AB. */
  function columnLetter(idx) {
    let n = idx + 1;
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /**
   * Tìm ô bất thường trong MỘT cột, đầu vào là công thức dạng R1C1 của từng ô
   * (kéo công thức xuống thì R1C1 giống hệt nhau).
   * - Công thức chiếm đa số (> 50% số công thức) là "chuẩn".
   * - Ô công thức khác chuẩn -> kind 'formula'.
   * - Ô nhập tay nằm giữa dòng công thức đầu và cuối -> kind 'manual' (nếu opts.manual).
   */
  function findFormulaOutliers(cells, opts) {
    const o = Object.assign({ manual: true, minFormulas: 3, minShare: 0.5 }, opts);
    const isFormula = (v,i) => o.formulaMask ? !!o.formulaMask[i] : typeof v === 'string' && v.startsWith('=') && (!o.values || v!==o.values[i]);
    const counts = new Map();
    let first = -1;
    let last = -1;
    let total = 0;
    cells.forEach((v, i) => {
      if (!isFormula(v,i)) return;
      total++;
      counts.set(v, (counts.get(v) || 0) + 1);
      if (first < 0) first = i;
      last = i;
    });
    if (total < o.minFormulas) return { status: 'few', total };
    let dom = null;
    let domCount = 0;
    counts.forEach((n, f) => { if (n > domCount) { dom = f; domCount = n; } });
    if (domCount / total <= o.minShare) return { status: 'mixed', total };
    const outliers = [];
    for (let i = first; i <= last; i++) {
      const v = cells[i];
      if (isFormula(v,i)) {
        if (v !== dom) outliers.push({ index: i, kind: 'formula' });
      } else if (o.blank && (v === '' || v === null || v === undefined)) {
        outliers.push({ index: i, kind: 'blank' });
      } else if (o.manual && v !== '' && v !== null && v !== undefined) {
        outliers.push({ index: i, kind: 'manual' });
      }
    }
    return { status: 'ok', total, domCount, sample: cells.findIndex((v,i)=>v===dom&&isFormula(v,i)), outliers };
  }

  // Conservative comparison: preserve punctuation, internal spaces and numeric types.
  function duplicateKey(row) {
    return JSON.stringify(row.map(v=>typeof v==='string'?v.normalize('NFC').trim().toLocaleLowerCase(LOCALE):v));
  }
  function safeSpreadsheetText(v) {
    return typeof v==='string' && (/^[\s\uFEFF]*[=+\-@]/.test(v)||/^[\t\r\n]/.test(v)) ? "'"+v : v;
  }
  function csvCell(v) {
    const text=safeSpreadsheetText(String(v??''));
    return '"'+text.replace(/"/g,'""')+'"';
  }
  // Whole rows for duplicate checks; whole columns for R1C1 comparison.
  function scanBlocks(rows,cols,mode,target=20000) {
    if(!Number.isInteger(rows)||!Number.isInteger(cols)||rows<1||cols<1)throw Error('Invalid range dimensions');
    const blocks=[];
    if(mode==='formula'){
      const width=Math.max(1,Math.floor(target/rows));
      for(let j=0;j<cols;j+=width)blocks.push({r:0,c:j,h:rows,w:Math.min(width,cols-j)});
    }else{
      const height=Math.max(1,Math.floor(target/cols));
      for(let i=0;i<rows;i+=height)blocks.push({r:i,c:0,h:Math.min(height,rows-i),w:cols});
    }
    return blocks;
  }
  // Rectangle coordinates are absolute worksheet indexes, not A1 strings.
  function formulaMask(rows,cols,rowIndex,columnIndex,areas){
    const mask=Array.from({length:rows},()=>new Uint8Array(cols));
    for(const a of areas){
      const r0=Math.max(0,a.rowIndex-rowIndex),r1=Math.min(rows,a.rowIndex-rowIndex+a.rowCount);
      const c0=Math.max(0,a.columnIndex-columnIndex),c1=Math.min(cols,a.columnIndex-columnIndex+a.columnCount);
      for(let i=r0;i<r1;i++)if(c1>c0)mask[i].fill(1,c0,c1);
    }
    return mask;
  }
  const api = { formulaMask, duplicateKey, safeSpreadsheetText, csvCell, scanBlocks, numberToWords, removeAccents, TEXT_OPS, formatDate, textStats, columnLetter, findFormulaOutliers };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.TextUtils = api;
})(typeof window !== 'undefined' ? window : globalThis);
