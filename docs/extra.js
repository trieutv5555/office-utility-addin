/* Additional tools: read-only audits and explicitly confirmed edits. */
const Extra = (() => {
  const $ = s => document.querySelector(s);
  const actions = {};
  const cap = 100000;
  const limit = 10000;
  let report = [];
  function show(items) {
    report = items;
    const box = $('#extra-results'); box.replaceChildren();
    const p = document.createElement('p'); p.textContent = `${items.length} kết quả (hiện tối đa 200; xuất CSV để xem đầy đủ).`; box.append(p);
    for (const it of items.slice(0,200)) {
      const b = document.createElement('button'); b.className='fx-item';
      b.textContent = `${it.sheet ? it.sheet+'!'+it.addr+' — ' : ''}${it.kind}: ${it.text}`;
      if (it.sheet) b.onclick = () => { if (!busy) excelGoto(it.sheet,it.addr).catch(e=>setStatus(e.message,true)); };
      else b.disabled=true;
      box.append(b);
    }
  }
  function csvCell(v) { let t=String(v??''); if (/^[\s]*[=+\-@]/.test(t)||/^[\t\r\n]/.test(t)) t="'"+t; return '"'+t.replace(/"/g,'""')+'"'; }
  actions['export-csv'] = async () => {
    if (!report.length) return 'Chưa có kết quả.';
    const text='\uFEFF'+[['Sheet','Ô','Loại','Nội dung'],...report.map(x=>[x.sheet,x.addr,x.kind,x.text])].map(r=>r.map(csvCell).join(',')).join('\r\n');
    const url=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'})); const a=document.createElement('a'); a.href=url; a.download='Bao-cao-tien-ich.csv'; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); return 'Đã yêu cầu tải báo cáo CSV.';
  };
  function audit(mode) { return Excel.run(async ctx=>{
    const all=$('#audit-all').checked;
    const term=$('#audit-query').value.trim().toLocaleLowerCase('vi-VN');
    if(mode==='search'&&!term) throw Error('Nhập từ khoá.');
    const sheets=ctx.workbook.worksheets; sheets.load('items/name,items/visibility');
    const active=ctx.workbook.worksheets.getActiveWorksheet(); active.load('name'); await ctx.sync();
    const items=[]; let skipped=0, scanned=0, truncated=false;
    for(const ws of sheets.items){
      if(ws.visibility!=='Visible'||(!all&&ws.name!==active.name)) continue;
      const r=all?ws.getUsedRangeOrNullObject(true):ctx.workbook.getSelectedRange();
      r.load(['rowCount','columnCount','rowIndex','columnIndex']); await ctx.sync();
      if(r.isNullObject) continue;
      if(r.rowCount*r.columnCount>cap){skipped++; continue;}
      const props=['values'];if(['search','blank','formula'].includes(mode))props.push('formulas');if(mode==='formula')props.push('formulasR1C1');if(mode==='errors')props.push('valueTypes');r.load(props);await ctx.sync();scanned++;
      const add=(i,j,kind,text)=>{if(items.length>=limit){truncated=true;return;} items.push({sheet:ws.name,addr:TextUtils.columnLetter(r.columnIndex+j)+(r.rowIndex+i+1),kind,text:String(text??'')});};
      if(mode==='duplicate'){
        const seen=new Map(); r.values.forEach((row,i)=>{if(row.every(v=>v===''))return;const key=JSON.stringify(row.map(v=>typeof v==='string'?v.trim().toLowerCase():v)); if(seen.has(key))add(i,0,'Dòng trùng',`Trùng dòng ${r.rowIndex+seen.get(key)+1}`);else seen.set(key,i);});
      } else if(mode==='formula'){
        for(let j=0;j<r.columnCount;j++){const res=TextUtils.findFormulaOutliers(r.formulasR1C1.map(row=>row[j]),{manual:true,blank:true,values:r.values.map(row=>row[j])}); if(!res.total)continue; if(res.status==='mixed'||res.status==='few')add(0,j,'Chưa đủ cơ sở',res.status==='mixed'?'Không có mẫu chiếm đa số':'Dưới 3 công thức'); if(res.status==='ok')res.outliers.forEach(o=>add(o.index,j,o.kind==='blank'?'Thiếu công thức':o.kind==='manual'?'Nhập tay':'Khác công thức',r.formulas[o.index][j]));}
      }else r.values.forEach((row,i)=>row.forEach((v,j)=>{if(mode==='search'&&(String(v).toLocaleLowerCase('vi-VN').includes(term)||String(r.formulas[i][j]).toLocaleLowerCase('vi-VN').includes(term)))add(i,j,'Tìm thấy',v); if(mode==='errors'&&r.valueTypes[i][j]==='Error')add(i,j,'Lỗi Excel',v); if(mode==='blank'&&v===''&&r.formulas[i][j]==='')add(i,j,'Ô trống','');}));
    }
    show(items); return `Đã quét ${scanned} sheet; bỏ qua ${skipped} vùng quá lớn; ${items.length} kết quả${truncated?' (đã chạm giới hạn 10.000)':''}.`;
  }); }
  for(const m of ['search','errors','blank','duplicate','formula']) actions['audit-'+m]=()=>audit(m);
  actions['xl-summary']=()=>Excel.run(async ctx=>{const r=ctx.workbook.getSelectedRange();r.load('rowCount,columnCount');await ctx.sync();if(r.rowCount*r.columnCount>cap)throw Error('Vùng quá lớn.');r.load('values');await ctx.sync();let sum=0,n=0,min=Infinity,max=-Infinity,blank=0;for(const row of r.values)for(const v of row){if(v==='')blank++;if(typeof v==='number'){n++;sum+=v;min=Math.min(min,v);max=Math.max(max,v);}}show([['Số ô số',n],['Tổng',sum],['Trung bình',n?sum/n:'—'],['Nhỏ nhất',n?min:'—'],['Lớn nhất',n?max:'—'],['Ô trống',blank]].map(([kind,text])=>({kind,text})));return 'Đã thống kê vùng chọn.';});
  actions['xl-format']=()=>Excel.run(async ctx=>{const r=ctx.workbook.getSelectedRange();r.load('rowCount,columnCount');await ctx.sync();if(r.rowCount*r.columnCount>cap)throw Error('Vùng quá lớn.');if(!await officeConfirm('Áp dụng Wrap Text, căn giữa theo chiều dọc và tự chỉnh chiều cao dòng?'))return 'Đã huỷ.';r.format.wrapText=true;r.format.verticalAlignment='Center';r.format.autofitRows();await ctx.sync();return 'Đã định dạng vùng chọn.';});
  actions['xl-autofit']=()=>Excel.run(async ctx=>{const r=ctx.workbook.getSelectedRange();if(!await officeConfirm('Tự chỉnh chiều rộng các cột đã chọn?'))return 'Đã huỷ.';r.format.autofitColumns();await ctx.sync();return 'Đã tự chỉnh chiều rộng cột.';});
  actions['xl-sheet-list']=()=>Excel.run(async ctx=>{const s=ctx.workbook.worksheets;s.load('items/name,items/visibility');await ctx.sync();show(s.items.map(w=>({sheet:w.visibility==='Visible'?w.name:'',addr:'A1',kind:w.visibility,text:w.name})));return `Có ${s.items.length} sheet.`;});
  actions['wd-find-count']=()=>Word.run(async ctx=>{const q=$('#wd-find-query').value.trim();if(!q)throw Error('Nhập từ khoá.');const r=ctx.document.body.search(q,{matchCase:$('#wd-case').checked});r.load('items');await ctx.sync();show([{kind:'Số vị trí tìm thấy',text:r.items.length}]);return `Tìm thấy ${r.items.length} vị trí.`;});
  actions['wd-insert-template']=()=>Word.run(async ctx=>{const t=$('#wd-template').value;if(!t.trim())throw Error('Nhập nội dung mẫu.');if(!await officeConfirm('Chèn mẫu tại vị trí con trỏ, không thay thế vùng chọn?'))return 'Đã huỷ.';ctx.document.getSelection().insertText(t,'End');await ctx.sync();return 'Đã chèn mẫu.';});
  function init(host){const search=$('#tool-search');search.addEventListener('input',()=>{const q=TextUtils.removeAccents(search.value.toLowerCase());document.querySelectorAll('main section').forEach(s=>{if(s.id==='report-section')return;s.hidden=(s.dataset.host&&s.dataset.host!==host)||!TextUtils.removeAccents(s.textContent.toLowerCase()).includes(q);});});document.querySelectorAll('main section:not(#report-section)').forEach(s=>{const h=s.querySelector('h2');if(!h)return;const d=document.createElement('details');d.open=false;const title=document.createElement('summary');title.textContent=h.textContent;h.remove();d.append(title,...Array.from(s.childNodes));s.append(d);});search.addEventListener('input',()=>{if(search.value)document.querySelectorAll('details').forEach(d=>d.open=true);});}
  return {actions,init,show,getReport:()=>report.slice()};
})();
