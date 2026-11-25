/* Additional tools: read-only audits and explicitly confirmed edits. */
const Extra = (() => {
  const $ = s => document.querySelector(s);
  const actions = {};
  const cap = 100000;
  const limit = 10000;
  let report = [];
  function show(items,total=items.length) {
    report = items;
    const box = $('#extra-results'); box.replaceChildren();
    const p = document.createElement('p'); p.textContent = `${total} kết quả; lưu ${items.length}, hiện tối đa 200. CSV chỉ xuất các kết quả đã lưu.`; box.append(p);
    for (const it of items.slice(0,200)) {
      const b = document.createElement('button'); b.className='fx-item';
      b.textContent = `${it.sheet ? it.sheet+'!'+it.addr+' — ' : ''}${it.kind}: ${it.text}`;
      if (it.sheet) b.onclick = () => { if (!busy) excelGoto(it.sheet,it.addr).catch(e=>setStatus(e.message,true)); };
      else b.disabled=true;
      box.append(b);
    }
  }
  const csvCell=TextUtils.csvCell;
  actions['export-csv'] = async () => {
    if (!report.length) return 'Chưa có kết quả.';
    const text='\uFEFF'+[['Sheet','Ô','Loại','Nội dung'],...report.map(x=>[x.sheet,x.addr,x.kind,x.text])].map(r=>r.map(csvCell).join(',')).join('\r\n');
    const url=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8'})); const a=document.createElement('a'); a.href=url; a.download='Bao-cao-tien-ich.csv'; document.body.append(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000); return 'Đã yêu cầu tải báo cáo CSV.';
  };
  // Read bounded blocks; do not silently drop a sheet just because it exceeds 100k cells.
  function audit(mode) { return Excel.run(async ctx=>{
    const all=$('#audit-all').checked;
    const term=$('#audit-query').value.trim().normalize('NFC').toLocaleLowerCase('vi-VN');
    if(mode==='search'&&!term) throw Error('Nhập từ khoá.');
    const sheets=ctx.workbook.worksheets; sheets.load('items/name,items/visibility');
    const active=ctx.workbook.worksheets.getActiveWorksheet(); active.load('name'); await ctx.sync();
    const targets=sheets.items.filter(ws=>ws.visibility==='Visible'&&(all||ws.name===active.name)).map(ws=>{
      const range=all?ws.getUsedRangeOrNullObject(true):ctx.workbook.getSelectedRange();
      range.load('rowCount,columnCount,rowIndex,columnIndex');return {ws,range};
    });
    await ctx.sync();
    const items=[];let scanned=0,matched=0,cells=0;const skipped=[];
    const sheetCap=1000000, workbookCap=5000000;
    const props=['values'];
    if(['search','blank','formula'].includes(mode))props.push('formulas');
    if(mode==='formula')props.push('formulasR1C1');
    if(mode==='errors')props.push('valueTypes');
    for(const {ws,range} of targets){
      if(range.isNullObject)continue;
      const count=range.rowCount*range.columnCount;
      // R1C1 scans need complete columns, otherwise the majority pattern changes at block boundaries.
      if(count>sheetCap||cells+count>workbookCap||(mode==='formula'&&range.rowCount>100000)){
        skipped.push(ws.name);continue;
      }
      const seen=mode==='duplicate'?new Map():null;
      const columns=Array.from({length:range.columnCount},(_,j)=>TextUtils.columnLetter(range.columnIndex+j));
      const blocks=TextUtils.scanBlocks(range.rowCount,range.columnCount,mode);
      const add=(i,j,kind,text)=>{
        matched++;if(items.length>=limit)return;
        items.push({sheet:ws.name,addr:columns[j]+(range.rowIndex+i+1),kind,text:String(text??'')});
      };
      setStatus('Đang rà soát '+ws.name+' ('+(scanned+1)+'/'+targets.length+')…');
      for(const b of blocks){
        const p=ws.getRangeByIndexes(range.rowIndex+b.r,range.columnIndex+b.c,b.h,b.w);
        p.load(props);await ctx.sync();
        if(mode==='duplicate'){
          p.values.forEach((row,i)=>{
            if(row.every(v=>v===''))return;
            const key=TextUtils.duplicateKey(row),idx=b.r+i;
            if(seen.has(key))add(idx,0,'Dòng trùng','Trùng dòng '+(range.rowIndex+seen.get(key)+1));
            else seen.set(key,idx);
          });
        }else if(mode==='formula'){
          const mask=await getFormulaMask(ctx,p);
          for(let j=0;j<b.w;j++){
            const res=TextUtils.findFormulaOutliers(p.formulasR1C1.map(row=>row[j]),{manual:true,blank:true,values:p.values.map(row=>row[j]),formulaMask:mask?mask.map(row=>row[j]):null});
            if(!res.total)continue;
            if(res.status==='mixed'||res.status==='few')add(0,b.c+j,'Chưa đủ cơ sở',res.status==='mixed'?'Không có mẫu chiếm đa số':'Dưới 3 công thức');
            if(res.status==='ok')res.outliers.forEach(o=>add(o.index,b.c+j,o.kind==='blank'?'Thiếu công thức':o.kind==='manual'?'Nhập tay':'Khác công thức',p.formulas[o.index][j]));
          }
        }else{
          p.values.forEach((row,i)=>row.forEach((v,j)=>{
            const ri=b.r+i,cj=b.c+j;
            if(mode==='search'&&(String(v).normalize('NFC').toLocaleLowerCase('vi-VN').includes(term)||String(p.formulas[i][j]).normalize('NFC').toLocaleLowerCase('vi-VN').includes(term)))add(ri,cj,'Tìm thấy',v);
            if(mode==='errors'&&p.valueTypes[i][j]==='Error')add(ri,cj,'Lỗi Excel',v);
            if(mode==='blank'&&v===''&&p.formulas[i][j]==='')add(ri,cj,'Ô trống','');
          }));
        }
      }
      scanned++;cells+=count;
    }
    show(items,matched);
    return 'Đã quét '+scanned+' sheet, '+cells.toLocaleString('vi-VN')+' ô; '+matched+' kết quả'+(matched>limit?' (chỉ lưu 10.000)':'')+(skipped.length?'; bỏ qua do giới hạn: '+skipped.join(', '):'')+'.';
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
