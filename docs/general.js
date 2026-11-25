/* General-purpose tools. No company-specific rules. */
const General = (() => {
 const a={}, $=s=>document.querySelector(s), val=id=>$('#'+id).value.trim();
 const yes=s=>officeConfirm(s+'? Hãy lưu bản sao trước khi thực hiện.');
 const cap=100000, safe=TextUtils.safeSpreadsheetText;
 // Formula strings can have text results; distinguish literal text beginning with '=' via value equality.
 const isFx=(f,v)=>typeof f==='string'&&f.startsWith('=')&&f!==v;
 function parseNumber(v,mode){if(typeof v!=='string')return null;let s=v.trim();if(!s)return null;const vi=mode==='vi';const re=vi?/^[+-]?(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d+)?$/:/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/;if(!re.test(s)||/^[+-]?0\d/.test(s))return null; s=vi?s.replace(/\./g,'').replace(',','.'):s.replace(/,/g,'');if(s.replace(/[^0-9]/g,'').length>15)return null;const n=Number(s);return Number.isFinite(n)?n:null;}
 async function read(c,r){r.load('rowCount,columnCount,rowIndex,columnIndex,address');await c.sync();if(r.rowCount*r.columnCount>cap)throw Error('Vùng vượt 100.000 ô.');r.load('values,formulas,valueTypes');await c.sync();return r;}
 const sel=(fn)=>()=>Excel.run(async c=>fn(c,await read(c,c.workbook.getSelectedRange())));
 function edit(mode){return sel(async(c,r)=>{
 const mask=await getFormulaMask(c,r),edits=[],numberMode=val('g-number-mode');
 for(let j=0;j<r.columnCount;j++){let last;for(let i=0;i<r.rowCount;i++){
   const v=r.values[i][j],f=r.formulas[i][j],fx=mask?!!mask[i][j]:isFx(f,v),err=r.valueTypes[i][j]==='Error';
   let n=null;
   if(mode==='fill'){if(err)last=undefined;else if(!fx&&v===''&&f===''&&last!==undefined)n=last;else if(v!=='')last=v;else if(fx)last=undefined;}
   else if(!fx&&!err&&!(typeof f==='string'&&f.startsWith('=')&&!mask))n=parseNumber(v,numberMode);
   if(n!==null&&n!==undefined&&n!==v)edits.push([i,j,n]);
 }}
 if(!edits.length)return 'Không có ô phù hợp.';
 if(edits.length>5000)throw Error('Tối đa 5.000 ô sửa mỗi lần.');
 if(!await yes('Sửa '+edits.length+' ô'))return 'Đã huỷ.';
 await protectArrays(c,r);await assertUnchanged(c,r);writeEdits(r,edits);await c.sync();return 'Đã sửa '+edits.length+' ô.';
 });}
 for(const m of ['number','fill'])a['g-'+m]=edit(m);
 async function dest(c,ws,addr,rows,cols){if(!/^[A-Z]{1,3}[1-9]\d*$/i.test(addr))throw Error('Nhập ô đích đơn, ví dụ H2.');if(rows*cols>cap)throw Error('Kết quả quá lớn.');const r=await read(c,ws.getRange(addr).getResizedRange(rows-1,cols-1));if(r.values.some((row,i)=>row.some((v,j)=>v!==''||r.formulas[i][j]!=='')))throw Error('Vùng đích không trống.');return r;}
 a['g-split']=sel(async(c,r)=>{if(r.columnCount!==1)throw Error('Chọn đúng một cột.');const sep=$('#g-sep').value;if(!sep)throw Error('Nhập ký tự phân tách.');const out=r.values.map(row=>String(row[0]).split(sep));const width=Math.max(...out.map(x=>x.length));const t=await dest(c,r.worksheet,val('g-dest'),out.length,width);if(!await yes('Ghi kết quả tách vào '+t.address))return 'Đã huỷ.';await protectArrays(c,t);await assertUnchanged(c,r);await assertUnchanged(c,t);t.values=out.map(row=>Array.from({length:width},(_,i)=>safe(row[i]??'')));await c.sync();return 'Đã tách cột.';});
 a['g-join']=sel(async(c,r)=>{const sep=$('#g-sep').value;const t=await dest(c,r.worksheet,val('g-dest'),r.rowCount,1);if(!await yes('Ghi kết quả ghép vào '+t.address))return 'Đã huỷ.';await protectArrays(c,t);await assertUnchanged(c,r);await assertUnchanged(c,t);t.values=r.values.map(row=>[safe(row.join(sep))]);await c.sync();return 'Đã ghép cột.';});
 const item=(r,i,j,kind,text)=>({sheet:r.address.slice(0,r.address.lastIndexOf('!')).replace(/^'|'$/g,'').replace(/''/g,"'"),addr:TextUtils.columnLetter(r.columnIndex+j)+(r.rowIndex+i+1),kind,text:String(text)});
 function display(items){Extra.show(items.slice(0,10000),items.length);return `${items.length} kết quả${items.length>10000?' (chỉ lưu 10.000)':''}.`;}
 a['g-compare']=sel(async(c,r)=>{const name=val('g-other-sheet');const ws=name?c.workbook.worksheets.getItem(name):r.worksheet;const addr=val('g-other-range');if(!/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/i.test(addr))throw Error('Nhập vùng A1 hợp lệ.');const b=await read(c,ws.getRange(addr));if(b.rowCount!==r.rowCount||b.columnCount!==r.columnCount)throw Error('Hai vùng phải cùng kích thước.');let out=[];r.values.forEach((row,i)=>row.forEach((v,j)=>{if(v!==b.values[i][j]||r.valueTypes[i][j]!==b.valueTypes[i][j])out.push(item(r,i,j,'Khác giá trị',`${v} ≠ ${b.values[i][j]}`));}));return display(out);});
 a['g-sequence']=sel(async(c,r)=>{if(r.columnCount!==1)throw Error('Chọn một cột số thứ tự, không gồm tiêu đề.');const seen=new Set(),out=[];let prev=null;r.values.forEach(([v],i)=>{if(typeof v!=='number'||!Number.isSafeInteger(v)){out.push(item(r,i,0,'Không phải số nguyên',v));return;}if(seen.has(v))out.push(item(r,i,0,'Trùng số',v));seen.add(v);if(prev!==null&&v>prev+1)out.push(item(r,i,0,'Khoảng nhảy',`${prev+1}–${v-1}`));if(prev!==null&&v<prev)out.push(item(r,i,0,'Giảm thứ tự',v));prev=v;});return display(out);});
 const sheets=async c=>{const s=c.workbook.worksheets;s.load('items/name,items/visibility,items/position');await c.sync();return s;};
 const requested=s=>{const names=val('g-sheets').split(/\r?\n/).map(x=>x.trim()).filter(Boolean);if(!names.length)throw Error('Nhập tên sheet, mỗi dòng một tên.');const list=[...new Set(names)].map(n=>s.items.find(w=>w.name===n));if(list.some(x=>!x))throw Error('Có tên sheet không tồn tại.');return list;};
 function newSheet(s,base){let n=base,k=1;const used=new Set(s.items.map(w=>w.name.toLowerCase()));while(used.has(n.toLowerCase()))n=base.slice(0,25)+' '+(++k);return s.add(n);}
 function link(ws,i,j,name,addr,label){ws.getCell(i,j).hyperlink={documentReference:"'"+name.replace(/'/g,"''")+"'!"+addr,textToDisplay:label};}
 function table(ws,rows){const r=ws.getRangeByIndexes(0,0,rows.length,rows[0].length);r.values=rows.map(row=>row.map(safe));r.format.wrapText=true;r.format.verticalAlignment='Center';ws.getRangeByIndexes(0,0,1,rows[0].length).format.font.bold=true;}
 a['g-report']=()=>Excel.run(async c=>{const list=Extra.getReport();if(!list.length)throw Error('Chưa có kết quả rà soát.');const s=await sheets(c);if(!await yes('Tạo sheet báo cáo mới'))return 'Đã huỷ.';const w=newSheet(s,'Báo cáo tiện ích');table(w,[['Sheet','Ô','Loại','Nội dung'],...list.map(x=>[x.sheet||'',x.addr||'',x.kind,x.text])]);list.forEach((x,i)=>{if(x.sheet&&/^[A-Z]+\d+$/.test(x.addr)&&s.items.some(w=>w.name===x.sheet))link(w,i+1,1,x.sheet,x.addr,x.addr);});w.activate();await c.sync();return 'Đã tạo báo cáo.';});
 a['g-index']=()=>Excel.run(async c=>{const s=await sheets(c);if(!await yes('Tạo mục lục sheet mới'))return 'Đã huỷ.';const w=newSheet(s,'Mục lục');table(w,[['Sheet','Trạng thái'],...s.items.map(x=>[x.name,x.visibility])]);s.items.forEach((x,i)=>{if(x.visibility==='Visible')link(w,i+1,0,x.name,'A1',x.name);});w.activate();await c.sync();return 'Đã tạo mục lục.';});
 a['g-merge']=()=>Excel.run(async c=>{
   const s=await sheets(c),list=requested(s);
   if(list.some(w=>w.visibility!=='Visible'))throw Error('Chỉ gộp sheet hiện.');
   const ranges=list.map(w=>{const r=w.getUsedRangeOrNullObject(true);r.load('rowCount,columnCount');return {w,r};});
   await c.sync();
   const valid=ranges.filter(x=>!x.r.isNullObject);
   if(!valid.length)return 'Không có dữ liệu.';
   const width=Math.max(...valid.map(x=>x.r.columnCount)),height=valid.reduce((n,x)=>n+x.r.rowCount,0);
   if((height+1)*(width+1)>cap)throw Error('Kết quả quá 100.000 ô.');
   if(!await yes('Gộp giá trị vào sheet mới, giữ tất cả dòng và thêm cột nguồn'))return 'Đã huỷ.';
   const rows=[['Sheet nguồn',...Array.from({length:width},(_,i)=>'Cột '+(i+1))]];
   // Read one source at a time; avoid oversized requests with many worksheets.
   for(const {w,r} of valid){r.load('values');await c.sync();for(const row of r.values)rows.push([w.name,...Array.from({length:width},(_,j)=>row[j]??'')]);}
   const w=newSheet(s,'Dữ liệu gộp');table(w,rows);w.activate();await c.sync();return 'Đã gộp '+height+' dòng.';
 });
 a['g-refs']=sel(async(c,r)=>{const mask=await getFormulaMask(c,r);let out=[];const expected=val('g-expected');r.formulas.forEach((row,i)=>row.forEach((f,j)=>{if(!(mask?mask[i][j]:isFx(f,r.values[i][j])))return;const stripped=f.replace(/"(?:[^"]|"")*"/g,'');if(/\[[^\]]+\][^!]*!/.test(stripped))out.push(item(r,i,j,'Có thể có liên kết ngoài',f));const refs=[...stripped.matchAll(/(?:'((?:[^']|'')+)'|([\p{L}\p{N}_.]+))!/gu)].map(m=>(m[1]||m[2]).replace(/''/g,"'"));if(expected&&refs.some(n=>n.toLowerCase()!==expected.toLowerCase()))out.push(item(r,i,j,'Tham chiếu khác sheet dự kiến',f));}));return display(out);});
 a['g-word-money']=()=>Word.run(async c=>{const r=c.document.getSelection();r.load('text');await c.sync();const n=parseNumber(r.text,val('g-word-number'));if(n===null||!Number.isSafeInteger(n))throw Error('Chọn một số tiền nguyên hợp lệ.');const text=TextUtils.numberToWords(n,true);if(Math.abs(n)>999999999999999)throw Error('Số vượt giới hạn 15 chữ số.');if(!await yes('Chèn số tiền bằng chữ sau vùng chọn'))return 'Đã huỷ.';r.insertText(' ('+text+')','End');await c.sync();return 'Đã chèn.';});
 a['g-word-punct']=()=>Word.run(async c=>{const r=$('#g-word-all').checked?c.document.body:c.document.getSelection();r.load('text');await c.sync();const out=[];r.text.split(/\r?\n|\r/).forEach((line,i)=>{if(/ {2,}/.test(line))out.push({kind:'Khoảng trắng',text:`Đoạn ${i+1}: ${line.slice(0,200)}`});if(/ +[,;:!?]/.test(line))out.push({kind:'Khoảng trắng trước dấu câu',text:`Đoạn ${i+1}: ${line.slice(0,200)}`});if(/[,;!?](?=\p{L})/u.test(line))out.push({kind:'Có thể thiếu khoảng trắng',text:`Đoạn ${i+1}: ${line.slice(0,200)}`});});return display(out);});
 // Explicit width is safe only for a single text column with known page geometry.
 function pageFitWidth(p,columnCount){
   if(columnCount!==1||![p.pageWidth,p.leftMargin,p.rightMargin,p.gutter].every(Number.isFinite))return null;
   if(p.gutter<0||p.leftMargin<0||p.rightMargin<0||!['Left','Right','Top'].includes(p.gutterPosition))return null;
   const width=p.pageWidth-p.leftMargin-p.rightMargin-(p.gutterPosition==='Top'?0:p.gutter);
   return width>=36&&width<=2000?width:null;
 }
 async function autoFitTables(c,scope,mode,supports,confirm,normalize=false){
   const labels={Window:'vừa chiều rộng trang',Content:'vừa nội dung',FixedSize:'cố định chiều rộng cột'};
   if(!Object.prototype.hasOwnProperty.call(labels,mode)||!['all','current'].includes(scope))throw Error('Tùy chọn AutoFit không hợp lệ.');
   const desktop=supports('WordApiDesktop','1.4'),geometry=supports('WordApiDesktop','1.3'),base=supports('WordApi','1.3');
   if(!base)throw Error('AutoFit bảng yêu cầu WordApi 1.3.');
   if(mode!=='Window'&&!desktop)throw Error('Chế độ này yêu cầu WordApiDesktop 1.4. Hãy cập nhật Word hoặc chọn Vừa chiều rộng trang.');
   let tables;
   if(scope==='all'){
     const collection=c.document.body.tables;
     collection.load('items/nestingLevel');await c.sync();tables=collection.items.slice();
   }else{
     const parent=c.document.getSelection().parentTableOrNullObject;
     parent.load('nestingLevel');await c.sync();tables=parent.isNullObject?[]:[parent];
   }
   const skipped=tables.filter(t=>t.nestingLevel!==1).length;
   tables=tables.filter(t=>t.nestingLevel===1);
   if(!tables.length)return skipped?'Không xử lý riêng bảng lồng.':scope==='all'?'Không có bảng trong nội dung chính.':'Đặt con trỏ bên trong bảng cần chỉnh.';
   const strict=mode==='Window'&&normalize;
   const plans=tables.map((table,i)=>({table,i:i+1,target:null}));
   if(strict&&geometry){
     // Resolve section for each table; do not assume all pages have identical margins.
     for(const p of plans){p.sections=p.table.getRange('Whole').sections;p.sections.load('items');}
     await c.sync();
     for(const p of plans){
       if(p.sections.items.length!==1)continue;
       p.setup=p.sections.items[0].pageSetup;
       p.setup.load('pageWidth,leftMargin,rightMargin,gutter,gutterPosition');
       p.columns=p.setup.textColumns;p.columns.load('items');
     }
     await c.sync();
     for(const p of plans)if(p.setup)p.target=pageFitWidth(p.setup,p.columns.items.length);
   }
   if(!await confirm('Áp dụng '+labels[mode]+' cho '+tables.length+' bảng? '+(strict?'Căn trái bảng; đặt chiều rộng theo lề trang khi xác định được. Không đổi lề trang hoặc liên kết Excel.':'Chỉ thay đổi bố cục bảng.')))return 'Đã huỷ.';
   for(const p of plans){
     if(strict)p.table.alignment='Left';
     if(desktop)p.table.autoFitBehavior(mode);
     else p.table.autoFitWindow();
   }
   await c.sync();
   if(strict){
     for(const p of plans)if(p.target!==null)p.table.width=p.target;
     if(plans.some(p=>p.target!==null))await c.sync();
   }
   for(const p of plans)p.table.load('width,alignment');
   await c.sync();
   const verified=plans.filter(p=>p.target!==null&&Number.isFinite(p.table.width)&&Math.abs(p.table.width-p.target)<=1&&p.table.alignment==='Left').length;
   const mismatches=plans.filter(p=>p.target!==null&&!(Number.isFinite(p.table.width)&&Math.abs(p.table.width-p.target)<=1&&p.table.alignment==='Left'));
   const unknown=plans.filter(p=>p.target===null).length;
   if(strict){
     return 'Đã gửi AutoFit cho '+tables.length+' bảng Word; '+verified+' bảng đạt chiều rộng mục tiêu (theo API)'+(mismatches.length?'; '+mismatches.length+' bảng chưa đạt: '+mismatches.map(p=>p.i).join(', '):'')+(unknown?'; '+unknown+' bảng chỉ dùng AutoFit do chưa xác định được chiều rộng trang':'')+(skipped?'; bỏ qua '+skipped+' bảng lồng':'')+'. Không xử lý đối tượng Excel liên kết.';
   }
   return 'Đã AutoFit '+tables.length+' bảng Word: '+labels[mode]+(skipped?'; bỏ qua '+skipped+' bảng lồng':'')+'. Không xử lý đối tượng Excel liên kết.';
 }
 a['g-word-autofit']=()=>Word.run(c=>autoFitTables(c,val('g-fit-scope'),val('g-fit-mode'),(set,version)=>Office.context.requirements.isSetSupported(set,version),yes,$('#g-fit-normalize').checked));
 function marginPoints(raw){
   const out={};
   for(const k of ['top','bottom','left','right']){
     const s=String(raw[k]??'').trim().replace(',','.');
     if(!/^\d+(?:\.\d+)?$/.test(s))throw Error('Nhập lề hợp lệ theo cm.');
     const n=Number(s);
     if(n<0.5||n>5)throw Error('Mỗi lề phải từ 0,5 đến 5 cm.');
     out[k+'Margin']=n*72/2.54;
   }
   return out;
 }
 async function applyPageMargins(c,raw,fit,supports,confirm){
   if(!supports('WordApiDesktop','1.3'))throw Error('Căn lề A4 yêu cầu WordApiDesktop 1.3.');
   const margins=marginPoints(raw);
   const sections=c.document.sections;sections.load('items');await c.sync();
   if(!sections.items.length)return 'Không tìm thấy section.';
   const setups=sections.items.map(s=>s.pageSetup);
   for(const p of setups)p.load('orientation');
   await c.sync();
   if(!await confirm('Đặt khổ A4 và lề đã nhập cho toàn bộ '+setups.length+' section, giữ hướng giấy hiện tại, bỏ lề đóng gáy và lề đối xứng'+(fit?', sau đó AutoFit bảng Word':'')+'? Bố cục và số trang có thể thay đổi.'))return 'Đã huỷ.';
   for(const p of setups){
     const orientation=p.orientation;
     p.paperSize='A4';p.orientation=orientation;
     p.mirrorMargins=false;p.gutter=0;
     for(const k of Object.keys(margins))p[k]=margins[k];
   }
   await c.sync();
   for(const p of setups)p.load('paperSize,topMargin,bottomMargin,leftMargin,rightMargin,gutter,mirrorMargins');
   await c.sync();
   const verified=setups.filter(p=>p.paperSize==='A4'&&p.gutter===0&&p.mirrorMargins===false&&Object.keys(margins).every(k=>Math.abs(p[k]-margins[k])<=0.1)).length;
   let result='Đã đặt lề A4: '+verified+'/'+setups.length+' section được xác minh.';
   if(fit){
     if(verified!==setups.length)return result+' Chưa AutoFit vì có section chưa đạt.';
     try{result+=' '+await autoFitTables(c,'all','Window',supports,async()=>true,true);}
     catch(e){result+=' Lề đã được áp dụng; AutoFit chưa hoàn tất: '+(e.message||e);}
   }
   return result;
 }
 a['g-word-margins']=()=>Word.run(c=>applyPageMargins(c,Object.fromEntries(['top','bottom','left','right'].map(k=>[k,val('g-margin-'+k)])),$('#g-margin-fit').checked,(set,version)=>Office.context.requirements.isSetSupported(set,version),yes));
 return {actions:a,parseNumber,autoFitTables,pageFitWidth,marginPoints,applyPageMargins};
})();
