/* Pure helpers + simulated range APIs; no real Office host. Run: node tests/regression.js */
const assert=require('node:assert/strict');
const U=require('../docs/textutils.js');
const fs=require('node:fs');
const root=require('node:path').join(__dirname,'..','docs');
for(const file of ['textutils.js','taskpane.js','extra.js','general.js'])new Function(fs.readFileSync(require('node:path').join(root,file),'utf8'));
function test(){
  assert.equal(U.numberToWords(-1.5,false),'Âm hai');
  assert.equal(U.duplicateKey(['  A',2]),U.duplicateKey(['a ',2]));
  assert.notEqual(U.duplicateKey(['2']),U.duplicateKey([2]));
  assert.equal(U.safeSpreadsheetText(' =1+1'),"' =1+1");
  assert.equal(U.csvCell('a,"b"'),'"a,""b"""');
  const mask=U.formulaMask(4,3,7,2,[{rowIndex:7,columnIndex:3,rowCount:2,columnCount:1},{rowIndex:10,columnIndex:4,rowCount:1,columnCount:1}]);
  assert.deepEqual(mask.map(r=>Array.from(r)),[[0,1,0],[0,1,0],[0,0,0],[0,0,1]]);
  assert.equal(U.findFormulaOutliers(['=R1','=R1','=R1','=R1'],{formulaMask:[1,0,1,1],values:['=R1','=R1',3,4]}).outliers[0].kind,'manual');
  assert.equal(U.findFormulaOutliers(['=RC','=RC','','=RC'],{formulaMask:[1,1,0,1],blank:true}).outliers[0].kind,'blank');
  assert.equal(U.findFormulaOutliers(['=X','=X','=X'],{formulaMask:[0,0,0]}).total,0);
  const blocks=U.scanBlocks(25001,2,'duplicate');assert.equal(blocks.reduce((n,b)=>n+b.h*b.w,0),50002);
  console.log('Regression checks passed (10 assertions; 4 syntax checks).');
}
test();

async function testAutoFit(){
  const G=new Function('TextUtils','document',fs.readFileSync(require('node:path').join(root,'general.js'),'utf8')+'\nreturn General;')(U,{});
  const calls=[];
  const table=()=>({load(){},isNullObject:false,nestingLevel:1,autoFitWindow(){calls.push('Window')},autoFitBehavior(m){calls.push(m)}});
  const first=table(),second=table(),collection={items:[first,second],load(){}};
  const ctx={document:{body:{tables:collection},getSelection:()=>({parentTableOrNullObject:second})},sync:async()=>{}};
  const desktop=()=>true,baseOnly=(set)=>set==='WordApi',confirm=async()=>true;
  for(const mode of ['Window','Content','FixedSize']){
    calls.length=0;await G.autoFitTables(ctx,'all',mode,desktop,confirm);assert.deepEqual(calls,[mode,mode]);
  }
  calls.length=0;await G.autoFitTables(ctx,'current','Content',desktop,confirm);assert.deepEqual(calls,['Content']);
  calls.length=0;await G.autoFitTables(ctx,'all','Window',baseOnly,confirm);assert.deepEqual(calls,['Window','Window']);
  calls.length=0;await G.autoFitTables(ctx,'all','Window',desktop,async()=>false);assert.equal(calls.length,0);
  await assert.rejects(()=>G.autoFitTables(ctx,'all','FixedSize',baseOnly,confirm));assert.equal(calls.length,0);
  collection.items=[];await G.autoFitTables(ctx,'all','Window',desktop,confirm);assert.equal(calls.length,0);
  const setup={pageWidth:600,leftMargin:40,rightMargin:50,gutter:10,gutterPosition:'Left',load(){},textColumns:{items:[{}],load(){}}};
  assert.equal(G.pageFitWidth(setup,1),500);
  assert.equal(G.pageFitWidth(setup,2),null);
  assert.equal(G.pageFitWidth({...setup,gutterPosition:'Top'},1),510);
  const fitTable={...table(),width:300,alignment:'Centered',getRange(){return {sections:{items:[{pageSetup:setup}],load(){}}}}};
  collection.items=[fitTable,{...table(),nestingLevel:2}];calls.length=0;
  const result=await G.autoFitTables(ctx,'all','Window',desktop,confirm,true);
  assert.equal(fitTable.width,500);assert.equal(fitTable.alignment,'Left');assert.deepEqual(calls,['Window']);
  assert.ok(result.includes('1 bảng đạt')&&result.includes('bỏ qua 1 bảng lồng'));
  fitTable.width=300;calls.length=0;
  const fallback=await G.autoFitTables(ctx,'all','Window',baseOnly,confirm,true);
  assert.equal(fitTable.width,300);assert.ok(fallback.includes('chưa xác định'));assert.deepEqual(calls,['Window']);
  console.log('AutoFit simulation passed (all modes, current table, compatibility, cancel, empty).');
}
testAutoFit().catch(e=>{console.error(e);process.exitCode=1;});
