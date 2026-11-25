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
