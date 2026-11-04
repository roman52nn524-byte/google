const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');const c=vm.createContext({});vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),c);
const cat=[{id:'T001',name:'Салют',factor:1,cost:3800,available:10},{id:'T002',name:'Петарды',factor:4,cost:50,available:40}];const p={client:'Тест',source:'Авито',delivery:300,courier:500,paid:0,items:[{id:'T001',unit:'шт',qty:1,price:7000},{id:'T002',unit:'уп',qty:2,price:350}]};let count=0;function test(fn){fn();count++}
test(()=>{const x=c.calculate(p,cat);assert.equal(x.total,7700);assert.equal(x.cost,4200);assert.equal(x.total-x.cost+x.delivery-x.courier,3300);assert.equal(x.lines[1][5],8)});
test(()=>{const x=c.calculate({...p,source:'Ozon',delivery:0,courier:0},cat);assert.equal(x.market,true)});
test(()=>assert.throws(()=>c.calculate({...p,source:'Ozon'},cat)));
test(()=>assert.throws(()=>c.calculate({...p,items:[{id:'T002',unit:'уп',qty:11,price:350}]},cat),/Недостаточно/));
test(()=>assert.throws(()=>c.calculate({...p,items:[{id:'T002',unit:'уп',qty:6,price:350},{id:'T002',unit:'уп',qty:5,price:350}]},cat),/Недостаточно/));
for(const qty of [0,-1,1.5,NaN])test(()=>assert.throws(()=>c.calculate({...p,items:[{id:'T001',unit:'шт',qty,price:7000}]},cat)));
test(()=>assert.throws(()=>c.calculate({...p,paid:9000},cat)));
test(()=>assert.equal(c.safe('=IMPORTXML("url")')[0],"'"));
test(()=>assert.match(c.formulasOrder(5)[12],/SUMIFS/));
test(()=>assert.throws(()=>c.calculate({...p,source:'Неизвестный'},cat)));
test(()=>assert.throws(()=>c.calculate({...p,items:[{id:'T001',unit:'коробка',qty:1,price:1}]},cat)));
// Проверяем синхронизацию по ID при перестановке строк и сохранение ручных расходов.
let data={'Заказы':[['A',new Date(),'Ozon','Клиент','','Салют',1000,400,'',0,0,'','','',0,'','','',true],['B',new Date(),'Ozon','Клиент 2','','Дым',500,200,'',0,0,'','','',0,'','','',true]],'Ozon':[['B',new Date(),'Клиент 2','Дым',500,200,true,'MP-B','TRACK-B','',50,20,0,0,0,'План','',500],['A',new Date(),'Клиент','Салют',1000,400,true,'MP-A','TRACK-A','',100,30,0,0,0,'План','',1000]],'Яндекс Маркет':[]};
function sheet(name){return {getDataRange:()=>({getValues:()=>[['header'],...data[name]]}),getLastRow:()=>data[name].length+1,getRange:(r,col,h=1,w=1)=>({setValues(vals){vals.forEach((row,j)=>{data[name][r-2+j]??=[];row.forEach((v,k)=>data[name][r-2+j][col-1+k]=v)});return this},setFormulas(vals){return this.setValues(vals)},setValue(v){return this.setValues([[v]])},setFormula(v){return this.setValues([[v]])},getValue(){return data[name][r-2]?.[col-1]}})}}
c.SpreadsheetApp={getActive:()=>({getSheetByName:sheet}),flush(){}};
test(()=>{c.syncInternal();assert.equal(data.Ozon[0][0],'B');assert.equal(data.Ozon[0][8],'TRACK-B');assert.equal(data.Ozon[1][8],'TRACK-A');assert.equal(data.Ozon[1][10],100)});
test(()=>{data['Заказы'][0][2]='Авито';c.syncInternal();assert.equal(data.Ozon[1][6],false);assert.equal(data.Ozon[1][8],'TRACK-A')});
console.log(count+' checks passed');
