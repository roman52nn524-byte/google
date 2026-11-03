const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict');const c=vm.createContext({});vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),c);
const products=[{id:'P1',name:'Салют',available:3,cost:100}];
const p={token:'test-token-1234',date:'2026-11-03',channel:'Авито',client:'Клиент',pickup:true,paid:50,delivery:100,items:[{id:'P1',qty:2,price:250}]};
let n=0;function test(fn){fn();n++}
test(()=>{let v=c.prepareSale(p,products);assert.equal(v.total,500);assert.equal(v.cost,200);assert.equal(v.due,550)});
test(()=>assert.equal(c.prepareSale({...p,full:true,deliveryPaid:true},products).due,0));
test(()=>assert.throws(()=>c.prepareSale({...p,items:[{id:'P1',qty:2,price:250},{id:'P1',qty:2,price:250}]},products),/Недостаточно/));
for(const qty of [0,-1,1.5,NaN])test(()=>assert.throws(()=>c.prepareSale({...p,items:[{id:'P1',qty,price:250}]},products)));
test(()=>assert.throws(()=>c.prepareSale({...p,paid:501},products)));
test(()=>assert.throws(()=>c.prepareSale({...p,date:'2026-02-30'},products)));
test(()=>assert.throws(()=>c.prepareSale({...p,pickup:false},products)));
test(()=>assert.equal(c.text('=IMPORTXML("url")')[0],"'"));
test(()=>assert.throws(()=>c.prepareSale(p,[{...products[0],available:-1}])));
console.log(n+' business logic checks passed');
