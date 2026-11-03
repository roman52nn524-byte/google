const SHEETS={orders:'Продажи',items:'Позиции',stock:'Склад'};
function onOpen(){SpreadsheetApp.getUi().createMenu('🎆 Салюты').addItem('Новая продажа','newSale').addItem('Маршрутный лист','routeSheet').addToUi();}
function newSale(){SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Sale').setWidth(1120).setHeight(760),'🎆 Новая продажа');}
function catalog(){return SpreadsheetApp.getActive().getSheetByName(SHEETS.stock).getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({id:r[0],sku:String(r[1]),name:r[2],variant:r[3],cost:Number(r[5]),price:Number(r[6]),available:Number(r[8])}));}
function num(v,label){const n=Number(v);if(!Number.isFinite(n)||n<0)throw Error('Проверьте поле: '+label);return n;}
function text(v){let s=String(v||'').trim();if(s.length>2000)throw Error('Слишком длинное поле');return /^[=+@-]/.test(s)?"'"+s:s;}
function dateValue(v){if(!/^\d{4}-\d{2}-\d{2}$/.test(v||''))throw Error('Укажите дату');let [y,m,d]=v.split('-').map(Number);let x=new Date(y,m-1,d,12);if(x.getMonth()!==m-1||x.getDate()!==d)throw Error('Некорректная дата');return x;}
function prepareSale(p,products){
 if(!p.token||!/^[a-zA-Z0-9-]{10,80}$/.test(p.token))throw Error('Некорректный ключ запроса');
 if(!p.client||!p.channel)throw Error('Укажите клиента и канал');
 const date=dateValue(p.date),deliveryDate=p.deliveryDate?dateValue(p.deliveryDate):'';
 if(!p.pickup&&(!deliveryDate||!String(p.address||'').trim()))throw Error('Укажите дату и адрес доставки');
 if(!Array.isArray(p.items)||!p.items.length||p.items.length>100)throw Error('Добавьте от 1 до 100 товаров');
 const map=new Map(products.map(x=>[x.id,x])), used={};let total=0,cost=0;
 const lines=p.items.map(v=>{const a=map.get(v.id);if(!a)throw Error('Товар не найден');const qty=num(v.qty,'Количество');if(!Number.isInteger(qty)||qty<1)throw Error('Количество должно быть целым и больше нуля');used[a.id]=(used[a.id]||0)+qty;if(!Number.isFinite(a.available)||used[a.id]>a.available)throw Error('Недостаточно на складе: '+a.name);const price=num(v.price,'Цена');if(price<=0)throw Error('Цена должна быть больше нуля');const c=num(a.cost,'Закупочная цена');total+=Math.round(price*qty*100)/100;cost+=Math.round(c*qty*100)/100;return [a.id,a.name,qty,price,c,Math.round(price*qty*100)/100,Math.round(c*qty*100)/100];});
 total=Math.round(total*100)/100;cost=Math.round(cost*100)/100;
 const paid=p.full?total:num(p.paid,'Оплачено товаров');if(paid>total)throw Error('Оплата больше стоимости товаров');const delivery=num(p.delivery||0,'Доставка');
 return {date,deliveryDate,lines,total,cost,paid,delivery,due:Math.round((total-paid+(p.deliveryPaid?0:delivery))*100)/100};
}
function saveSale(p){
 const lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
 const ss=SpreadsheetApp.getActive(),o=ss.getSheetByName(SHEETS.orders),i=ss.getSheetByName(SHEETS.items);
 if(o.getLastRow()>1){const tokens=o.getRange(2,23,o.getLastRow()-1,1).getValues().flat();const pos=tokens.indexOf(p.token);if(pos>=0)return {id:o.getRange(pos+2,1).getValue(),duplicate:true};}
 SpreadsheetApp.flush();const v=prepareSale(p,catalog());
 if(i.getLastRow()+v.lines.length>10000||o.getLastRow()>=10000)throw Error('Достигнут предел формул. Расширьте диапазоны.');
 const id='SAL-'+Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8);
 const row=[id,v.date,text(p.channel),text(p.client),text(p.phone),text(p.recipient||p.client),text(p.recipientPhone||p.phone),'Продажа',v.lines.map(x=>x[1]+' × '+x[2]).join('; '),v.total,v.cost,Math.round((v.total-v.cost)*100)/100,v.paid,v.due,v.deliveryDate,text(p.time),text(p.pickup?'Самовывоз':p.address),text(p.courier),text(p.courierPhone),v.delivery,!!p.deliveryPaid,text(p.comment),p.token];
 const ir=i.getLastRow()+1,or=o.getLastRow()+1;
 try{i.getRange(ir,1,v.lines.length,8).setValues(v.lines.map(x=>[id,...x.map(a=>typeof a==='string'?text(a):a)]));o.getRange(or,1,1,row.length).setValues([row]);SpreadsheetApp.flush();}catch(e){i.getRange(ir,1,v.lines.length,8).clearContent();o.getRange(or,1,1,23).clearContent();throw e;}
 return {id};
 }finally{lock.releaseLock();}
}
function esc(v){return String(v||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function routeSheet(){
 const rows=SpreadsheetApp.getActive().getSheetByName(SHEETS.orders).getDataRange().getValues();
 const html='<style>body{font:15px Arial;color:#294a40;padding:24px}article{border:1px solid #ddd;border-radius:14px;padding:20px;margin:16px 0}h2{color:#a64f70}@media print{button{display:none}article{break-inside:avoid}}</style><button onclick="window.print()">Распечатать</button><h1>Маршрутный лист</h1><p>Новые продажи с доставкой; сверяйте выполнение перед выездом.</p>'+rows.slice(1).filter(r=>r[16]&&r[16]!=='Самовывоз').sort((a,b)=>new Date(a[14])-new Date(b[14])||String(a[15]).localeCompare(String(b[15]))).map(r=>'<article><h2>'+esc(r[0])+'</h2><p>'+esc(r[14] instanceof Date?Utilities.formatDate(r[14],SpreadsheetApp.getActive().getSpreadsheetTimeZone(),'dd.MM.yyyy'):'Дата не указана')+' · '+esc(r[15])+'</p><h3>'+esc(r[16])+'</h3><p>'+esc(r[5])+' · '+esc(r[6])+'</p><p>'+esc(r[8])+'</p><p>Получить: '+esc(r[13])+' ₽</p><p>'+esc(r[21])+'</p><p>□ Доставлено &nbsp; Время: __________</p></article>').join('');
 SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(850).setHeight(700),'Печать доставки');
}
