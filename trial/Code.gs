const TZ='Europe/Moscow';
const LIMIT=1001;
const SOURCES=['Авито','ТГ','МАКС','ВК','Ozon','Яндекс Маркет','По звонку','Карты','Оффлайн','Разное'];
function onOpen(){SpreadsheetApp.getUi().createMenu('🎆 CRM').addItem('Подготовить тестовую таблицу','setupDemo').addItem('Новая продажа','newSale').addItem('Обновить площадки','syncMarkets').addToUi();}
function active(){return SpreadsheetApp.getActive();}
function rows(name){const s=active().getSheetByName(name);if(!s)throw Error('Нет листа '+name);return s.getDataRange().getValues().slice(1);}
function setupDemo(){
 const ss=active();ss.setSpreadsheetTimeZone(TZ);
 ['Товары','Закупка','Склад','Заказы','Позиции','Ozon','Яндекс Маркет'].forEach(n=>{const s=ss.getSheetByName(n);if(!s)throw Error('Загрузите demo-crm.xlsx: нет листа '+n);if(s.getMaxRows()<LIMIT)s.insertRowsAfter(s.getMaxRows(),LIMIT-s.getMaxRows());});
 const rule=SpreadsheetApp.newDataValidation().requireCheckbox().build();
 [['Заказы',19],['Закупка',10],['Ozon',7],['Яндекс Маркет',7]].forEach(([n,c])=>ss.getSheetByName(n).getRange(2,c,1000,1).setDataValidation(rule));
 syncMarkets();SpreadsheetApp.getUi().alert('Готово. Выберите «Новая продажа». Тестовые данные сохранены.');
}
function newSale(){SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Sale').setWidth(1100).setHeight(760),'🎆 Новая продажа · тест');}
function getCatalog(){SpreadsheetApp.flush();const stock=new Map(rows('Склад').filter(r=>r[0]).map(r=>[r[0],r[4]]));return rows('Товары').filter(r=>r[0]).map(r=>{const available=stock.get(r[0]);if(typeof available!=='number'||!Number.isFinite(available))throw Error('Не рассчитан остаток '+r[0]+'. Запустите «Подготовить тестовую таблицу» и проверьте формулы склада.');return {id:r[0],sku:r[1],name:r[2],factor:Number(r[3]),price:Number(r[4]),packPrice:Number(r[5]),cost:Number(r[6]),available};});}
function n(v,label){if(v===''||v===null||v===undefined)throw Error('Заполните '+label);const x=Number(v);if(!Number.isFinite(x)||x<0)throw Error('Некорректное поле: '+label);return x;}
function safe(v){const s=String(v??'').trim();if(s.length>1500)throw Error('Слишком длинное текстовое поле');return /^[=+@-]/.test(s)?"'"+s:s;}
function round(v){return Math.round((v+Number.EPSILON)*100)/100;}
function calculate(p,catalog){
 if(!p||!SOURCES.includes(p.source)||!String(p.client||'').trim())throw Error('Укажите клиента и источник');
 if(!Array.isArray(p.items)||!p.items.length||p.items.length>50)throw Error('Добавьте от 1 до 50 позиций');
 const map=new Map(catalog.map(x=>[x.id,x])),used={};let total=0,cost=0;
 const lines=p.items.map(x=>{const a=map.get(x.id);if(!a)throw Error('Товар не найден');const qty=n(x.qty,'количество');if(!Number.isInteger(qty)||qty<1)throw Error('Количество должно быть целым');if(!['шт','уп'].includes(x.unit))throw Error('Неизвестная единица');const factor=x.unit==='уп'?a.factor:1;if(!Number.isInteger(factor)||factor<1)throw Error('Неверный размер упаковки');const units=qty*factor;used[a.id]=(used[a.id]||0)+units;if(!Number.isFinite(a.available)||used[a.id]>a.available)throw Error('Недостаточно товара: '+a.name);const price=n(x.price,'цену'),purchase=n(a.cost,'закупочную цену');if(price<=0)throw Error('Цена должна быть больше нуля');const sum=round(qty*price),c=round(units*purchase);total+=sum;cost+=c;return [a.id,a.name,x.unit,qty,factor,units,price,sum,purchase,c];});
 const delivery=n(p.delivery,'доставку с клиента'),courier=n(p.courier,'расход курьер'),paid=n(p.paid,'оплачено');total=round(total);cost=round(cost);
 const market=['Ozon','Яндекс Маркет'].includes(p.source);if(market&&(delivery!==0||courier!==0||paid!==0))throw Error('Для площадок доставка, удержания и выплаты вводятся на отдельном листе. Поля прямой оплаты должны быть нулевыми.');
 if(!market&&paid>round(total+delivery))throw Error('Оплата больше суммы заказа');return {lines,total,cost,delivery,courier,paid,market};
}
function formulasOrder(r){return {9:`=IF(S${r},G${r}-H${r},0)`,12:`=SUMIFS(Ozon!S$2:S$1001,Ozon!A$2:A$1001,A${r},Ozon!G$2:G$1001,TRUE)+SUMIFS('Яндекс Маркет'!S$2:S$1001,'Яндекс Маркет'!A$2:A$1001,A${r},'Яндекс Маркет'!G$2:G$1001,TRUE)`,13:`=IF(S${r},G${r}-H${r}+J${r}-K${r}-L${r},0)`,14:`=G${r}+J${r}`,16:`=IF(OR(C${r}="Ozon",C${r}="Яндекс Маркет"),0,MAX(0,N${r}-O${r}))`,21:`=IFERROR(IF(C${r}="Ozon",VLOOKUP(A${r},Ozon!A$2:V$1001,8,FALSE),IF(C${r}="Яндекс Маркет",VLOOKUP(A${r},'Яндекс Маркет'!A$2:V$1001,8,FALSE),"")),"")`,22:`=IFERROR(IF(C${r}="Ozon",VLOOKUP(A${r},Ozon!A$2:V$1001,16,FALSE),IF(C${r}="Яндекс Маркет",VLOOKUP(A${r},'Яндекс Маркет'!A$2:V$1001,16,FALSE),"Не применяется")),"План")`};}
function formulasMarket(s,r){s.getRange(r,19,1,4).setFormulas([[`=K${r}+L${r}+M${r}-N${r}`,`=IF(G${r},E${r}-F${r}-S${r},0)`,`=R${r}-S${r}`,`=U${r}-O${r}`]]);}
function syncInternal(){
 const ss=active(),orders=rows('Заказы').filter(r=>r[0]);
 for(const name of ['Ozon','Яндекс Маркет']){
 const s=ss.getSheetByName(name),existing=new Map();rows(name).forEach((r,j)=>{if(r[0]){if(existing.has(r[0]))throw Error('Дубли ID на листе '+name);existing.set(r[0],j+2);}});
 // Не удаляем строки с введенными расходами: заказ, сменивший источник, помечаем неактивным.
 const current=new Set(orders.filter(r=>r[2]===name).map(r=>r[0]));existing.forEach((r,id)=>{if(!current.has(id))s.getRange(r,7).setValue(false);});
 for(const o of orders.filter(r=>r[2]===name)){
 let r=existing.get(o[0]);if(!r){r=s.getLastRow()+1;if(r>LIMIT)throw Error('Достигнут предел площадки');s.getRange(r,8,1,11).setValues([['','','',0,0,0,0,0,'План','',`=E${r}`]]);}
 s.getRange(r,1,1,7).setValues([[o[0],o[1],o[3],o[5],o[6],o[7],o[18]===true]]);formulasMarket(s,r);
 }
 }
 const o=ss.getSheetByName('Заказы');for(let r=2;r<=o.getLastRow();r++){if(o.getRange(r,1).getValue())Object.entries(formulasOrder(r)).forEach(([col,f])=>o.getRange(r,Number(col)).setFormula(f));}
 SpreadsheetApp.flush();
}
function syncMarkets(){const lock=LockService.getDocumentLock();lock.waitLock(30000);try{syncInternal();}finally{lock.releaseLock();}}
function onEdit(e){if(!e||!e.range)return;const name=e.range.getSheet().getName();if(name==='Заказы'&&e.range.getRow()>1&&[3,19].includes(e.range.getColumn()))syncMarkets();}
function saveSale(p){
 const lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
 if(!p||!/^[a-zA-Z0-9-]{10,80}$/.test(p.token||''))throw Error('Неверный ключ запроса');
 const ss=active(),o=ss.getSheetByName('Заказы'),l=ss.getSheetByName('Позиции'),old=rows('Заказы').find(r=>r[19]===p.token);
 if(old){let warning='';try{syncInternal();}catch(e){warning='Заказ сохранен. Обновите площадки через меню: '+e.message;}return {id:old[0],warning};}
 const v=calculate(p,getCatalog()),r=o.getLastRow()+1,lr=l.getLastRow()+1;if(r>LIMIT||lr+v.lines.length-1>LIMIT)throw Error('Предел тестовой версии — 1000 строк');
 const id='TEST-'+Utilities.formatDate(new Date(),TZ,'yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8);
 const row=[id,new Date(),p.source,safe(p.client),safe(p.phone),safe(v.lines.map(x=>x[1]+' × '+x[3]+' '+x[2]).join('; ')),v.total,v.cost,'',v.delivery,v.courier,'','','',v.paid,'',safe(p.address),safe(p.comment),true,p.token,'','','',''];
 const fs=formulasOrder(r);Object.entries(fs).forEach(([col,f])=>row[Number(col)-1]=f);
 try{l.getRange(lr,1,v.lines.length,12).setValues(v.lines.map((x,j)=>[id,...x.map(a=>typeof a==='string'?safe(a):a),`=IF(COUNTIFS(Заказы!A$2:A$1001,A${lr+j},Заказы!S$2:S$1001,TRUE)>0,G${lr+j},0)`]));o.getRange(r,1,1,24).setValues([row]);o.getRange(r,2).setNumberFormat('dd.MM.yyyy HH:mm');SpreadsheetApp.flush();}catch(e){l.getRange(lr,1,v.lines.length,12).clearContent();o.getRange(r,1,1,24).clearContent();throw e;}
 let warning='';try{syncInternal();}catch(e){warning='Продажа сохранена, но лист площадки не обновился. Выберите «Обновить площадки»: '+e.message;}
 return {id,warning};
 }finally{lock.releaseLock();}
}
