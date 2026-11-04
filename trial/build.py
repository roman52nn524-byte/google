from pathlib import Path
import openpyxl,datetime
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
p=Path(__file__).parent;w=openpyxl.Workbook();w.remove(w.active)
def sheet(n,h):
 s=w.create_sheet(n);s.append(h);return s
s=sheet('Старт',['САЛЮТЫ · ПРОБНАЯ CRM','Описание'])
for r in [('Это тестовая копия','5 вымышленных товаров, 3 демонстрационных заказа. Данные реального бизнеса не использованы.'),('Начать','Загрузите demo-crm.xlsx в Google Диск → сохраните как Google Таблицы. Установите Code.gs и Sale.html по README.'),('Продать','Меню 🎆 CRM → Новая продажа. Выберите товары, штуки/упаковки, клиента и источник.'),('Склад','Закупка увеличивает остаток только с отметкой Принято. Заказы с отметкой Учитывать списывают товар сразу.'),('Маркетплейсы','Меню Обновить площадки создает строки по ID заказа. Заполняйте H:M и O:R на Ozon/Яндекс Маркет. Не меняйте ID.'),('Расходы площадок','K — комиссии, L — логистика, M — прочие удержания, N — компенсации. Общие расходы указываются один раз на заказ.'),('Обычная доставка','Оплата доставки клиентом и ваши расходы на курьера — отдельные поля в заказе. Не дублируйте логистику площадки.'),('Статусы','Для отмены снимите флажок Учитывать в Заказах: товар вернется в доступный остаток. Частичные возвраты пока не поддержаны.'),('Петарды','1 упаковка = 4 штуки; цена упаковки отдельная. Эта версия считает общий запас в штуках, целостность упаковок пока не отслеживает.'),('Предел теста','До 1000 строк данных на каждом рабочем листе. Расходы и выплаты площадок вводятся вручную. Интеграций с площадками нет.'),('Услуги','Тяжелый дым, бронирование оборудования и мероприятия добавим следующим этапом; здесь тестируем товарные продажи.'),('Безопасность теста','Не используйте для рабочего учета до проверки импорта, сохранения заказов и пересчета формул в своем Google-аккаунте.')]:s.append(r)
s=sheet('Справочники',['Источники'])
channels=['Авито','ТГ','МАКС','ВК','Ozon','Яндекс Маркет','По звонку','Карты','Оффлайн','Разное']
for v in channels:s.append([v])
s=sheet('Товары',['ID','Артикул','Товар','Штук в упаковке','Цена штуки','Цена упаковки','Закупка за штуку','Минимум'])
for r in [('T001','DEMO-S16','Салют «Полярная ночь» · 16 залпов',1,7000,7000,3800,2),('T002','DEMO-P4','Петарды «Гром» · 4 шт/уп',4,100,350,50,8),('T003','DEMO-F','Фонтан «Искры»',1,800,800,450,3),('T004','DEMO-D','Дым «Облако» · синий',1,800,800,350,3),('T005','DEMO-B','Бенгальские огни · упаковка',1,300,300,150,5)]:s.append(r)
s=sheet('Закупка',['ID приемки','Дата','ID товара','Единица','Кол-во единиц','Штук в единице','Всего штук','Цена закупки единицы','Сумма','Принято','Комментарий'])
for i,(id,unit,qty,factor,price) in enumerate([('T001','шт',10,1,3800),('T002','уп',10,4,200),('T003','шт',20,1,450),('T004','шт',15,1,350),('T005','шт',30,1,150)],2):s.append([f'IN-{i-1:03}',datetime.date(2026,11,4),id,unit,qty,factor,f'=E{i}*F{i}',price,f'=E{i}*H{i}',True,'Начальный тестовый приход'])
s=sheet('Склад',['ID товара','Товар','Получено, шт','Списано, шт','Доступно, шт','Стоимость по учетной цене'])
for i in range(2,7):
 s.append([f"='Товары'!A{i}",f"='Товары'!C{i}",f'=SUMIFS(Закупка!G$2:G$1001,Закупка!C$2:C$1001,A{i},Закупка!J$2:J$1001,TRUE)',f'=SUMIF(Позиции!B$2:B$1001,A{i},Позиции!L$2:L$1001)',f'=C{i}-D{i}',f'=E{i}*Товары!G{i}'])
s=sheet('Заказы',['ID','Дата','Источник','Клиент','Телефон','Состав','Товары после скидок','Себестоимость','Валовая прибыль товаров','Доставка с клиента','Расход курьер','Расходы площадки нетто','Прибыль после прямых расходов','К оплате клиенту','Оплачено напрямую','Осталось напрямую','Адрес','Комментарий','Учитывать','Ключ запроса','ID заказа площадки','Статус сверки','Дата доставки','Время'])
orders=[('DEMO-001','Авито','Тестовый клиент А','Салют × 1; Петарды × 2 уп',7700,4200,300,500,8000,True),('DEMO-002','Ozon','Тестовый клиент Б','Фонтан × 2; Дым × 1',2400,1250,0,0,0,True),('DEMO-003','Яндекс Маркет','Тестовый клиент В','Бенгальские огни × 3',900,450,0,0,0,True)]
for i,(id,channel,client,composition,total,cost,delivery,expense,paid,active) in enumerate(orders,2):
 s.append([id,datetime.date(2026,11,4),channel,client,'',composition,total,cost,f'=IF(S{i},G{i}-H{i},0)',delivery,expense,f'=SUMIFS(Ozon!S$2:S$1001,Ozon!A$2:A$1001,A{i},Ozon!G$2:G$1001,TRUE)+SUMIFS(\'Яндекс Маркет\'!S$2:S$1001,\'Яндекс Маркет\'!A$2:A$1001,A{i},\'Яндекс Маркет\'!G$2:G$1001,TRUE)',f'=IF(S{i},G{i}-H{i}+J{i}-K{i}-L{i},0)',f'=G{i}+J{i}',paid,f'=IF(OR(C{i}="Ozon",C{i}="Яндекс Маркет"),0,MAX(0,N{i}-O{i}))','Самовывоз' if channel=='Авито' else 'Тестовый адрес','Демонстрация',active,'demo-token-'+id,'','План',None,''])
s=sheet('Позиции',['ID заказа','ID товара','Товар','Единица','Количество','Штук в единице','Всего штук','Цена единицы','Продажа','Закупка за штуку','Себестоимость','Списание, шт'])
lines=[('DEMO-001','T001','Салют «Полярная ночь»','шт',1,1,7000,3800),('DEMO-001','T002','Петарды «Гром»','уп',2,4,350,50),('DEMO-002','T003','Фонтан «Искры»','шт',2,1,800,450),('DEMO-002','T004','Дым «Облако»','шт',1,1,800,350),('DEMO-003','T005','Бенгальские огни','шт',3,1,300,150)]
for i,(id,product,name,unit,qty,factor,price,cost) in enumerate(lines,2):s.append([id,product,name,unit,qty,factor,qty*factor,price,qty*price,cost,qty*factor*cost,f'=IF(COUNTIFS(Заказы!A$2:A$1001,A{i},Заказы!S$2:S$1001,TRUE)>0,G{i},0)'])
headers=['ID заказа','Дата','Клиент','Состав','Продажа товаров','Себестоимость','Учитывать','Номер площадки','Трек','Отгрузка / статус','Комиссия / размещение','Логистика','Другие удержания','Компенсации','Выплачено площадкой','Сверка','Комментарий','Начисления продавцу до удержаний','Расходы нетто','Прибыль заказа','Ожидаемая выплата','Осталось получить']
for name,order,commission,logistics,other in [('Ozon',orders[1],360,180,30),('Яндекс Маркет',orders[2],135,120,0)]:
 s=sheet(name,headers);id,_,client,composition,total,cost,*_=order
 s.append([id,datetime.date(2026,11,4),client,composition,total,cost,True,'TEST-MP-001','TEST-TRACK-001','Тест: готов к отгрузке',commission,logistics,other,0,0,'План','Вымышленные расходы для проверки',total,'=K2+L2+M2-N2','=IF(G2,E2-F2-S2,0)','=R2-S2','=U2-O2'])
s=sheet('Обзор',['Показатель','Сумма / количество'])
for r in [('Продажи товаров','=SUMIF(Заказы!S2:S1001,TRUE,Заказы!G2:G1001)'),('Валовая прибыль товаров','=SUM(Заказы!I2:I1001)'),('Прибыль после прямых расходов','=SUM(Заказы!M2:M1001)'),('Долг прямых клиентов','=SUMIF(Заказы!S2:S1001,TRUE,Заказы!P2:P1001)'),('Ожидаемые выплаты площадок','=SUMIF(Ozon!G2:G1001,TRUE,Ozon!V2:V1001)+SUMIF(\'Яндекс Маркет\'!G2:G1001,TRUE,\'Яндекс Маркет\'!V2:V1001)'),('Остаток в штуках','=SUM(Склад!E2:E6)')]:s.append(r)
for s in w:
 s.freeze_panes='A2';s.sheet_view.showGridLines=False;s.sheet_properties.tabColor='BA6686' if s.title in ['Заказы','Обзор','Ozon','Яндекс Маркет'] else '80AF9C'
 s.row_dimensions[1].height=38
 for c in s[1]:c.fill=PatternFill('solid',fgColor='F5E1E9');c.font=Font(name='Calibri',bold=True,color='8D405E');c.alignment=Alignment(wrap_text=True,vertical='center')
 for col in s.columns:s.column_dimensions[col[0].column_letter].width=min(36,max(18,len(str(col[0].value))+2))
 for row in s.iter_rows(min_row=2):
  for c in row:
   c.font=Font(name='Calibri',color='294A40');c.alignment=Alignment(vertical='center',wrap_text=True)
   if c.row%2==0:c.fill=PatternFill('solid',fgColor='EFF7F2')
   if isinstance(c.value,datetime.date):c.number_format='dd.mm.yyyy'
 if s.max_row>1:s.auto_filter.ref=s.dimensions
w['Старт'].column_dimensions['B'].width=110
for row in w['Старт'].iter_rows(min_row=2):w['Старт'].row_dimensions[row[0].row].height=42
for name in ['Ozon','Яндекс Маркет']:
 for row in w[name].iter_rows(min_row=2,min_col=8,max_col=18):
  for c in row:c.fill=PatternFill('solid',fgColor='FFF3D9')
for name,col in [('Закупка','J'),('Заказы','S')]:
 d=DataValidation(type='list',formula1='"TRUE,FALSE"');w[name].add_data_validation(d);d.add(f'{col}2:{col}1001')
w['Склад'].conditional_formatting.add('E2:E6',CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FFC7CE')))
w.save(p/'demo-crm.xlsx');print('Workbook generated')
