import sys, re, json, gzip
import os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
import build_pm06_meta as B
path, name, out = sys.argv[1:4]
z=B._open(path); rows=B._rows(z,B._sheet(z,'PM-06'))
for r in rows:
    if any(isinstance(v,str) and v.strip()=='Заказ' for v in r.values()): hdr=r; break
H={v.strip():k for k,v in hdr.items() if isinstance(v,str) and v.strip()}
g=lambda r,n,o=0: (r.get(H[n]+o) or '').strip() if n in H else ''
N=lambda r,n: B._num(r.get(H[n])) if n in H else 0.0
O={}
for r in rows:
    o=g(r,'Заказ')
    if not o or o=='#': continue
    x=O.get(o)
    if x is None:
        x=O[o]=dict(be=g(r,'Балансовая единица'),eo=g(r,'ЕО'),eon=g(r,'ЕО',1),tm=g(r,'Техническое место'),tmn=g(r,'Техническое место',1),
            gar=g(r,'Гаражный номер'),txt=g(r,'Заказ',1),kind=g(r,'Заказ Вид заказа'),w=g(r,'Заказ Вид работы ТОРО'),wt=g(r,'Заказ Вид работы ТОРО',1),
            bs=g(r,'Заказ Базисный срок начала (дата)'),sys=g(r,'Заказ Системный статус'),usr=g(r,'Заказ Пользовательский статус'),
            pp=g(r,'Заказ Завод, планирующий ТОРО'),pg=g(r,'Заказ Группа планирования ТОРО'),mu=g(r,'Заказ признак ХС/УСО/Комб'),
            mvzg=g(r,'МВЗ Группа МВЗ'),omvz=g(r,'Заказ Место возникновения затрат'),emvz=g(r,'ЕО Место возникновения затрат'),
            mark=g(r,'ЕО Марка заводская'),mfr=g(r,'ЕО изготовитель'),din=g(r,'ЕО Дата ввода в эксплуатацию'),
            dplan=g(r,'ЕО Плановая дата ввода в эксплуатацию'),dout=g(r,'ЕО Плановый срок вывода из эксплуатации'),
            wc=g(r,'Операция заказа Рабочее место'),spp=g(r,'СПП-элемент Центр затрат для ремонтов'),inv=g(r,'СПП-элемент Причина инвестиций'),
            year=g(r,'Заказ Год утверждения'),
            p=0.0,a=0.0,up=0.0,uf=0.0,n=0,multi=0)
    if g(r,'ЕО',1)!=x['eon'] or g(r,'Техническое место',1)!=x['tmn']: x['multi']+=1
    x['p']+=N(r,'Стоимость МТР, План'); x['a']+=N(r,'Стоимость МТР, Факт'); x['up']+=N(r,'Стоимость УСО, План'); x['uf']+=N(r,'Стоимость УСО, Факт'); x['n']+=1
json.dump({'name':name,'cols':sorted(H),'hdr':{k:v for k,v in H.items()},'orders':O}, gzip.open(out,'wt',encoding='utf-8'), ensure_ascii=False)
print(name, 'заказов', len(O))
