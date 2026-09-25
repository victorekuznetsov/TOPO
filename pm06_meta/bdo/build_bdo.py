#!/usr/bin/env python3
"""
БДО (база данных оборудования) из всех заказов ТОРО PM-06.

  python3 build_bdo.py <папка со сканами *.json.gz> <папка вывода>

Скан — scan_pm06_bdo.py (по одной выгрузке, см. build_all.sh). Выгрузка PM-06 даёт у заказа
техническое место (код + имя) и единицу оборудования (номер + имя), поэтому справочник покрывает
только объекты, на которые были заказы 2022–2027; промежуточные уровни ТМ без заказов известны
только кодом (имя пустое).

Выход:
  tm.json — узлы ТМ: код → [имя, уровень, родитель, первый год, последний год, заказов, ЕО]
  eo.json — ЕО: номер → [имя, {год: ТМ}, марка заводская, изготовитель, дата ввода, заказов]
  wk.json — ветка экскаваторов WK: машина (ТМ) → цепочка предков, ЕО на ТМ, заказы на узлы
            (ЕО не экскаватор: ковш, ЭД, рукоять…) — их отчёт по ЕО «WK» раньше не видел
"""
import json, gzip, glob, os, re, sys, collections

WK = re.compile(r'WK-?\d', re.I)
real = lambda c: bool(c) and c != '#'
BE_OF = {'03': 'АО «Полюс Вернинское»', '04': 'АО «Полюс Алдан»', '05': 'АО «Полюс Магадан»',
         '06': 'ООО «Полюс Сухой Лог»', '11': 'АО «Полюс Красноярск»'}


def main(src, out):
    os.makedirs(out, exist_ok=True)
    X = []
    files = sorted(glob.glob(os.path.join(src, '*.json.gz')))
    for fn in files:
        d = json.load(gzip.open(fn, 'rt', encoding='utf-8'))
        if 'hdr' not in d:
            raise SystemExit(f'{fn}: старый формат скана — пересканируйте')
        y = d['name'].split('_')[1]
        for o, x in d['orders'].items():
            x.update(file=d['name'], fy=y, order=o)
            X.append(x)
    # ── ТМ
    tm = {}
    for x in X:
        c = x['tmn']
        if not real(c):
            continue
        t = tm.setdefault(c, {'name': x['tm'], 'years': set(), 'orders': 0, 'eo': set()})
        t['years'].add(x['fy']); t['orders'] += 1
        if real(x['eon']):
            t['eo'].add(x['eon'])
    nodes = {}
    for c in list(tm):
        p = c.split('-')
        for i in range(1, len(p) + 1):
            k = '-'.join(p[:i])
            if k not in nodes:
                t = tm.get(k)
                nodes[k] = [t['name'] if t else '', i, '-'.join(p[:i - 1]) if i > 1 else '',
                            min(t['years']) if t else '', max(t['years']) if t else '',
                            t['orders'] if t else 0, len(t['eo']) if t else 0]
    # ── ЕО
    eo = {}
    for x in X:
        e = x['eon']
        if not real(e):
            continue
        r = eo.setdefault(e, {'names': collections.Counter(), 'tm': {}, 'mark': '', 'mfr': '', 'din': '', 'orders': 0})
        r['names'][x['eo']] += 1; r['orders'] += 1
        if real(x['tmn']):
            r['tm'][x['fy']] = x['tmn']
        for k in ('mark', 'mfr', 'din'):
            if x.get(k) and x[k] != '#' and not r[k]:
                r[k] = x[k]
    eo_out = {e: [r['names'].most_common(1)[0][0], dict(sorted(r['tm'].items())), r['mark'], r['mfr'], r['din'], r['orders']]
              for e, r in eo.items()}
    moved = {e: v for e, v in eo_out.items() if len(set(v[1].values())) > 1}
    # ── WK
    wk_tm = {c: t['name'] for c, t in tm.items() if WK.search(t['name'])}

    def wk_root(c):
        p = c.split('-')
        for i in range(len(p), 0, -1):
            k = '-'.join(p[:i])
            if k in wk_tm:
                return k
        return None
    machines = {}
    for c, n in sorted(wk_tm.items()):
        p = c.split('-')
        machines[c] = {'name': n, 'be': BE_OF.get(p[0], p[0]), 'chain': ['-'.join(p[:i]) for i in range(1, len(p))],
                       'eo': {}, 'componentOrders': []}
    keep = ('order', 'file', 'fy', 'be', 'eo', 'eon', 'tm', 'tmn', 'txt', 'kind', 'w', 'wt', 'bs', 'sys', 'usr', 'pg', 'mu', 'p', 'a', 'up', 'uf', 'n')
    for x in X:
        if not real(x['tmn']):
            continue
        r = wk_root(x['tmn'])
        if not r:
            continue
        m = machines[r]
        if real(x['eon']):
            e = m['eo'].setdefault(x['eon'], {'name': x['eo'], 'isMachine': bool(WK.search(x['eo'])), 'orders': 0, 'years': set()})
            e['orders'] += 1; e['years'].add(x['fy'])
        if not WK.search(x['eo']):
            m['componentOrders'].append({k: (round(x[k], 2) if isinstance(x[k], float) else x[k]) for k in keep})
    for m in machines.values():
        for e in m['eo'].values():
            e['years'] = sorted(e['years'])
        m['componentOrders'].sort(key=lambda o: (o['fy'], o['order']))
    comp = [o for m in machines.values() for o in m['componentOrders']]
    no_obj = [x for x in X if not real(x['tmn']) and not real(x['eon'])]
    meta = {
        'source': 'TOPO ветка rawdata: M06_{завод}_{год} (PM-06), ' + ', '.join(os.path.basename(f).split('.')[0] for f in files),
        'orders': len(X), 'tm': len(tm), 'nodes': len(nodes), 'eo': len(eo),
        'noObjectOrders': len(no_obj), 'noObjectKinds': dict(collections.Counter(x['kind'] for x in no_obj).most_common()),
        'eoWithoutTm': sum(1 for x in X if real(x['eon']) and not real(x['tmn'])),
        'tmWithSeveralEo': sum(1 for t in tm.values() if len(t['eo']) > 1),
        'eoMoved': len(moved),
        'levels': dict(sorted(collections.Counter(v[1] for c, v in nodes.items() if c in tm).items())),
        'be': BE_OF,
        'rule': 'Код ТМ — сегменты через «-»: 1-й — БЕ (03 Вернинское, 04 Алдан, 05 Магадан, 06 Сухой Лог, 11 Красноярск), '
                'далее производство (ZIF, RUD, KAR, ATC…), цех/карьер, участок, технологическая позиция (машина), группирующие ТМ. '
                'Родитель — код без последнего сегмента. ЕО монтируется на ТМ; заказ ведётся на ЕО и/или ТМ.',
    }
    dump = lambda n, o: json.dump(o, open(os.path.join(out, n), 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    dump('tm.json', {'meta': meta, 'columns': ['name', 'level', 'parent', 'firstYear', 'lastYear', 'orders', 'eo'], 'nodes': nodes})
    dump('eo.json', {'meta': meta, 'columns': ['name', 'tmByYear', 'mark', 'mfr', 'commissioned', 'orders'], 'eo': eo_out})
    dump('wk.json', {'meta': {**meta, 'machines': len(machines), 'componentOrders': len(comp),
                              'componentPlan': round(sum(o['p'] + o['up'] for o in comp), 2),
                              'componentFact': round(sum(o['a'] + o['uf'] for o in comp), 2)},
                     'machines': machines}, )
    print(f"заказов {len(X)}; ТМ {len(tm)} (узлов с предками {len(nodes)}); ЕО {len(eo)}; перемонтаж ЕО {len(moved)}; "
          f"машин WK {len(machines)}; заказов на узлы WK {len(comp)} на {sum(o['p'] + o['up'] for o in comp) / 1e6:.1f} млн ₽ плана")


if __name__ == '__main__':
    main(*sys.argv[1:3])
