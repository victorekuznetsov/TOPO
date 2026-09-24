#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Метаданные исходных выгрузок PM-06, которых нет в data/*.json:
статусы и фаза заказа, признак копии при переносе в АО «Развитие» и
признак ППМ компонентов («Резерв./заявка»).

Два шага, чтобы не держать все выгрузки на диске одновременно:

  extract  — одна выгрузка (xlsx или склеенный zip из ветки rawdata) →
             промежуточный JSON по заказам в рабочей папке;
  assemble — все промежуточные JSON → order_status/ и ppm_flags/.

    python3 build_pm06_meta.py extract M06_1200_2026.zip 1200_2026 work/
    python3 build_pm06_meta.py assemble work/ .

Правила (подтверждены документацией SAP и ответами бизнеса, см.
README.md и pm06_status_catalog.json → meta.userAnswers):

  * фаза заказа — один из системных кодов: ЗАКР > ТЗКР > ДЕБЛ/ЧДЕБ > ОТКР;
    строки без статуса (#) — позиции графика ППР без заказа SAP, в план
    не входят;
  * копия переноса — заказ планирующего завода АО «Развитие» и заказ
    планирующего завода БЕ на ту же ЕО, вид работ и базисную дату начала.
    Оригинал БЕ помечается: его факт настоящий, но неисполненный план
    дублирует копию;
  * признак ППМ компонента: «Немедленно» (передан в ППМ сразу) — значение
    по умолчанию и не хранится; «Начиная с деблок.» и «Никогда» хранятся
    как исключения по паре заказ × № ЕКМТР.
"""
import io
import json
import os
import re
import subprocess
import sys
import zipfile
from collections import defaultdict
from xml.etree.ElementTree import iterparse

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
BE_SITE = {'АО "Полюс Красноярск"': "1100", 'АО "Полюс Вернинское"': "1200",
           'АО "Полюс Алдан"': "1300", 'АО "Полюс Магадан"': "1400",
           'ООО "Полюс Сухой Лог"': "2400"}
PHASES = ["нет статуса", "ОТКР", "ДЕБЛ", "ТЗКР", "ЗАКР"]
PPM_FLAGS = ["Начиная с деблокир.", "Никогда"]    # «Немедленно» — по умолчанию


def ppm_index(value):
    """Индекс признака «Резерв./заявка» в PPM_FLAGS; None — «Немедленно»
    или пусто. Сравнение по началу текста: BW режет подписи по длине
    («Начиная с деблокир.»), точное совпадение хрупко."""
    v = (value or "").strip().lower()
    if v.startswith("никогда"):
        return 1
    if v.startswith("начиная с деблок"):
        return 0
    return None


def phase_of(sys_status):
    s = set(sys_status.split())
    if not s or s == {"#"}:
        return 0
    if "ЗАКР" in s:
        return 4
    if "ТЗКР" in s:
        return 3
    if s & {"ДЕБЛ", "ЧДЕБ"}:
        return 2
    return 1


# ---------- чтение xlsx потоком ------------------------------------------
def _col(ref):
    n = 0
    for ch in re.match(r"([A-Z]+)", ref).group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def _open(path):
    if path.lower().endswith(".xlsx"):
        return zipfile.ZipFile(path)
    outer = zipfile.ZipFile(path)
    inner = [n for n in outer.namelist() if n.lower().endswith(".xlsx")][0]
    try:
        data = outer.read(inner)
    except NotImplementedError:          # Deflate64 — системный unzip
        data = subprocess.run(["unzip", "-p", path, inner], check=True,
                              capture_output=True).stdout
    return zipfile.ZipFile(io.BytesIO(data))


def _sheet(z, name):
    wb = z.read("xl/workbook.xml").decode()
    rels = z.read("xl/_rels/workbook.xml.rels").decode()
    tgt = {}
    for m in re.finditer(r"<Relationship [^>]*>", rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(0)))
        tgt[a.get("Id")] = a.get("Target")
    for m in re.finditer(r"<sheet [^>]*>", wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(0)))
        if a.get("name") == name:
            return "xl/" + tgt[a["r:id"]].lstrip("/").replace("xl/", "")
    raise SystemExit(f"нет листа «{name}»")


def _rows(z, sheet):
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for _, el in iterparse(z.open("xl/sharedStrings.xml")):
            if el.tag == NS + "si":
                shared.append("".join(t.text or "" for t in el.iter(NS + "t")))
                el.clear()
    for _, el in iterparse(z.open(sheet)):
        if el.tag != NS + "row":
            continue
        row = {}
        for c in el.findall(NS + "c"):
            t, v, inl = c.get("t"), c.find(NS + "v"), c.find(NS + "is")
            if t == "s" and v is not None:
                val = shared[int(v.text)]
            elif t == "inlineStr" and inl is not None:
                val = "".join(x.text or "" for x in inl.iter(NS + "t"))
            elif v is not None:
                val = v.text
            else:
                continue
            row[_col(c.get("r"))] = val
        yield row
        el.clear()


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


# ---------- extract --------------------------------------------------------
def extract(path, name, work):
    site0, year = name.split("_")
    z = _open(path)
    rows = _rows(z, _sheet(z, "PM-06"))
    hdr = None
    for r in rows:
        if any(isinstance(v, str) and v.strip() == "Заказ" for v in r.values()):
            hdr = r
            break
    if hdr is None:
        raise SystemExit("нет строки заголовка с колонкой «Заказ»")
    H = {v.strip(): k for k, v in hdr.items() if isinstance(v, str) and v.strip()}

    def col(n, off=0):
        return H[n] + off if n in H else None
    C = dict(o=col("Заказ"), be=col("Балансовая единица"), eo=col("ЕО"), eon=col("ЕО", 1),
             sys=col("Заказ Системный статус"), usr=col("Заказ Пользовательский статус"),
             kind=col("Заказ Вид заказа"), kindt=col("Заказ Вид заказа", 1),
             w=col("Заказ Вид работы ТОРО"), bs=col("Заказ Базисный срок начала (дата)"),
             pp=col("Заказ Завод, планирующий ТОРО"),
             ek=col("Компонент Заказа/Заявки", 1), rz=col("Компонент Заказа Резерв./заявка"),
             p=col("Стоимость МТР, План"), a=col("Стоимость МТР, Факт"),
             up=col("Стоимость УСО, План"), uf=col("Стоимость УСО, Факт"),
             qf=col("Количество МТР, Факт"), sv=col("Стоимость услуг, Факт"))
    get = lambda r, k: r.get(C[k]) if C[k] is not None else None
    orders, comps = {}, defaultdict(dict)
    rz_seen = defaultdict(int)
    for r in rows:
        o = get(r, "o")
        if not o or o == "#":
            continue
        g = orders.get(o)
        if g is None:
            be = (get(r, "be") or "").strip()
            g = orders[o] = {
                "site": BE_SITE.get(be, site0), "sys": (get(r, "sys") or "#").strip(),
                "usr": (get(r, "usr") or "#").strip(), "kind": get(r, "kind") or "",
                "kindText": get(r, "kindt") or "", "w": get(r, "w") or "",
                "bs": get(r, "bs") or "", "pp": get(r, "pp") or "",
                "eon": get(r, "eon") or "", "wk": False, "plan": 0.0, "fact": 0.0}
        if "WK" in (get(r, "eo") or "").upper().replace("-", ""):
            g["wk"] = True
        g["plan"] += _num(get(r, "p")) + _num(get(r, "up"))
        g["fact"] += _num(get(r, "a")) + _num(get(r, "uf")) + _num(get(r, "sv"))
        rz, code = get(r, "rz"), get(r, "ek")
        rz_seen[rz or ""] += 1
        idx = ppm_index(rz)
        if idx is not None and code and code != "#":
            # если у материала в заказе несколько строк с разными признаками,
            # оставляем «худший» для закупки: Никогда > Начиная с деблокир.
            prev = comps[o].get(code)
            comps[o][code] = max(idx, prev) if prev is not None else idx
    os.makedirs(work, exist_ok=True)
    with open(os.path.join(work, name + ".json"), "w", encoding="utf-8") as f:
        json.dump({"file": name, "year": year, "orders": orders, "comps": comps},
                  f, ensure_ascii=False, separators=(",", ":"))
    unknown = {k: v for k, v in rz_seen.items() if k and ppm_index(k) is None and not k.lower().startswith("немедленно")}
    print(f"{name}: заказов {len(orders)}, заказов с признаком ППМ ≠ «Немедленно» {len(comps)}; "
          f"значения «Резерв./заявка»: {dict(rz_seen)}" + (f"; НЕИЗВЕСТНЫЕ: {unknown}" if unknown else ""))


# ---------- assemble -------------------------------------------------------
def migration_originals(orders):
    """Номера заказов БЕ, у которых есть копия АО «Развитие» (та же ЕО,
    вид работ и базисная дата начала). Поле «планирующий завод» есть
    только в выгрузках с 2024 года."""
    groups = defaultdict(list)
    for o, v in orders.items():
        if v["eon"] and v["eon"] != "#" and v["bs"] and v["pp"]:
            groups[(v["eon"], v["w"], v["bs"])].append(o)
    originals = set()
    for lst in groups.values():
        rz = [o for o in lst if "Развити" in orders[o]["pp"]]
        be = [o for o in lst if o not in rz]
        if rz and be:
            originals.update(be)
    return originals


def assemble(work, out):
    by_key = defaultdict(dict)
    comps_by_key = defaultdict(dict)
    sources = defaultdict(set)
    for fn in sorted(os.listdir(work)):
        if not fn.endswith(".json"):
            continue
        with open(os.path.join(work, fn), encoding="utf-8") as f:
            d = json.load(f)
        originals = migration_originals(d["orders"])
        for o, v in d["orders"].items():
            key = (v["site"], d["year"])
            v["copyOriginal"] = o in originals
            by_key[key][o] = v
            sources[key].add(d["file"])
            if o in d["comps"]:
                comps_by_key[key][o] = d["comps"][o]
    for sub in ("order_status", "ppm_flags"):
        os.makedirs(os.path.join(out, sub), exist_ok=True)
    norm = lambda s: " ".join(sorted(s.split()))
    total = 0
    for (site, year), O in sorted(by_key.items()):
        sysd = sorted({norm(v["sys"]) for v in O.values()})
        usrd = sorted({norm(v["usr"]) for v in O.values()})
        kd = sorted({v["kind"] for v in O.values()})
        si = {x: i for i, x in enumerate(sysd)}
        ui = {x: i for i, x in enumerate(usrd)}
        ki = {x: i for i, x in enumerate(kd)}
        kt = {v["kind"]: v["kindText"] for v in O.values()}
        src = "TOPO ветка rawdata: " + ", ".join("M06_" + f for f in sorted(sources[(site, year)]))
        status = {
            "meta": {"site": site, "year": year, "orders": len(O), "source": src,
                     "columns": "o[заказ] = [индекс в sys, индекс в usr, индекс в kind, фаза, копия-оригинал]",
                     "phases": PHASES,
                     "phaseRule": "ЗАКР > ТЗКР > ДЕБЛ/ЧДЕБ > ОТКР; нет статуса (#) — позиция графика ППР без заказа SAP, в план не входит. Заказ закрыт при фазе ТЗКР или ЗАКР.",
                     "copyRule": "1 — заказ БЕ, у которого есть копия АО «Развитие» (та же ЕО, вид работ, базисная дата начала): факт учитывать, неисполненный план — нет."},
            "sys": sysd, "usr": usrd, "kind": kd, "kindText": {k: kt[k] for k in kd},
            "o": {o: [si[norm(v["sys"])], ui[norm(v["usr"])], ki[v["kind"]],
                      phase_of(v["sys"]), 1 if v["copyOriginal"] else 0]
                  for o, v in sorted(O.items())}}
        with open(os.path.join(out, "order_status", f"{site}_{year}.json"), "w", encoding="utf-8") as f:
            json.dump(status, f, ensure_ascii=False, separators=(",", ":"))
        flags = {"meta": {"site": site, "year": year, "source": src,
                          "flags": PPM_FLAGS, "default": "Немедленно",
                          "columns": "o[заказ][№ ЕКМТР] = индекс в flags; пары, которых нет, — «Немедленно»",
                          "meaning": "«Резерв./заявка» — релевантность для ППМ: «Немедленно» — потребность передана в ППМ сразу; «Начиная с деблок.» — только после деблокирования заказа; «Никогда» — в ППМ и закупочную заявку не попадает."},
                 "o": dict(sorted(comps_by_key[(site, year)].items()))}
        with open(os.path.join(out, "ppm_flags", f"{site}_{year}.json"), "w", encoding="utf-8") as f:
            json.dump(flags, f, ensure_ascii=False, separators=(",", ":"))
        total += len(O)
        print(f"{site}_{year}: заказов {len(O)}, копий-оригиналов {sum(v['copyOriginal'] for v in O.values())}, "
              f"заказов с исключениями ППМ {len(comps_by_key[(site, year)])}")
    print("всего заказов", total)


if __name__ == "__main__":
    if len(sys.argv) < 2 or sys.argv[1] not in ("extract", "assemble"):
        print(__doc__, file=sys.stderr)
        sys.exit(1)
    if sys.argv[1] == "extract":
        extract(*sys.argv[2:5])
    else:
        assemble(*sys.argv[2:4])
