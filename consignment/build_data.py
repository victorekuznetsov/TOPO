# -*- coding: utf-8 -*-
"""Сборка данных для модели внешней консигнации подрядчика CAT.

    python consignment/build_data.py --src "<CAT>/Остатки-закупки-аналоги" \
        [--stock-js "<CAT>/data/stock.js"] [--out consignment/cons_data.json]

Одна строка на позицию площадки (поле и цех вместе):
  * план 2027 — предварительные заказы ТОРО (data/<БЕ>_2027.json этого репозитория);
  * спрос вне плана — факт 2024 – авг 2026 по видам работ вне ТО и КР;
  * признак CAPEX/OPEX — поле «Причина инвестиций» заказа;
  * остатки (MM-M03) и закупка АО «Развитие» — только по заводам площадки,
    без складов подрядчиков (категория склада 5), консигнации и АТЗ;
  * замены номеров Caterpillar — из data/stock.js каталога CAT (необязательно).
Подробности — consignment/README.md и навык .claude/skills/consignment.
"""
import argparse, glob, json, re, collections, openpyxl, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
ap.add_argument("--src", required=True, help="папка с выгрузками ЕКМТР_*.xlsx, Остатки_*.xlsx, Закупка*.XLSX")
ap.add_argument("--topo", default=os.path.join(os.path.dirname(HERE), "data"), help="папка data/ репозитория TOPO")
ap.add_argument("--stock-js", default="", help="data/stock.js каталога CAT — для замен номеров Cat (необязательно)")
ap.add_argument("--out", default=os.path.join(HERE, "cons_data.json"))
A = ap.parse_args()
TOPO, SRC = A.topo, A.src

def one(*masks):
    for m in masks:
        got = sorted(glob.glob(os.path.join(SRC, m)))
        if got: return got[-1]
    raise SystemExit("не нашёл %s в %s" % (" / ".join(masks), SRC))
F_EK, F_ST, F_BUY = one("ЕКМТР*.xlsx"), one("Остатки*.xlsx"), one("Закупка*.XLSX", "Закупка*.xlsx")
m = re.search(r"(\d{2})_(\d{2})_(\d{4})", os.path.basename(F_BUY))
ASOF = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else datetime.date.today()
print("читаю:", *(os.path.basename(f) for f in (F_EK, F_ST, F_BUY)), "| дата закупки", ASOF, flush=True)
norm = lambda s: re.sub(r"[\s\-]", "", str(s or "")).upper()
HIST = ["2024", "2025", "2026"]

SITES = {
  "1100": {"eo": re.compile(r"(Самосвал карьерный|Кузов) 793D|^ДВС 3516"),
           "be": 'АО "Полюс Красноярск"', "razv": 'АО "Развитие" Еруда\\Благ.', "razv2": 'АО "Развитие" ПБ КБЕ', "plants": ("7101",), "other_plants": ("7106",), "stock_plants": ("1100", "7101")},
  "1300": {"eo": re.compile(r"Самосвал карьерный 77[37][EG]|Экскаватор гидравл\S* 395|^ДВС C(27|32)"),
           "be": 'АО "Полюс Алдан"', "razv": 'АО "Развитие" Алдан', "plants": ("7103",), "other_plants": (), "stock_plants": ("1300", "7103")},
}
def model_of(e):
    for m, rx in (("793D", r"793D|3516"), ("773E", r"773E|C27"), ("777G", r"777G|C32"), ("395", r"395")):
        if re.search(rx, e): return m
    return "?"
ENGINE = re.compile(r"^ДВС")
RE_TO = re.compile(r"(?i)^(техническое обслуж|ежесмен|сезонн|годовое)")
RE_KR = re.compile(r"(?i)^(капитальный ремонт|капитализируемый|кр |средний ремонт)")
def wclass(w):
    if RE_TO.search(w): return "ТО"
    if RE_KR.search(w): return "КР"
    return "ВНЕ"          # текущий ремонт, замена, монтаж, восстановление, прочие, не присвоено
def part_of(e, w):
    return "Цех" if (ENGINE.search(e) or RE_KR.search(w) or re.match(r"(?i)восстановление", w)) else "Поле"
def capex_of(reason):
    r = reason or ""
    if re.search(r"CAPEX|Капитализ|Компонент", r): return "CAPEX"
    return "OPEX"
EXCL = [
  ("масла и смазки", re.compile(r"(?i)^(масло|смазка|жидкость|антифриз|охлаждающая жидк)")),
  ("шины", re.compile(r"(?i)^(шина(?! заземл)|покрышка|камера шин)")),
  ("GET", re.compile(r"(?i)^(коронка|палец коронки|фиксатор коронки|зуб\b|нож\b|нож |резак|накладка ковша|протектор ковша|угол отвала|кромка)")),
  ("агрегат", re.compile(r"(?i)^(двигатель|дв\.\s*сб|дв\.сб|кпп|пер\.конечн|передача конечн|диф\.\s*сб|дифференциал|гидр/трансф|гидротрансформ|группа колеса|ступица сб|подвеска|круг поворотный|привод поворота|редуктор поворот|цилиндр подъемника|цилиндр стрелы|цилиндр ковша|гидроцилиндр|насос гл\.|главная передача|мост )")),
]
def excl_reason(name):
    for why, rx in EXCL:
        if rx.search(name): return why
    return ""

print("ЕК МТР…", flush=True)
EK = {}
wb = openpyxl.load_workbook(F_EK, read_only=True, data_only=True)
ws = wb["Лист1"] if "Лист1" in wb.sheetnames else wb.worksheets[0]
for i, r in enumerate(ws.iter_rows(values_only=True)):
    if not i: continue
    name, code, no, _g, maker = (list(r) + [None]*5)[:5]
    if code: EK[str(code).strip()] = (str(name or "").strip(), str(no or "").strip(), str(maker or "").strip())
wb.close()

def month_of(d, r, y):
    s = d["nd"][d["ndi"][r]] or (d["od"].get(d["o"][r]) or [""])[0] or ""
    if not s: return 6
    yy, mm = int(s[:4]), int(s[5:7])
    return 1 if yy < int(y) else 12 if yy > int(y) else mm

def new_item(site, ek, name, eno):
    return {"site": site, "ek": ek, "name": name,
            "no": eno or (re.findall(r"\b\d?[A-Z0-9]{2,4}-?\d{3,4}\b", name) or [""])[-1],
            "models": set(), "wk": collections.Counter(), "rs": collections.Counter(),
            "q27": 0.0, "v27": 0.0, "m27": [0.0]*12, "q27_part": collections.Counter(), "q27_cap": collections.Counter(),
            "q27_vne": 0.0, "picks": 0, "maxpick": 0.0,
            "h_ord": {y: set() for y in HIST}, "h_q": {y: 0.0 for y in HIST}, "h_v": {y: 0.0 for y in HIST},
            "h_part": collections.Counter(), "h_cap": collections.Counter(), "h_maxpick": 0.0, "h_models": set(),
            "h_wk": collections.Counter(), "hall_q": 0.0, "hall_v": 0.0}

items, excluded = {}, collections.defaultdict(lambda: {"q": 0.0, "v": 0.0})
for site, S in SITES.items():
    for y in ["2027"] + HIST:
        print(site, y, flush=True)
        d = json.load(open(f"{TOPO}/{site}_{y}.json"))
        eo = {i: e for i, e in enumerate(d["e"]) if S["eo"].search(e)}
        pick = collections.defaultdict(float)
        for r in range(d["n"]):
            e = eo.get(d["ei"][r])
            if e is None: continue
            ek = d["cek"][d["ci"][r]]
            if not ek or ek.startswith("SRV"): continue
            name = d["c"][d["ci"][r]]
            ekn, eno, emk = EK.get(ek, ("", "", ""))
            maker = d["mf"][d["mfi"][r]] or ""
            if not (maker.lower().startswith("caterp") or emk.lower().startswith("caterp")): continue
            w = d["w"][d["wi"][r]]
            q = (d["qp"][r] if y == "2027" else d["qf"][r]) or 0
            v = d["p"][r] if y == "2027" else d["a"][r]
            if q <= 0: continue
            why = excl_reason(name)
            if why:
                if y == "2027":
                    k = (site, ek, name, why); excluded[k]["q"] += q; excluded[k]["v"] += v
                continue
            key = (site, ek)
            it = items.get(key) or items.setdefault(key, new_item(site, ek, name, eno))
            reason = d["orr"].get(d["o"][r], "") or "Не присвоено"
            part, cap, wc = part_of(e, w), capex_of(reason), wclass(w)
            if y == "2027":
                it["q27"] += q; it["v27"] += v; it["m27"][month_of(d, r, y) - 1] += q
                it["models"].add(model_of(e)); it["wk"][w] += q; it["rs"][reason] += q
                it["q27_part"][part] += q; it["q27_cap"][cap] += q
                if wc == "ВНЕ": it["q27_vne"] += q
                pick[(key, d["o"][r])] += q
            else:
                it["hall_q"] += q; it["hall_v"] += v
                if wc != "ВНЕ": continue
                it["h_ord"][y].add(d["o"][r]); it["h_q"][y] += q; it["h_v"][y] += v
                it["h_part"][part] += q; it["h_cap"][cap] += q; it["h_models"].add(model_of(e))
                it["h_wk"][w] += q
                pick[(key, d["o"][r])] += q
        for (key, o), q in pick.items():
            it = items[key]
            if y == "2027":
                it["picks"] += 1; it["maxpick"] = max(it["maxpick"], q)
            else:
                it["h_maxpick"] = max(it["h_maxpick"], q)

items = {k: v for k, v in items.items() if v["q27"] > 0 or sum(v["h_q"].values()) > 0}
print("позиций: план или вне плана:", len(items))

# ---------- остатки: категория склада «5» — склады подрядчиков, не учитываем
print("остатки…", flush=True)
reg_of = {}
for site, S in SITES.items():
    reg_of[S["be"]] = (site, "БЕ"); reg_of[S["razv"]] = (site, "Развитие")
    if S.get("razv2"): reg_of[S["razv2"]] = (site, "Развитие")
stock = collections.defaultdict(collections.Counter)
stock_no = collections.defaultdict(collections.Counter)
wb = openpyxl.load_workbook(F_ST, read_only=True, data_only=True)
ws = wb["MM-M03"]
for i, r in enumerate(ws.iter_rows(values_only=True)):
    if i <= 11 or len(r) < 21: continue
    reg = reg_of.get(r[2]) or reg_of.get(r[4])
    if not reg: continue
    try: q = float(r[19] or 0)
    except (TypeError, ValueError): continue
    if q <= 0: continue
    site, owner = reg
    stcat, store = str(r[16] or ""), str(r[18] or "")
    if str(r[3] or "").strip() not in SITES[site]["stock_plants"]: b = "Другой завод"
    elif r[8] == "АТЗ": b = "АТЗ"
    elif stcat == "5": b = "Подрядчик"
    elif "онсигнац" in store: b = "Консигнация"
    elif store in ("Виртуальный", "Склад разниц") or stcat == "99": b = "Прочее"
    else: b = owner
    stock[(site, str(r[9] or "").strip())][b] += q
    no = norm(r[11]) if r[11] not in (None, "#") else ""
    if no: stock_no[(site, no)][b] += q
wb.close()

# ---------- закупка: открытые заказы и заявки АО «Развитие»
print("закупка…", flush=True)
P2S = {p: site for site, S in SITES.items() for p in S["plants"] + S["other_plants"]}
def dt(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    return None
sup = collections.defaultdict(lambda: {"other": 0.0, "open": 0.0, "late": 0.0, "m": [0.0]*12, "after": 0.0, "pend": 0.0,
                                       "sups": collections.Counter(), "docs": set()})
wb = openpyxl.load_workbook(F_BUY, read_only=True, data_only=True)
it_ = wb.worksheets[0].iter_rows(values_only=True); next(it_)
lead_analog = collections.Counter()
for r in it_:
    site = P2S.get(str(r[2] or "").strip())
    if not site: continue
    code = str(r[16] or r[14] or "").strip()
    doc = str(r[23] or "").strip()
    status = str(r[30] or "").strip()
    try:
        left, tr, qty = float(r[32] or 0), float(r[33] or 0), float(r[22] or 0)
    except (TypeError, ValueError):
        continue
    x = sup[(site, code)]
    lead = str(r[12] or "").strip()
    if lead and lead != code: lead_analog[(site, lead)] += 1
    firm = bool(doc) and status in ("Действующий", "Верифицирован", "")
    if str(r[2] or "").strip() not in SITES[site]["plants"]:
        if firm: x["other"] += max(left, tr)
        continue
    if firm:
        q = max(left, tr)
        if q <= 0: continue
        d = dt(r[36]) or dt(r[37]) or dt(r[13])
        if d is None or d.year < 2027:
            x["open"] += q
            if d is None or d < ASOF: x["late"] += q
        elif d.year == 2027: x["m"][d.month - 1] += q
        else: x["after"] += q
        x["sups"][str(r[27] or "").strip()] += q; x["docs"].add(doc)
    else:
        q = qty if not doc else max(left, tr, qty)
        d = dt(r[13])
        if q > 0 and (d is None or d.year <= 2027): x["pend"] += q
wb.close()

REP = {}
if A.stock_js:
    t = open(A.stock_js, encoding="utf-8").read()
    REP = json.loads(t[t.index("= ") + 2:].rstrip().rstrip(";\n")).get("rep", {})
    print("замены Cat:", len(REP), "номеров", flush=True)
else:
    print("замены Cat не заданы (--stock-js) — остаток по заменам будет нулевым", flush=True)

def mix(counter, a, b):
    ta, tb = counter.get(a, 0), counter.get(b, 0)
    if ta and tb: return a + "+" + b
    return a if ta else b if tb else ""

rows = []
for (site, ek), it in items.items():
    s = stock.get((site, ek), collections.Counter())
    fam = [x[0] for x in REP.get(norm(it["no"]), [])]
    rep_q = sum(stock_no.get((site, n), collections.Counter())["БЕ"] + stock_no.get((site, n), collections.Counter())["Развитие"] for n in fam)
    hq = sum(it["h_q"].values())
    if it["q27"]:
        price = it["v27"] / it["q27"]
    else:
        hv = sum(it["h_v"].values()); price = hv / hq if hq else (it["hall_v"] / it["hall_q"] if it["hall_q"] else 0)
    cap_tot = it["q27_cap"] + it["h_cap"]; part_tot = it["q27_part"] + it["h_part"]
    rows.append({
        "site": site, "ek": ek, "no": it["no"], "name": it["name"],
        "models": ", ".join(sorted(it["models"] | it["h_models"])),
        "part": mix(part_tot, "Поле", "Цех"), "capex": mix(cap_tot, "CAPEX", "OPEX"),
        "rs": "; ".join(k for k, _ in (it["rs"] if it["rs"] else collections.Counter()).most_common(2)),
        "src": "План + вне плана" if it["q27"] and hq else "План" if it["q27"] else "Вне плана",
        "work": ", ".join(k for k, _ in (it["wk"] + it["h_wk"]).most_common(3)),
        "q27": round(it["q27"], 3), "q27_pole": round(it["q27_part"]["Поле"], 3), "q27_ceh": round(it["q27_part"]["Цех"], 3),
        "q27_capex": round(it["q27_cap"]["CAPEX"], 3), "q27_vne": round(it["q27_vne"], 3),
        "m27": [round(x, 3) for x in it["m27"]], "picks": it["picks"], "maxpick": round(it["maxpick"], 3),
        "h_o": [len(it["h_ord"][y]) for y in HIST], "h_q": [round(it["h_q"][y], 3) for y in HIST],
        "h_maxpick": round(it["h_maxpick"], 3),
        "h_pole_share": round(it["h_part"]["Поле"] / hq, 4) if hq else 0,
        "h_capex_share": round(it["h_cap"]["CAPEX"] / hq, 4) if hq else 0,
        "price": round(price, 2),
        "st_be": s["БЕ"], "st_razv": s["Развитие"], "st_rep": rep_q, "rep": " ".join(fam[:4]),
        "st_contr": s["Подрядчик"], "st_cons": s["Консигнация"], "st_atz": s["АТЗ"],
        "del_open": round(sup[(site, ek)]["open"], 3), "del_late": round(sup[(site, ek)]["late"], 3),
        "del_m": [round(v, 3) for v in sup[(site, ek)]["m"]], "del_after": round(sup[(site, ek)]["after"], 3),
        "del_pend": round(sup[(site, ek)]["pend"], 3),
        "del_sup": "; ".join(k for k, _ in sup[(site, ek)]["sups"].most_common(2)),
        "del_docs": len(sup[(site, ek)]["docs"]), "lead_analog": lead_analog[(site, ek)],
        "st_other": s["Другой завод"], "del_other": round(sup[(site, ek)]["other"], 3)})

exc = [{"site": k[0], "ek": k[1], "name": k[2], "why": k[3], "q": round(v["q"], 3), "v": round(v["v"])} for k, v in excluded.items()]
json.dump({"rows": rows, "excluded": exc, "hist": HIST, "built": datetime.date.today().isoformat(),
           "asof": ASOF.isoformat(), "files": [os.path.basename(f) for f in (F_EK, F_ST, F_BUY)]},
          open(A.out, "w"), ensure_ascii=False)
print("записано", A.out)
for site in SITES:
    rs = [r for r in rows if r["site"] == site]
    c = collections.Counter(r["src"] for r in rs)
    print(site, dict(c), "| план 2027 %.1f млн" % (sum(r["q27"] * r["price"] for r in rs) / 1e6),
          "| вне плана/год %.1f млн" % (sum(sum(r["h_q"]) * r["price"] for r in rs) / 2.667 / 1e6),
          "| остаток у подрядчиков по позициям %.1f млн" % (sum(r["st_contr"] * r["price"] for r in rs) / 1e6))
    print("   исключено как другие заводы: остаток %.1f млн, поставки по заказам %.1f млн" % (
        sum(r["st_other"] * r["price"] for r in rs) / 1e6, sum(r["del_other"] * r["price"] for r in rs) / 1e6))
    print("   поставки: до 2027 %.1f млн (просрочено %.1f), в 2027 %.1f млн, после %.1f, неподтверждённые %.1f; позиций с поставками %d, заказан аналог вместо позиции %d" % (
        sum(r["del_open"] * r["price"] for r in rs) / 1e6, sum(r["del_late"] * r["price"] for r in rs) / 1e6,
        sum(sum(r["del_m"]) * r["price"] for r in rs) / 1e6, sum(r["del_after"] * r["price"] for r in rs) / 1e6,
        sum(r["del_pend"] * r["price"] for r in rs) / 1e6,
        sum(1 for r in rs if r["del_open"] or sum(r["del_m"])), sum(1 for r in rs if r["lead_analog"])))
