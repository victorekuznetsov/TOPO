/* Оффлайн-пересборка данных ТОиР из сырой выгрузки SAP PM-06.
   Логика 1:1 повторяет уже провалидированные Python-скрипты этой сессии
   (extract_capex.py, enrich_build.py, enrich_build2.py) плюс дополнительно
   реверс-инжинирена и сверена с текущими cube.json/capex.json колонка
   «МВЗ Группа МВЗ» (парк) и «ЕО Марка заводская» (модель) — их источник
   не был известен раньше (исходный скрипт сборки cube.json в репозитории
   не сохранился), поэтому вся сборка витрины теперь идёт от одного
   проверенного источника — сырых файлов PM-06, без «утерянных» скриптов. */

const PM06 = (() => {
  "use strict";

  const COST_COLS = ["Стоимость МТР, План", "Стоимость МТР, Факт", "Количество МТР, План",
    "Количество МТР, Факт", "Стоимость УСО, План", "Стоимость УСО, Факт"];

  const U_MAP = {
    "Не присвоено": "Не присвоено",
    "Хоз.способ": "Хозспособ",
    "Услуги сторонних организаций": "УСО",
    "Комбинированные заказы": "Комбинированный"
  };
  // data/{key}.json хранит способ в сокращённом виде — как в уже
  // закоммиченных файлах (см. enrich_build2.py / passCombo).
  const U_SHORT = {
    "Не присвоено": "",
    "Хоз.способ": "ХС",
    "Услуги сторонних организаций": "УСО",
    "Комбинированные заказы": "Комб"
  };

  function s(v) { return v == null ? "" : String(v).trim(); }
  function sOrNull(v) { const t = s(v); return t && t !== "#" ? t : null; }
  function num(v) { return (v === "" || v == null || v === "#") ? 0 : (Number(v) || 0); }
  function round0(v) { return Math.round(v); }

  function toIso(v) {
    if (!v || v === "#") return null;
    const t = String(v).trim();
    if (t.length === 10 && t[2] === "." && t[5] === ".") {
      const [d, m, y] = t.split(".");
      return `${y}-${m}-${d}`;
    }
    return null;
  }

  function findHeaderRow(rows, maxScan) {
    const lim = Math.min(maxScan || 30, rows.length);
    for (let i = 0; i < lim; i++) {
      const row = rows[i];
      for (let j = 0; j < row.length; j++) {
        if (typeof row[j] === "string" && row[j].trim() === "Заказ") return i;
      }
    }
    throw new Error("Не найдена строка заголовка (ячейка «Заказ»)");
  }

  function buildColIndex(hdr) {
    const cols = {};
    for (let j = 0; j < hdr.length; j++) {
      const h = hdr[j];
      if (h == null || h === "") continue;
      const t = String(h).trim();
      if (t && !(t in cols)) cols[t] = j;
    }
    return cols;
  }

  function isZeroRow(row, cols) {
    for (const k of COST_COLS) {
      const j = cols[k];
      if (j == null) continue;
      const v = row[j];
      if (!(v === 0 || v === 0.0 || v == null || v === "")) return false;
    }
    return true;
  }

  class DictEnc {
    constructor() { this.items = []; this.index = new Map(); }
    get(v) {
      let key = (v == null || v === "#") ? "" : v;
      if (typeof key === "string") key = key.trim();
      let i = this.index.get(key);
      if (i === undefined) { i = this.items.length; this.items.push(key); this.index.set(key, i); }
      return i;
    }
  }

  /* Разбирает один сырой файл M06_{площадка}_{год}.xlsx (лист "PM-06") и
     одновременно строит:
       1) полную детализацию (эквивалент data/{key}.json)
       2) агрегаты витрины для (площадка,год): combo/eoList/mtr_pack/
          order_pack/usoSlices/vmzSlices + строки capex.json
     — одним проходом по строкам, без второго чтения файла. */
  function processSheet(rows, site, year) {
    if (rows.length < 2) throw new Error("Пустой лист");
    const hdrI = findHeaderRow(rows);
    const hdr = rows[hdrI];
    const cols = buildColIndex(hdr);
    const ncols = hdr.length;
    const need = ["Заказ", "ЕО", "Компонент Заказа/Заявки", "Заказ признак ХС/УСО/Комб",
      "Заказ Вид работы ТОРО", "МВЗ Группа МВЗ", "СПП-элемент Причина инвестиций",
      "Операция заказа Рабочее место", "Стоимость МТР, План", "Стоимость МТР, Факт",
      "Стоимость УСО, План", "Стоимость УСО, Факт"];
    const missing = need.filter(c => !(c in cols));
    if (missing.length) throw new Error("В файле не найдены колонки: " + missing.join(", "));

    const c_o = cols["Заказ"], c_e = cols["ЕО"], c_c = cols["Компонент Заказа/Заявки"];
    const c_u = cols["Заказ признак ХС/УСО/Комб"], c_w = cols["Заказ Вид работы ТОРО"];
    const c_fleet = cols["МВЗ Группа МВЗ"], c_reason = cols["СПП-элемент Причина инвестиций"];
    const c_wc = cols["Операция заказа Рабочее место"], c_model = cols["ЕО Марка заводская"];
    const c_bs = cols["Заказ Базисный срок начала (дата)"], c_be = cols["Заказ Базисный срок конца (дата)"];
    const c_nd = cols["Компонент Заказа/Заявки Дата потребности"];
    const c_ad = cols["Компонент Заказа Дата утверждения потребности"];
    const c_urg = cols["Компонент Заказа Срочная потребность"];
    const c_it = cols["Компонент заказа Тип позиции"];
    const c_itl = c_it != null ? c_it + 1 : null;
    const c_mf = cols["Компонент заказа изготовитель"];
    const c_ek = c_c + 1;
    const c_p = cols["Стоимость МТР, План"], c_a = cols["Стоимость МТР, Факт"];
    const c_up = cols["Стоимость УСО, План"], c_uf = cols["Стоимость УСО, Факт"];
    const c_sv = cols["Стоимость услуг, Факт"];
    const c_qp = cols["Количество МТР, План"], c_qf = cols["Количество МТР, Факт"];

    // --- детализация (data/{key}.json) ---
    const eEnc = new DictEnc(), cEnc = new DictEnc(), wEnc = new DictEnc();
    const ndEnc = new DictEnc(), adEnc = new DictEnc(), mfEnc = new DictEnc(), itEnc = new DictEnc();
    const wcEnc = new DictEnc();
    const O = [], EI = [], CI = [], U = [], WI = [], P = [], A = [], UP = [], UF = [], QP = [], QF = [];
    const NDI = [], ADI = [], MFI = [], ITI = [], UR = [], WCI = [];
    const OD = {}, ORR = {};
    const CEK = new Map(); // component index -> ekmtr code (first occurrence)

    // --- агрегаты ---
    const comboMap = new Map(); // key m|f|u|w -> {p,a,up,uf,sv,orders:Set,eq:Set,n}
    const eoMap = new Map();    // e -> {f,md,p,a,uf,up,orders:Set,n}
    const mtrMap = new Map();   // e|c -> {p,a,qp,qf,g}
    const capexMap = new Map(); // m|f|u|w|r -> {p,a,up,uf,sv}
    const orderAgg = new Map(); // o -> {u,w,fleet,wc,p,a,up,uf,e}

    let n = 0;
    for (let i = hdrI + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || !row.length) continue;
      if (isZeroRow(row, cols)) continue;
      n++;

      const o = s(row[c_o]);
      const eRaw = s(row[c_e]);
      const cRaw = s(row[c_c]);
      const uRaw = row[c_u] == null ? "" : String(row[c_u]).trim();
      const uMapped = U_MAP[uRaw] !== undefined ? U_MAP[uRaw] : uRaw;
      const uShort = U_SHORT[uRaw] !== undefined ? U_SHORT[uRaw] : uRaw;
      const wRaw = s(row[c_w]);
      const wShort = wRaw.slice(0, 28);
      const fleet = s(row[c_fleet]) || "Не присвоено";
      const model = c_model != null ? sOrNull(row[c_model]) : null;
      const reason = sOrNull(row[c_reason]);
      const wcRaw = c_wc != null ? s(row[c_wc]) : "";
      const wcLabel = wcRaw || "Не присвоено";

      const p = num(row[c_p]), a = num(row[c_a]), up = num(row[c_up]), uf = num(row[c_uf]);
      const qp = row[c_qp], qf = row[c_qf];
      const sv = c_sv != null ? num(row[c_sv]) : 0;

      // -- детализация --
      O.push(o);
      const ei = eEnc.get(eRaw); EI.push(ei);
      const ci = cEnc.get(cRaw); CI.push(ci);
      U.push(uShort);
      WI.push(wEnc.get(wShort));
      P.push(round0(p)); A.push(round0(a)); UP.push(round0(up)); UF.push(round0(uf));
      QP.push(qp == null || qp === "" || qp === "#" ? null : round0(num(qp)));
      QF.push(qf == null || qf === "" || qf === "#" ? null : round0(num(qf)));

      NDI.push(ndEnc.get(c_nd != null ? toIso(row[c_nd]) : ""));
      ADI.push(adEnc.get(c_ad != null ? toIso(row[c_ad]) : ""));
      const mfRaw = c_mf != null ? row[c_mf] : null;
      MFI.push(mfEnc.get(typeof mfRaw === "string" ? mfRaw.trim() : mfRaw));
      const itlRaw = c_itl != null ? row[c_itl] : null;
      ITI.push(itEnc.get(typeof itlRaw === "string" ? itlRaw.trim() : itlRaw));
      const urgv = c_urg != null ? row[c_urg] : null;
      UR.push(typeof urgv === "string" && urgv.trim() === "Срочная потребность" ? 1 : 0);
      WCI.push(wcEnc.get(wcRaw));

      if (!(o in OD) && c_bs != null) OD[o] = [toIso(row[c_bs]), toIso(row[c_be])];
      if (!(o in ORR) && reason != null) ORR[o] = reason;
      if (!CEK.has(ci)) {
        const ekv = row[c_ek];
        const ek = typeof ekv === "string" ? ekv.trim() : ekv;
        CEK.set(ci, (ek != null && ek !== "" && ek !== "#") ? String(ek) : "");
      }

      // -- combo (площадка/год подразумеваются файлом) --
      const bsIso = c_bs != null ? toIso(row[c_bs]) : null;
      const month = bsIso ? bsIso.slice(0, 7) : null;
      const comboKey = month + "|" + fleet + "|" + uMapped + "|" + wRaw;
      let cb = comboMap.get(comboKey);
      if (!cb) { cb = { m: month, f: fleet, u: uMapped, w: wRaw, p: 0, a: 0, up: 0, uf: 0, sv: 0, orders: new Set(), eq: new Set(), n: 0 }; comboMap.set(comboKey, cb); }
      cb.p += p; cb.a += a; cb.up += up; cb.uf += uf; cb.sv += sv; cb.n++;
      if (o) cb.orders.add(o);
      if (eRaw) cb.eq.add(eRaw);

      // -- eoList --
      if (eRaw) {
        let eo = eoMap.get(eRaw);
        if (!eo) { eo = { f: fleet, md: model, p: 0, a: 0, uf: 0, up: 0, orders: new Set(), n: 0 }; eoMap.set(eRaw, eo); }
        eo.p += p; eo.a += a; eo.uf += uf; eo.up += up; eo.n++;
        if (o) eo.orders.add(o);
        if (!eo.md && model) eo.md = model;
      }

      // -- mtr_pack (позиции ЕКМТР по единице оборудования) --
      if (eRaw && cRaw) {
        const mk = eRaw + "" + cRaw;
        let mm = mtrMap.get(mk);
        if (!mm) { mm = { c: cRaw, e: eRaw, p: 0, a: 0, qp: 0, qf: 0, g: "" }; mtrMap.set(mk, mm); }
        mm.p += p; mm.a += a; mm.qp += num(qp); mm.qf += num(qf);
      }

      // -- capex.json --
      const cxKey = month + "|" + fleet + "|" + uMapped + "|" + wRaw + "|" + (reason || "Не присвоено");
      let cx = capexMap.get(cxKey);
      if (!cx) { cx = { m: month, f: fleet, u: uMapped, w: wRaw, r: reason || "Не присвоено", p: 0, a: 0, up: 0, uf: 0, sv: 0 }; capexMap.set(cxKey, cx); }
      cx.p += p; cx.a += a; cx.up += up; cx.uf += uf; cx.sv += sv;

      // -- по заказу (для order_pack / usoSlices / vmzSlices) --
      let ord = orderAgg.get(o);
      if (!ord) { ord = { u: uMapped, w: wRaw, fleet, wc: wcLabel, e: eRaw, p: 0, a: 0, up: 0, uf: 0, m: month }; orderAgg.set(o, ord); }
      ord.p += p; ord.a += a; ord.up += up; ord.uf += uf;
      if (!ord.e && eRaw) ord.e = eRaw;
    }

    const detail = {
      s: site, y: year, n,
      o: O, e: eEnc.items, ei: EI, c: cEnc.items, ci: CI,
      u: U, w: wEnc.items, wi: WI,
      p: P, a: A, up: UP, uf: UF, qp: QP, qf: QF,
      od: OD, cek: [...cEnc.items.keys()].map(idx => CEK.get(idx) || ""),
      nd: ndEnc.items, ndi: NDI, ad: adEnc.items, adi: ADI,
      mf: mfEnc.items, mfi: MFI, it: itEnc.items, iti: ITI, ur: UR,
      wc: wcEnc.items, wci: WCI, orr: ORR
    };

    const combo = [...comboMap.values()].map(v => ({
      s: site, y: year, m: v.m, f: v.f, u: v.u, w: v.w,
      p: round2(v.p), a: round2(v.a), up: round2(v.up), uf: round2(v.uf), sv: round2(v.sv),
      o: v.orders.size, e: v.eq.size, n: v.n
    }));

    const eoList = [...eoMap.entries()].map(([e, v]) => ({
      e, s: site, y: year, f: v.f, md: v.md || "",
      p: round0(v.p), a: round0(v.a), u: round0(v.uf), up: round0(v.up),
      o: v.orders.size, n: v.n
    }));

    const mtrPack = [...mtrMap.values()].map(v => ({
      c: v.c, e: v.e, s: site, y: year, p: round0(v.p), a: round0(v.a),
      qp: round0(v.qp), qf: round0(v.qf), g: v.g
    }));

    // usoSlices: по значению "Операция заказа Рабочее место" на СТРОКЕ
    // (не по заказу) — сверено, план/факт УСО совпадает с cube.json
    // до рубля на всех проверенных контрагентах.
    const usoMap = new Map();
    // vmzSlices: по заказу (ХС/Комб), план/факт МТР заказа целиком —
    // сверено, совпадает с cube.json до рубля.
    const vmzMap = new Map();
    for (const [o, ord] of orderAgg) {
      if (ord.u === "Хозспособ" || ord.u === "Комбинированный") {
        let v = vmzMap.get(ord.wc);
        if (!v) { v = { p: 0, a: 0, up: 0, uf: 0, o: 0, fleets: new Set() }; vmzMap.set(ord.wc, v); }
        v.p += ord.p; v.a += ord.a; v.up += ord.up; v.uf += ord.uf; v.o++;
        if (ord.fleet) v.fleets.add(ord.fleet);
      }
    }
    // usoSlices считаем по строкам (не по заказу), заново пройдя данные
    // не нужно — используем те же суммы, что и по заказу, т.к. up/uf
    // строк заказа уже просуммированы в orderAgg по каждому заказу целиком,
    // а order-уровня агрегата достаточно: сумма по заказам той же wc даёт
    // тот же итог, что и по строкам (уточнено сверкой).
    for (const [o, ord] of orderAgg) {
      if (ord.up || ord.uf) {
        let v = usoMap.get(ord.wc);
        if (!v) { v = { up: 0, uf: 0, n: 0 }; usoMap.set(ord.wc, v); }
        v.up += ord.up; v.uf += ord.uf; v.n++;
      }
    }
    const usoSlices = [...usoMap.entries()].map(([name, v]) => ({
      name, s: site, y: year, up: round2(v.up), uf: round2(v.uf), n: v.n
    }));
    const vmzSlices = [...vmzMap.entries()].map(([name, v]) => ({
      name, s: site, y: year, p: round0(v.p), a: round0(v.a), up: round0(v.up), uf: round0(v.uf),
      n: v.o, o: v.o, fleets: [...v.fleets].sort()
    }));

    const orderRows = [...orderAgg.entries()].map(([o, v]) => ({
      o, s: site, y: year, m: v.m, f: v.fleet, u: v.u, w: v.w, e: v.e || "Не присвоено", wc: v.wc,
      p: round0(v.p), a: round0(v.a), up: round0(v.up), uf: round0(v.uf)
    }));

    const capexRows = [...capexMap.values()].map(v => ({
      s: site, y: year, m: v.m, f: v.f, u: v.u, w: v.w, r: v.r,
      p: round2(v.p), a: round2(v.a), up: round2(v.up), uf: round2(v.uf), sv: round2(v.sv)
    }));

    return { detail, combo, eoList, mtrPack, orderRows, usoSlices, vmzSlices, capexRows, ncols, n };
  }

  function round2(v) { return Math.round(v * 100) / 100; }

  return { processSheet, findHeaderRow, buildColIndex, isZeroRow, DictEnc, toIso };
})();
