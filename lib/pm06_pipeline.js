/* Пересборка данных ТОиР из сырой выгрузки SAP PM-06 — прямо в браузере.

   Логика извлечения полей 1:1 повторяет уже провалидированные Python-скрипты
   (extract_capex.py, enrich_build.py, enrich_build2.py), плюс реверс-
   инжинирены и сверены с текущими cube.json/capex.json колонки
   «МВЗ Группа МВЗ» (парк) и «ЕО Марка заводская» (модель) — исходный
   скрипт сборки cube.json в репозитории не сохранился, поэтому вся
   витрина теперь собирается из одного проверяемого источника — сырых
   файлов PM-06.

   Строки подаются по одной (createProcessor), а не массивом: у выгрузки
   площадки 1100 около 320 тыс. строк по 60 колонок — 19 млн ячеек
   одновременно в памяти не держим. */

const PM06 = (() => {
  "use strict";

  const COST_COLS = ["Стоимость МТР, План", "Стоимость МТР, Факт", "Количество МТР, План",
    "Количество МТР, Факт", "Стоимость УСО, План", "Стоимость УСО, Факт"];

  // Способ исполнения хранится в трёх видах: сырой текст выгрузки,
  // подпись витрины (cube.json/capex.json) и сокращение в data/*.json.
  const U_MAP = {
    "Не присвоено": "Не присвоено",
    "Хоз.способ": "Хозспособ",
    "Услуги сторонних организаций": "УСО",
    "Комбинированные заказы": "Комбинированный"
  };
  const U_SHORT = {
    "Не присвоено": "",
    "Хоз.способ": "ХС",
    "Услуги сторонних организаций": "УСО",
    "Комбинированные заказы": "Комб"
  };

  // Размер выборок в cube.json: сколько позиций ЕКМТР и заказов
  // оставлять на единицу оборудования (полные данные — в data/*.json).
  const MTR_PER_EO = 2;
  const ORD_PER_EO = 3;

  const REQUIRED = ["Заказ", "ЕО", "Компонент Заказа/Заявки", "Заказ признак ХС/УСО/Комб",
    "Заказ Вид работы ТОРО", "МВЗ Группа МВЗ", "СПП-элемент Причина инвестиций",
    "Операция заказа Рабочее место", "Стоимость МТР, План", "Стоимость МТР, Факт",
    "Стоимость УСО, План", "Стоимость УСО, Факт"];

  function s(v) { return v == null ? "" : String(v).trim(); }
  function sOrNull(v) { const t = s(v); return t && t !== "#" ? t : null; }
  function num(v) { return (v === "" || v == null || v === "#") ? 0 : (Number(v) || 0); }

  /* Округление «половина к чётному», как в Python round() — именно им
     собраны уже опубликованные data/*.json. Math.round округляет
     половину вверх и на 117 тыс. строк даёт расхождение в сотни рублей
     против действующих файлов. */
  function r0(v) {
    const f = Math.floor(v), d = v - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }
  function r2(v) { return Math.round(v * 100) / 100; }

  /* Словари в data/*.json отсортированы — так их строил исходный
     сборщик. Пересортировка на выходе делает результат ещё и
     детерминированным: повторный прогон того же файла даёт побайтово
     тот же JSON, и diff показывает только реальные изменения данных. */
  function sortDict(items, idxArrays) {
    const order = items.map((v, i) => i).sort((a, b) => (items[a] < items[b] ? -1 : items[a] > items[b] ? 1 : 0));
    const remap = new Array(items.length);
    order.forEach((oldI, newI) => { remap[oldI] = newI; });
    const sorted = order.map(i => items[i]);
    idxArrays.forEach(arr => { for (let i = 0; i < arr.length; i++) arr[i] = remap[arr[i]]; });
    return { sorted, remap };
  }

  function toIso(v) {
    if (!v || v === "#") return null;
    const t = String(v).trim();
    if (t.length === 10 && t[2] === "." && t[5] === ".") {
      const p = t.split(".");
      return p[2] + "-" + p[1] + "-" + p[0];
    }
    return null;
  }

  class DictEnc {
    constructor() { this.items = []; this.index = new Map(); }
    get(v) {
      let k = (v == null || v === "#") ? "" : v;
      if (typeof k === "string") k = k.trim();
      let i = this.index.get(k);
      if (i === undefined) { i = this.items.length; this.items.push(k); this.index.set(k, i); }
      return i;
    }
  }

  function isHeaderRow(cells) {
    for (let i = 0; i < cells.length; i++) {
      if (typeof cells[i] === "string" && cells[i].trim() === "Заказ") return true;
    }
    return false;
  }

  function buildColIndex(cells) {
    const cols = {};
    for (let j = 0; j < cells.length; j++) {
      const h = cells[j];
      if (h == null || h === "") continue;
      const t = String(h).trim();
      if (t && !(t in cols)) cols[t] = j;
    }
    return cols;
  }

  /* Обработчик одного файла: принимает строки по одной, на выходе —
     полная детализация data/{площадка}_{год}.json и агрегаты витрины. */
  function createProcessor(site, year) {
    let cols = null, headerSeen = false, scanned = 0;
    let c = {};                       // индексы нужных колонок

    const eEnc = new DictEnc(), cEnc = new DictEnc(), wEnc = new DictEnc();
    const ndEnc = new DictEnc(), adEnc = new DictEnc(), mfEnc = new DictEnc();
    const itEnc = new DictEnc(), wcEnc = new DictEnc();
    const O = [], EI = [], CI = [], U = [], WI = [], P = [], A = [], UP = [], UF = [], QP = [], QF = [];
    const NDI = [], ADI = [], MFI = [], ITI = [], UR = [], WCI = [];
    const OD = {}, ORR = {}, CEK = [];

    const comboMap = new Map(), eoMap = new Map(), mtrMap = new Map();
    const capexMap = new Map(), orderAgg = new Map(), usoMap = new Map();
    let n = 0;

    function setHeader(cells) {
      cols = buildColIndex(cells);
      const missing = REQUIRED.filter(k => !(k in cols));
      if (missing.length) throw new Error("В выгрузке нет колонок: " + missing.join(", "));
      c = {
        o: cols["Заказ"], e: cols["ЕО"], c: cols["Компонент Заказа/Заявки"],
        u: cols["Заказ признак ХС/УСО/Комб"], w: cols["Заказ Вид работы ТОРО"],
        fleet: cols["МВЗ Группа МВЗ"], reason: cols["СПП-элемент Причина инвестиций"],
        wc: cols["Операция заказа Рабочее место"], model: cols["ЕО Марка заводская"],
        bs: cols["Заказ Базисный срок начала (дата)"], be: cols["Заказ Базисный срок конца (дата)"],
        nd: cols["Компонент Заказа/Заявки Дата потребности"],
        ad: cols["Компонент Заказа Дата утверждения потребности"],
        urg: cols["Компонент Заказа Срочная потребность"],
        it: cols["Компонент заказа Тип позиции"],
        mf: cols["Компонент заказа изготовитель"],
        grp: cols["Компонент Заказа Группа материалов"],
        p: cols["Стоимость МТР, План"], a: cols["Стоимость МТР, Факт"],
        up: cols["Стоимость УСО, План"], uf: cols["Стоимость УСО, Факт"],
        sv: cols["Стоимость услуг, Факт"],
        qp: cols["Количество МТР, План"], qf: cols["Количество МТР, Факт"]
      };
      // № ЕКМТР и текст типа позиции лежат в безымянной колонке сразу
      // за соответствующей именованной — так в выгрузке PM-06.
      c.ek = c.c + 1;
      c.itl = c.it != null ? c.it + 1 : null;
      headerSeen = true;
    }

    function isZeroRow(cells) {
      for (let i = 0; i < COST_COLS.length; i++) {
        const j = cols[COST_COLS[i]];
        if (j == null) continue;
        const v = cells[j];
        if (!(v === 0 || v == null || v === "")) return false;
      }
      return true;
    }

    function pushRow(cells) {
      if (!headerSeen) {
        if (++scanned > 40) throw new Error("Не найдена строка заголовка с колонкой «Заказ» — это не выгрузка PM-06");
        if (isHeaderRow(cells)) setHeader(cells);
        return;
      }
      if (!cells || !cells.length) return;

      /* Витрина оборудования считает и строки с нулевой стоимостью:
         иначе из парка пропадает техника, у которой в этом году есть
         заказы, но нет затрат (в действующем cube.json таких единиц
         6 только по площадке 1400 за 2026). Всё остальное — включая
         полную детализацию — строится по ненулевым строкам. */
      const zeroRow = isZeroRow(cells);
      if (zeroRow) {
        const e0 = s(cells[c.e]);
        const o0 = s(cells[c.o]);
        const f0 = s(cells[c.fleet]) || "Не присвоено";
        if (e0) {
          let eo0 = eoMap.get(e0);
          if (!eo0) {
            eo0 = { f: f0, md: c.model != null ? sOrNull(cells[c.model]) : null, p: 0, a: 0, uf: 0, up: 0, orders: new Set(), n: 0 };
            eoMap.set(e0, eo0);
          }
          eo0.n++;
          if (o0) eo0.orders.add(o0);
        }
        // Витрина затрат тоже считает нулевые строки: без них пропадают
        // целые срезы вроде ежесменного ТО, где заказы есть, а затрат нет
        // (20 таких групп только по площадке 1100 за 2026).
        const uR = cells[c.u] == null ? "" : String(cells[c.u]).trim();
        const b0 = c.bs != null ? toIso(cells[c.bs]) : null;
        const k0 = (b0 ? b0.slice(0, 7) : null) + "|" + f0 + "|" + (U_MAP[uR] !== undefined ? U_MAP[uR] : uR) + "|" + s(cells[c.w]);
        let cb0 = comboMap.get(k0);
        if (!cb0) {
          cb0 = { m: b0 ? b0.slice(0, 7) : null, f: f0, u: U_MAP[uR] !== undefined ? U_MAP[uR] : uR, w: s(cells[c.w]), p: 0, a: 0, up: 0, uf: 0, sv: 0, orders: new Set(), eq: new Set(), n: 0 };
          comboMap.set(k0, cb0);
        }
        cb0.n++;
        if (o0) cb0.orders.add(o0);
        if (e0) cb0.eq.add(e0);
        return;
      }
      n++;

      const o = s(cells[c.o]);
      const eRaw = s(cells[c.e]);
      const cRaw = s(cells[c.c]);
      const uRaw = cells[c.u] == null ? "" : String(cells[c.u]).trim();
      const uMapped = U_MAP[uRaw] !== undefined ? U_MAP[uRaw] : uRaw;
      const uShort = U_SHORT[uRaw] !== undefined ? U_SHORT[uRaw] : uRaw;
      const wRaw = s(cells[c.w]);
      const fleet = s(cells[c.fleet]) || "Не присвоено";
      const model = c.model != null ? sOrNull(cells[c.model]) : null;
      const reason = sOrNull(cells[c.reason]);
      const wcRaw = c.wc != null ? s(cells[c.wc]) : "";
      const wcLabel = wcRaw || "Не присвоено";

      const p = num(cells[c.p]), a = num(cells[c.a]);
      const up = num(cells[c.up]), uf = num(cells[c.uf]);
      const sv = c.sv != null ? num(cells[c.sv]) : 0;
      const qpRaw = cells[c.qp], qfRaw = cells[c.qf];

      // --- полная детализация ---
      O.push(o);
      const ei = eEnc.get(eRaw); EI.push(ei);
      const ci = cEnc.get(cRaw); CI.push(ci);
      U.push(uShort);
      WI.push(wEnc.get(wRaw.slice(0, 28)));
      P.push(r0(p)); A.push(r0(a)); UP.push(r0(up)); UF.push(r0(uf));
      QP.push(qpRaw == null || qpRaw === "" || qpRaw === "#" ? null : r0(num(qpRaw)));
      QF.push(qfRaw == null || qfRaw === "" || qfRaw === "#" ? null : r0(num(qfRaw)));
      NDI.push(ndEnc.get(c.nd != null ? toIso(cells[c.nd]) : ""));
      ADI.push(adEnc.get(c.ad != null ? toIso(cells[c.ad]) : ""));
      MFI.push(mfEnc.get(c.mf != null ? cells[c.mf] : null));
      ITI.push(itEnc.get(c.itl != null ? cells[c.itl] : null));
      const urg = c.urg != null ? cells[c.urg] : null;
      UR.push(typeof urg === "string" && urg.trim() === "Срочная потребность" ? 1 : 0);
      WCI.push(wcEnc.get(wcRaw));

      if (!(o in OD) && c.bs != null) OD[o] = [toIso(cells[c.bs]), toIso(cells[c.be])];
      if (!(o in ORR) && reason != null) ORR[o] = reason;
      if (CEK[ci] === undefined) {
        const ekv = cells[c.ek];
        const ek = typeof ekv === "string" ? ekv.trim() : ekv;
        CEK[ci] = (ek != null && ek !== "" && ek !== "#") ? String(ek) : "";
      }

      // --- агрегаты витрины ---
      const bsIso = c.bs != null ? toIso(cells[c.bs]) : null;
      const month = bsIso ? bsIso.slice(0, 7) : null;

      const ck = month + "|" + fleet + "|" + uMapped + "|" + wRaw;
      let cb = comboMap.get(ck);
      if (!cb) { cb = { m: month, f: fleet, u: uMapped, w: wRaw, p: 0, a: 0, up: 0, uf: 0, sv: 0, orders: new Set(), eq: new Set(), n: 0 }; comboMap.set(ck, cb); }
      cb.p += p; cb.a += a; cb.up += up; cb.uf += uf; cb.sv += sv; cb.n++;
      if (o) cb.orders.add(o);
      if (eRaw) cb.eq.add(eRaw);

      if (eRaw) {
        let eo = eoMap.get(eRaw);
        if (!eo) { eo = { f: fleet, md: model, p: 0, a: 0, uf: 0, up: 0, orders: new Set(), n: 0 }; eoMap.set(eRaw, eo); }
        eo.p += p; eo.a += a; eo.uf += uf; eo.up += up; eo.n++;
        if (o) eo.orders.add(o);
        if (!eo.md && model) eo.md = model;

        if (cRaw) {
          const mk = eRaw + " " + cRaw;
          let mm = mtrMap.get(mk);
          if (!mm) { mm = { c: cRaw, e: eRaw, p: 0, a: 0, qp: 0, qf: 0, g: "" }; mtrMap.set(mk, mm); }
          mm.p += p; mm.a += a; mm.qp += num(qpRaw); mm.qf += num(qfRaw);
          if (!mm.g && c.grp != null) { const g = sOrNull(cells[c.grp]); if (g) mm.g = g; }
        }
      }

      const xk = ck + "|" + (reason || "Не присвоено");
      let cx = capexMap.get(xk);
      if (!cx) { cx = { m: month, f: fleet, u: uMapped, w: wRaw, r: reason || "Не присвоено", p: 0, a: 0, up: 0, uf: 0, sv: 0 }; capexMap.set(xk, cx); }
      cx.p += p; cx.a += a; cx.up += up; cx.uf += uf; cx.sv += sv;

      let ord = orderAgg.get(o);
      if (!ord) { ord = { u: uMapped, w: wRaw, f: fleet, wc: wcLabel, e: eRaw, m: month, p: 0, a: 0, up: 0, uf: 0 }; orderAgg.set(o, ord); }
      ord.p += p; ord.a += a; ord.up += up; ord.uf += uf;
      if (!ord.e && eRaw) ord.e = eRaw;

      /* Контрагенты УСО считаются по СТРОКЕ, а не по заказу: «Операция
         заказа Рабочее место» меняется внутри одного заказа, и услуга
         должна попасть тому рабочему месту, которое её выполнило.
         По заказу целиком итог тот же, но деньги уходят не тому имени. */
      if (up || uf) {
        let uv = usoMap.get(wcLabel);
        if (!uv) { uv = { up: 0, uf: 0, n: 0 }; usoMap.set(wcLabel, uv); }
        uv.up += up; uv.uf += uf; uv.n++;
      }
    }

    function finish() {
      if (!headerSeen) throw new Error("Пустой лист: строка заголовка не найдена");

      const cekRaw = new Array(cEnc.items.length);
      for (let i = 0; i < cekRaw.length; i++) cekRaw[i] = CEK[i] || "";

      const eS = sortDict(eEnc.items, [EI]);
      const cS = sortDict(cEnc.items, [CI]);
      const wS = sortDict(wEnc.items, [WI]);
      const ndS = sortDict(ndEnc.items, [NDI]);
      const adS = sortDict(adEnc.items, [ADI]);
      const mfS = sortDict(mfEnc.items, [MFI]);
      const itS = sortDict(itEnc.items, [ITI]);
      const wcS = sortDict(wcEnc.items, [WCI]);
      const cek = new Array(cekRaw.length);
      cS.remap.forEach((newI, oldI) => { cek[newI] = cekRaw[oldI]; });

      const detail = {
        s: site, y: year, n,
        o: O, e: eS.sorted, ei: EI, c: cS.sorted, ci: CI,
        u: U, w: wS.sorted, wi: WI,
        p: P, a: A, up: UP, uf: UF, qp: QP, qf: QF,
        od: OD, cek,
        nd: ndS.sorted, ndi: NDI, ad: adS.sorted, adi: ADI,
        mf: mfS.sorted, mfi: MFI, it: itS.sorted, iti: ITI, ur: UR,
        wc: wcS.sorted, wci: WCI, orr: ORR
      };

      const combo = [];
      comboMap.forEach(v => combo.push({
        s: site, y: year, m: v.m, f: v.f, u: v.u, w: v.w,
        p: r2(v.p), a: r2(v.a), up: r2(v.up), uf: r2(v.uf), sv: r2(v.sv),
        o: v.orders.size, e: v.eq.size, n: v.n
      }));

      const eoList = [];
      eoMap.forEach((v, e) => eoList.push({
        e, s: site, y: year, f: v.f, md: v.md || "",
        p: r0(v.p), a: r0(v.a), u: r0(v.uf), up: r0(v.up), o: v.orders.size, n: v.n
      }));

      /* mtr_pack и order_pack — не полные витрины, а выборки: cube.json
         грузится при каждом открытии страницы, поэтому исходный сборщик
         оставлял в них только крупнейшее (полные данные лежат в
         data/{площадка}_{год}.json и подгружаются по требованию).
         Правило восстановлено сверкой с действующим cube.json:
         mtr_pack — две самые дорогие позиции ЕКМТР на единицу
         оборудования (совпало 1313 из 1313 по площадке 1400 за 2026).
         Покрытие дашборд считает сам и показывает честно. */
      const mtrByEo = new Map();
      mtrMap.forEach(v => {
        if (!(v.p + v.a)) return;
        let lst = mtrByEo.get(v.e);
        if (!lst) { lst = []; mtrByEo.set(v.e, lst); }
        lst.push(v);
      });
      const mtrPack = [];
      mtrByEo.forEach(lst => {
        lst.sort((x, y) => (y.p + y.a) - (x.p + x.a));
        for (let i = 0; i < Math.min(MTR_PER_EO, lst.length); i++) {
          const v = lst[i];
          mtrPack.push({ c: v.c, e: v.e, s: site, y: year, p: r0(v.p), a: r0(v.a), qp: r0(v.qp), qf: r0(v.qf), g: v.g });
        }
      });

      // usoSlices — план/факт УСО по «Операция заказа Рабочее место».
      // vmzSlices — план/факт МТР по нему же, по всем заказам, КРОМЕ УСО:
      // у хозспособа, комбинированных и заказов без назначенного способа
      // в этом поле стоит внутреннее подразделение. Сверено с cube.json —
      // совпадение до рубля и по всем 49 именам (площадка 1400, 2026).
      const vmzMap = new Map();
      orderAgg.forEach(ord => {
        if (ord.u !== "УСО") {
          let v = vmzMap.get(ord.wc);
          if (!v) { v = { p: 0, a: 0, up: 0, uf: 0, o: 0, fleets: new Set() }; vmzMap.set(ord.wc, v); }
          v.p += ord.p; v.a += ord.a; v.up += ord.up; v.uf += ord.uf; v.o++;
          if (ord.f) v.fleets.add(ord.f);
        }
      });
      const usoSlices = [];
      usoMap.forEach((v, name) => usoSlices.push({ name, s: site, y: year, up: r2(v.up), uf: r2(v.uf), n: v.n }));
      const vmzSlices = [];
      vmzMap.forEach((v, name) => vmzSlices.push({
        name, s: site, y: year, p: r0(v.p), a: r0(v.a), up: r0(v.up), uf: r0(v.uf),
        n: v.o, o: v.o, fleets: [...v.fleets].sort()
      }));

      /* order_pack — тоже выборка (крупнейшие заказы на единицу
         оборудования). Полный список заказов живёт в детализации;
         здесь важно удержать размер cube.json. Дополнительно
         гарантируем, что в выборку попадёт хотя бы один заказ каждого
         подрядчика — иначе он исчезнет из списка фильтра «Подрядчик». */
      const ordByEo = new Map();
      orderAgg.forEach((v, o) => {
        if (!(v.p + v.a + v.up + v.uf)) return;
        const k = v.e || "Не присвоено";
        let lst = ordByEo.get(k);
        if (!lst) { lst = []; ordByEo.set(k, lst); }
        lst.push([o, v]);
      });
      const picked = new Map();
      ordByEo.forEach(lst => {
        lst.sort((x, y) => (y[1].p + y[1].a + y[1].up + y[1].uf) - (x[1].p + x[1].a + x[1].up + x[1].uf));
        for (let i = 0; i < Math.min(ORD_PER_EO, lst.length); i++) picked.set(lst[i][0], lst[i][1]);
      });
      const seenWc = new Set();
      picked.forEach(v => seenWc.add(v.wc));
      const byWc = new Map();
      orderAgg.forEach((v, o) => {
        if (seenWc.has(v.wc) || picked.has(o)) return;
        const cur = byWc.get(v.wc);
        if (!cur || (v.p + v.a + v.up + v.uf) > (cur[1].p + cur[1].a + cur[1].up + cur[1].uf)) byWc.set(v.wc, [o, v]);
      });
      byWc.forEach(([o, v]) => picked.set(o, v));

      const orderRows = [];
      picked.forEach((v, o) => orderRows.push({
        o, s: site, y: year, m: v.m, f: v.f, u: v.u, w: v.w, e: v.e || "Не присвоено", wc: v.wc,
        p: r0(v.p), a: r0(v.a), up: r0(v.up), uf: r0(v.uf)
      }));

      // capex.json собран без разреза по базис-месяцу — вкладке
      // CAPEX/OPEX месяц не нужен, а файл без него в 27 раз меньше.
      const cxAgg = new Map();
      capexMap.forEach(v => {
        const k = v.f + "|" + v.u + "|" + v.w + "|" + v.r;
        let x = cxAgg.get(k);
        if (!x) { x = { f: v.f, u: v.u, w: v.w, r: v.r, p: 0, a: 0, up: 0, uf: 0, sv: 0 }; cxAgg.set(k, x); }
        x.p += v.p; x.a += v.a; x.up += v.up; x.uf += v.uf; x.sv += v.sv;
      });
      const capexRows = [];
      cxAgg.forEach(v => {
        if (!(v.p || v.a || v.up || v.uf || v.sv)) return;   // пустые срезы не храним
        capexRows.push({
          s: site, y: year, f: v.f, u: v.u, w: v.w, r: v.r,
          p: r2(v.p), a: r2(v.a), up: r2(v.up), uf: r2(v.uf), sv: r2(v.sv)
        });
      });

      return { detail, combo, eoList, mtrPack, orderRows, usoSlices, vmzSlices, capexRows, n };
    }

    return { pushRow, finish, get rows() { return n; } };
  }

  return { createProcessor, U_MAP, U_SHORT, toIso, DictEnc };
})();

if (typeof module !== "undefined" && module.exports) module.exports = PM06;
