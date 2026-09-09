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

  /* Размер выборок в cube.json: полные данные лежат в data/*.json и
     подгружаются по требованию, а cube.json читается при каждом открытии
     страницы, поэтому в нём хранится только крупнейшее. Правила сверены
     с действующей витриной: позиции ЕКМТР — самые дорогие на единицу
     оборудования, не дешевле 20 тыс. ₽. В опубликованных данных глубина
     разная: по 2 позиции на единицу, а за 2027 год — по 5 (его когда-то
     собирали подробнее под отдельный разбор плана). Здесь единая
     глубина 5 для всех лет: пересборка 2027-го иначе обеднила бы уже
     опубликованные данные, а годы остались бы несопоставимы между собой.
     Покрытие дашборд считает сам и показывает честно. */
  const MTR_PER_EO = 5;
  const MTR_MIN_COST = 20000;
  const ORD_PER_EO = 3;
  // Позиции МТР по срезам — глобально крупнейшие по всей витрине.
  const MAT_LIMIT = 6000;

  /* Справочные позиции (тип «S») — информационные строки заказа, у них
     нет ни количества, ни собственной стоимости: 141 496 строк на
     482 тыс ₽ плана из 127,9 млрд. В составах заказов и списках ЕКМТР
     они создавали пустые строки. Исключаем их, но только пока по ним
     нет факта: те немногие, где освоение реально прошло, остаются. */
  const REF_ITEM_CODE = "S";
  /* Причины инвестиций, относящиеся к CAPEX, — тот же список, что и в
     дашборде: витрина материалов хранит признак прямо в ключе ячейки,
     чтобы фильтр CAPEX/OPEX на вкладке «Материалы» работал без
     подгрузки полной детализации. */
  const CAPEX_REASONS = new Set(["Капитализируемые ремонты", "Компоненты"]);

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

    const matUse = new Map();          // ЕКМТР + признак CAPEX → расход
    const matName = new Map();         // ЕКМТР → наименование из выгрузки
    const eEnc = new DictEnc(), cEnc = new DictEnc(), wEnc = new DictEnc();
    const ndEnc = new DictEnc(), adEnc = new DictEnc(), mfEnc = new DictEnc();
    const itEnc = new DictEnc(), wcEnc = new DictEnc();
    const O = [], EI = [], CI = [], U = [], WI = [], P = [], A = [], UP = [], UF = [], QP = [], QF = [];
    const NDI = [], ADI = [], MFI = [], ITI = [], UR = [], WCI = [];
    const OD = {}, ORR = {}, CEK = [];

    const comboMap = new Map(), eoMap = new Map(), mtrMap = new Map();
    const capexMap = new Map(), orderAgg = new Map(), usoMap = new Map();
    // Все номера заказов, включая встречающиеся только в нулевых строках:
    // именно так считается «заказов в источнике» для раскрытия покрытия.
    const allOrders = new Set(), orderMonth = new Map(), oemPairs = new Map();
    const matMap = new Map();
    let n = 0, rawRows = 0, refSkipped = 0;

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
        oem: cols["Номер детали производителя"],
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
      rawRows++;

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
      const bsIso = c.bs != null ? toIso(cells[c.bs]) : null;

      if (c.it != null && s(cells[c.it]) === REF_ITEM_CODE
        && !num(cells[c.a]) && !num(cells[c.uf]) && !(c.sv != null && num(cells[c.sv]))) {
        refSkipped++;
        return;
      }

      const zeroRow = isZeroRow(cells);

      /* Парк, оборудование и состав заказов собираем по ВСЕМ строкам,
         включая нулевые: иначе из витрин пропадает техника с заказами
         без затрат, целые срезы вроде ежесменного ТО и подразделения,
         у которых в этом году не было списаний. Деньги и полная
         детализация — только по ненулевым строкам. */
      if (o) {
        allOrders.add(o);
        // Базисный срок — реквизит заказа, а не строки: в части строк он
        // пуст, и если брать месяц построчно, деньги уезжают в срез «без
        // месяца» вместо месяца своего заказа. Если даты начала нет вовсе
        // (в 2025 году по площадке 1100 таких заказов семь), заказ
        // относится к январю своего года — так собрана действующая
        // витрина, в ней нет ни одной строки без месяца.
        if (bsIso && !orderMonth.has(o)) orderMonth.set(o, bsIso.slice(0, 7));
      }

      if (eRaw) {
        let eo = eoMap.get(eRaw);
        if (!eo) { eo = { f: fleet, md: model, p: 0, a: 0, uf: 0, up: 0, orders: new Set(), n: 0 }; eoMap.set(eRaw, eo); }
        eo.n++;
        if (o) eo.orders.add(o);
        if (!eo.md && model) eo.md = model;
      }

      const ck = o + "\u0000" + fleet + "\u0000" + uMapped + "\u0000" + wRaw;
      let cb = comboMap.get(ck);
      if (!cb) { cb = { o, f: fleet, u: uMapped, w: wRaw, p: 0, a: 0, up: 0, uf: 0, sv: 0, eq: new Set(), n: 0 }; comboMap.set(ck, cb); }
      cb.n++;
      if (eRaw) cb.eq.add(eRaw);

      let ord = orderAgg.get(o);
      if (!ord) { ord = { u: uMapped, w: wRaw, f: fleet, wc: wcLabel, e: eRaw, p: 0, a: 0, up: 0, uf: 0, nz: !zeroRow }; orderAgg.set(o, ord); }
      else if (!ord.nz && !zeroRow) { ord.u = uMapped; ord.w = wRaw; ord.f = fleet; ord.wc = wcLabel; ord.e = eRaw; ord.nz = true; }
      if (!ord.e && eRaw) ord.e = eRaw;

      if (zeroRow) return;
      n++;

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
      if (c.oem != null && cRaw && !oemPairs.has(cRaw)) {
        const pn = sOrNull(cells[c.oem]);
        if (pn) oemPairs.set(cRaw, pn);
      }
      if (CEK[ci] === undefined) {
        const ekv = cells[c.ek];
        const ek = typeof ekv === "string" ? ekv.trim() : ekv;
        CEK[ci] = (ek != null && ek !== "" && ek !== "#") ? String(ek) : "";
      }

      /* Витрина использования материалов (mtr.json): ключ — код ЕКМТР и
         признак CAPEX/OPEX, площадка и год у файла одни. Складываем те
         же округлённые величины, что легли в полную детализацию, — иначе
         итоги вкладки «Материалы» разошлись бы с составами заказов. */
      const cxKey = reason == null || reason === "Не присвоено" ? 2 : (CAPEX_REASONS.has(reason) ? 1 : 0);
      const ekCode = CEK[ci];
      const mk = ekCode + "\u0000" + cxKey;
      let mu = matUse.get(mk);
      if (!mu) { mu = { k: ekCode, cx: cxKey, qp: 0, qf: 0, p: 0, a: 0, n: 0, ord: new Set(), eo: new Set() }; matUse.set(mk, mu); }
      mu.qp += QP[QP.length - 1] || 0; mu.qf += QF[QF.length - 1] || 0;
      mu.p += P[P.length - 1]; mu.a += A[A.length - 1]; mu.n++;
      if (o) mu.ord.add(o);
      mu.eo.add(ei);
      if (ekCode && !matName.has(ekCode)) matName.set(ekCode, cRaw);

      // --- деньги в уже созданные срезы ---
      cb.p += p; cb.a += a; cb.up += up; cb.uf += uf; cb.sv += sv;

      if (eRaw) {
        const eo = eoMap.get(eRaw);
        eo.p += p; eo.a += a; eo.uf += uf; eo.up += up;

        if (cRaw) {
          const mk = eRaw + "\u0000" + cRaw;
          let mm = mtrMap.get(mk);
          if (!mm) { mm = { c: cRaw, e: eRaw, p: 0, a: 0, qp: 0, qf: 0, g: "" }; mtrMap.set(mk, mm); }
          mm.p += p; mm.a += a; mm.qp += num(qpRaw); mm.qf += num(qfRaw);
          if (!mm.g && c.grp != null) { const g = sOrNull(cells[c.grp]); if (g) mm.g = g; }
        }
      }

      if (cRaw) {
        let mt = matMap.get(cRaw);
        if (!mt) { mt = { p: 0, a: 0, n: 0 }; matMap.set(cRaw, mt); }
        mt.p += p; mt.a += a; mt.n++;
      }

      const xk = fleet + " " + uMapped + " " + wRaw + " " + (reason || "Не присвоено");
      let cx = capexMap.get(xk);
      if (!cx) { cx = { f: fleet, u: uMapped, w: wRaw, r: reason || "Не присвоено", p: 0, a: 0, up: 0, uf: 0, sv: 0 }; capexMap.set(xk, cx); }
      cx.p += p; cx.a += a; cx.up += up; cx.uf += uf; cx.sv += sv;

      ord.p += p; ord.a += a; ord.up += up; ord.uf += uf;

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

      /* Сворачиваем срезы затрат: копили по заказу, теперь проставляем
         месяц заказа и объединяем в (месяц × парк × способ × вид работ). */
      const monthOf = o => orderMonth.get(o) || (year + "-01");
      const comboFinal = new Map();
      comboMap.forEach(v => {
        const m = monthOf(v.o);
        const k = m + " " + v.f + " " + v.u + " " + v.w;
        let g = comboFinal.get(k);
        if (!g) { g = { m, f: v.f, u: v.u, w: v.w, p: 0, a: 0, up: 0, uf: 0, sv: 0, orders: new Set(), eq: new Set(), n: 0 }; comboFinal.set(k, g); }
        g.p += v.p; g.a += v.a; g.up += v.up; g.uf += v.uf; g.sv += v.sv; g.n += v.n;
        if (v.o) g.orders.add(v.o);
        v.eq.forEach(e => g.eq.add(e));
      });
      const combo = [];
      comboFinal.forEach(v => combo.push({
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
        if (v.p + v.a < MTR_MIN_COST) return;
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
        o, s: site, y: year, m: monthOf(o), f: v.f, u: v.u, w: v.w, e: v.e || "Не присвоено", wc: v.wc,
        p: r0(v.p), a: r0(v.a), up: r0(v.up), uf: r0(v.uf)
      }));

      // capex.json собран без разреза по базис-месяцу — вкладке
      // CAPEX/OPEX месяц не нужен, а файл без него в 27 раз меньше.
      const matSlices = [];
      matMap.forEach((v, name) => {
        if (!(v.p || v.a)) return;
        matSlices.push({ name, s: site, y: year, p: r2(v.p), a: r2(v.a), u: 0, n: v.n });
      });

      const capexRows = [];
      capexMap.forEach(v => {
        if (!(v.p || v.a || v.up || v.uf || v.sv)) return;   // пустые срезы не храним
        capexRows.push({
          s: site, y: year, f: v.f, u: v.u, w: v.w, r: v.r,
          p: r2(v.p), a: r2(v.a), up: r2(v.up), uf: r2(v.uf), sv: r2(v.sv)
        });
      });

      const matRows = [];
      matUse.forEach(v => matRows.push({
        k: v.k, cx: v.cx, qp: Math.round(v.qp * 1000) / 1000, qf: Math.round(v.qf * 1000) / 1000,
        p: Math.round(v.p), a: Math.round(v.a), n: v.n, o: v.ord.size, e: v.eo.size
      }));

      return {
        detail, combo, eoList, mtrPack, matSlices, orderRows, usoSlices, vmzSlices, capexRows, n,
        oemPairs: [...oemPairs],
        matRows, matNames: [...matName],
        meta: { site, year, rows: rawRows, orders: allOrders.size, refSkipped }
      };
    }

    return { pushRow, finish, get rows() { return n; } };
  }

  /* Слияние пересобранных площадко-лет с действующей витриной: строки
     затронутых площадка-год заменяются целиком, всё остальное остаётся
     нетронутым. Так обновление одной выгрузки не трогает соседние годы. */
  function mergePack(pack, cols, newRows, touched) {
    const out = {};
    cols.forEach(c => { out[c] = []; });
    const n = pack && pack[cols[0]] ? pack[cols[0]].length : 0;
    for (let i = 0; i < n; i++) {
      if (touched.has(pack.s[i] + "|" + pack.y[i])) continue;
      cols.forEach(c => out[c].push(pack[c][i]));
    }
    newRows.forEach(r => cols.forEach(c => out[c].push(r[c])));
    return out;
  }

  /* Указатель заказов (order_idx.json) — плоский список всех номеров с
     ссылкой на площадко-год. По нему работает поиск заказа и, главное,
     определяется, для каких площадко-лет вообще есть полная детализация:
     без обновления новый год не подхватится вкладками «Заказы» и
     «Оборудование». */
  function mergeOrderIndex(idx, results) {
    const oldCombos = (idx && idx.combos) || [];
    const oldO = (idx && idx.o) || [], oldC = (idx && idx.c) || [];
    const touched = new Set(results.map(r => r.meta.site + "_" + r.meta.year));
    const o = [], keys = [];
    for (let i = 0; i < oldO.length; i++) {
      const k = oldCombos[oldC[i]];
      if (touched.has(k)) continue;
      o.push(oldO[i]); keys.push(k);
    }
    results.forEach(r => {
      const k = r.meta.site + "_" + r.meta.year;
      const seen = new Set();
      r.detail.o.forEach(x => { if (x && !seen.has(x)) { seen.add(x); o.push(x); keys.push(k); } });
    });
    const combos = [...new Set(keys)].sort();
    const pos = new Map(combos.map((k, i) => [k, i]));
    return { combos, o, c: keys.map(k => pos.get(k)), n: o.length };
  }

  /* Каталог каталожных номеров (oem_idx.json): пары «компонент → номер
     детали производителя» из новых выгрузок дописываем к существующим,
     иначе у новых позиций колонка каталожного номера останется пустой. */
  function mergeOemIndex(idx, results) {
    const oem = ((idx && idx.oem) || []).slice();
    const c = ((idx && idx.c) || []).slice();
    const have = new Set(c);
    let added = 0;
    results.forEach(r => {
      (r.oemPairs || []).forEach(([name, part]) => {
        if (!name || !part || have.has(name)) return;
        have.add(name); c.push(name); oem.push(part); added++;
      });
    });
    return { idx: { oem, c }, added };
  }

  function mergeIntoCube(cube, capexRows, results) {
    const touched = new Set(results.map(r => r.meta.site + "|" + r.meta.year));
    const keep = r => !touched.has(r.s + "|" + r.y);
    const cat = (arr, pick) => (arr || []).filter(keep).concat(...results.map(pick));

    const out = Object.assign({}, cube);
    out.combo = cat(cube.combo, r => r.combo);
    out.eoList = cat(cube.eoList, r => r.eoList);
    out.usoSlices = cat(cube.usoSlices, r => r.usoSlices);
    /* Позиции МТР по срезам — не «топ на площадко-год», а глобально
       крупнейшие по всей витрине (сверено: 6000 из 6000 совпадают).
       Поэтому после пересборки одной выгрузки список переотбирается
       целиком, иначе обновлённый год занял бы в нём чужое место. */
    out.matSlices = cat(cube.matSlices, r => r.matSlices)
      .sort((x, y) => (y.p + y.a) - (x.p + x.a))
      .slice(0, MAT_LIMIT);
    out.vmzSlices = cat(cube.vmzSlices, r => r.vmzSlices);
    out.mtr_pack = mergePack(cube.mtr_pack, cube.mtr_cols, [].concat(...results.map(r => r.mtrPack)), touched);
    out.order_pack = mergePack(cube.order_pack, cube.order_cols, [].concat(...results.map(r => r.orderRows)), touched);

    const files = (cube.files || []).slice();
    results.forEach(r => {
      const name = `M06_${r.meta.site}_${r.meta.year}.xlsx`;
      const rec = { file: name, site: r.meta.site, year: Number(r.meta.year), rows: r.meta.rows, orders: r.meta.orders, refSkipped: r.meta.refSkipped || 0 };
      const i = files.findIndex(f => f.file === name || (String(f.site) === r.meta.site && String(f.year) === r.meta.year));
      if (i >= 0) files[i] = rec; else files.push(rec);
    });
    out.files = files;
    out.rows = files.reduce((a, f) => a + (f.rows || 0), 0);
    // «Заказов в источнике» суммируем по выгрузкам; если у какой-то
    // выгрузки счётчик ещё не записан, берём прежний итог, чтобы не
    // занизить знаменатель в раскрытии покрытия.
    out.order_source = files.every(f => f.orders != null)
      ? files.reduce((a, f) => a + f.orders, 0)
      : (cube.order_source || 0);
    out.eo_source = out.eoList.length;

    out.sites = Object.assign({}, cube.sites || {});
    results.forEach(r => { if (!out.sites[r.meta.site]) out.sites[r.meta.site] = r.meta.site; });

    const capex = (capexRows || []).filter(keep).concat(...results.map(r => r.capexRows));
    return { cube: out, capex, touched: [...touched] };
  }


  /* Слияние витрины использования материалов (mtr.json) с результатами
     пересборки. На вход — объект витрины, у которого поле k уже развёрнуто
     из дельт (так его держит дашборд после загрузки), на выход — новый
     объект, готовый к сохранению: k снова дельта-кодирован.

     Ячейка строки закодирована как ((площадка × число лет) + год) × 4 +
     признак CAPEX, поэтому появление новой площадки или нового года
     сдвигает всю кодировку. Чтобы этого не бояться, сначала разбираем
     существующие строки в тройки (площадка, год, признак), потом
     кодируем заново уже с новыми измерениями.

     Справочная часть — наименования, каталожные номера, ППЗ, изготовители
     и группы аналогов — собрана из выгрузок папки ЕКМТР и здесь не
     меняется: у новых кодов ЕКМТР заполняется только наименование из
     самой выгрузки PM-06, остальное остаётся пустым до обновления
     справочников. */
  function mergeMatUse(mtr, results) {
    const NYold = mtr.years.length;
    const sites = mtr.sites.slice(), years = mtr.years.slice();
    const touched = new Set();
    results.forEach(r => {
      const st = String(r.meta.site), yr = String(r.meta.year);
      if (!sites.includes(st)) sites.push(st);
      if (!years.includes(yr)) years.push(yr);
      touched.add(st + "|" + yr);
    });
    sites.sort(); years.sort();
    const NY = years.length;
    const cellOf = (st, yr, cx) => ((sites.indexOf(st) * NY) + years.indexOf(yr)) * 4 + cx;

    /* словарь материалов: старый + коды, которых в нём ещё не было */
    const ek = mtr.ek.slice(), nm = mtr.nm.slice(), cat = mtr.cat.slice();
    const ppzi = mtr.ppzi.slice(), mfi = mtr.mfi.slice();
    const ix = new Map(ek.map((c, i) => [c, i]));
    const emptyPpz = mtr.ppz.indexOf("") >= 0 ? mtr.ppz.indexOf("") : (mtr.ppz.push(""), mtr.ppzc.push(""), mtr.ppz.length - 1);
    const emptyMf = mtr.mf.indexOf("") >= 0 ? mtr.mf.indexOf("") : (mtr.mf.push(""), mtr.mf.length - 1);
    let added = 0;
    results.forEach(r => {
      (r.matNames || []).forEach(([code, name]) => {
        if (!code || ix.has(code)) return;
        ix.set(code, ek.length); ek.push(code); nm.push(name || ""); cat.push("");
        ppzi.push(emptyPpz); mfi.push(emptyMf); added++;
      });
      (r.matRows || []).forEach(row => {
        if (!row.k || ix.has(row.k)) return;
        ix.set(row.k, ek.length); ek.push(row.k); nm.push(""); cat.push("");
        ppzi.push(emptyPpz); mfi.push(emptyMf); added++;
      });
    });

    /* существующие строки, кроме затронутых площадко-лет */
    const rows = [];
    for (let i = 0; i < mtr.k.length; i++) {
      const c = mtr.c[i], cx = c & 3, t = c >> 2;
      const st = mtr.sites[(t / NYold) | 0], yr = mtr.years[t % NYold];
      if (touched.has(st + "|" + yr)) continue;
      rows.push({ k: ix.get(mtr.ek[mtr.k[i]]), c: cellOf(st, yr, cx), qp: mtr.qp[i], qf: mtr.qf[i], p: mtr.p[i], a: mtr.a[i], n: mtr.n[i], o: mtr.o[i], e: mtr.e[i] });
    }
    /* и заново собранные */
    let addedRows = 0;
    results.forEach(r => {
      const st = String(r.meta.site), yr = String(r.meta.year);
      (r.matRows || []).forEach(row => {
        rows.push({ k: ix.get(row.k), c: cellOf(st, yr, row.cx), qp: row.qp, qf: row.qf, p: row.p, a: row.a, n: row.n, o: row.o, e: row.e });
        addedRows++;
      });
    });
    rows.sort((x, y) => (x.k - y.k) || (x.c - y.c));

    const K = rows.map(r => r.k);
    const out = {
      v: mtr.v, built: new Date().toISOString().slice(0, 10),
      sites, years, ek, nm, cat,
      ppz: mtr.ppz, ppzi, ppzc: mtr.ppzc, mf: mtr.mf, mfi,
      k: K.map((v, i) => i ? v - K[i - 1] : v),
      c: rows.map(r => r.c), qp: rows.map(r => r.qp), qf: rows.map(r => r.qf),
      p: rows.map(r => r.p), a: rows.map(r => r.a),
      n: rows.map(r => r.n), o: rows.map(r => r.o), e: rows.map(r => r.e),
      an: mtr.an,
      meta: {
        srcRows: rows.reduce((s2, r) => s2 + r.n, 0),
        usedCodes: new Set(rows.map(r => r.k)).size,
        refCodes: mtr.meta.refCodes, groups: mtr.meta.groups
      }
    };
    return { mtr: out, addedCodes: added, addedRows, touched: [...touched] };
  }

  return { createProcessor, mergeIntoCube, mergeOrderIndex, mergeOemIndex, mergeMatUse, U_MAP, U_SHORT, toIso, DictEnc, MTR_PER_EO, MTR_MIN_COST, ORD_PER_EO };
})();

if (typeof module !== "undefined" && module.exports) module.exports = PM06;
