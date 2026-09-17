/* Разбор и пересборка «МТР УСО {год} all.xlsx» — материалов, которые
   подрядчики списывают при исполнении заказов УСО (отдельная выгрузка
   BW, не PM-06 «Гибкий отчёт»; лист внутри файла тоже называется
   PM-06 — не путать с основной выгрузкой). Логика 1:1 повторяет
   scratchpad-скрипт build_uso_mtr2.py, которым собран текущий
   uso_mtr.json: то же исключение строк-услуг, то же лечение порчи
   «Количество факт», та же построчная упаковка (формат v2).

   Формат v2: словари + индексы. Свойства, постоянные внутри заказа
   (площадка, балансовая единица-заказчик, производитель, модель
   заводская) и подтянутые по номеру
   заказа из полной детализации ТОиР (контрагент, вид работы, способ,
   причина инвестиций, базис-месяц, единица оборудования) лежат один раз
   на заказ; на строке — материал, № ЕКМТР, количества и суммы. Так
   витрина вдвое компактнее предагрегатов и при этом фильтруется по
   всем измерениям дашборда.

   Схема листа: именованные колонки находим по заголовку (buildColIndex),
   номер заказа лежит в безымянной колонке сразу за «Заказ» — тот же
   приём, что и в pm06_pipeline.js для № ЕКМТР. */

const USOPIPE = (() => {
  "use strict";

  const FMT = 2;

  /* Структура: заказчик (балансовая единица — 1100 Полюс Красноярск,
     1200 Полюс Вернинское, 1300 Полюс Алдан, 1400 Полюс Магадан,
     2400 Полюс Сухой Лог) → подрядчик АО «Развитие», который ведёт
     заказ (коды планирующего завода 7101-7104) → субподрядчики,
     выполняющие работы и поставки (контрагент заказа).

     Площадка строки — это заказчик (BE_SITE). Код завода в колонке
     «Завод, планирующий ТОРО» для фильтра не годится: у 96,6% плана
     там стоит подрядчик. SITE_REGION — только запасной путь, если
     заказчик не распознан; 7102 обслуживает сразу Вернинское и Сухой
     Лог, поэтому по коду завода их не различить. */
  const SITE_LABELS = {
    "1100": "Красноярск / Еруда", "1200": "Вернинское", "1300": "Алдан", "1400": "Магадан",
    "2400": "Сухой Лог",
    "7101": "Развитие Красноярск", "7102": "Развитие Иркутск", "7103": "Развитие Алдан", "7104": "Развитие Магадан"
  };
  const SITE_REGION = { "7101": "1100", "7102": "1200", "7103": "1300", "7104": "1400" };
  const BE_SITE = {
    "АО \"Полюс Красноярск\"": "1100",
    "АО \"Полюс Вернинское\"": "1200",
    "АО \"Полюс Алдан\"": "1300",
    "АО \"Полюс Магадан\"": "1400",
    "ООО \"Полюс Сухой Лог\"": "2400"
  };

  function s(v) { return v == null ? "" : String(v).trim(); }
  function num(v) {
    if (v === "" || v == null) return 0;
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }
  function r2(v) { return Math.round(v * 100) / 100; }
  /* Округление «половина к чётному», как в Python round(): именно им
     собран опубликованный uso_mtr.json, и без этого пересборка того же
     файла в браузере разошлась бы с ним на сотни рублей. */
  function r0(v) {
    const f = Math.floor(v), d = v - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }
  function q2(v) {
    const x = v * 100, f = Math.floor(x), d = x - f;
    const r = d > 0.5 ? f + 1 : d < 0.5 ? f : (f % 2 === 0 ? f : f + 1);
    return r / 100;
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
  function isHeaderRow(cells) {
    for (let i = 0; i < cells.length; i++) {
      if (typeof cells[i] === "string" && cells[i].trim() === "Материал подрядчика Краткий текст МТР") return true;
    }
    return false;
  }

  const NEEDED = ["Завод, планирующий ТОРО", "Заказ", "Материал подрядчика Краткий текст МТР",
    "ЕО изготовитель", "ЕО Марка заводская", "Компонент Заказа/Заявки",
    "Количество план (подр.)", "Количество факт (подр.)", "Цена план, в руб. (подр.)",
    "Стоимость план, в руб. (подр.)", "Стоимость факт, в руб. (подр.)"];

  /* Один файл = один год, но сразу все площадки — в отличие от основной
     выгрузки PM-06, где один файл = одна площадка и один год. */
  function createProcessor(year) {
    let cols = null, headerSeen = false, scanned = 0, c = {};
    const rows = [];
    let nTotal = 0, nService = 0, nCorrupt = 0, corruptCost = 0;

    function setHeader(cells) {
      cols = buildColIndex(cells);
      const missing = NEEDED.filter(k => !(k in cols));
      if (missing.length) throw new Error("В выгрузке нет колонок: " + missing.join(", ") + " — это не файл «МТР УСО»?");
      c = {
        site: cols["Завод, планирующий ТОРО"],
        order: cols["Заказ"] + 1,
        name: cols["Материал подрядчика Краткий текст МТР"],
        mfr: cols["ЕО изготовитель"],
        model: cols["ЕО Марка заводская"],
        comp: cols["Компонент Заказа/Заявки"],
        /* Балансовая единица — заказчик (чья техника и чей бюджет). В
           заказах сервисных площадок АО «Развитие» (7101-7104) он и
           объясняет, за каким добывающим активом числится работа.
           Колонка необязательная: в старых выгрузках её может не быть. */
        be: cols["Балансовая единица"],
        qp: cols["Количество план (подр.)"],
        qf: cols["Количество факт (подр.)"],
        unitp: cols["Цена план, в руб. (подр.)"],
        p: cols["Стоимость план, в руб. (подр.)"],
        a: cols["Стоимость факт, в руб. (подр.)"]
      };
      headerSeen = true;
    }

    function pushRow(cells) {
      if (!headerSeen) {
        if (++scanned > 60) throw new Error("Не найдена строка заголовка с колонкой «Материал подрядчика Краткий текст МТР» — это не выгрузка «МТР УСО»");
        if (isHeaderRow(cells)) setHeader(cells);
        return;
      }
      if (!cells || !cells.length) return;
      const site = s(cells[c.site]);
      if (!site || site === "Завод, планирующий ТОРО") return;
      const mat = s(cells[c.name]);
      if (!mat || mat === "#") return;
      nTotal++;
      const mfr = s(cells[c.mfr]);
      if (!mfr || mfr === "Не определен") { nService++; return; }
      const model = s(cells[c.model]);
      const comp = s(cells[c.comp]);
      let qf = num(cells[c.qf]), a = num(cells[c.a]);
      /* Порча источника: «Количество факт» иногда содержит не количество,
         а цену за единицу (та же цифра, что и «Цена план, в руб.») —
         стоимость факт при этом = кол-во(=цена) × цена, раздувая факт в
         тысячи раз на одной строке. План не трогаем. */
      const unitp = num(cells[c.unitp]);
      if (unitp && qf > 10 && Math.abs(qf - unitp) < unitp * 0.01) {
        nCorrupt++; corruptCost += a;
        qf = 0; a = 0;
      }
      rows.push({
        site, order: s(cells[c.order]), name: mat, mfr,
        be: c.be == null ? "" : s(cells[c.be]),
        model: model && model !== "#" ? model : "",
        ek: comp && comp !== "#" ? comp : "",
        y: year, qp: q2(num(cells[c.qp])), qf: q2(qf),
        p: r0(num(cells[c.p])), a: r0(a)
      });
    }

    function finish() {
      return { year, rows, meta: { totalRows: nTotal, serviceRowsDropped: nService, materialRows: rows.length, corruptFactRows: nCorrupt, corruptFactValue: r2(corruptCost) } };
    }
    return { pushRow, finish };
  }

  /* Разворачивает упакованную витрину обратно в плоские строки — так
     слияние нового года со старыми сводится к «выкинуть строки этих лет,
     досыпать новые, упаковать заново», без ручной правки индексов. */
  function unpack(data) {
    if (!data || !data.row || !data.dict) return [];
    const D = data.dict, O = data.ord, R = data.row, out = [];
    for (let i = 0; i < R.o.length; i++) {
      const oi = R.o[i];
      out.push({
        site: D.st[O.s[oi]], order: O.o[oi], name: D.nm[R.n[i]], ek: D.ek[R.ek[i]],
        mfr: D.mf[O.mf[oi]], model: D.md[O.md[oi]], y: D.yr[O.y[oi]],
        be: (D.be && O.be) ? D.be[O.be[oi]] : "",
        qp: R.qp[i], qf: R.qf[i], p: R.p[i], a: R.a[i],
        wc: D.wc[O.wc[oi]], wk: D.wk[O.wk[oi]], mu: D.mu[O.mu[oi]],
        rs: D.rs[O.rs[oi]], mo: D.mo[O.mo[oi]], eo: D.eo[O.eo[oi]]
      });
    }
    return out;
  }

  function Dict() {
    const items = [], index = new Map();
    return {
      items,
      get(v) {
        v = v || "";
        let i = index.get(v);
        if (i === undefined) { i = items.length; items.push(v); index.set(v, i); }
        return i;
      }
    };
  }

  /* Упаковка плоских строк в формат v2. Ключ заказа — номер И год: один
     и тот же заказ встречается в выгрузках соседних лет, а год строки
     должен остаться годом её собственной выгрузки. */
  function pack(rows, meta) {
    const D = {
      nm: Dict(), ek: Dict(), mf: Dict(), md: Dict(), st: Dict(), yr: Dict(),
      wc: Dict(), wk: Dict(), mu: Dict(), rs: Dict(), mo: Dict(), eo: Dict(), be: Dict()
    };
    const orders = new Map();
    const O = { o: [], s: [], y: [], mf: [], md: [], wc: [], wk: [], mu: [], rs: [], mo: [], eo: [], be: [] };
    const R = { o: [], n: [], ek: [], qp: [], qf: [], p: [], a: [] };

    rows.forEach(r => {
      const okey = r.order + "|" + r.y;
      let oi = orders.get(okey);
      if (oi === undefined) {
        oi = O.o.length;
        orders.set(okey, oi);
        O.o.push(r.order);
        O.s.push(D.st.get(r.site));
        O.y.push(D.yr.get(r.y));
        O.mf.push(D.mf.get(r.mfr));
        O.md.push(D.md.get(r.model));
        O.wc.push(D.wc.get(r.wc));
        O.wk.push(D.wk.get(r.wk));
        O.mu.push(D.mu.get(r.mu));
        O.rs.push(D.rs.get(r.rs));
        O.mo.push(D.mo.get(r.mo));
        O.eo.push(D.eo.get(r.eo));
        O.be.push(D.be.get(r.be));
      }
      R.o.push(oi);
      R.n.push(D.nm.get(r.name));
      R.ek.push(D.ek.get(r.ek));
      R.qp.push(r.qp);
      R.qf.push(r.qf);
      R.p.push(r.p);
      R.a.push(r.a);
    });

    const dict = {};
    Object.keys(D).forEach(k => { dict[k] = D[k].items; });
    return { v: FMT, meta, dict, ord: O, row: R };
  }

  /* Свойства заказа (контрагент, вид работы, способ, причина, месяц, ЕО)
     приходят из index.html — он берёт их из полной детализации ТОиР. */
  function enrich(rows, props) {
    let matched = 0;
    rows.forEach(r => {
      const pr = props.get(r.order);
      if (pr) {
        r.wc = pr.wc || ""; r.wk = pr.w || ""; r.mu = pr.u || "";
        r.rs = pr.orr || ""; r.mo = (pr.bs || "").slice(0, 7); r.eo = pr.e || "";
        if (r.wc) matched++;
      } else {
        r.wc = r.wk = r.mu = r.rs = r.mo = r.eo = "";
      }
    });
    return matched;
  }

  /* Слияние: годы из processedResults заменяются целиком, остальные
     переносятся из текущей витрины как есть. */
  function mergeIntoData(current, results, props) {
    const updated = new Set(results.map(r => String(r.year)));
    const kept = unpack(current).filter(r => !updated.has(String(r.y)));

    const fresh = [];
    results.forEach(r => fresh.push(...r.rows));
    const matched = enrich(fresh, props);

    const byYear = Object.assign({}, (current.meta && current.meta.byYear) || {});
    results.forEach(r => {
      const m = Object.assign({}, r.meta);
      m.matchedToContractor = r.rows.filter(x => x.wc).length;
      byYear[r.year] = m;
    });
    const sums = k => Object.values(byYear).reduce((a, v) => a + (v[k] || 0), 0);
    const meta = Object.assign({}, current.meta, {
      years: [...new Set([...((current.meta && current.meta.years) || []), ...updated])].sort(),
      byYear,
      totalRows: sums("totalRows"),
      serviceRowsDropped: sums("serviceRowsDropped"),
      materialRows: sums("materialRows"),
      matchedToContractor: sums("matchedToContractor"),
      corruptFactRows: sums("corruptFactRows"),
      corruptFactValue: r2(sums("corruptFactValue")),
      siteLabels: Object.assign({}, (current.meta && current.meta.siteLabels) || {}, SITE_LABELS),
      siteRegion: Object.assign({}, (current.meta && current.meta.siteRegion) || {}, SITE_REGION),
      beSite: Object.assign({}, (current.meta && current.meta.beSite) || {}, BE_SITE)
    });

    return { data: pack(kept.concat(fresh), meta), updatedYears: updated, matched, freshRows: fresh.length };
  }

  /* Суммы плана и факта по годам — для таблицы «что изменится». */
  function yearTotals(data) {
    const out = new Map();
    if (!data || !data.row) return out;
    const D = data.dict, O = data.ord, R = data.row;
    for (let i = 0; i < R.o.length; i++) {
      const y = D.yr[O.y[R.o[i]]];
      let x = out.get(y);
      if (!x) out.set(y, x = { y, p: 0, a: 0, n: 0 });
      x.p += R.p[i]; x.a += R.a[i]; x.n++;
    }
    return out;
  }

  return { FMT, SITE_LABELS, SITE_REGION, BE_SITE, createProcessor, unpack, pack, enrich, mergeIntoData, yearTotals };
})();

if (typeof module !== "undefined" && module.exports) module.exports = USOPIPE;
