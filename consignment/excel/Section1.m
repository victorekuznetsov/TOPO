section Section1;

// ---------- служебные функции ----------
shared фнТекст = (v as any) as text =>
    let t = if v = null then "" else Text.Trim(Text.From(v))
    in if t = "#" then "" else t;

shared фнЧисло = (v as any) as number =>
    if v = null then 0
    else if Value.Is(v, type number) then v
    else
        let t = Text.Remove(Text.Trim(Text.From(v)), {" ", "#(00A0)"})
        in if t = "" or t = "#" then 0 else (try Number.FromText(Text.Replace(t, ",", "."), "en-US") otherwise 0);

shared фнДата = (v as any) as nullable date =>
    if v = null then null
    else if Value.Is(v, type date) then v
    else if Value.Is(v, type datetime) then Date.From(v)
    else if Value.Is(v, type number) then (try Date.From(v) otherwise null)
    else
        let t = фнТекст(v), p = Text.Split(t, ".")
        in if t = "" then null
           else if List.Count(p) = 3 then (try #date(Number.FromText(Text.Start(p{2}, 4)), Number.FromText(p{1}), Number.FromText(p{0})) otherwise null)
           else (try Date.From(DateTime.FromText(t)) otherwise null);

shared фнНомер = (v as any) as text => Text.Upper(Text.Remove(фнТекст(v), {" ", "-", "#(00A0)"}));

shared фнСумма = (l as list) as number => List.Sum({0} & List.RemoveNulls(l));

shared фнЯчейка = (name as text) as any => Excel.CurrentWorkbook(){[Name = name]}[Content]{0}[Column1];

// ---------- параметры и настройки из книги ----------
shared ПапкаДанных = let p = фнТекст(фнЯчейка("ПапкаДанных")) in if Text.EndsWith(p, "\") then p else p & "\";

shared ГодПлана = Number.From(фнЯчейка("ГодПлана"));

shared ДатаЗакупки = фнДата(фнЯчейка("ДатаВыгрузкиЗакупки"));

shared Техника = Table.Buffer(Table.SelectRows(
    Table.TransformColumns(Excel.CurrentWorkbook(){[Name = "Техника"]}[Content],
        {{"Площадка", фнТекст}, {"Начало наименования ЕО", фнТекст}, {"Модель", фнТекст}}),
    each [Площадка] <> "" and [#"Начало наименования ЕО"] <> ""));

shared Заводы = Table.Buffer(Table.SelectRows(
    Table.TransformColumns(Excel.CurrentWorkbook(){[Name = "Заводы"]}[Content],
        {{"Площадка", фнТекст}, {"Код завода", фнТекст}, {"Роль", фнТекст}}),
    each [Площадка] <> "" and [#"Код завода"] <> ""));

shared Исключения = Table.Buffer(Table.SelectRows(
    Table.TransformColumns(Excel.CurrentWorkbook(){[Name = "Исключения"]}[Content],
        {{"Начало наименования", each Text.Lower(фнТекст(_))}, {"Кроме", each Text.Lower(фнТекст(_))}, {"Причина", фнТекст}}),
    each [#"Начало наименования"] <> ""));

// ---------- файлы в папке ----------
shared Файлы = Table.Buffer(Table.SelectRows(Folder.Files(ПапкаДанных), each not Text.StartsWith([Name], "~$")));

shared фнФайл = (prefix as text) as binary =>
    let f = Table.Sort(
            Table.SelectRows(Файлы, each Text.StartsWith(Text.Lower([Name]), Text.Lower(prefix))
                and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))),
            {{"Date modified", Order.Descending}})
    in if Table.IsEmpty(f)
       then error Error.Record("Нет файла", "В папке «" & ПапкаДанных & "» нет файла, имя которого начинается с «" & prefix & "»")
       else f{0}[Content];

shared фнЛист = (content as binary, name as text, anyIfMissing as logical) as table =>
    let wb = Excel.Workbook(content, null, true),
        sheets = Table.SelectRows(wb, each [Kind] = "Sheet"),
        hit = Table.SelectRows(sheets, each [Name] = name)
    in if not Table.IsEmpty(hit) then hit{0}[Data]
       else if anyIfMissing then sheets{0}[Data]
       else error Error.Record("Нет листа", "В файле нет листа «" & name & "»");

shared фнСтрокаШапки = (t as table, marks as list) as number =>
    let rows = Table.ToRows(Table.FirstN(t, 60)),
        i = List.PositionOf(List.Transform(rows, (r) => List.ContainsAll(List.Transform(r, фнТекст), marks)), true)
    in if i < 0 then error Error.Record("Нет шапки", "Не найдена строка шапки с «" & Text.Combine(marks, "», «") & "»") else i;

// ---------- справочники ----------
shared ЕКМТР = let
    t = Table.PromoteHeaders(фнЛист(фнФайл("ЕКМТР"), "Лист1", true), [PromoteAllScalars = true]),
    s = Table.SelectColumns(t, {"ЕКМТР№", "Каталожный номер", "Изготовитель"}, MissingField.UseNull),
    x = Table.TransformColumns(s, {{"ЕКМТР№", фнТекст}, {"Каталожный номер", фнТекст}, {"Изготовитель", фнТекст}}),
    r = Table.RenameColumns(x, {{"ЕКМТР№", "Код"}, {"Каталожный номер", "Номер"}}),
    nz = Table.SelectRows(r, each [Код] <> "")
in Table.Buffer(Table.Distinct(nz, {"Код"}));

shared Замены = let
    wb = Excel.Workbook(фнФайл("Замены"), null, true),
    sheets = Table.SelectRows(wb, each [Kind] = "Sheet"),
    one = (data as table) as table =>
        let t = Table.PromoteHeaders(data, [PromoteAllScalars = true]),
            s = Table.SelectColumns(t, {"ItemId", "ReplacingNumber"}, MissingField.UseNull)
        in Table.RenameColumns(s, {{"ItemId", "Старый"}, {"ReplacingNumber", "Новый"}}),
    all = Table.Combine(List.Transform(sheets[Data], one)),
    x = Table.TransformColumns(all, {{"Старый", фнНомер}, {"Новый", фнНомер}}),
    ok = Table.SelectRows(x, each Text.Length([Старый]) >= 3 and Text.Length([Новый]) >= 3 and [Старый] <> [Новый])
in Table.Buffer(Table.Distinct(ok));

shared ЗаменыВперёд = let
    g = Table.Group(Замены, {"Старый"}, {{"Новые", each List.Sort(List.Distinct([Новый])), type list}})
in Record.FromList(g[Новые], g[Старый]);

shared ЗаменыНазад = let
    g = Table.Group(Замены, {"Новый"}, {{"Старые", each List.Sort(List.Distinct([Старый])), type list}})
in Record.FromList(g[Старые], g[Новый]);

// Взаимозаменяемые номера: вперёд цепочкой до 4 шагов (развилка больше 4 — дальше не идём), назад — на шаг.
shared фнСемьяЗамен = (no as text) as list =>
    let
        fwd = ЗаменыВперёд,
        back = ЗаменыНазад,
        step = (edge as list, depth as number, seen as list, acc as list) as list =>
            if List.IsEmpty(edge) or depth >= 4 then acc
            else
                let kids = List.Combine(List.Transform(edge, (x) =>
                        let k = Record.FieldOrDefault(fwd, x, {}) in if List.Count(k) > 4 and x <> no then {} else k)),
                    fresh = List.Distinct(List.Select(kids, (k) => not List.Contains(seen, k)))
                in @step(fresh, depth + 1, seen & fresh, acc & fresh),
        f = if no = "" then {} else step({no}, 0, {no}, {}),
        b = if no = "" then {} else List.Select(Record.FieldOrDefault(back, no, {}), (k) => not List.Contains({no} & f, k))
    in List.FirstN(f & b, 16);

// ---------- ТОРО (PM-06) ----------
shared фнКлассРабот = (w as text) as text =>
    let l = Text.Lower(w)
    in if List.AnyTrue(List.Transform({"техническое обслуж", "ежесмен", "сезонн", "годовое"}, (p) => Text.StartsWith(l, p))) then "ТО"
       else if List.AnyTrue(List.Transform({"капитальный ремонт", "капитализируемый", "кр ", "средний ремонт"}, (p) => Text.StartsWith(l, p))) then "КР"
       else "ВНЕ";

shared фнИсключение = (name as text) as text =>
    let l = Text.Lower(name),
        hit = Table.SelectRows(Исключения, (r) =>
            (Text.StartsWith(l, r[#"Начало наименования"]) or l = Text.Trim(r[#"Начало наименования"]))
            and (r[Кроме] = "" or not Text.StartsWith(l, r[Кроме])))
    in if Table.IsEmpty(hit) then "" else hit{0}[Причина];

shared фнPM06 = (content as binary) as table =>
    let
        sh = фнЛист(content, "PM-06", true),
        hi = фнСтрокаШапки(sh, {"Заказ"}),
        hdr = List.Transform(Record.FieldValues(sh{hi}), фнТекст),
        names = Table.ColumnNames(sh),
        spec = {
            {"Заказ", "Заказ", 0}, {"ЕО", "ЕО", 0},
            {"Компонент", "Компонент Заказа/Заявки", 0}, {"ЕКМТР", "Компонент Заказа/Заявки", 1},
            {"ВидРабот", "Заказ Вид работы ТОРО", 0}, {"Причина", "СПП-элемент Причина инвестиций", 0},
            {"Изготовитель", "Компонент заказа изготовитель", 0},
            {"СтПлан", "Стоимость МТР, План", 0}, {"СтФакт", "Стоимость МТР, Факт", 0},
            {"УсоПлан", "Стоимость УСО, План", 0}, {"УсоФакт", "Стоимость УСО, Факт", 0},
            {"КолПлан", "Количество МТР, План", 0}, {"КолФакт", "Количество МТР, Факт", 0},
            {"ДатаПотр", "Компонент Заказа/Заявки Дата потребности", 0},
            {"БазСрок", "Заказ Базисный срок начала (дата)", 0}},
        found = List.Transform(spec, (s) => let p = List.PositionOf(hdr, s{1}) in {s{0}, if p < 0 then null else names{p + s{2}}}),
        missing = List.Select(found, (f) => f{1} = null),
        mustHave = {"Заказ", "ЕО", "Компонент", "ЕКМТР", "ВидРабот", "КолПлан", "КолФакт"},
        broken = List.Select(missing, (f) => List.Contains(mustHave, f{0})),
        have = List.Select(found, (f) => f{1} <> null),
        body = Table.SelectColumns(Table.Skip(sh, hi + 1), List.Transform(have, (f) => f{1})),
        ren = Table.RenameColumns(body, List.Transform(have, (f) => {f{1}, f{0}})),
        full = List.Accumulate(missing, ren, (t, m) => Table.AddColumn(t, m{0}, each null))
    in if List.IsEmpty(broken) then full
       else error Error.Record("PM-06", "В выгрузке нет колонок: " & Text.Combine(List.Transform(broken, (f) => f{0}), ", "));

shared фнТОРОфайл = (content as binary, site as text, year as number) as table =>
    let
        raw = фнPM06(content),
        tech = Table.SelectRows(Техника, (r) => r[Площадка] = site),
        prefixes = List.Buffer(tech[#"Начало наименования ЕО"]),
        models = List.Buffer(tech[Модель]),
        eo = Table.TransformColumns(raw, {{"ЕО", фнТекст}}),
        withModel = Table.AddColumn(eo, "Модель", (r) =>
            let i = List.PositionOf(List.Transform(prefixes, (p) => Text.StartsWith(r[ЕО], p)), true)
            in if i < 0 then null else models{i}),
        fleet = Table.SelectRows(withModel, (r) => r[Модель] <> null),
        tx = Table.TransformColumns(fleet, {
            {"Заказ", фнТекст}, {"Компонент", фнТекст}, {"ЕКМТР", фнТекст}, {"ВидРабот", фнТекст},
            {"Причина", фнТекст}, {"Изготовитель", фнТекст},
            {"СтПлан", фнЧисло}, {"СтФакт", фнЧисло}, {"УсоПлан", фнЧисло}, {"УсоФакт", фнЧисло},
            {"КолПлан", фнЧисло}, {"КолФакт", фнЧисло}}),
        ordInfo = Table.Group(tx, {"Заказ"}, {
            {"ПричинаЗаказа", each List.First(List.Select([Причина], (x) => x <> ""), ""), type text},
            {"СрокЗаказа", each List.First(List.RemoveNulls(List.Transform([БазСрок], фнДата)), null), type nullable date}}),
        nz = Table.SelectRows(tx, (r) => (r[СтПлан] <> 0 or r[СтФакт] <> 0 or r[УсоПлан] <> 0 or r[УсоФакт] <> 0)
            and r[ЕКМТР] <> "" and not Text.StartsWith(r[ЕКМТР], "SRV")),
        isPlan = year = ГодПлана,
        qv = Table.AddColumn(Table.AddColumn(nz, "Кол", (r) => if isPlan then r[КолПлан] else r[КолФакт], type number),
            "Сумма", (r) => if isPlan then r[СтПлан] else r[СтФакт], type number),
        pos = Table.SelectRows(qv, (r) => r[Кол] > 0),
        j = Table.ExpandTableColumn(Table.NestedJoin(pos, {"Заказ"}, ordInfo, {"Заказ"}, "о", JoinKind.LeftOuter), "о", {"ПричинаЗаказа", "СрокЗаказа"}),
        k = Table.ExpandTableColumn(Table.NestedJoin(j, {"ЕКМТР"}, ЕКМТР, {"Код"}, "е", JoinKind.LeftOuter), "е", {"Изготовитель"}, {"ИзготовительЕК"}),
        cat = Table.SelectRows(k, (r) => Text.StartsWith(Text.Lower(r[Изготовитель]), "caterp")
            or Text.StartsWith(Text.Lower(фнТекст(r[ИзготовительЕК])), "caterp")),
        c1 = Table.AddColumn(cat, "Класс", (r) => фнКлассРабот(r[ВидРабот]), type text),
        c2 = Table.AddColumn(c1, "Часть", (r) =>
            if Text.StartsWith(r[ЕО], "ДВС") or r[Класс] = "КР" or Text.StartsWith(Text.Lower(r[ВидРабот]), "восстановление")
            then "Цех" else "Поле", type text),
        c3 = Table.AddColumn(c2, "Причина2", (r) => if фнТекст(r[ПричинаЗаказа]) = "" then "Не присвоено" else r[ПричинаЗаказа], type text),
        c4 = Table.AddColumn(c3, "CAPEX", (r) => Text.Contains(r[Причина2], "CAPEX") or Text.Contains(r[Причина2], "Капитализ")
            or Text.Contains(r[Причина2], "Компонент"), type logical),
        c5 = Table.AddColumn(c4, "Исключение", (r) => фнИсключение(r[Компонент]), type text),
        c6 = Table.AddColumn(c5, "Месяц", (r) =>
            if not isPlan then null
            else let d = фнДата(r[ДатаПотр]), d2 = if d = null then r[СрокЗаказа] else d
                 in if d2 = null then 6 else if Date.Year(d2) < year then 1 else if Date.Year(d2) > year then 12 else Date.Month(d2),
            type nullable number),
        c7 = Table.AddColumn(Table.AddColumn(c6, "Площадка", each site, type text), "Год", each year, type number)
    in Table.SelectColumns(c7, {"Площадка", "Год", "Заказ", "ЕО", "Модель", "Компонент", "ЕКМТР", "ВидРабот", "Класс",
        "Часть", "CAPEX", "Причина2", "Исключение", "Месяц", "Кол", "Сумма"});

shared ТОРО_строки = let
    годы = List.Transform({ГодПлана - 3, ГодПлана - 2, ГодПлана - 1, ГодПлана}, Text.From),
    площадки = List.Distinct(Техника[Площадка]),
    parts = Table.AddColumn(Файлы, "Части", each Text.Split(Text.BeforeDelimiter([Name], "."), "_")),
    m06 = Table.SelectRows(parts, each List.Count([Части]) >= 3 and Text.Upper([Части]{0}) = "M06"
        and List.Contains(площадки, [Части]{1}) and List.Contains(годы, [Части]{2})
        and List.Contains({".xlsx", ".xlsm"}, Text.Lower([Extension]))),
    keyed = Table.AddColumn(m06, "Ключ", each [Части]{1} & "_" & [Части]{2}),
    latest = Table.Distinct(Table.Buffer(Table.Sort(keyed, {{"Date modified", Order.Descending}})), {"Ключ"}),
    withRows = Table.AddColumn(latest, "Строки", each фнТОРОфайл([Content], [Части]{1}, Number.FromText([Части]{2}))),
    all = if Table.IsEmpty(withRows)
        then error Error.Record("Нет выгрузок ТОРО", "В папке нет файлов вида M06_<площадка>_<год>.xlsx")
        else Table.Combine(withRows[Строки])
in Table.Buffer(all);

shared План = let
    t = Table.SelectRows(ТОРО_строки, each [Год] = ГодПлана and [Исключение] = ""),
    byOrder = Table.Group(t, {"Площадка", "ЕКМТР", "Заказ"}, {{"Кол", each List.Sum([Кол]), type number}}),
    ordStat = Table.Group(byOrder, {"Площадка", "ЕКМТР"}, {
        {"Заказов план", each Table.RowCount(_), Int64.Type}, {"Макс отбор план", each List.Max([Кол]), type number}}),
    aggs = {
        {"Наименование план", each [Компонент]{0}, type text},
        {"План шт", each List.Sum([Кол]), type number},
        {"План руб", each List.Sum([Сумма]), type number},
        {"План поле", each фнСумма(Table.SelectRows(_, (r) => r[Часть] = "Поле")[Кол]), type number},
        {"План цех", each фнСумма(Table.SelectRows(_, (r) => r[Часть] = "Цех")[Кол]), type number},
        {"План CAPEX", each фнСумма(Table.SelectRows(_, (r) => r[CAPEX])[Кол]), type number},
        {"План ВНЕ", each фнСумма(Table.SelectRows(_, (r) => r[Класс] = "ВНЕ")[Кол]), type number},
        {"Модели план", each List.Distinct([Модель]), type list},
        {"Виды план", each Table.Group(_, {"ВидРабот"}, {{"q", (g) => List.Sum(g[Кол]), type number}}), type table},
        {"Причины", each Table.Group(_, {"Причина2"}, {{"q", (g) => List.Sum(g[Кол]), type number}}), type table}}
        & List.Transform({1..12}, (m) => {"П" & Text.From(m), (g) => фнСумма(Table.SelectRows(g, (r) => r[Месяц] = m)[Кол]), type number}),
    g = Table.Group(t, {"Площадка", "ЕКМТР"}, aggs),
    j = Table.ExpandTableColumn(Table.NestedJoin(g, {"Площадка", "ЕКМТР"}, ordStat, {"Площадка", "ЕКМТР"}, "о", JoinKind.LeftOuter),
        "о", {"Заказов план", "Макс отбор план"})
in Table.Buffer(j);

shared ВнеПлана = let
    t = Table.SelectRows(ТОРО_строки, each [Год] < ГодПлана and [Класс] = "ВНЕ" and [Исключение] = ""),
    byOrder = Table.Group(t, {"Площадка", "ЕКМТР", "Заказ"}, {{"Кол", each List.Sum([Кол]), type number}}),
    ordMax = Table.Group(byOrder, {"Площадка", "ЕКМТР"}, {{"Макс отбор вне", each List.Max([Кол]), type number}}),
    perYear = List.Combine(List.Transform({3, 2, 1}, (k) => {
        {"Заказов -" & Text.From(k), (g) => List.Count(List.Distinct(Table.SelectRows(g, (r) => r[Год] = ГодПлана - k)[Заказ])), Int64.Type},
        {"Шт -" & Text.From(k), (g) => фнСумма(Table.SelectRows(g, (r) => r[Год] = ГодПлана - k)[Кол]), type number}})),
    aggs = perYear & {
        {"Наименование вне", each [Компонент]{0}, type text},
        {"Вне руб", each List.Sum([Сумма]), type number},
        {"Вне шт", each List.Sum([Кол]), type number},
        {"Вне поле", each фнСумма(Table.SelectRows(_, (r) => r[Часть] = "Поле")[Кол]), type number},
        {"Вне CAPEX", each фнСумма(Table.SelectRows(_, (r) => r[CAPEX])[Кол]), type number},
        {"Модели вне", each List.Distinct([Модель]), type list},
        {"Виды вне", each Table.Group(_, {"ВидРабот"}, {{"q", (g) => List.Sum(g[Кол]), type number}}), type table}},
    g = Table.Group(t, {"Площадка", "ЕКМТР"}, aggs),
    j = Table.ExpandTableColumn(Table.NestedJoin(g, {"Площадка", "ЕКМТР"}, ordMax, {"Площадка", "ЕКМТР"}, "о", JoinKind.LeftOuter),
        "о", {"Макс отбор вне"})
in Table.Buffer(j);

shared Исключено = let
    t = Table.SelectRows(ТОРО_строки, each [Год] = ГодПлана and [Исключение] <> ""),
    g = Table.Group(t, {"Площадка", "ЕКМТР", "Компонент", "Исключение"}, {
        {"План, шт", each List.Sum([Кол]), type number}, {"План, руб", each List.Sum([Сумма]), type number}})
in Table.Sort(g, {{"План, руб", Order.Descending}});

// ---------- остатки MM-M03 ----------
shared Остатки_строки = let
    sh = фнЛист(фнФайл("Остатки"), "MM-M03", false),
    hi = фнСтрокаШапки(sh, {"Материал", "Завод", "Склад"}),
    hdr = List.Transform(Record.FieldValues(sh{hi}), фнТекст),
    names = Table.ColumnNames(sh),
    p = (h as text) as number => let i = List.PositionOf(hdr, h) in if i < 0 then error ("В остатках нет колонки «" & h & "»") else i,
    pNo = List.PositionOf(List.Transform(hdr, (h) => Text.StartsWith(h, "Номер детали")), true),
    cols = {
        {names{p("Балансовая единица") + 1}, "БЕ"}, {names{p("Завод")}, "Завод"},
        {names{p("Категория запаса") + 1}, "КатЗапаса"}, {names{p("Материал")}, "Материал"},
        {names{pNo}, "НомерПроизв"}, {names{p("Категория склада")}, "КатСклада"},
        {names{p("Склад") + 1}, "Склад"}, {names{p("RUB") - 1}, "Кол"}},
    body = Table.RenameColumns(Table.SelectColumns(Table.Skip(sh, hi + 1), List.Transform(cols, (c) => c{0})), cols),
    tx = Table.TransformColumns(body, {{"БЕ", фнТекст}, {"Завод", фнТекст}, {"КатЗапаса", фнТекст}, {"Материал", фнТекст},
        {"НомерПроизв", фнНомер}, {"КатСклада", фнТекст}, {"Склад", фнТекст}, {"Кол", фнЧисло}}),
    pos = Table.SelectRows(tx, each [Кол] > 0 and [Материал] <> ""),
    z = Table.ExpandTableColumn(Table.NestedJoin(pos, {"Завод"}, Заводы, {"Код завода"}, "з", JoinKind.Inner), "з", {"Площадка", "Роль"}),
    b = Table.AddColumn(z, "Корзина", each
        if [Роль] = "Справочно" then "Другой завод"
        else if not Text.StartsWith([Роль], "Остатки") then ""
        else if [КатЗапаса] = "АТЗ" then "АТЗ"
        else if [КатСклада] = "5" then "Подрядчик"
        else if Text.Contains(Text.Lower([Склад]), "онсигнац") then "Консигнация"
        else if List.Contains({"Виртуальный", "Склад разниц"}, [Склад]) or [КатСклада] = "99" then "Прочее"
        else if Text.Contains(Text.Upper([БЕ]), "РАЗВИТ") then "Развитие"
        else "БЕ", type text)
in Table.Buffer(Table.SelectRows(b, each [Корзина] <> ""));

shared Остатки = Table.Buffer(Table.Group(Остатки_строки, {"Площадка", "Материал"},
    List.Transform({"БЕ", "Развитие", "Подрядчик", "Консигнация", "АТЗ", "Другой завод"}, (k) =>
        {"Ост " & k, (g) => фнСумма(Table.SelectRows(g, (r) => r[Корзина] = k)[Кол]), type number})));

shared ОстаткиПоНомеру = let
    t = Table.SelectRows(Остатки_строки, each ([Корзина] = "БЕ" or [Корзина] = "Развитие") and [НомерПроизв] <> ""),
    g = Table.Group(t, {"Площадка", "НомерПроизв"}, {{"Кол", each List.Sum([Кол]), type number}})
in Record.FromList(g[Кол], List.Transform(Table.ToRecords(g), (r) => r[Площадка] & "|" & r[НомерПроизв]));

// ---------- закупка АО «Развитие» ----------
shared Закупка_строки = let
    t = Table.PromoteHeaders(фнЛист(фнФайл("Закупка"), "Sheet1", true), [PromoteAllScalars = true]),
    need = {"Завод", "Документ закупки", "Описание", "еще поставить (количество)", "Количество в пути", "Количество",
        "Материал", "Код услуги/ЕК МТР", "Дата поставки по заказу", "Статистическая дата поставки", "Дата поставки", "Имя поставщика"},
    s = Table.SelectColumns(t, need, MissingField.UseNull),
    tx = Table.TransformColumns(s, {{"Завод", фнТекст}, {"Документ закупки", фнТекст}, {"Описание", фнТекст},
        {"еще поставить (количество)", фнЧисло}, {"Количество в пути", фнЧисло}, {"Количество", фнЧисло},
        {"Материал", фнТекст}, {"Код услуги/ЕК МТР", фнТекст}, {"Имя поставщика", фнТекст}}),
    code = Table.AddColumn(tx, "Код", each if [Материал] <> "" then [Материал] else [#"Код услуги/ЕК МТР"], type text),
    z = Table.ExpandTableColumn(Table.NestedJoin(code, {"Завод"}, Заводы, {"Код завода"}, "з", JoinKind.Inner), "з", {"Площадка", "Роль"}),
    rel = Table.SelectRows(z, each [Роль] = "Справочно" or Text.Contains([Роль], "закупка") or [Роль] = "Закупка"),
    firm = Table.AddColumn(rel, "Твёрдый", each [Документ закупки] <> "" and List.Contains({"Действующий", "Верифицирован", ""}, [Описание]), type logical),
    q = Table.AddColumn(firm, "К поставке", each
        if [Твёрдый] then List.Max({[#"еще поставить (количество)"], [Количество в пути]})
        else if [Документ закупки] = "" then [Количество]
        else List.Max({[#"еще поставить (количество)"], [Количество в пути], [Количество]}), type number),
    d = Table.AddColumn(q, "Дата", each
        if [Твёрдый] then
            let a = фнДата([Дата поставки по заказу]), b = фнДата([Статистическая дата поставки]), c = фнДата([Дата поставки])
            in if a <> null then a else if b <> null then b else c
        else фнДата([Дата поставки]), type nullable date),
    k = Table.AddColumn(d, "Корзина", each
        if [Роль] = "Справочно" then (if [Твёрдый] then "Другой завод" else "")
        else if [Твёрдый] then
            (if [Дата] = null or Date.Year([Дата]) < ГодПлана then "До года"
             else if Date.Year([Дата]) = ГодПлана then "М" & Text.From(Date.Month([Дата]))
             else "После")
        else (if [Дата] = null or Date.Year([Дата]) <= ГодПлана then "Неподтв" else ""), type text),
    ok = Table.SelectRows(k, each [Корзина] <> "" and [К поставке] > 0),
    late = Table.AddColumn(ok, "Просрочено", each [Корзина] = "До года"
        and ([Дата] = null or (ДатаЗакупки <> null and [Дата] < ДатаЗакупки)), type logical)
in Table.Buffer(late);

shared Закупка = let
    sumOf = (g as table, k as text) as number => фнСумма(Table.SelectRows(g, (r) => r[Корзина] = k)[К поставке]),
    own = (g as table) as table => Table.SelectRows(g, (r) => r[Твёрдый] and r[Роль] <> "Справочно"),
    aggs = {
        {"Пост до года", (g) => sumOf(g, "До года"), type number},
        {"Пост просрочено", (g) => фнСумма(Table.SelectRows(g, (r) => r[Просрочено])[К поставке]), type number}}
        & List.Transform({1..12}, (m) => {"Пост м" & Text.From(m), (g) => sumOf(g, "М" & Text.From(m)), type number})
        & {
        {"Пост после", (g) => sumOf(g, "После"), type number},
        {"Пост неподтв", (g) => sumOf(g, "Неподтв"), type number},
        {"Пост другой завод", (g) => sumOf(g, "Другой завод"), type number},
        {"Поставщики", (g) =>
            let s = own(g),
                gs = Table.Sort(Table.Group(s, {"Имя поставщика"}, {{"q", (x) => List.Sum(x[К поставке]), type number}}), {{"q", Order.Descending}})
            in Text.Combine(List.FirstN(gs[Имя поставщика], 2), "; "), type text},
        {"Заказов закупки", (g) => List.Count(List.Distinct(own(g)[Документ закупки])), Int64.Type}}
in Table.Buffer(Table.Group(Закупка_строки, {"Площадка", "Код"}, aggs));

// ---------- итоговая таблица позиций ----------
shared Позиции = let
    n = (v) => if v = null then 0 else v,
    ключи = Table.Distinct(Table.Combine({Table.SelectColumns(План, {"Площадка", "ЕКМТР"}), Table.SelectColumns(ВнеПлана, {"Площадка", "ЕКМТР"})})),
    планКол = List.RemoveItems(Table.ColumnNames(План), {"Площадка", "ЕКМТР"}),
    внеКол = List.RemoveItems(Table.ColumnNames(ВнеПлана), {"Площадка", "ЕКМТР"}),
    остКол = List.RemoveItems(Table.ColumnNames(Остатки), {"Площадка", "Материал"}),
    постКол = List.RemoveItems(Table.ColumnNames(Закупка), {"Площадка", "Код"}),
    j1 = Table.ExpandTableColumn(Table.NestedJoin(ключи, {"Площадка", "ЕКМТР"}, План, {"Площадка", "ЕКМТР"}, "x", JoinKind.LeftOuter), "x", планКол),
    j2 = Table.ExpandTableColumn(Table.NestedJoin(j1, {"Площадка", "ЕКМТР"}, ВнеПлана, {"Площадка", "ЕКМТР"}, "x", JoinKind.LeftOuter), "x", внеКол),
    j3 = Table.ExpandTableColumn(Table.NestedJoin(j2, {"ЕКМТР"}, ЕКМТР, {"Код"}, "x", JoinKind.LeftOuter), "x", {"Номер"}),
    j4 = Table.ExpandTableColumn(Table.NestedJoin(j3, {"Площадка", "ЕКМТР"}, Остатки, {"Площадка", "Материал"}, "x", JoinKind.LeftOuter), "x", остКол),
    j5 = Table.ExpandTableColumn(Table.NestedJoin(j4, {"Площадка", "ЕКМТР"}, Закупка, {"Площадка", "Код"}, "x", JoinKind.LeftOuter), "x", постКол),
    c1 = Table.AddColumn(j5, "Наименование", each if [Наименование план] <> null then [Наименование план] else [Наименование вне], type text),
    c2 = Table.AddColumn(c1, "Номер Cat", each фнТекст([Номер]), type text),
    c3 = Table.AddColumn(c2, "Модели", each Text.Combine(List.Sort(List.Distinct(
        (if [Модели план] = null then {} else [Модели план]) & (if [Модели вне] = null then {} else [Модели вне]))), ", "), type text),
    c4 = Table.AddColumn(c3, "Часть", each
        let pole = Number.Round(n([План поле]) + n([Вне поле]), 6),
            ceh = Number.Round(n([План цех]) + n([Вне шт]) - n([Вне поле]), 6)
        in if pole > 0 and ceh > 0 then "Поле+Цех" else if pole > 0 then "Поле" else if ceh > 0 then "Цех" else "", type text),
    c5 = Table.AddColumn(c4, "Признак", each
        let cap = Number.Round(n([План CAPEX]) + n([Вне CAPEX]), 6),
            op = Number.Round(n([План шт]) - n([План CAPEX]) + n([Вне шт]) - n([Вне CAPEX]), 6)
        in if cap > 0 and op > 0 then "CAPEX+OPEX" else if cap > 0 then "CAPEX" else if op > 0 then "OPEX" else "", type text),
    c6 = Table.AddColumn(c5, "Причина", each if [Причины] = null then ""
        else Text.Combine(List.FirstN(Table.Sort([Причины], {{"q", Order.Descending}})[Причина2], 2), "; "), type text),
    c7 = Table.AddColumn(c6, "Источник", each
        if n([План шт]) > 0 and n([Вне шт]) > 0 then "План + вне плана" else if n([План шт]) > 0 then "План" else "Вне плана", type text),
    c8 = Table.AddColumn(c7, "Виды работ", each
        let v = Table.Combine(List.RemoveNulls({[Виды план], [Виды вне]}))
        in if Table.IsEmpty(v) then ""
           else Text.Combine(List.FirstN(Table.Sort(Table.Group(v, {"ВидРабот"}, {{"q", (g) => List.Sum(g[q]), type number}}),
               {{"q", Order.Descending}})[ВидРабот], 3), ", "), type text),
    c9 = Table.AddColumn(c8, "Цена", each
        if n([План шт]) > 0 then n([План руб]) / [План шт] else if n([Вне шт]) > 0 then n([Вне руб]) / [Вне шт] else 0, type number),
    c10 = Table.AddColumn(Table.AddColumn(c9, "Доля поле вне", each if n([Вне шт]) > 0 then n([Вне поле]) / [Вне шт] else 0, type number),
        "Доля CAPEX вне", each if n([Вне шт]) > 0 then n([Вне CAPEX]) / [Вне шт] else 0, type number),
    c11 = Table.AddColumn(c10, "Семья", each фнСемьяЗамен(фнНомер([Номер Cat])), type list),
    c12 = Table.AddColumn(c11, "Номера-замены", each Text.Combine(List.FirstN([Семья], 4), " "), type text),
    c13 = Table.AddColumn(c12, "Ост замены", each
        List.Sum({0} & List.Transform([Семья], (x) => Record.FieldOrDefault(ОстаткиПоНомеру, [Площадка] & "|" & x, 0))), type number),
    c14 = Table.AddColumn(c13, "Ключ", each (n([План шт]) + n([Вне шт]) / 2.667) * [Цена], type number),
    byPlace = Table.Group(c14, {"Площадка"}, {{"t", each Table.AddIndexColumn(Table.Sort(_, {{"Ключ", Order.Descending}}), "№", 1, 1, Int64.Type), type table}}),
    all = Table.Combine(Table.Sort(byPlace, {{"Площадка", Order.Ascending}})[t]),
    zeros = Table.TransformColumns(all, List.Transform(List.Select(Table.ColumnNames(all), (c) =>
        List.Contains({"План шт", "План поле", "План цех", "План CAPEX", "План ВНЕ", "Заказов план", "Макс отбор план",
            "Заказов -3", "Заказов -2", "Заказов -1", "Шт -3", "Шт -2", "Шт -1", "Макс отбор вне",
            "Ост БЕ", "Ост Развитие", "Ост Подрядчик", "Ост Консигнация", "Ост АТЗ", "Ост Другой завод",
            "Пост до года", "Пост просрочено", "Пост после", "Пост неподтв", "Пост другой завод", "Заказов закупки"}, c)
        or Text.StartsWith(c, "П") and Text.Length(c) <= 3 or Text.StartsWith(c, "Пост м")), (c) => {c, n})),
    out = Table.SelectColumns(zeros, {"Площадка", "№", "Номер Cat", "ЕКМТР", "Наименование", "Модели", "Часть", "Признак", "Причина", "Источник", "Виды работ", "План шт", "План поле", "План цех", "План CAPEX", "План ВНЕ", "Заказов план", "Макс отбор план", "П1", "П2", "П3", "П4", "П5", "П6", "П7", "П8", "П9", "П10", "П11", "П12", "Заказов -3", "Заказов -2", "Заказов -1", "Шт -3", "Шт -2", "Шт -1", "Макс отбор вне", "Доля поле вне", "Доля CAPEX вне", "Цена", "Ост БЕ", "Ост Развитие", "Ост замены", "Номера-замены", "Ост Подрядчик", "Ост Консигнация", "Ост АТЗ", "Ост Другой завод", "Пост до года", "Пост просрочено", "Пост м1", "Пост м2", "Пост м3", "Пост м4", "Пост м5", "Пост м6", "Пост м7", "Пост м8", "Пост м9", "Пост м10", "Пост м11", "Пост м12", "Пост после", "Пост неподтв", "Поставщики", "Заказов закупки", "Пост другой завод"}),
    renamed = Table.RenameColumns(out, {{"ЕКМТР", "Код ЕК МТР"}, {"Часть", "Часть (поле / цех)"}, {"Признак", "Признак CAPEX / OPEX"}, {"Причина", "Причина инвестиций SAP (план)"}, {"Источник", "Источник спроса"}, {"План шт", "План, шт"}, {"План поле", "в т.ч. поле, шт"}, {"План цех", "в т.ч. цех, шт"}, {"План CAPEX", "в т.ч. CAPEX, шт"}, {"План ВНЕ", "в т.ч. внеплановые виды работ в плане, шт"}, {"Заказов план", "Заказов в плане"}, {"Макс отбор план", "Макс. разовый отбор в плане, шт"}, {"П1", "План янв, шт"}, {"П2", "План фев, шт"}, {"П3", "План мар, шт"}, {"П4", "План апр, шт"}, {"П5", "План май, шт"}, {"П6", "План июн, шт"}, {"П7", "План июл, шт"}, {"П8", "План авг, шт"}, {"П9", "План сен, шт"}, {"П10", "План окт, шт"}, {"П11", "План ноя, шт"}, {"П12", "План дек, шт"}, {"Заказов -3", "Вне плана: заказов (год −3)"}, {"Заказов -2", "Вне плана: заказов (год −2)"}, {"Заказов -1", "Вне плана: заказов (год −1)"}, {"Шт -3", "Вне плана: шт (год −3)"}, {"Шт -2", "Вне плана: шт (год −2)"}, {"Шт -1", "Вне плана: шт (год −1)"}, {"Макс отбор вне", "Вне плана: макс. разовый отбор, шт"}, {"Доля поле вне", "Вне плана: доля поле"}, {"Доля CAPEX вне", "Вне плана: доля CAPEX"}, {"Цена", "Цена, руб/шт"}, {"Ост БЕ", "Остаток БЕ площадки, шт"}, {"Ост Развитие", "Остаток АО «Развитие» без подрядчиков, шт"}, {"Ост замены", "Остаток по заменам Cat, шт"}, {"Ост Подрядчик", "Справочно: на складах подрядчиков, шт"}, {"Ост Консигнация", "Справочно: действующая консигнация, шт"}, {"Ост АТЗ", "Справочно: АТЗ, шт"}, {"Ост Другой завод", "Справочно: остаток на других заводах региона, шт"}, {"Пост до года", "Поставки: до года плана и просроченные, шт"}, {"Пост просрочено", "в т.ч. просрочено на дату выгрузки, шт"}, {"Пост м1", "Поставка янв, шт"}, {"Пост м2", "Поставка фев, шт"}, {"Пост м3", "Поставка мар, шт"}, {"Пост м4", "Поставка апр, шт"}, {"Пост м5", "Поставка май, шт"}, {"Пост м6", "Поставка июн, шт"}, {"Пост м7", "Поставка июл, шт"}, {"Пост м8", "Поставка авг, шт"}, {"Пост м9", "Поставка сен, шт"}, {"Пост м10", "Поставка окт, шт"}, {"Пост м11", "Поставка ноя, шт"}, {"Пост м12", "Поставка дек, шт"}, {"Пост после", "Поставки после года плана (справочно), шт"}, {"Пост неподтв", "Неподтверждённые: заявки и заказы на согласовании, шт"}, {"Поставщики", "Поставщики по заказам"}, {"Пост другой завод", "Справочно: заказы на другие заводы региона, шт"}}),
    typed = Table.TransformColumnTypes(renamed, {{"Площадка", type text}, {"№", Int64.Type}, {"Номер Cat", type text}, {"Код ЕК МТР", type text}, {"Наименование", type text}, {"Модели", type text}, {"Часть (поле / цех)", type text}, {"Признак CAPEX / OPEX", type text}, {"Причина инвестиций SAP (план)", type text}, {"Источник спроса", type text}, {"Виды работ", type text}, {"План, шт", type number}, {"в т.ч. поле, шт", type number}, {"в т.ч. цех, шт", type number}, {"в т.ч. CAPEX, шт", type number}, {"в т.ч. внеплановые виды работ в плане, шт", type number}, {"Заказов в плане", type number}, {"Макс. разовый отбор в плане, шт", type number}, {"План янв, шт", type number}, {"План фев, шт", type number}, {"План мар, шт", type number}, {"План апр, шт", type number}, {"План май, шт", type number}, {"План июн, шт", type number}, {"План июл, шт", type number}, {"План авг, шт", type number}, {"План сен, шт", type number}, {"План окт, шт", type number}, {"План ноя, шт", type number}, {"План дек, шт", type number}, {"Вне плана: заказов (год −3)", type number}, {"Вне плана: заказов (год −2)", type number}, {"Вне плана: заказов (год −1)", type number}, {"Вне плана: шт (год −3)", type number}, {"Вне плана: шт (год −2)", type number}, {"Вне плана: шт (год −1)", type number}, {"Вне плана: макс. разовый отбор, шт", type number}, {"Вне плана: доля поле", type number}, {"Вне плана: доля CAPEX", type number}, {"Цена, руб/шт", type number}, {"Остаток БЕ площадки, шт", type number}, {"Остаток АО «Развитие» без подрядчиков, шт", type number}, {"Остаток по заменам Cat, шт", type number}, {"Номера-замены", type text}, {"Справочно: на складах подрядчиков, шт", type number}, {"Справочно: действующая консигнация, шт", type number}, {"Справочно: АТЗ, шт", type number}, {"Справочно: остаток на других заводах региона, шт", type number}, {"Поставки: до года плана и просроченные, шт", type number}, {"в т.ч. просрочено на дату выгрузки, шт", type number}, {"Поставка янв, шт", type number}, {"Поставка фев, шт", type number}, {"Поставка мар, шт", type number}, {"Поставка апр, шт", type number}, {"Поставка май, шт", type number}, {"Поставка июн, шт", type number}, {"Поставка июл, шт", type number}, {"Поставка авг, шт", type number}, {"Поставка сен, шт", type number}, {"Поставка окт, шт", type number}, {"Поставка ноя, шт", type number}, {"Поставка дек, шт", type number}, {"Поставки после года плана (справочно), шт", type number}, {"Неподтверждённые: заявки и заказы на согласовании, шт", type number}, {"Поставщики по заказам", type text}, {"Заказов закупки", Int64.Type}, {"Справочно: заказы на другие заводы региона, шт", type number}})
in typed;
