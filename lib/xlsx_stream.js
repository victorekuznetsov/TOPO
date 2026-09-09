/* Потоковое чтение больших .xlsx прямо в браузере.

   Зачем своё, а не SheetJS: SheetJS распаковывает лист целиком в одну
   JS-строку, а у движка V8 предел длины строки ~512 МБ. Внутренний XML
   листа PM-06 площадки 1100 занимает 600–800 МБ — библиотека падает с
   «Invalid array length» ещё до разбора. Здесь лист читается потоком
   (нативный DecompressionStream('deflate-raw')), в памяти держится
   только текущий кусок — размер файла перестаёт иметь значение.

   Поддерживается обычный ZIP со сжатием deflate (метод 8) и без
   сжатия (метод 0) — так пишут и SAP BW, и Excel. ZIP64 (файлы >4 ГБ)
   не поддерживается осознанно: выгрузки PM-06 на порядок меньше, а
   молчаливая порча данных хуже честной ошибки. */

const XLSXStream = (() => {
  "use strict";

  const td = new TextDecoder("utf-8");

  function u16(dv, o) { return dv.getUint16(o, true); }
  function u32(dv, o) { return dv.getUint32(o, true); }

  async function sliceBuf(file, start, end) {
    return new Uint8Array(await file.slice(start, end).arrayBuffer());
  }

  /* --- каталог zip: последние 64 КБ → EOCD → центральный каталог --- */
  async function openZip(file) {
    const size = file.size;
    const tailLen = Math.min(65557, size);
    const tail = await sliceBuf(file, size - tailLen, size);
    const dv = new DataView(tail.buffer);
    let eocd = -1;
    for (let i = tail.length - 22; i >= 0; i--) {
      if (tail[i] === 0x50 && tail[i + 1] === 0x4b && tail[i + 2] === 0x05 && tail[i + 3] === 0x06) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error("Это не .xlsx: не найден каталог zip-архива");
    const nEntries = u16(dv, eocd + 10);
    const cdSize = u32(dv, eocd + 12);
    const cdOff = u32(dv, eocd + 16);
    if (cdOff === 0xffffffff || nEntries === 0xffff) throw new Error("Архив в формате ZIP64 не поддерживается");

    const cd = await sliceBuf(file, cdOff, cdOff + cdSize);
    const cdv = new DataView(cd.buffer);
    const entries = new Map();
    let p = 0;
    for (let i = 0; i < nEntries && p + 46 <= cd.length; i++) {
      if (u32(cdv, p) !== 0x02014b50) break;
      const method = u16(cdv, p + 10);
      const compSize = u32(cdv, p + 20);
      const rawSize = u32(cdv, p + 24);
      const nameLen = u16(cdv, p + 28);
      const extraLen = u16(cdv, p + 30);
      const cmtLen = u16(cdv, p + 32);
      const localOff = u32(cdv, p + 42);
      const name = td.decode(cd.subarray(p + 46, p + 46 + nameLen));
      entries.set(name, { method, compSize, rawSize, localOff });
      p += 46 + nameLen + extraLen + cmtLen;
    }
    return { file, entries };
  }

  /* Смещение самих данных: локальный заголовок объявляет свои длины
     имени и extra-поля, они могут отличаться от центрального каталога. */
  async function dataStart(zip, e) {
    const h = await sliceBuf(zip.file, e.localOff, e.localOff + 30);
    const hv = new DataView(h.buffer);
    if (u32(hv, 0) !== 0x04034b50) throw new Error("Повреждён локальный заголовок zip");
    return e.localOff + 30 + u16(hv, 26) + u16(hv, 28);
  }

  function inflateStream(blob, method) {
    if (method === 0) return blob.stream();
    if (method !== 8) throw new Error("Неподдерживаемый метод сжатия zip: " + method);
    return blob.stream().pipeThrough(new DecompressionStream("deflate-raw"));
  }

  async function entryBlob(zip, name) {
    const e = zip.entries.get(name);
    if (!e) throw new Error("В архиве нет " + name);
    const start = await dataStart(zip, e);
    return { blob: zip.file.slice(start, start + e.compSize), method: e.method, rawSize: e.rawSize };
  }

  async function readText(zip, name) {
    const { blob, method } = await entryBlob(zip, name);
    return await new Response(inflateStream(blob, method)).text();
  }

  /* --- XML: сущности и мелкие помощники --- */
  const ENT = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
  function unesc(s) {
    if (s.indexOf("&") < 0) return s;
    return s.replace(/&(#x?[0-9a-fA-F]+|[a-z]+);/g, (m, g) => {
      if (g[0] === "#") return String.fromCodePoint(parseInt(g[1] === "x" || g[1] === "X" ? g.slice(2) : g.slice(1), g[1] === "x" || g[1] === "X" ? 16 : 10));
      return ENT[g] !== undefined ? ENT[g] : m;
    });
  }
  function attr(tag, name) {
    const k = name + '="';
    const i = tag.indexOf(k);
    if (i < 0) return null;
    const j = tag.indexOf('"', i + k.length);
    return j < 0 ? null : tag.slice(i + k.length, j);
  }
  function colOf(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i);
      if (c < 65 || c > 90) break;
      n = n * 26 + (c - 64);
    }
    return n - 1;
  }

  /* --- словарь общих строк (у PM-06 это единицы мегабайт) --- */
  async function readSharedStrings(zip) {
    if (!zip.entries.has("xl/sharedStrings.xml")) return [];
    const xml = await readText(zip, "xl/sharedStrings.xml");
    const out = [];
    let p = 0;
    for (;;) {
      const si = xml.indexOf("<si", p);
      if (si < 0) break;
      let end = xml.indexOf("</si>", si);
      if (end < 0) { end = xml.length; }
      const chunk = xml.slice(si, end);
      // <si> может состоять из нескольких <t> (форматированные куски)
      let s = "", q = 0;
      for (;;) {
        const t = chunk.indexOf("<t", q);
        if (t < 0) break;
        const gt = chunk.indexOf(">", t);
        if (gt < 0) break;
        if (chunk[gt - 1] === "/") { q = gt + 1; continue; }
        const te = chunk.indexOf("</t>", gt);
        if (te < 0) break;
        s += chunk.slice(gt + 1, te);
        q = te + 4;
      }
      out.push(unesc(s));
      p = end + 5;
    }
    return out;
  }

  /* --- какой лист соответствует имени вкладки --- */
  async function sheetPath(zip, sheetName) {
    const wb = await readText(zip, "xl/workbook.xml");
    const rels = await readText(zip, "xl/_rels/workbook.xml.rels");
    const re = /<sheet\b[^>]*\/?>/g;
    let m, rid = null;
    while ((m = re.exec(wb))) {
      if (unesc(attr(m[0], "name") || "") === sheetName) { rid = attr(m[0], "r:id") || attr(m[0], "id"); break; }
    }
    if (!rid) throw new Error(`В книге нет листа «${sheetName}»`);
    const rre = /<Relationship\b[^>]*\/?>/g;
    while ((m = rre.exec(rels))) {
      if (attr(m[0], "Id") === rid) {
        let t = attr(m[0], "Target") || "";
        t = t.replace(/^\/?xl\//, "").replace(/^\.\//, "");
        return "xl/" + t;
      }
    }
    throw new Error("Не найден файл листа " + sheetName);
  }

  /* Читает лист построчно. onRow(cells, rowNumber) — cells это разрежённый
     массив значений (строки/числа), индекс = номер колонки с нуля.
     onProgress({done, total}) — по распакованным байтам. */
  async function streamSheet(zip, sheetName, onRow, onProgress) {
    const shared = await readSharedStrings(zip);
    const path = await sheetPath(zip, sheetName);
    const { blob, method, rawSize } = await entryBlob(zip, path);
    const reader = inflateStream(blob, method).getReader();
    const dec = new TextDecoder("utf-8");

    let buf = "";
    let done = 0;
    let nRows = 0;
    let lastTick = 0;

    const flushRows = (final) => {
      let pos = 0;
      for (;;) {
        const rs = buf.indexOf("<row", pos);
        if (rs < 0) break;
        const tagEnd = buf.indexOf(">", rs);
        if (tagEnd < 0) break;
        const rowTag = buf.slice(rs, tagEnd + 1);
        let body, next;
        if (rowTag.endsWith("/>")) {           // пустая строка <row .../>
          body = ""; next = tagEnd + 1;
        } else {
          const re2 = buf.indexOf("</row>", tagEnd);
          if (re2 < 0) break;                  // строка не дочитана
          body = buf.slice(tagEnd + 1, re2); next = re2 + 6;
        }
        const rowNum = +(attr(rowTag, "r") || ++nRows);
        onRow(parseRow(body, shared), rowNum);
        nRows++;
        pos = next;
      }
      buf = pos ? buf.slice(pos) : buf;
      if (final && buf.length > 1e7) buf = "";
    };

    for (;;) {
      const { value, done: fin } = await reader.read();
      if (fin) break;
      done += value.length;
      buf += dec.decode(value, { stream: true });
      flushRows(false);
      if (onProgress && done - lastTick > 8e6) { lastTick = done; onProgress({ done, total: rawSize, rows: nRows }); }
    }
    buf += dec.decode();
    flushRows(true);
    if (onProgress) onProgress({ done: rawSize, total: rawSize, rows: nRows });
    return nRows;
  }

  function parseRow(body, shared) {
    const cells = [];
    let p = 0;
    for (;;) {
      const cs = body.indexOf("<c", p);
      if (cs < 0) break;
      const ce = body.indexOf(">", cs);
      if (ce < 0) break;
      const tag = body.slice(cs, ce + 1);
      const selfClosed = tag.endsWith("/>");
      const ref = attr(tag, "r");
      const ci = ref ? colOf(ref) : cells.length;
      let val = null;
      let next = ce + 1;
      if (!selfClosed) {
        const cEnd = body.indexOf("</c>", ce);
        const inner = body.slice(ce + 1, cEnd < 0 ? body.length : cEnd);
        next = cEnd < 0 ? body.length : cEnd + 4;
        const t = attr(tag, "t");
        if (t === "inlineStr") {
          const ts = inner.indexOf("<t");
          if (ts >= 0) {
            const tg = inner.indexOf(">", ts), te = inner.indexOf("</t>", tg);
            if (tg >= 0 && te >= 0) val = unesc(inner.slice(tg + 1, te));
          }
        } else {
          const vs = inner.indexOf("<v>");
          if (vs >= 0) {
            const ve = inner.indexOf("</v>", vs);
            const raw = inner.slice(vs + 3, ve < 0 ? inner.length : ve);
            if (t === "s") { const idx = +raw; val = shared[idx] !== undefined ? shared[idx] : ""; }
            else if (t === "str") val = unesc(raw);
            else if (t === "b") val = raw === "1";
            else if (t === "e") val = null;
            else { const n = +raw; val = raw === "" ? null : (Number.isNaN(n) ? unesc(raw) : n); }
          }
        }
      }
      if (ci >= 0) cells[ci] = val;
      p = next;
    }
    return cells;
  }

  return { openZip, readText, streamSheet, readSharedStrings, sheetPath };
})();

if (typeof module !== "undefined" && module.exports) module.exports = XLSXStream;
