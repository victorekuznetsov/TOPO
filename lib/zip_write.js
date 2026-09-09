/* Минимальная сборка .zip в браузере — чтобы отдать обновлённые данные
   одним файлом. Сжатие нативное (CompressionStream), без библиотек:
   JSON ужимается примерно в десять раз, иначе выгрузка весила бы
   сотни мегабайт. Пишем обычный zip (не ZIP64) — этого достаточно,
   пока итог меньше 4 ГБ; при превышении честно сообщаем об ошибке. */

const ZipWrite = (() => {
  "use strict";

  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(chunks) {
    let c = 0xffffffff;
    for (const b of chunks) {
      for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8);
    }
    return (c ^ 0xffffffff) >>> 0;
  }

  async function deflateRaw(blob) {
    const cs = new CompressionStream("deflate-raw");
    return new Uint8Array(await new Response(blob.stream().pipeThrough(cs)).arrayBuffer());
  }

  function dosTime(d) {
    const t = ((d.getHours() & 31) << 11) | ((d.getMinutes() & 63) << 5) | ((d.getSeconds() / 2) & 31);
    const dt = (((d.getFullYear() - 1980) & 127) << 9) | (((d.getMonth() + 1) & 15) << 5) | (d.getDate() & 31);
    return { t, dt };
  }

  function put(dv, o, v, n) {
    if (n === 2) dv.setUint16(o, v, true); else dv.setUint32(o, v, true);
  }

  /* files: [{name, text}] — имена с прямыми слэшами, как пути в архиве.
     onProgress(name, i, total) вызывается перед упаковкой каждого файла. */
  async function build(files, onProgress) {
    const enc = new TextEncoder();
    const parts = [];
    const central = [];
    let offset = 0;
    const now = new Date();
    const { t, dt } = dosTime(now);

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (onProgress) await onProgress(f.name, i, files.length);
      const nameBytes = enc.encode(f.name);
      const blob = f.blob || new Blob([f.text]);
      const rawChunks = [new Uint8Array(await blob.arrayBuffer())];
      const rawSize = rawChunks[0].length;
      const crc = crc32(rawChunks);
      const comp = await deflateRaw(new Blob(rawChunks));

      const lh = new Uint8Array(30 + nameBytes.length);
      const ldv = new DataView(lh.buffer);
      put(ldv, 0, 0x04034b50, 4); put(ldv, 4, 20, 2); put(ldv, 6, 0x0800, 2);
      put(ldv, 8, 8, 2); put(ldv, 10, t, 2); put(ldv, 12, dt, 2);
      put(ldv, 14, crc, 4); put(ldv, 18, comp.length, 4); put(ldv, 22, rawSize, 4);
      put(ldv, 26, nameBytes.length, 2); put(ldv, 28, 0, 2);
      lh.set(nameBytes, 30);
      parts.push(lh, comp);

      const ch = new Uint8Array(46 + nameBytes.length);
      const cdv = new DataView(ch.buffer);
      put(cdv, 0, 0x02014b50, 4); put(cdv, 4, 20, 2); put(cdv, 6, 20, 2);
      put(cdv, 8, 0x0800, 2); put(cdv, 10, 8, 2); put(cdv, 12, t, 2); put(cdv, 14, dt, 2);
      put(cdv, 16, crc, 4); put(cdv, 20, comp.length, 4); put(cdv, 24, rawSize, 4);
      put(cdv, 28, nameBytes.length, 2); put(cdv, 30, 0, 2); put(cdv, 32, 0, 2);
      put(cdv, 34, 0, 2); put(cdv, 36, 0, 2); put(cdv, 38, 0, 4);
      put(cdv, 42, offset, 4);
      ch.set(nameBytes, 46);
      central.push(ch);

      offset += lh.length + comp.length;
      if (offset > 0xfffffff0) throw new Error("Архив вышел за 4 ГБ — выгрузите файлы по частям");
    }

    const cdStart = offset;
    let cdSize = 0;
    for (const ch of central) { parts.push(ch); cdSize += ch.length; }

    const eocd = new Uint8Array(22);
    const edv = new DataView(eocd.buffer);
    put(edv, 0, 0x06054b50, 4); put(edv, 4, 0, 2); put(edv, 6, 0, 2);
    put(edv, 8, central.length, 2); put(edv, 10, central.length, 2);
    put(edv, 12, cdSize, 4); put(edv, 16, cdStart, 4); put(edv, 20, 0, 2);
    parts.push(eocd);

    return new Blob(parts, { type: "application/zip" });
  }

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }

  return { build, download, crc32 };
})();

if (typeof module !== "undefined" && module.exports) module.exports = ZipWrite;
