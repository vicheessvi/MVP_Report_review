(function () {
  'use strict';

  const decoder = new TextDecoder('utf-8', { fatal: true });
  const parser = new DOMParser();
  const maxXmlSize = 8 * 1024 * 1024;
  const maxTotalXmlSize = 48 * 1024 * 1024;

  function archiveError() {
    return new Error('Не удалось прочитать структуру PPTX. Откройте файл в PowerPoint и сохраните его заново.');
  }

  function indexArchive(buffer) {
    const view = new DataView(buffer);
    let end = view.byteLength - 22;
    for (; end >= Math.max(0, view.byteLength - 65557); end--) {
      if (view.getUint32(end, true) === 0x06054b50 && end + 22 + view.getUint16(end + 20, true) === view.byteLength) break;
    }
    if (end < Math.max(0, view.byteLength - 65557)) throw archiveError();
    const count = view.getUint16(end + 10, true);
    const start = view.getUint32(end + 16, true);
    const directoryEnd = start + view.getUint32(end + 12, true);
    if (view.getUint32(end + 4, true) || view.getUint16(end + 8, true) !== count || directoryEnd > end) throw archiveError();

    const entries = new Map();
    let offset = start;
    for (let i = 0; i < count; i++) {
      if (offset + 46 > directoryEnd || view.getUint32(offset, true) !== 0x02014b50) throw archiveError();
      const nameSize = view.getUint16(offset + 28, true);
      const next = offset + 46 + nameSize + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
      if (next > directoryEnd) throw archiveError();
      const name = new TextDecoder().decode(new Uint8Array(buffer, offset + 46, nameSize));
      entries.set(name, {
        flags: view.getUint16(offset + 8, true),
        method: view.getUint16(offset + 10, true),
        compressedSize: view.getUint32(offset + 20, true),
        size: view.getUint32(offset + 24, true),
        offset: view.getUint32(offset + 42, true)
      });
      offset = next;
    }

    let totalSize = 0;
    return async function readXml(name) {
      const entry = entries.get(name);
      if (!entry || entry.flags & 1 || entry.offset + 30 > start || view.getUint32(entry.offset, true) !== 0x04034b50) throw archiveError();
      totalSize += entry.size;
      if (entry.size > maxXmlSize || totalSize > maxTotalXmlSize) throw new Error('В презентации слишком много текстовых данных. Разделите её на несколько файлов.');
      const dataOffset = entry.offset + 30 + view.getUint16(entry.offset + 26, true) + view.getUint16(entry.offset + 28, true);
      if (dataOffset + entry.compressedSize > start) throw archiveError();
      const compressed = new Uint8Array(buffer, dataOffset, entry.compressedSize);
      let data = compressed;
      if (entry.method === 8) {
        let inflater;
        try {
          inflater = new DecompressionStream('deflate-raw');
        } catch {
          throw new Error('Для открытия PPTX нужен актуальный Chrome, Edge, Firefox или Safari.');
        }
        const reader = new Blob([compressed]).stream().pipeThrough(inflater).getReader();
        data = new Uint8Array(entry.size);
        let written = 0;
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;
          if (written + value.length > data.length) {
            await reader.cancel();
            throw archiveError();
          }
          data.set(value, written);
          written += value.length;
        }
        if (written !== data.length) throw archiveError();
      } else if (entry.method !== 0 || data.length !== entry.size) {
        throw archiveError();
      }
      let xml;
      try {
        const utf16 = data[0] === 0xff && data[1] === 0xfe ? 'utf-16le' : data[0] === 0xfe && data[1] === 0xff ? 'utf-16be' : '';
        xml = (utf16 ? new TextDecoder(utf16, { fatal: true }) : decoder).decode(data);
      } catch {
        throw archiveError();
      }
      if (/<!DOCTYPE/i.test(xml)) throw archiveError();
      const doc = parser.parseFromString(xml, 'application/xml');
      if (doc.getElementsByTagNameNS('*', 'parsererror').length) throw archiveError();
      return doc.documentElement;
    };
  }

  function children(node, name) {
    return Array.from(node?.children || []).filter(child => child.localName === name);
  }

  function first(node, name) {
    return children(node, name)[0];
  }

  function resolvePart(base, target) {
    const url = new URL(target, 'https://pptx.invalid/' + base);
    if (url.origin !== 'https://pptx.invalid' || url.search || url.hash) throw archiveError();
    return decodeURIComponent(url.pathname.slice(1));
  }

  function readParagraph(node) {
    return Array.from(node.children, child => {
      if (child.localName === 'br') return '\n';
      if (child.localName === 'tab') return '\t';
      if (['r', 'fld'].includes(child.localName)) return first(child, 't')?.textContent || '';
      return '';
    }).join('').trim();
  }

  function readText(body) {
    return children(body, 'p').flatMap(p => readParagraph(p).split('\n')).map(text => text.trim()).filter(Boolean);
  }

  function makeSlide(root, number) {
    const tree = first(first(root, 'cSld'), 'spTree');
    if (root.localName !== 'sld' || !tree) throw archiveError();
    const blocks = [];
    const tables = [];
    let title = '';
    let titleBlock;

    function visit(container) {
      for (const shape of container.children) {
        if (shape.localName === 'grpSp' || shape.localName === 'Choice' || shape.localName === 'Fallback') {
          visit(shape);
        } else if (shape.localName === 'AlternateContent') {
          const branch = first(shape, 'Choice') || first(shape, 'Fallback');
          if (branch) visit(branch);
        } else if (shape.localName === 'sp' || shape.localName === 'cxnSp') {
          const lines = readText(first(shape, 'txBody'));
          if (!lines.length) continue;
          const properties = first(first(shape, 'nvSpPr'), 'nvPr');
          const type = first(properties, 'ph')?.getAttribute('type');
          const block = { lines };
          if (!title && (type === 'title' || type === 'ctrTitle')) {
            title = lines.join(' ');
            titleBlock = block;
          }
          blocks.push(block);
        } else if (shape.localName === 'graphicFrame') {
          const table = first(first(first(shape, 'graphic'), 'graphicData'), 'tbl');
          if (!table) continue;
          const rows = children(table, 'tr').map(row => children(row, 'tc').map(cell => {
            const merged = ['hMerge', 'vMerge'].some(key => ['1', 'true'].includes(cell.getAttribute(key)));
            return merged ? '' : readText(first(cell, 'txBody')).join(' ');
          }));
          tables.push(rows);
        }
      }
    }

    visit(tree);
    if (!title && blocks.length && blocks[0].lines.length === 1 && blocks[0].lines[0].length <= 100) {
      title = blocks[0].lines[0];
    }
    return { number, title, blocks, tables, titleBlock };
  }

  async function read(file) {
    if (!/\.pptx$/i.test(file.name)) throw new Error('Выберите файл с расширением .pptx.');
    if (file.size > 50 * 1024 * 1024) throw new Error('Максимальный размер одного файла — 50 МБ.');
    try {
      const readXml = indexArchive(await file.arrayBuffer());
      const packageRelationships = await readXml('_rels/.rels');
      const office = children(packageRelationships, 'Relationship').find(rel => rel.getAttribute('Type')?.endsWith('/officeDocument'));
      if (!office || office.getAttribute('TargetMode') === 'External') throw archiveError();
      const part = resolvePart('', office.getAttribute('Target'));
      const presentation = await readXml(part);
      if (presentation.localName !== 'presentation') throw archiveError();
      const order = children(first(presentation, 'sldIdLst'), 'sldId');
      if (!order.length) throw new Error('В презентации нет слайдов.');
      if (order.length > 1000) throw new Error('Можно проверить до 1000 слайдов за один раз.');
      const separator = part.lastIndexOf('/') + 1;
      const relationships = await readXml(part.slice(0, separator) + '_rels/' + part.slice(separator) + '.rels');
      const slidesById = new Map(children(relationships, 'Relationship')
        .filter(rel => rel.getAttribute('Type')?.endsWith('/slide') && rel.getAttribute('TargetMode') !== 'External')
        .map(rel => [rel.getAttribute('Id'), rel.getAttribute('Target')]));
      const slides = [];
      for (const [index, entry] of order.entries()) {
        const relId = Array.from(entry.attributes).find(attr => attr.localName === 'id' && attr.namespaceURI?.endsWith('/relationships'));
        const target = slidesById.get(relId?.value);
        if (!target) throw archiveError();
        slides.push(makeSlide(await readXml(resolvePart(part, target)), index + 1));
      }
      return slides;
    } catch (error) {
      if (error instanceof RangeError || error instanceof TypeError || error instanceof DOMException) throw archiveError();
      throw error;
    }
  }

  window.Pptx = { read };
})();
