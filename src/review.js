(function () {
  'use strict';

  const monthNames = [
    'января|январь|янв', 'февраля|февраль|фев', 'марта|март|мар',
    'апреля|апрель|апр', 'мая|май', 'июня|июнь|июн', 'июля|июль|июл',
    'августа|август|авг', 'сентября|сентябрь|сент|сен', 'октября|октябрь|окт',
    'ноября|ноябрь|ноя', 'декабря|декабрь|дек'
  ];
  const months = new Map(monthNames.flatMap((names, i) => names.split('|').map(name => [name, i + 1])));
  const writtenMonth = [...months.keys()].join('|');
  const datePattern = new RegExp(
    '(?<![\\p{L}\\p{N}])(?:' +
    '(?<isoYear>\\d{4})(?<isoSep>[-./])(?<isoMonth>\\d{1,2})\\k<isoSep>(?<isoDay>\\d{1,2})' +
    '|(?<day>\\d{1,2})(?<sep>[./-])(?<month>\\d{1,2})\\k<sep>(?<year>\\d{4}|\\d{2})' +
    '|(?<wordDay>\\d{1,2})\\s+(?<wordMonth>' + writtenMonth + ')(?![а-яё])(?:\\.?\\s+(?<wordYear>\\d{4}|\\d{2})(?!\\d))?' +
    ')(?!\\d|[./-]\\d)', 'giu'
  );
  const shortDatePattern = /(?<![\p{L}\p{N}./-])(\d{1,2})[./](\d{1,2})(?!\d|[./-]\d)/gu;
  const dateContext = /(?:^|[^\p{L}])(?:кс|срок(?:и|а)?|дата(?:ы)?|дедлайн)(?=$|[^\p{L}])/iu;
  const idHeader = /^(?:id|идентификатор|код|№|n\s*o\.?|номер|№\s*(?:п\/п|пункта|задачи))$/iu;
  const titleHeader = /^(?:наименование(?:\s+(?:мероприятия|пункта|задачи|работы))?|описание|мероприятие|пункт|задача|вопрос|действие|работа|тема|название|содержание|поручение)$/iu;
  const statusHeader = /^(?:статус(?:\s+(?:исполнения|выполнения))?|состояние|ход\s+выполнения)$/iu;
  const deadlineHeader = /^(?:кс|контрольный\s+срок(?:\s*\(кс\))?|срок(?:\s+выполнения)?|дата(?:\s+выполнения)?|дедлайн)$/iu;
  const statusLabel = /^(?:статус(?:\s+(?:исполнения|выполнения))?|состояние)\s*[:–—-]\s*/iu;
  const completedStatus = /^(?:выполнено|выполнен[аы]?|завершено|заверш[её]н[аы]?|исполнено|исполнен[аы]?|закрыто|закрыт[аы]?)(?=$|[\s.,;:–—()])/iu;
  const otherStatus = /^(?:не\s+выполнено|в\s+работе|в\s+процессе|выполняется|открыто|планируется|просрочено|отложено|отменено|на\s+согласовании|не\s+начато)(?=$|[\s.,;:–—()])/iu;
  const metadata = /^(?:(?:контрольный\s+)?срок|кс|дата|ответственный|исполнитель|примечание|комментарий|прогресс|статус)\s*[:–—]/iu;
  const idLabel = /^(?:id|код|№|идентификатор|номер(?:\s+(?:пункта|задачи))?)\s*[:№]?\s*([\p{L}\d][\p{L}\d._/-]*)(?=$|\s)/iu;
  const itemLabel = /^(?:наименование|пункт|задача|мероприятие|поручение|описание)\s*[:–—]\s*/iu;

  function dateValue(year, month, day) {
    if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null;
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCMonth() === month - 1 && date.getUTCDate() === day ? year * 10000 + month * 100 + day : null;
  }

  function expandYear(value, fallback) {
    if (!value) return fallback;
    return value.length === 2 ? 2000 + Number(value) : Number(value);
  }

  function extractDates(text, year, allowShort = false) {
    const dates = [];
    for (const match of text.matchAll(datePattern)) {
      const group = match.groups;
      const explicitYear = group.isoYear || group.year || group.wordYear;
      const month = Number(group.isoMonth || group.month) || months.get(group.wordMonth?.toLocaleLowerCase('ru'));
      const day = Number(group.isoDay || group.day || group.wordDay);
      dates.push({ raw: match[0], start: match.index, value: dateValue(expandYear(explicitYear, year), month || 0, day), assumed: !explicitYear });
    }
    if (allowShort || dateContext.test(text)) {
      for (const match of text.matchAll(shortDatePattern)) {
        dates.push({ raw: match[0], start: match.index, value: dateValue(year, Number(match[2]), Number(match[1])), assumed: true });
      }
    }
    return dates.sort((a, b) => a.start - b.start);
  }

  function clean(text) {
    return text.normalize('NFKC').toLocaleLowerCase('ru').replaceAll('ё', 'е').replace(/[^\p{L}\p{N}]+/gu, ' ').trim();
  }

  function status(text) {
    const value = text.trim().replace(/^[✓✔✅☑●•]\s*/u, '').replace(statusLabel, '');
    if (completedStatus.test(value)) {
      const rest = value.replace(completedStatus, '').trim();
      if (/^[\p{L}]/u.test(rest) && !/^\d+\s/u.test(rest)) return null;
      return 'done';
    }
    if (otherStatus.test(value)) return 'open';
    return null;
  }

  function stripTitle(text) {
    return text.trim().replace(/^[•●▪]\s*/u, '').replace(itemLabel, '')
      .replace(/^(?:№\s*)?\d{1,4}[.)]\s+/u, '').trim();
  }

  function makeItem(rawTitle, id, state, slide, source, excerpt, row) {
    const title = stripTitle(rawTitle || '');
    const key = clean(title);
    if (!id && (!/\p{L}/u.test(key) || key.length < 3)) return null;
    return { title: title || id, key, id: id ? id.normalize('NFKC').toLocaleUpperCase('ru').trim() : '', state, slide, source, excerpt, row };
  }

  function headerFor(row) {
    const names = row.map(text => text.trim().replace(/[:.]+$/, ''));
    const id = names.findIndex(name => idHeader.test(name));
    const title = names.findIndex(name => titleHeader.test(name));
    const state = names.findIndex(name => statusHeader.test(name));
    const deadlines = names.map(name => deadlineHeader.test(name));
    if (state === -1 && !(title !== -1 && (id !== -1 || deadlines.some(Boolean)))) return null;
    return { id, numbered: id >= 0 && !/^(?:id|идентификатор|код)$/iu.test(names[id]), title, state, deadlines };
  }

  function guessTitle(cells, excluded, year) {
    return cells.filter((text, i) => i !== excluded && text.length >= 3 && /\p{L}/u.test(text)
      && !status(text) && !metadata.test(text) && !extractDates(text, year).some(date => date.raw === text))
      .sort((a, b) => b.length - a.length)[0] || '';
  }

  function scanSlide(slide, limit, year) {
    const items = [];
    const findings = [];
    let datesChecked = 0;
    let textFound = false;

    function checkText(text, source, shortDates) {
      if (!text) return;
      textFound = true;
      for (const date of extractDates(text, year, shortDates)) {
        datesChecked++;
        if (date.value !== null && date.value >= limit) continue;
        findings.push({
          type: date.value === null ? 'invalid' : 'date',
          slide: slide.number, slideTitle: slide.title, source, excerpt: text,
          subject: date.raw, start: date.start, assumed: date.assumed
        });
      }
    }

    for (const [tableIndex, rows] of slide.tables.entries()) {
      let header;
      for (const [rowIndex, cells] of rows.entries()) {
        const source = 'Таблица ' + (tableIndex + 1) + ' · строка ' + (rowIndex + 1);
        for (const [index, text] of cells.entries()) {
          checkText(text, source + ' · столбец ' + (index + 1), header?.deadlines[index]);
        }
        const detectedHeader = headerFor(cells);
        if (detectedHeader) {
          header = detectedHeader;
          continue;
        }
        const stateIndex = header?.state >= 0 ? header.state : cells.findIndex(cell => status(cell));
        const state = status(cells[stateIndex] || '');
        const idText = header?.id >= 0 ? cells[header.id] || '' : '';
        const id = header?.numbered && /^\d+\.?$/.test(idText) ? '' : idText;
        const title = header?.title >= 0 ? cells[header.title] : guessTitle(cells, stateIndex, year);
        const item = makeItem(title, id, state, slide.number, source, cells.filter(Boolean).join(' · '), true);
        if (item) items.push(item);
      }
    }

    let standaloneTitle = '';
    let standaloneId = '';
    let itemCount = 0;
    for (const [index, block] of slide.blocks.entries()) {
      const source = block === slide.titleBlock ? 'Заголовок слайда' : 'Текстовый блок ' + (index + 1);
      for (const [i, line] of block.lines.entries()) {
        checkText(line, source, i > 0 && dateContext.test(block.lines[i - 1]) && /^[\d./\s]+$/u.test(line));
      }
      if (block === slide.titleBlock) continue;

      let title = standaloneTitle;
      let id = standaloneId;
      for (const [lineIndex, line] of block.lines.entries()) {
        const inline = line.match(/^(.+?)(?:\s+[-–—]\s+|\s*[|·;]\s*|\t+)(.+)$/u)
          || line.match(/^(.+?)\s+\(([^()]+)\)$/u);
        const inlineState = inline && status(inline[2]);
        const state = inlineState || status(line);
        if (state) {
          const name = inlineState ? inline[1] : title || slide.title;
          const item = makeItem(name, id, state, slide.number, source,
            inlineState ? line : [name, line].filter(Boolean).join(' · '), false);
          if (item) {
            items.push(item);
            itemCount++;
          }
          title = '';
          id = '';
          continue;
        }
        const identifier = line.match(idLabel);
        if (identifier) {
          id = identifier[1];
          const afterId = line.slice(identifier[0].length).replace(/^\s*[-–—:]\s*/u, '').trim();
          if (afterId) title = afterId;
        } else if (!metadata.test(line) && !extractDates(line, year).some(date => date.raw === line)) {
          if (!title || lineIndex === 0 || itemLabel.test(line) || /^(?:[•●▪]|\d+[.)]\s)/u.test(line)) title = line;
          if (!id) {
            const candidate = makeItem(line, '', null, slide.number, source, line, false);
            if (candidate) items.push(candidate);
          }
        }
      }
      standaloneTitle = title;
      standaloneId = id;
    }
    if (!itemCount && !slide.tables.length && slide.title) {
      const item = makeItem(slide.title, standaloneId, null, slide.number, 'Заголовок слайда', slide.title, false);
      if (item) items.push(item);
    }
    return { items, findings, datesChecked, textFound };
  }

  function indexItems(items, field) {
    const index = new Map();
    for (const item of items) {
      if (!item[field]) continue;
      const list = index.get(item[field]) || [];
      list.push(item);
      index.set(item[field], list);
    }
    return index;
  }

  function findPrevious(item, byId, byTitle) {
    const idMatch = item.id && byId.get(item.id);
    if (idMatch) return idMatch.every(candidate => candidate.state === 'done') ? { item: idMatch[0], method: 'по ID' } : null;
    const matches = (byTitle.get(item.key) || []).filter(candidate => !candidate.id || !item.id);
    const identities = new Set(matches.map(candidate => candidate.id).filter(Boolean));
    if (matches.length && identities.size <= 1 && matches.every(candidate => candidate.state === 'done')) {
      return { item: matches[0], method: 'по названию' };
    }
    return null;
  }

  function analyze(current, previous, reportDate) {
    const [year, month, day] = reportDate.split('-').map(Number);
    const limit = dateValue(year, month, day);
    if (limit === null) throw new Error('Укажите корректную дату отчёта.');
    const checked = current.map(slide => scanSlide(slide, limit, year));
    const history = previous?.flatMap(slide => scanSlide(slide, limit, year).items).filter(item => item.state || item.row) || [];
    const byId = indexItems(history, 'id');
    const byTitle = indexItems(history, 'key');
    const findings = checked.flatMap(slide => slide.findings);
    for (const [index, slide] of checked.entries()) {
      const repeated = new Set();
      const identifiedTitles = new Set(slide.items.filter(item => item.id && item.key).map(item => item.key));
      for (const item of slide.items) {
        if (!item.id && !item.row && identifiedTitles.has(item.key)) continue;
        const match = findPrevious(item, byId, byTitle);
        const identity = item.id || item.key;
        if (!match || repeated.has(identity)) continue;
        repeated.add(identity);
        findings.push({
          type: 'repeat', slide: item.slide, slideTitle: current[index].title, subject: item.title,
          source: item.source, excerpt: item.excerpt, previous: match.item, method: match.method
        });
      }
    }
    findings.sort((a, b) => a.slide - b.slide);
    return {
      findings,
      slides: current.length,
      datesChecked: checked.reduce((count, slide) => count + slide.datesChecked, 0),
      textless: current.filter((slide, i) => !checked[i].textFound).map(slide => slide.number),
      completed: history.filter(item => item.state === 'done').length,
      previousSlides: previous?.length || 0
    };
  }

  window.ReportReview = { analyze };
})();
