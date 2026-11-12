(function () {
  'use strict';

  const $ = id => document.getElementById(id);
  const inputDate = $('report-date');
  const state = { current: null, previous: null, result: null, filter: 'all', query: '', version: 0, busy: false };
  const dateFormatter = new Intl.DateTimeFormat('ru-RU', { dateStyle: 'short', timeZone: 'UTC' });
  const numberFormatter = new Intl.NumberFormat('ru-RU');
  const labels = {
    date: { title: 'Дата раньше отчётной', filter: 'Критичные даты', icon: 'calendar', color: 'red' },
    repeat: { title: 'Пункт уже был выполнен', filter: 'Выполненные пункты', icon: 'repeat', color: 'purple' },
    invalid: { title: 'Некорректная дата', filter: 'Некорректные даты', icon: 'warning', color: 'orange' }
  };

  function element(tag, className, value) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (value !== undefined) node.textContent = value;
    return node;
  }

  function icon(name) {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', 'icon');
    svg.setAttribute('aria-hidden', 'true');
    const use = document.createElementNS(svg.namespaceURI, 'use');
    use.setAttribute('href', '#' + name);
    svg.append(use);
    return svg;
  }

  function formatDate(value) {
    return dateFormatter.format(new Date(value + 'T00:00:00Z'));
  }

  function fileSize(bytes) {
    if (bytes < 1024) return bytes + ' Б';
    if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' КБ';
    return (bytes / (1024 * 1024)).toFixed(1).replace('.', ',') + ' МБ';
  }

  function fileError(key, message) {
    const field = $(key + '-error');
    field.textContent = message;
    field.hidden = !message;
    $(key + '-card').classList.toggle('has-error', Boolean(message));
    if (key === 'previous') $('previous-skip').hidden = !message;
  }

  function renderFile(key) {
    const source = state[key];
    $(key + '-drop').hidden = Boolean(source);
    $(key + '-chip').hidden = !source;
    $(key + '-name').textContent = source?.name || '';
    $(key + '-name').title = source?.name || '';
    $(key + '-meta').textContent = source?.file ? fileSize(source.file.size) : 'Демонстрационный отчёт';
  }

  function renderControls() {
    $('date-card').classList.toggle('has-value', Boolean(inputDate.value));
    $('reset').hidden = !(state.current || state.previous || inputDate.value);
    $('run').disabled = state.busy;
    $('run').classList.toggle('is-busy', state.busy);
    $('demo').disabled = state.busy;
  }

  function message(value, isError = false) {
    const node = $('form-status');
    node.hidden = !value;
    node.textContent = value;
    node.classList.toggle('error', isError);
  }

  function updateInput() {
    state.version++;
    state.busy = false;
    state.result = null;
    renderControls();
    renderResults();
    message('');
  }

  function setFile(key, file, error = '') {
    if (file && [state.current, state.previous].some(source => source && !source.file)) {
      const other = key === 'current' ? 'previous' : 'current';
      if (state[other] && !state[other].file) {
        state[other] = null;
        fileError(other, '');
        renderFile(other);
      }
      if (inputDate.value === ReviewDemo.date) inputDate.value = '';
    }
    state[key] = null;
    $(key + '-file').value = '';
    if (file) {
      if (!/\.pptx$/i.test(file.name)) error = 'Выберите файл формата .pptx.';
      else if (file.size > 50 * 1024 * 1024) error = 'Размер файла превышает 50 МБ.';
      else state[key] = { file, name: file.name };
    }
    fileError(key, error);
    renderFile(key);
    updateInput();
  }

  for (const key of ['current', 'previous']) {
    const card = $(key + '-card');
    const picker = $(key + '-file');
    $(key + '-drop').addEventListener('click', () => picker.click());
    $(key + '-remove').addEventListener('click', () => {
      setFile(key, null);
      $(key + '-drop').focus();
    });
    picker.addEventListener('change', () => {
      if (picker.files.length) setFile(key, picker.files[0]);
    });
    card.addEventListener('dragover', event => {
      if (!event.dataTransfer.types.includes('Files')) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = 'copy';
      card.classList.add('dragging');
    });
    card.addEventListener('dragleave', event => {
      if (!card.contains(event.relatedTarget)) card.classList.remove('dragging');
    });
    card.addEventListener('drop', event => {
      event.preventDefault();
      card.classList.remove('dragging');
      const files = event.dataTransfer.files;
      if (files.length > 1) setFile(key, null, 'Перетащите один файл в эту область.');
      else if (files.length) setFile(key, files[0]);
    });
  }
  $('previous-skip').addEventListener('click', () => {
    setFile('previous', null);
    $('run').focus();
  });

  window.addEventListener('dragover', event => {
    if (event.dataTransfer.types.includes('Files')) event.preventDefault();
  });
  window.addEventListener('drop', event => {
    if (event.dataTransfer.types.includes('Files')) event.preventDefault();
  });
  inputDate.addEventListener('input', updateInput);
  $('review-form').noValidate = true;
  $('review-form').addEventListener('submit', event => {
    event.preventDefault();
    run();
  });

  $('reset').addEventListener('click', () => {
    inputDate.value = '';
    setFile('current', null);
    setFile('previous', null);
    $('current-drop').focus();
  });

  $('demo').addEventListener('click', () => {
    for (const key of ['current', 'previous']) {
      state[key] = ReviewDemo[key];
      fileError(key, '');
      renderFile(key);
    }
    inputDate.value = ReviewDemo.date;
    updateInput();
    run();
  });

  async function readSource(source) {
    if (!source) return null;
    if (!source.slides) source.slides = await Pptx.read(source.file);
    return source.slides;
  }

  async function run() {
    if (state.busy) return;
    if (!state.current) {
      fileError('current', 'Добавьте текущий отчёт в формате .pptx.');
      $('current-drop').focus();
      return;
    }
    if (!inputDate.reportValidity()) return;
    if (!state.previous && !$('previous-error').hidden) {
      $('previous-drop').focus();
      return;
    }
    const version = ++state.version;
    state.busy = true;
    renderControls();
    message('Читаем презентации и проверяем текст слайдов…');
    const sources = [state.current, state.previous];
    const results = await Promise.allSettled(sources.map(readSource));
    if (version !== state.version) return;
    try {
      let failed = false;
      for (const [index, result] of results.entries()) {
        if (result.status !== 'rejected') continue;
        const key = index === 0 ? 'current' : 'previous';
        fileError(key, result.reason?.message || 'Не удалось открыть презентацию.');
        failed = true;
      }
      if (failed) {
        message('Замените файл с ошибкой и повторите проверку.', true);
        return;
      }
      const result = ReportReview.analyze(results[0].value, results[1].value, inputDate.value);
      state.result = { ...result, date: inputDate.value, name: sources[0].name, demo: !sources[0].file };
      state.filter = 'all';
      state.query = '';
      renderResults();
      message('');
      $('results-title').focus({ preventScroll: true });
      $('results').scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'instant' : 'smooth', block: 'start' });
      selectSection('#results');
    } catch (error) {
      message(error.message || 'Не удалось завершить проверку.', true);
    } finally {
      state.busy = false;
      renderControls();
    }
  }

  function stat(value, label, iconName, color) {
    const node = element('div', 'stat');
    const tile = element('span', 'tile tile-' + color);
    tile.append(icon(iconName));
    const content = element('div', 'stat-content');
    content.append(element('strong', '', numberFormatter.format(value)), element('span', '', label));
    node.append(tile, content);
    return node;
  }

  function notice(text, name = 'info') {
    const box = element('div', 'notice');
    box.append(icon(name), element('p', '', text));
    return box;
  }

  function renderFinding(finding) {
    const config = labels[finding.type];
    const item = element('article', 'finding');
    const marker = element('span', 'tile tile-' + config.color);
    marker.append(icon(config.icon));
    const content = element('div', 'finding-content');
    const heading = element('div', 'finding-heading');
    const title = element('h4', '', config.title);
    heading.append(title, element('span', 'source', finding.source));
    const excerpt = element('div', 'excerpt');
    const start = finding.start ?? finding.excerpt.toLocaleLowerCase('ru').indexOf(finding.subject.toLocaleLowerCase('ru'));
    const end = start + finding.subject.length;
    if (start >= 0) {
      excerpt.append(finding.excerpt.slice(0, start), element('mark', 'mark-' + config.color, finding.excerpt.slice(start, end)), finding.excerpt.slice(end));
    } else {
      excerpt.textContent = finding.excerpt;
    }
    if (finding.type === 'repeat' && start < 0) heading.insertBefore(element('strong', 'finding-subject', finding.subject), heading.lastChild);
    content.append(heading, excerpt);
    if (finding.assumed) {
      content.append(element('p', 'finding-note', 'Дата без года · использован год отчёта (' + state.result.date.slice(0, 4) + ')'));
    }
    if (finding.previous) {
      const details = element('details', 'evidence');
      const summary = element('summary');
      summary.append(icon('history'), 'Прошлый отчёт · слайд ' + finding.previous.slide + ' · совпадение ' + finding.method);
      details.append(summary, element('div', 'evidence-text', finding.previous.excerpt));
      content.append(details);
    }
    item.append(marker, content);
    return item;
  }

  function renderList(parent, filters) {
    const { findings } = state.result;
    const normalize = text => text.toLocaleLowerCase('ru').replaceAll('ё', 'е').trim();
    const query = normalize(state.query);
    const filtered = findings.filter(finding => {
      if (state.filter !== 'all' && finding.type !== state.filter) return false;
      const content = [finding.excerpt, finding.slideTitle, finding.subject, 'слайд ' + finding.slide, finding.previous?.excerpt].join(' ');
      return !query || normalize(content).includes(query);
    });
    for (const button of filters.children) button.setAttribute('aria-pressed', String(button.dataset.filter === state.filter));
    parent.replaceChildren();
    if (!filtered.length) {
      parent.append(notice('По выбранным фильтрам замечаний нет.', 'search'));
      return;
    }
    const count = element('p', 'findings-count', 'Показано ' + filtered.length + ' из ' + findings.length);
    count.setAttribute('aria-live', 'polite');
    parent.append(count);
    let section;
    let currentSlide;
    for (const finding of filtered) {
      if (currentSlide !== finding.slide) {
        currentSlide = finding.slide;
        section = element('section', 'slide-group');
        const heading = element('h3', 'slide-heading');
        heading.append(element('span', 'slide-number', 'Слайд ' + finding.slide));
        if (finding.slideTitle) heading.append(element('span', 'slide-title', finding.slideTitle));
        section.append(heading);
        parent.append(section);
      }
      section.append(renderFinding(finding));
    }
  }

  function renderResults() {
    const parent = $('report-results');
    const result = state.result;
    parent.hidden = !result;
    $('initial-results').hidden = Boolean(result);
    $('export').hidden = !result?.findings.length;
    $('results-caption').textContent = result
      ? (result.demo ? 'Пример · ' : '') + result.name + ' · дата отчёта ' + formatDate(result.date)
      : 'Замечания с номерами слайдов и фрагментами отчёта.';
    if (!result) {
      parent.replaceChildren();
      return;
    }
    const stats = element('div', 'stats-grid');
    const critical = result.findings.filter(item => item.type === 'date').length;
    const repeats = result.findings.filter(item => item.type === 'repeat').length;
    const invalid = result.findings.filter(item => item.type === 'invalid').length;
    stats.append(
      stat(result.slides, 'Слайды', 'slides', 'blue'),
      stat(critical, labels.date.filter, labels.date.icon, labels.date.color),
      stat(repeats, 'Повторные пункты', 'repeat', 'purple'),
      stat(invalid, labels.invalid.filter, labels.invalid.icon, labels.invalid.color)
    );
    parent.replaceChildren(stats);
    const datesText = 'Распознано дат: ' + result.datesChecked + '. ';
    let comparison = 'Сравнение пропущено: прошлый отчёт не загружен.';
    if (result.previousSlides) comparison = 'Прошлый отчёт: слайдов ' + result.previousSlides + ', выполненных пунктов ' + result.completed + '.';
    parent.append(notice(datesText + comparison));
    if (result.textless.length) {
      const slideList = result.textless.slice(0, 20).join(', ') + (result.textless.length > 20 ? ' и ещё ' + (result.textless.length - 20) : '');
      parent.append(notice('Не найден редактируемый текст на слайдах: ' + slideList + '. Проверьте эти слайды вручную.', 'warning'));
    }
    if (!result.findings.length) {
      const empty = element('div', 'empty-state empty-success');
      const mark = element('span', 'empty-icon');
      mark.append(icon('check'));
      const copy = element('div');
      copy.append(element('strong', '', 'Замечаний не найдено'), element('p', '', 'Проверка распознанного текста завершена.'));
      empty.append(mark, copy);
      parent.append(empty);
      return;
    }
    const toolbar = element('div', 'results-toolbar');
    const filters = element('div', 'filters');
    filters.setAttribute('aria-label', 'Тип замечания');
    const list = element('div', 'findings-list');
    for (const [type, label, count] of [
      ['all', 'Все', result.findings.length], ['date', labels.date.filter, critical],
      ['repeat', labels.repeat.filter, repeats], ['invalid', labels.invalid.filter, invalid]
    ]) {
      if (!count) continue;
      const button = element('button', 'filter', label);
      button.type = 'button';
      button.dataset.filter = type;
      button.append(element('span', '', String(count)));
      button.addEventListener('click', () => {
        state.filter = type;
        renderList(list, filters);
      });
      filters.append(button);
    }
    const search = element('label', 'search-box');
    const input = element('input');
    input.type = 'search';
    input.placeholder = 'Поиск по замечаниям';
    input.setAttribute('aria-label', 'Поиск по замечаниям');
    input.addEventListener('input', () => {
      state.query = input.value;
      renderList(list, filters);
    });
    search.append(icon('search'), input);
    toolbar.append(filters, search);
    parent.append(toolbar, list);
    renderList(list, filters);
  }

  function csvCell(value) {
    let text = String(value ?? '');
    if (/^[\s\uFEFF]*[=+\-@]/u.test(text) || /^[\t\r\n]/u.test(text)) text = "'" + text;
    return '"' + text.replaceAll('"', '""') + '"';
  }

  $('export').addEventListener('click', () => {
    if (!state.result?.findings.length) return;
    const { findings, date, name } = state.result;
    const rows = [
      ['Отчёт', 'Дата отчёта', 'Слайд', 'Заголовок слайда', 'Тип замечания', 'Дата или пункт', 'Расположение',
        'Текст', 'Слайд прошлого отчёта', 'Совпадение', 'Текст прошлого отчёта', 'Примечание'],
      ...findings.map(item => [name, formatDate(date), item.slide, item.slideTitle, labels[item.type].title, item.subject, item.source,
        item.excerpt, item.previous?.slide, item.method, item.previous?.excerpt, item.assumed ? 'Год отчёта: ' + date.slice(0, 4) : ''])
    ];
    const csv = rows.map(row => row.map(csvCell).join(';')).join('\r\n');
    const link = element('a');
    link.href = URL.createObjectURL(new Blob(['\uFEFF', csv], { type: 'text/csv;charset=utf-8' }));
    link.download = 'Замечания-' + date + '.csv';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 30000);
  });

  const help = $('help-dialog');
  for (const button of document.querySelectorAll('[data-help]')) button.addEventListener('click', () => help.showModal());
  help.addEventListener('click', event => {
    if (event.target !== help) return;
    const bounds = help.getBoundingClientRect();
    if (event.clientX < bounds.left || event.clientX > bounds.right || event.clientY < bounds.top || event.clientY > bounds.bottom) help.close();
  });

  function selectSection(hash) {
    for (const link of document.querySelectorAll('a.nav-link')) {
      if (link.hash === hash) link.setAttribute('aria-current', 'location');
      else link.removeAttribute('aria-current');
    }
  }

  window.addEventListener('hashchange', () => selectSection(location.hash === '#results' ? '#results' : '#check'));
  for (const link of document.querySelectorAll('a.nav-link, .brand')) link.addEventListener('click', () => selectSection(link.hash));
  selectSection(location.hash === '#results' ? '#results' : '#check');
  renderControls();
})();
