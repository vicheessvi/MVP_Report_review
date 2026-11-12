import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';

const source = new URL('./src/', import.meta.url);
const output = new URL('./Контроль отчётов.html', import.meta.url);

function read(file) {
  return readFileSync(new URL(file, source), 'utf8').replace(/\r\n?/g, '\n').trimEnd();
}

function inline(content, tag) {
  if (new RegExp('</' + tag, 'i').test(content)) {
    throw new Error('Unexpected closing ' + tag + ' tag in bundled content.');
  }
  return '\n' + content + '\n';
}

function hash(content) {
  return "'sha256-" + createHash('sha256').update(content).digest('base64') + "'";
}

const css = inline(read('styles.css'), 'style');
const scripts = [];
const favicon = Buffer.from(read('favicon.svg')).toString('base64');
let html = read('index.html')
  .replace('href="favicon.svg"', 'href="data:image/svg+xml;base64,' + favicon + '"')
  .replace('<link rel="stylesheet" href="styles.css">', () => '<style>' + css + '</style>')
  .replaceAll(/^ {2}<script src="([^"]+)" defer><\/script>\n/gm, (_, file) => {
    scripts.push(inline(read(file), 'script'));
    return '';
  });

if (!scripts.length || /<(?:script|link)\b[^>]*(?:src|href)="(?!data:)/i.test(html)) {
  throw new Error('The offline HTML must contain all styles and scripts.');
}

const policy = [
  "default-src 'none'",
  "connect-src 'none'",
  'script-src ' + scripts.map(hash).join(' '),
  'style-src ' + hash(css),
  'img-src data:',
  "base-uri 'none'",
  "form-action 'none'"
].join('; ');

html = html
  .replace('<meta charset="utf-8">', '<meta charset="utf-8">\n  <meta http-equiv="Content-Security-Policy" content="' + policy + '">')
  .replace('</body>', () => scripts.map(script => '<script>' + script + '</script>').join('\n') + '\n</body>');

writeFileSync(output, html + '\n', 'utf8');
