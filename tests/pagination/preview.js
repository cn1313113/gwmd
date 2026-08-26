// 预览分页（确定性模型）：node preview_pages.js <file.gwmd>
// 输出: PREVIEW_PAGES:<n> P1:first..last | P2:first..last ...
const fs = require('fs');
const path = require('path');
const APP_JS = path.join(__dirname, '..', '..', 'app.js');
const appSrc = fs.readFileSync(APP_JS, 'utf8');

function extract(name) {
  const i = appSrc.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('not found: ' + name);
  let j = appSrc.indexOf('{', i), depth = 0;
  for (; j < appSrc.length; j++) {
    if (appSrc[j] === '{') depth++;
    else if (appSrc[j] === '}') { depth--; if (depth === 0) break; }
  }
  return appSrc.slice(i, j + 1);
}
for (const n of ['CHARS_PER_LINE', 'FN_CHARS_PER_LINE', 'LINE_BODY_PT', 'LINE_TITLE_PT',
                 'LINE_CC_PT', 'LINE_TABLE_PT', 'LINE_FN_PT', 'PAGE_HEIGHT_PT']) {
  eval('var ' + n + ' = ' + appSrc.match(new RegExp('const ' + n + ' = ([^;]+);'))[1] + ';');
}
eval('var EAW_WIDE = ' + appSrc.match(/const EAW_WIDE = (\[[\s\S]*?\n\]);/)[1] + ';');
for (const fn of ['isWideChar', 'textWidth', 'calcLines', 'splitBodyLines',
                  'collectFootnotesFrom', 'blockHeightPt', 'fnReservePt',
                  'paginateDeterministic']) eval(extract(fn));

const GWMD = require(APP_JS);
const text = fs.readFileSync(process.argv[2], 'utf8');
const { blocks, footnotes } = GWMD.parse(text);
const pages = paginateDeterministic(blocks, footnotes);
const brief = pages.map((pg) => {
  const first = pg.items[0] || { type: '空' };
  const last = pg.items[pg.items.length - 1] || { type: '空' };
  return `P${first.type}..${last.type}`;
}).join(' | ');
console.log('PREVIEW_PAGES:' + pages.length + ' ' + brief);