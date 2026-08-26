/* ============================================================
 * GWMD 公文编辑器前端逻辑
 * - 左侧编辑 GWMD 文本，右侧实时预览 A4 公文效果
 * - 保存 .gwmd 文本文件 / 导出 .docx（调用 api.php）
 * ============================================================ */
'use strict';

/* ---------------- 解析器：GWMD → 块列表 ---------------- */
const GWMD = (() => {
  const RE_DIRECTIVE = /^@([a-z]+)\s*(.*)$/;
  const RE_HEADING = /^(#{1,6})\s+(.*)$/;
  const RE_ORDERED = /^(\d+)[\.、]\s+(.*)$/;
  const RE_UNORDERED = /^[-*]\s+(.*)$/;
  const RE_QUOTE = /^>\s?(.*)$/;
  const RE_TABLE_ROW = /^\|.*\|$/;
  const RE_TABLE_SEP = /^\|[\s:\-|]+\|$/;
  const RE_FOOTNOTE_DEF = /^\[\^(\d+)\]:\s*(.*)$/;
  const RE_ATTACH_ITEM = /\[([^\]]+)\]/g;
  const RE_INLINE_TITLE = /^(#{1,5})\s+(.+?)\s+\1\s+(.+)$/;
  const HEADING_MAP = { 1: 'h1', 2: 'h2', 3: 'h3', 4: 'h4', 5: 'h4', 6: 'h4' };
  const HEADER_DIRECTIVES = { header: 'masthead', docno: 'docno', secret: 'secret', urgent: 'urgent', serial: 'serial', signer: 'signer' };
  const BODY_DIRECTIVES = { to: 'to', sign: 'sign', date: 'date', stamp: 'stamp', note: 'note', subtitle: 'subtitle', title: 'title' };
  const RECORD_DIRECTIVES = { cc: 'cc', print: 'print', issue: 'issue' };

  // ---- 标题自动编号（与后端 gwmd2docx.py 一致）----
  const CN_DIGITS = '零一二三四五六七八九';
  const RE_CN_NUM = /^([一二三四五六七八九十]+)、/;
  const RE_CN_PAREN = /^（([一二三四五六七八九十]+)）/;
  const RE_ARABIC_NUM = /^(\d+)[\.．]/;
  const RE_ARABIC_PAREN = /^（(\d+)）/;

  function intToCn(n) {
    if (n <= 0) return '零';
    if (n < 10) return CN_DIGITS[n];
    if (n < 20) return '十' + (n > 10 ? CN_DIGITS[n - 10] : '');
    const tens = Math.floor(n / 10), ones = n % 10;
    return CN_DIGITS[tens] + '十' + (ones ? CN_DIGITS[ones] : '');
  }

  function cnToInt(s) {
    if (!s) return 0;
    const units = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
    if (!s.includes('十')) return units[s] || 0;
    const parts = s.split('十');
    const tens = parts[0] ? (units[parts[0]] || 1) : 1;
    const ones = parts[1] ? (units[parts[1]] || 0) : 0;
    return tens * 10 + ones;
  }

  function numPrefix(level, n) {
    if (level === 2) return intToCn(n) + '、';
    if (level === 3) return '（' + intToCn(n) + '）';
    if (level === 4) return n + '.';
    if (level === 5) return '（' + n + '）';
    return '';
  }

  function detectManualNumber(text, level) {
    let m;
    if (level === 2) { m = text.match(RE_CN_NUM); if (m) return [true, cnToInt(m[1])]; }
    else if (level === 3) { m = text.match(RE_CN_PAREN); if (m) return [true, cnToInt(m[1])]; }
    else if (level === 4) { m = text.match(RE_ARABIC_NUM); if (m) return [true, parseInt(m[1], 10)]; }
    else if (level === 5) { m = text.match(RE_ARABIC_PAREN); if (m) return [true, parseInt(m[1], 10)]; }
    return [false, 0];
  }

  function numberHeadings(blocks) {
    const counters = { 2: 0, 3: 0, 4: 0, 5: 0 };
    for (const b of blocks) {
      let level;
      const t = b.type;
      if (['h1', 'h2', 'h3', 'h4'].includes(t)) level = { h1: 2, h2: 3, h3: 4, h4: 5 }[t];
      else if (t === 'inline_title') level = b.level || 2;
      else continue;
      const text = (t === 'inline_title' ? (b.title || '') : (b.text || '')).trim();
      const [manual, num] = detectManualNumber(text, level);
      if (manual) {
        counters[level] = Math.max(counters[level], num);
      } else {
        counters[level] += 1;
        const prefix = numPrefix(level, counters[level]);
        if (t === 'inline_title') b.title = prefix + text;
        else b.text = prefix + text;
      }
      for (let lv = level + 1; lv <= 5; lv++) counters[lv] = 0;
    }
    return blocks;
  }

  function extractTrailingIndent(blocks) {
    // @sign/@date 行末空格 → right 字段（1/100 字符宽：半角空格=50，全角空格=100）
    for (const b of blocks) {
      if (b.type !== 'sign' && b.type !== 'date') continue;
      const t = b.text || '';
      const m = t.match(/[ \u3000]+$/);
      if (m) {
        const trail = m[0];
        b.right = ((trail.match(/ /g) || []).length * 50) +
                  ((trail.match(/\u3000/g) || []).length * 100);
        b.text = t.slice(0, m.index);
      }
    }
    return blocks;
  }

  function isListStart(line) {
    return RE_ORDERED.test(line) || RE_UNORDERED.test(line);
  }

  function parse(text) {
    const lines = text.split('\n');
    const blocks = [];
    const footnotes = {};
    let pendingTable = null;
    let pendingList = null;
    let i = 0;

    const flushTable = () => {
      if (pendingTable) { blocks.push({ type: 'table', rows: pendingTable }); pendingTable = null; }
    };
    const flushList = () => {
      if (pendingList) { blocks.push({ type: 'list', ordered: pendingList.ordered, items: pendingList.items }); pendingList = null; }
    };

    while (i < lines.length) {
      const raw = lines[i];
      const line = raw.trim();
      i += 1;

      if (!line) { flushTable(); flushList(); continue; }

      // 行级转义：\X 开头 → 字面正文（去掉反斜杠）
      if (line.startsWith('\\')) {
        flushTable(); flushList();
        blocks.push({ type: 'body', text: line.slice(1) });
        continue;
      }

      // 注释行：<!-- 内容 --> 不导出
      if (line.startsWith('<!--')) { continue; }

      // 脚注定义
      let m = line.match(RE_FOOTNOTE_DEF);
      if (m) { footnotes[m[1]] = m[2].trim(); continue; }

      // 指令行
      if (line.startsWith('@')) {
        flushTable(); flushList();
        const rawLine = raw.replace(/\r?\n$/, '');
        m = rawLine.trimStart().match(RE_DIRECTIVE);
        if (m) {
          const key = m[1];
          const content = (key === 'sign' || key === 'date') ? m[2] : m[2].trim();
          if (key === 'comment') { /* 注释，不导出 */ }
          else if (key === 'redline') blocks.push({ type: 'redline' });
          else if (key === 'pagebreak') blocks.push({ type: 'pagebreak' });
          else if (key === 'config') blocks.push({ type: 'config', text: content });
          else if (key === 'blank' || key === 'spacer' || key === 'space' || key === 'sp') {
            const n = parseInt(content.trim(), 10);
            blocks.push({ type: 'blank', count: Math.max(1, Math.min(isNaN(n) ? 1 : n, 20)) });
          }
          else if (key === 'attach') {
            const items = [...content.matchAll(/\[([^\]]+)\]/g)].map(x => x[1].trim());
            if (items.length) blocks.push({ type: 'attach', items });
            else {
              const name = content.replace(/^附件\s*[:：]?\s*/, '').trim();
              blocks.push({ type: 'attach', items: [name || '附件'] });
            }
          }
          else if (HEADER_DIRECTIVES[key]) blocks.push({ type: HEADER_DIRECTIVES[key], text: content });
          else if (BODY_DIRECTIVES[key]) blocks.push({ type: BODY_DIRECTIVES[key], text: content });
          else if (RECORD_DIRECTIVES[key]) blocks.push({ type: RECORD_DIRECTIVES[key], text: content });
          else blocks.push({ type: 'body', text: line });
        }
        continue;
      }

      // 表格
      if (RE_TABLE_ROW.test(line)) {
        flushList();
        if (RE_TABLE_SEP.test(line)) continue;
        const cells = protectEscapes(line.replace(/^\||\|$/g, '')).split('|').map(c => c.trim());
        if (!pendingTable) pendingTable = [];
        pendingTable.push(cells);
        continue;
      }

      // 配对标题：# 标题 # 正文
      m = line.match(RE_INLINE_TITLE);
      if (m) {
        flushTable(); flushList();
        blocks.push({ type: 'inline_title', level: Math.min(m[1].length + 1, 5), title: m[2].trim(), body: m[3].trim() });
        continue;
      }

      // 标题
      m = line.match(RE_HEADING);
      if (m) {
        flushTable(); flushList();
        const level = m[1].length;
        blocks.push({ type: HEADING_MAP[level] || 'h4', text: m[2].trim() });
        continue;
      }

      // 有序列表
      m = line.match(RE_ORDERED);
      if (m) {
        flushTable();
        if (!pendingList || !pendingList.ordered) { flushList(); pendingList = { ordered: true, items: [] }; }
        pendingList.items.push(m[2].trim());
        continue;
      }
      // 无序列表
      m = line.match(RE_UNORDERED);
      if (m) {
        flushTable();
        if (!pendingList || pendingList.ordered) { flushList(); pendingList = { ordered: false, items: [] }; }
        pendingList.items.push(m[1].trim());
        continue;
      }

      // 引用
      m = line.match(RE_QUOTE);
      if (m) {
        flushTable(); flushList();
        blocks.push({ type: 'quote', text: m[1].trim() });
        continue;
      }

      // 正文（连续行合并为一段）
      flushTable(); flushList();
      const paraLines = [line];
      while (i < lines.length) {
        const next = lines[i].trim();
        if (!next) break;
        if (next.startsWith('#') || next.startsWith('@') || next.startsWith('|') ||
            next.startsWith('>') || isListStart(next) || RE_FOOTNOTE_DEF.test(next) ||
            RE_TABLE_ROW.test(next)) break;
        paraLines.push(next);
        i += 1;
      }
      blocks.push({ type: 'body', text: paraLines.join(' ') });
    }
    flushTable(); flushList();
    const numbered = numberHeadings(blocks);
    stripUndefinedFootnotes(numbered, footnotes);
    extractTrailingIndent(numbered);
    return { blocks: numbered, footnotes };
  }

  function stripUndefinedFootnotes(blocks, footnotes) {
    // 剔除未定义（无 [^n]: 脚注定义）的引用，与后端逻辑保持一致
    const clean = s => String(s == null ? '' : s).replace(/\[\^(\d+)\]/g, (m, id) => (id in footnotes ? m : ''));
    for (const b of blocks) {
      if (b.type === 'inline_title') { b.title = clean(b.title); b.body = clean(b.body); }
      else if (b.type === 'table') { b.rows = b.rows.map(r => r.map(c => clean(c))); }
      else if (b.type === 'list') { b.items = b.items.map(clean); }
      else b.text = clean(b.text || '');
    }
  }

  function validate(src) {
    const warnings = [];
    const { blocks, footnotes } = parse(src);
    // 大标题
    const titles = blocks.filter(b => b.type === 'title');
    if (!titles.length) warnings.push('缺少公文大标题（@title）');
    else if (titles.length > 1) warnings.push('检测到多个 @title，公文大标题应只有一个');
    // 未定义脚注（原始文本扫描，排除转义 \[^n\]）
    const defined = new Set(Object.keys(footnotes));
    const used = new Set();
    for (const m of src.matchAll(/(?<!\\)\[\^(\d+)\]/g)) used.add(m[1]);
    for (const n of [...used].filter(x => !defined.has(x)).sort((a, b) => +a - +b)) {
      warnings.push(`脚注 [^${n}] 无定义，导出时已忽略`);
    }
    // 未知指令（行级扫描）
    const KNOWN = new Set([...Object.keys(HEADER_DIRECTIVES), ...Object.keys(BODY_DIRECTIVES),
      ...Object.keys(RECORD_DIRECTIVES), 'redline', 'pagebreak', 'config', 'blank',
      'spacer', 'space', 'sp', 'attach', 'comment']);
    for (const rawLine of src.split('\n')) {
      const l = rawLine.trimStart();
      if (l.startsWith('@')) {
        const m = l.match(RE_DIRECTIVE);
        if (m && !KNOWN.has(m[1])) warnings.push(`未知指令 @${m[1]}，已按正文处理`);
      }
    }
    // 标题层级跳跃
    let prev = 1;
    for (const b of blocks) {
      let lv = 0, title = '';
      if (['h1', 'h2', 'h3', 'h4'].includes(b.type)) { lv = { h1: 2, h2: 3, h3: 4, h4: 5 }[b.type]; title = b.text || ''; }
      else if (b.type === 'inline_title') { lv = b.level || 2; title = b.title || ''; }
      if (!lv) continue;
      if (lv > prev + 1) warnings.push(`标题层级跳跃："${title}"，前面缺少上级标题`);
      prev = Math.max(prev, lv);
    }
    // @stamp 之后仍有正文
    let stampSeen = false;
    for (const b of blocks) {
      if (b.type === 'stamp') { stampSeen = true; continue; }
      if (stampSeen && ['body', 'title', 'h1', 'h2', 'h3', 'h4', 'inline_title', 'to', 'subtitle', 'attach'].includes(b.type)) {
        warnings.push('@stamp 之后仍有正文内容，印章应位于落款之后、附注之前');
        break;
      }
    }
    return warnings;
  }

  return { parse, validate };
})();

/* ---------------- 转义：\X 输出字面 X（特殊字符表顺序与后端 gwmd2docx.py ESC_CHARS 严格一致） ---------------- */
const ESC_CHARS = '\\*[]^@#|>(';
const ESC_PUA = '\uE000\uE001\uE002\uE003\uE004\uE005\uE006\uE007\uE008\uE009\uE00A';
function protectEscapes(s) {
  let out = '', i = 0;
  const n = s.length;
  while (i < n) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < n) {
      const idx = ESC_CHARS.indexOf(s[i + 1]);
      if (idx >= 0) { out += ESC_PUA[idx]; i += 2; continue; }
    }
    out += ch; i += 1;
  }
  return out;
}
function restoreEscapes(s) {
  return s.replace(/[\uE000-\uE00A]/g, c => ESC_CHARS.charAt(c.charCodeAt(0) - 0xE000));
}

/* ---------------- 行内渲染：**加粗** *斜体* [^1] [链接] ---------------- */
function renderInline(text, footnotes) {
  // 转义保护：\X → 私有区占位（先于一切标记解析，最后恢复）
  text = protectEscapes(text);
  // 先转义 HTML，防止注入（标记符号不受转义影响）
  text = escapeHtml(text);
  // 脚注引用 [^n] → 上标
  let html = text.replace(/\[\^(\d+)\]/g, (_, n) =>
    `<sup class="gw-footnote-ref" title="${escapeHtml((footnotes && footnotes[n]) || '')}">${n}</sup>`);
  // 链接 [text](url)
  html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, (_, t, u) =>
    `<a href="${u}" target="_blank">${t}</a>`);
  // 加粗 **x**（先保护斜体内的）
  html = html.replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>');
  // 斜体 *x*
  html = html.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, '$1<i>$2</i>');
  return restoreEscapes(html);
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/* ---------------- 块 → HTML ---------------- */
/* ---------------- 分页渲染（确定性行模型，与 docx 固定行距一一对应） ---------------- */
const CHARS_PER_LINE = 28;    // 每行 28 字（正文 3 号，版心 156mm）
const FN_CHARS_PER_LINE = 31; // 脚注 4 号（14pt）每行字数（156mm / 14pt）
const LINE_BODY_PT = 29;      // 正文行距（docx LINE_SPACING）
const LINE_TITLE_PT = 33;     // 大标题行距（docx TITLE_LINE_SPACING）
const LINE_CC_PT = 24;        // 版记行距（docx set_spacing(p, 24)）
const LINE_TABLE_PT = 20;     // 表格单元格行距（docx set_spacing(p, 20)）
const LINE_FN_PT = 20;        // 脚注行距（docx 注入固定 400 twips exact）
const PAGE_HEIGHT_PT = 638;   // 版心高 225mm（22 行 × 29pt）

// East Asian Width 全宽区间表（F/W/A），由 tests/gen_eaw_table.py 生成
// 与后端 gwmd2docx.py 的 unicodedata.east_asian_width 判定保持一致
// 修改：重跑 tests/gen_eaw_table.py（勿手改，否则 tests/consistency.py 会失败）
const EAW_WIDE = [
  [161,161], [164,164], [167,168], [170,170], [173,174], [176,180], [182,186], [188,191],
  [198,198], [208,208], [215,216], [222,225], [230,230], [232,234], [236,237], [240,240],
  [242,243], [247,250], [252,252], [254,254], [257,257], [273,273], [275,275], [283,283],
  [294,295], [299,299], [305,307], [312,312], [319,322], [324,324], [328,331], [333,333],
  [338,339], [358,359], [363,363], [462,462], [464,464], [466,466], [468,468], [470,470],
  [472,472], [474,474], [476,476], [593,593], [609,609], [708,708], [711,711], [713,715],
  [717,717], [720,720], [728,731], [733,733], [735,735], [768,879], [913,929], [931,937],
  [945,961], [963,969], [1025,1025], [1040,1103], [1105,1105], [4352,4447], [8208,8208], [8211,8214],
  [8216,8217], [8220,8221], [8224,8226], [8228,8231], [8240,8240], [8242,8243], [8245,8245], [8251,8251],
  [8254,8254], [8308,8308], [8319,8319], [8321,8324], [8364,8364], [8451,8451], [8453,8453], [8457,8457],
  [8467,8467], [8470,8470], [8481,8482], [8486,8486], [8491,8491], [8531,8532], [8539,8542], [8544,8555],
  [8560,8569], [8585,8585], [8592,8601], [8632,8633], [8658,8658], [8660,8660], [8679,8679], [8704,8704],
  [8706,8707], [8711,8712], [8715,8715], [8719,8719], [8721,8721], [8725,8725], [8730,8730], [8733,8736],
  [8739,8739], [8741,8741], [8743,8748], [8750,8750], [8756,8759], [8764,8765], [8776,8776], [8780,8780],
  [8786,8786], [8800,8801], [8804,8807], [8810,8811], [8814,8815], [8834,8835], [8838,8839], [8853,8853],
  [8857,8857], [8869,8869], [8895,8895], [8978,8978], [8986,8987], [9001,9002], [9193,9196], [9200,9200],
  [9203,9203], [9312,9449], [9451,9547], [9552,9587], [9600,9615], [9618,9621], [9632,9633], [9635,9641],
  [9650,9651], [9654,9655], [9660,9661], [9664,9665], [9670,9672], [9675,9675], [9678,9681], [9698,9701],
  [9711,9711], [9725,9726], [9733,9734], [9737,9737], [9742,9743], [9748,9749], [9756,9756], [9758,9758],
  [9792,9792], [9794,9794], [9800,9811], [9824,9825], [9827,9829], [9831,9834], [9836,9837], [9839,9839],
  [9855,9855], [9875,9875], [9886,9887], [9889,9889], [9898,9899], [9917,9919], [9924,9953], [9955,9955],
  [9960,9983], [9989,9989], [9994,9995], [10024,10024], [10045,10045], [10060,10060], [10062,10062], [10067,10069],
  [10071,10071], [10102,10111], [10133,10135], [10160,10160], [10175,10175], [11035,11036], [11088,11088], [11093,11097],
  [11904,11929], [11931,12019], [12032,12245], [12272,12283], [12288,12350], [12353,12438], [12441,12543], [12549,12591],
  [12593,12686], [12688,12771], [12784,12830], [12832,19903], [19968,42124], [42128,42182], [43360,43388], [44032,55203],
  [57344,64255], [65024,65049], [65072,65106], [65108,65126], [65128,65131], [65281,65376], [65504,65510], [65533,65533],
  [94176,94180], [94192,94193], [94208,100343], [100352,101589], [101632,101640], [110576,110579], [110581,110587], [110589,110590],
  [110592,110882], [110898,110898], [110928,110930], [110933,110933], [110948,110951], [110960,111355], [126980,126980], [127183,127183],
  [127232,127242], [127248,127277], [127280,127337], [127344,127404], [127488,127490], [127504,127547], [127552,127560], [127568,127569],
  [127584,127589], [127744,127776], [127789,127797], [127799,127868], [127870,127891], [127904,127946], [127951,127955], [127968,127984],
  [127988,127988], [127992,128062], [128064,128064], [128066,128252], [128255,128317], [128331,128334], [128336,128359], [128378,128378],
  [128405,128406], [128420,128420], [128507,128591], [128640,128709], [128716,128716], [128720,128722], [128725,128727], [128732,128735],
  [128747,128748], [128756,128764], [128992,129003], [129008,129008], [129292,129338], [129340,129349], [129351,129535], [129648,129660],
  [129664,129672], [129680,129725], [129727,129733], [129742,129755], [129760,129768], [129776,129784], [131072,196605], [196608,262141],
  [917760,917999], [983040,1048573], [1048576,1114109],
];
function isWideChar(ch) {
  const c = ch.codePointAt(0);
  let lo = 0, hi = EAW_WIDE.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const r = EAW_WIDE[mid];
    if (c < r[0]) hi = mid - 1;
    else if (c > r[1]) lo = mid + 1;
    else return true;
  }
  return false;
}

function textWidth(str) {
  let w = 0;
  for (const ch of String(str)) w += isWideChar(ch) ? 1 : 0.5;
  return w;
}
function calcLines(str, indent) {
  // 首行缩进 indent 字，后续行满行 28 字（与 Word 排版一致）
  const w = textWidth(str);
  const first = CHARS_PER_LINE - (indent || 0);
  if (w <= first) return 1;
  return 1 + Math.ceil((w - first) / CHARS_PER_LINE);
}
function splitBodyLines(text) {
  const lines = [];
  let cur = '', w = 0;
  for (const ch of String(text)) {
    const cw = isWideChar(ch) ? 1 : 0.5;
    if (cur && w + cw > CHARS_PER_LINE) { lines.push(cur); cur = ch; w = cw; }
    else { cur += ch; w += cw; }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}
function collectFootnotesFrom(b) {
  const txt = b.type === 'inline_title'
    ? ((b.title || '') + ' ' + (b.body || ''))
    : (b.text || '');
  const ids = new Set();
  for (const m of String(txt).matchAll(/\[\^(\d+)\]/g)) ids.add(m[1]);
  return ids;
}
function blockHeightPt(b) {
  // 确定性行模型：行数 × 固定行距（与 docx 各块 set_spacing 一一对应）
  const t = b.type;
  const txt = t === 'inline_title' ? ((b.title || '') + (b.body || '')) : (b.text || '');
  const lines = indent => calcLines(txt, indent);
  switch (t) {
    case 'title': {
      let n = 0;
      for (const part of (b.text || '').split('|')) n += calcLines(part, 0);
      return n * LINE_TITLE_PT;
    }
    case 'subtitle': case 'masthead': return lines(0) * LINE_TITLE_PT;
    case 'list': {
      // 列表项逐项计行（含编号/项目符号前缀），行距 29pt——避免低估导致内容溢出页面
      let n = 0;
      (b.items || []).forEach((it, i) => {
        n += calcLines((b.ordered ? `${i + 1}. ` : '- ') + it, 2);
      });
      return Math.max(1, n) * LINE_BODY_PT;
    }
    case 'h1': case 'h2': case 'h3': case 'h4':
    case 'inline_title': case 'quote': case 'note': case 'body':
      return lines(2) * LINE_BODY_PT;
    case 'redline': return 2;  // 红线占位 2pt（与 docx set_spacing(p, 2) 一致）
    case 'to': case 'docno': case 'secret': case 'urgent': case 'serial':
    case 'signer': case 'sign': case 'date': case 'stamp':
      return lines(0) * LINE_BODY_PT;
    case 'blank': return (b.count || 1) * LINE_BODY_PT;
    case 'attach': {
      const items = b.items || [b.text || '附件'];
      let n = 0;
      items.forEach((it, i) => {
        const label = (i === 0 ? '附件：' : '') + (i + 1) + '. ' + it;
        n += calcLines(label, i === 0 ? 2 : 5);
      });
      return n * LINE_BODY_PT;
    }
    case 'table': return (b.rows || []).length * LINE_TABLE_PT + LINE_BODY_PT; // + 表格后空段
    case 'cc': case 'print': case 'issue': return lines(0) * LINE_CC_PT;
    case 'config': case 'pagebreak': return 0;
    default: return lines(2) * LINE_BODY_PT;
  }
}
function fnReservePt(ids, footnotes) {
  // 脚注区：分隔符行 20pt + 内容行 20pt × 行数（docx 已固定 exact 20pt）
  let sum = 0;
  for (const id of ids) {
    sum += Math.max(1, Math.ceil(textWidth(footnotes[id] || '') / FN_CHARS_PER_LINE)) * LINE_FN_PT;
  }
  return LINE_FN_PT + sum;
}
function paginateDeterministic(blocks, footnotes) {
  const pages = [];
  let curItems = [], curFoot = new Set(), curH = 0;
  const flush = () => {
    pages.push({ items: curItems, footnotes: [...curFoot].sort((a, b) => +a - +b) });
    curItems = []; curFoot = new Set(); curH = 0;
  };
  for (const b of blocks) {
    if (b.type === 'pagebreak') { flush(); continue; }
    const ids = [...collectFootnotesFrom(b)];
    const h = blockHeightPt(b);
    if (curH > 0 && curH + h + fnReservePt([...curFoot, ...ids], footnotes) > PAGE_HEIGHT_PT) {
      flush();
    }
    // 超长正文按行拆分跨页
    if (b.type === 'body' && h > PAGE_HEIGHT_PT) {
      const lines = splitBodyLines(b.text || '');
      let i = 0;
      while (i < lines.length) {
        const space = PAGE_HEIGHT_PT - curH - fnReservePt([...curFoot, ...ids], footnotes);
        const canFit = Math.max(1, Math.floor(space / LINE_BODY_PT));
        const part = { type: 'body', text: lines.slice(i, i + canFit).join('') };
        curItems.push(part);
        ids.forEach(id => curFoot.add(id));
        curH += canFit * LINE_BODY_PT;
        i += canFit;
        if (i < lines.length) flush();
      }
      continue;
    }
    curItems.push(b);
    ids.forEach(id => curFoot.add(id));
    curH += h;
  }
  if (curItems.length || curFoot.size) flush();
  if (!pages.length) pages.push({ items: [], footnotes: [] });
  return pages;
}
/* ---------------- 块 → HTML ---------------- */
function renderBlockItem(b, footnotes) {
  const t = b.type;
  const txt = b.text || '';
  switch (t) {
    case 'blank': {
      const bd = Math.max(1, b.count || 1);
      return `<div class="gw-blank" style="height:${bd * 29}pt"></div>`;
    }
    case 'title': {
      const parts = protectEscapes(txt || '').split('|');
      const titleHtml = parts.map((part, i) => (i > 0 ? '<br>' : '') + renderInline(part)).join('');
      return `<div class="gw-title">${titleHtml}</div>`;
    }
    case 'subtitle': return `<div class="gw-subtitle">${renderInline(txt)}</div>`;
    case 'masthead': return `<div class="gw-masthead">${renderInline(txt)}</div>`;
    case 'docno': return `<div class="gw-docno">${renderInline(txt)}</div>`;
    case 'secret': return `<div class="gw-secret">${renderInline(txt)}</div>`;
    case 'urgent': return `<div class="gw-urgent">${renderInline(txt)}</div>`;
    case 'serial': return `<div class="gw-serial">${renderInline(txt)}</div>`;
    case 'signer': return `<div class="gw-signer">${renderInline(txt)}</div>`;
    case 'redline': return '<div class="gw-redline"></div>';
    case 'pagebreak': return '<div class="gw-pagebreak"></div>';
    case 'config': return '';
    case 'h1': return `<div class="gw-h1">${renderInline(txt)}</div>`;
    case 'h2': return `<div class="gw-h2">${renderInline(txt)}</div>`;
    case 'h3': return `<div class="gw-h3">${renderInline(txt)}</div>`;
    case 'h4': return `<div class="gw-h4">${renderInline(txt)}</div>`;
    case 'inline_title': {
      const level = b.level || 2;
      const headClass = level === 2 ? 'gw-h1-inline' : level === 3 ? 'gw-h2-inline' : level === 4 ? 'gw-h3-inline' : 'gw-h4-inline';
      const headHtml = renderInline(b.title || '', footnotes);
      const bodyHtml = b.body ? renderInline(b.body, footnotes) : '';
      let html = `<div class="gw-inline-title"><span class="${headClass}">${headHtml}</span>`;
      if (bodyHtml) html += `<span class="gw-inline-title-body"> ${bodyHtml}</span>`;
      html += '</div>';
      return html;
    }
    case 'to': return `<div class="gw-to">${renderInline(txt)}</div>`;
    case 'attach': {
      const items = b.items || [txt || '附件'];
      return items.map((name, i) => {
        if (i === 0) return `<div class="gw-attach">附件：${i + 1}. ${renderInline(name)}</div>`;
        return `<div class="gw-attach gw-attach-sub">${i + 1}. ${renderInline(name)}</div>`;
      }).join('\n');
    }
    case 'note': return `<div class="gw-note">${renderInline(txt)}</div>`;
    case 'stamp': return `<div class="gw-stamp">${renderInline(txt || '（印章）')}</div>`;
    case 'sign': {
      const pad = b.right ? ` style="padding-right:${b.right / 100}em"` : '';
      return `<div class="gw-sign"${pad}>${renderInline(txt)}</div>`;
    }
    case 'date': {
      const pad = b.right ? ` style="padding-right:${b.right / 100}em"` : '';
      return `<div class="gw-date"${pad}>${renderInline(txt)}</div>`;
    }
    case 'quote': return `<div class="gw-quote">${renderInline(txt)}</div>`;
    case 'cc': case 'print': case 'issue':
      return `<div class="gw-${t}">${renderInline(txt)}</div>`;
    case 'list': {
      return b.items.map((it, idx) =>
        `<div class="gw-list">${b.ordered ? `${idx + 1}. ` : '- '}${renderInline(it)}</div>`).join('\n');
    }
    case 'table': {
      const rows = b.rows.map((r, ri) => {
        const tag = ri === 0 ? 'th' : 'td';
        return `<tr>${r.map(c => `<${tag}>${escapeHtml(c)}</${tag}>`).join('')}</tr>`;
      }).join('');
      return `<div class="gw-table-wrap"><table class="gw-table"><tbody>${rows}</tbody></table></div>`;
    }
    default: return `<div class="gw-body">${renderInline(txt, footnotes)}</div>`;
  }
}
function renderBlocks(parsed) {
  const { blocks, footnotes } = parsed;
  const pages = paginateDeterministic(blocks, footnotes);
  return pages.map((pg, idx) => {
    const pageNo = idx + 1;
    const itemsHtml = pg.items.map(b => renderBlockItem(b, footnotes)).join('\n');
    let fnHtml = '';
    if (pg.footnotes.length) {
      fnHtml = '<div class="gw-footnotes">' +
        pg.footnotes.map(n => `<div>${n}. ${escapeHtml(footnotes[n] || '')}</div>`).join('') +
        '</div>';
    }
    const numCls = pageNo % 2 === 1 ? 'odd' : 'even';
    return `<div class="page"><div class="gw-content">${itemsHtml}</div>${fnHtml}<div class="page-num ${numCls}">— ${pageNo} —</div></div>`;
  }).join('\n');
}

function renderWarnings(warnings) {
  if (!warnings.length) return '';
  const items = warnings.map(w => `<div class="gw-warn-item">⚠ ${escapeHtml(w)}</div>`).join('');
  return `<div class="gw-warnings"><div class="gw-warn-head">格式校验提示（${warnings.length}）</div>${items}</div>`;
}

/* ---------------- 编辑器绑定 ---------------- */
const $ = id => document.getElementById(id);

function init() {
  const editor = $('editor');
  const preview = $('preview');
  const statusText = $('status-text');
  const editMeta = $('edit-meta');
  const previewMeta = $('preview-meta');
  const toastEl = $('toast');
  let toastTimer = null;

  function showToast(msg, isError) {
    toastEl.textContent = msg;
    toastEl.className = 'toast' + (isError ? ' error' : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.classList.add('hidden'); }, 2600);
  }

  let lastSource = '';
  let previewZoom = 1;
  function fitPreview() {
    const scroll = document.querySelector('.preview-scroll');
    if (!scroll) return;
    const avail = scroll.clientWidth - 32;
    const pageW = 210 * 96 / 25.4;
    previewZoom = Math.max(0.5, Math.min(1, (avail - 8) / pageW));
    preview.style.zoom = previewZoom;
  }
  function refresh() {
    const src = editor.value;
    if (src === lastSource) return;
    lastSource = src;
    preview.style.zoom = 1; // 测量/渲染用真实尺寸，保证分页准确
    let parsed;
    try {
      parsed = GWMD.parse(src);
    } catch (e) {
      preview.style.zoom = previewZoom;
      preview.innerHTML = `<div style="color:#c33;padding:12px">解析错误：${escapeHtml(e.message)}</div>`;
      return;
    }
    preview.innerHTML = renderWarnings(GWMD.validate(src)) + renderBlocks(parsed);
    preview.style.zoom = previewZoom;
    const chars = src.length;
    const lines = src ? src.split('\n').length : 0;
    editMeta.textContent = `${lines} 行 · ${chars} 字`;
    const blockCount = parsed.blocks.length;
    previewMeta.textContent = `${blockCount} 个块 · ${Object.keys(parsed.footnotes).length} 条脚注`;
  }

  let debounceTimer = null;
  editor.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(refresh, 120);
  });

  /* ---- 编辑区字号调整（12–28px，记忆上次设置） ---- */
  const FONT_MIN = 12;
  const FONT_MAX = 28;
  const FONT_KEY = 'gwmd.editorFontSize';
  function clampFontSize(px) { return Math.max(FONT_MIN, Math.min(FONT_MAX, px)); }
  function applyFontSize(px) {
    const s = clampFontSize(px);
    editor.style.fontSize = s + 'px';
    $('font-dec').disabled = s <= FONT_MIN;
    $('font-inc').disabled = s >= FONT_MAX;
    try { localStorage.setItem(FONT_KEY, String(s)); } catch (_) {}
    return s;
  }
  (function initFontSize() {
    let saved = parseInt(localStorage.getItem(FONT_KEY), 10);
    if (!Number.isFinite(saved)) saved = 14;
    applyFontSize(saved);
  })();
  $('font-dec').addEventListener('click', () => {
    applyFontSize(clampFontSize(parseFloat(editor.style.fontSize) - 1));
  });
  $('font-inc').addEventListener('click', () => {
    applyFontSize(clampFontSize(parseFloat(editor.style.fontSize) + 1));
  });

  /* ---- 导入 .gwmd ---- */
  $('btn-import-gwmd').addEventListener('click', () => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.gwmd,.txt,text/plain';
    input.onchange = () => {
      const file = input.files && input.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        editor.value = String(reader.result || '');
        lastSource = '';
        refresh();
        showToast('已导入 ' + file.name);
      };
      reader.onerror = () => showToast('读取文件失败', true);
      reader.readAsText(file, 'utf-8');
    };
    input.click();
  });

  /* ---- 保存 .gwmd ---- */
  $('btn-save-gwmd').addEventListener('click', () => {
    const src = editor.value;
    const blob = new Blob([src], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const name = guessTitle(src) + '.gwmd';
    a.download = name;
    a.click();
    URL.revokeObjectURL(a.href);
    showToast(`已保存 ${name}`);
  });

  /* ---- 导出 .docx ---- */
  $('btn-export-docx').addEventListener('click', async () => {
    const btn = $('btn-export-docx');
    const src = editor.value;
    if (!src.trim()) { showToast('内容为空，无法导出', true); return; }
    const warnCount = GWMD.validate(src).length;
    if (warnCount) showToast(`有 ${warnCount} 条格式校验提示，见预览顶部`);
    btn.disabled = true;
    btn.textContent = '生成中…';
    statusText.textContent = '正在生成 docx…';
    try {
      const resp = await fetch('api.php', {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: src,
      });
      if (!resp.ok) {
        let detail = '';
        try {
          const j = await resp.json();
          detail = j.detail || j.error || '';
        } catch (_) { /* 非 JSON */ }
        throw new Error(detail || `HTTP ${resp.status}`);
      }
      const blob = await resp.blob();
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      const name = guessTitle(src) + '.docx';
      a.download = name;
      a.click();
      URL.revokeObjectURL(a.href);
      showToast(`已导出 ${name}`);
      statusText.textContent = '导出完成';
    } catch (e) {
      showToast('导出失败：' + e.message, true);
      statusText.textContent = '导出失败';
    } finally {
      btn.disabled = false;
      btn.textContent = '导出 .docx';
    }
  });

  /* ---- 示例 ---- */
  $('btn-example').addEventListener('click', () => {
    editor.value = SAMPLE;
    lastSource = '';
    refresh();
    showToast('已载入示例文档');
  });

  /* ---- 语法弹窗 ---- */
  const modal = $('syntax-modal');
  $('btn-syntax').addEventListener('click', () => {
    $('syntax-body').innerHTML = SYNTAX_HTML;
    modal.classList.remove('hidden');
  });
  $('syntax-close').addEventListener('click', () => modal.classList.add('hidden'));
  modal.addEventListener('click', e => { if (e.target === modal) modal.classList.add('hidden'); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') modal.classList.add('hidden'); });

  /* ---- 初始：空白文档 ---- */
  editor.value = '';
  refresh();
  window.addEventListener('resize', fitPreview);
  fitPreview();
  statusText.textContent = '就绪 — 左侧编辑，右侧实时预览（空文档可点「示例」载入参考）';
}

function guessTitle(src) {
  const m = src.match(/^@title\s+(.+)$/m) || src.match(/^#\s+(.+)$/m);
  if (m) {
    // 剔除标题换行分隔符（| 或 \\|），再清理非法文件名字符
    return m[1].trim()
      .replace(/[\\|]/g, '')
      .replace(/[/:*?"<>]/g, '_')
      .slice(0, 50);
  }
  const d = new Date();
  return `公文_${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

/* ---------------- 示例文档 ---------------- */
const SAMPLE = `@serial 000001
@secret 内部资料★注意保存
@urgent 特急
@header ××集团有限公司文件
@docno ××集团发〔2026〕15号
@signer 签发人：×××
@redline
@blank
@title 关于举办2026年度员工技能|培训的通知
@subtitle （征求意见稿）
@to 各部门、各子公司：

为全面提升员工专业素养和岗位胜任能力，支撑公司业务高质量发展，经集团研究决定，举办2026年度员工技能培训。现将有关事项通知如下：

# 总体要求

通过系统化培训，使员工掌握岗位核心技能，提升跨部门协作效率。**本次培训将与年度绩效考核挂钩**，请各部门高度重视、认真组织[^1]。

> 学以致用、知行合一，是本次培训的基本要求。

# 培训安排

## 培训对象： ## 集团总部各部门及所属子公司全体员工。

## 培训时间： ## 2026年9月至11月，分三期开展，每期5天。

## 培训地点： ## 集团培训中心（总部大楼三层），具体教室另行通知。

# 培训内容

本次培训设置以下模块：

- 项目管理实务
- 高效沟通与协作
- 数据分析基础
- 职业素养与安全合规

# 实施步骤

1. 报名阶段：8月25日至9月5日，各部门汇总报名表；
2. 培训阶段：9月至11月，按批次组织实施；
3. 考核阶段：每期培训结束后统一考核，成绩记入个人培训档案。

# 有关要求

**（一）** 合理安排工作，确保参训人员按时参加，不得无故缺席。

**（二）** 培训期间严格遵守课堂纪律，手机调至静音。

**（三）** 如因工作需要调整参训人员，须提前3天向人力资源部报备。

# 材料报送

请各部门于9月5日前报送以下材料：

| 序号 | 材料名称 | 报送时限 |
| --- | --- | --- |
| 1 | 报名汇总表 | 9月5日前 |
| 2 | 培训需求问卷 | 9月5日前 |
| 3 | 部门联络人信息 | 9月5日前 |

@blank 1
@attach [员工培训报名表] [课程安排表] [考核评分标准]

@sign ××集团有限公司    
@date 二〇二六年八月二十五日
@stamp

@note （此件发至各部门、各子公司）

@cc 集团办公室，人力资源部，财务管理部，各子公司
@print ××集团有限公司办公室 2026年8月25日印发
@issue 共印20份

[^1]: 培训事务联系人：王××，联系电话：010-88880000，邮箱：training@corp.example.com。
`;

/* ---------------- 语法说明 ---------------- */
const SYNTAX_HTML = `
<p><b>GWMD</b> 是一种面向公文写作的纯文本格式：左侧编写、右侧实时预览 A4 公文效果，点击「导出 .docx」生成符合《党政机关公文格式》（GB/T 9704-2012）的 Word 文档。除特殊说明外，所有标记均在<u>行首</u>书写，一行一个要素。</p>

<h4>文档结构</h4>
<table class="syn">
<tr><th>区域</th><th>要素</th></tr>
<tr><td>版头</td><td>@header 发文机关标志 · @docno 发文字号 · @secret 密级 · @urgent 紧急程度 · @serial 份号 · @signer 签发人 · @redline 红线</td></tr>
<tr><td>主体</td><td>@subtitle 副标题 · @title 大标题 · @to 主送机关 · 正文 · @attach 附件说明 · @sign 署名 · @date 日期 · @stamp 印章 · @note 附注</td></tr>
<tr><td>版记</td><td>@cc 抄送机关 · @print 印发机关和日期 · @issue 发行信息（置于文档末尾）</td></tr>
</table>

<h4>标题层级（自动编号）</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">@title 公文标题</td><td>大标题：小标宋 2号 居中；标题内可用 <code>|</code> 手工换行</td></tr>
<tr><td class="code"># 一级标题</td><td>黑体 3号，自动编号「一、」</td></tr>
<tr><td class="code">## 二级标题</td><td>楷体 3号，自动编号「（一）」</td></tr>
<tr><td class="code">### 三级标题</td><td>仿宋 3号 加粗，自动编号「1.」</td></tr>
<tr><td class="code">#### 四级标题</td><td>仿宋 3号，自动编号「（1）」</td></tr>
<tr><td class="code"># 标题 # 紧接正文</td><td>配对标题：标题与正文同一段（行内标题）</td></tr>
</table>

<h4>版头要素</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">@header 发文机关标志</td><td>红头：红色小标宋 加粗 居中</td></tr>
<tr><td class="code">@docno ××〔2026〕8号</td><td>发文字号：仿宋 3号 居中</td></tr>
<tr><td class="code">@secret 密级★1年</td><td>密级：黑体 3号，左上角</td></tr>
<tr><td class="code">@urgent 特急</td><td>紧急程度：黑体 3号</td></tr>
<tr><td class="code">@serial 000001</td><td>份号：黑体 3号 数字，置于左上角</td></tr>
<tr><td class="code">@signer 签发人：×××</td><td>签发人：右上角</td></tr>
<tr><td class="code">@redline</td><td>红色分隔线（套红）</td></tr>
</table>

<h4>主体要素</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">@subtitle （送审稿）</td><td>副标题：楷体 居中</td></tr>
<tr><td class="code">@title 公文标题</td><td>大标题（见「标题层级」）</td></tr>
<tr><td class="code">@to 各单位：</td><td>主送机关：顶格，冒号结尾</td></tr>
<tr><td class="code">（无标记行）</td><td>正文：仿宋 3号，首行缩进 2 字符；连续行自动合并为一段，空行分段</td></tr>
<tr><td class="code">@attach [附件名]</td><td>附件说明，附件名用方括号包裹：<br>单附件 <code>@attach [任务分工表]</code><br>多附件 <code>@attach [任务分工表][时间安排表]</code><br>自动编号「附件：1. ×××」，续行缩进 5 字符</td></tr>
<tr><td class="code">@sign 发文机关署名</td><td>落款：右对齐；行末半角空格左移半字、全角空格左移一字</td></tr>
<tr><td class="code">@date 成文日期</td><td>日期：右对齐；多行署名/日期自动补空格对齐</td></tr>
<tr><td class="code">@stamp</td><td>印章占位，显示「（印章）」</td></tr>
<tr><td class="code">@note （此件公开发布）</td><td>附注：首行缩进 2 字符</td></tr>
</table>

<h4>版记要素</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">@cc 抄送：×××</td><td>抄送机关：仿宋 12pt，顶部单线</td></tr>
<tr><td class="code">@print ××办公室 2026年8月25日印发</td><td>印发机关和日期：保留连续空格，可手工对齐（如 2026 前加空格推位）</td></tr>
<tr><td class="code">@issue 发行信息</td><td>发行信息：样式同上</td></tr>
</table>

<h4>行内格式</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">**加粗文字**</td><td>加粗</td></tr>
<tr><td class="code">*斜体文字*</td><td>斜体</td></tr>
<tr><td class="code">[^1]</td><td>脚注引用（角标）</td></tr>
<tr><td class="code">[^1]: 脚注内容</td><td>脚注定义：导出为 Word 页面底部真脚注；无定义的引用会被忽略</td></tr>
<tr><td class="code">[链接文字](https://…)</td><td>超链接</td></tr>
</table>

<h4>列表 · 引用 · 表格</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">- 无序列表</td><td>无序列表</td></tr>
<tr><td class="code">1. 有序列表</td><td>有序列表（也支持「1、」）</td></tr>
<tr><td class="code">&gt; 引用内容</td><td>引用：楷体，左侧竖线</td></tr>
<tr><td class="code">| 表头1 | 表头2 |</td><td>表格：Markdown 语法；<code>|----|</code> 分隔行自动忽略</td></tr>
</table>

<h4>空行 · 分页 · 注释 · 转义</h4>
<table class="syn">
<tr><th>写法</th><th>说明</th></tr>
<tr><td class="code">@blank 2</td><td>插入 2 个空行（范围 1–20；别名 @spacer / @space / @sp）</td></tr>
<tr><td class="code">@pagebreak</td><td>分页符：强制分页</td></tr>
<tr><td class="code">@comment 草稿批注</td><td>注释：不导出（<code>&lt;!-- 批注 --&gt;</code> 同样不导出）</td></tr>
<tr><td class="code">\\* \\@ \\# \\| \\[ \\] \\^ \\&gt; \\( \\\\</td><td>转义：输出字面字符。<br>例：<code>\\@title</code> 显示为 <code>@title</code>，转义后不再作为指令</td></tr>
</table>

<h4>自动编号规则</h4>
<ul class="syn-list">
<li>四个层级分别连续编号：一、→（一）→ 1. →（1）；进入下一级后子级序号自动复位</li>
<li>若自己写了编号（如「一、」「（一）」「1.」「（1）」），会被识别并接续，不会重复添加</li>
<li>标题层级不可跳级（# 后直接 ### 会提示），建议逐级使用</li>
</ul>

<h4>格式校验提示</h4>
<ul class="syn-list">
<li>缺少 @title 大标题 / 存在多个 @title</li>
<li>未知指令（如 @xxx）会按正文处理并提示</li>
<li>标题层级跳跃</li>
<li>脚注 [^n] 写了引用但无定义</li>
<li>@stamp 之后仍有正文内容</li>
</ul>

<h4>导出 .docx 排版标准</h4>
<table class="syn">
<tr><th>项目</th><th>标准</th></tr>
<tr><td>纸型版面</td><td>A4；版心 156×225mm（上 3.7 / 下 3.5 / 左 2.8 / 右 2.6 cm）</td></tr>
<tr><td>正文</td><td>仿宋 3号（16pt），行距 29 磅，每面 22 行、每行 28 字，首行缩进 2 字符</td></tr>
<tr><td>大标题</td><td>小标宋 2号（22pt）居中，行距 33 磅；标题内 <code>|</code> 换行保持居中</td></tr>
<tr><td>页码</td><td>4号宋体，一字线（— n —），奇数页右空一字、偶数页左空一字</td></tr>
<tr><td>落款/日期</td><td>右对齐；多行自动用全角空格补齐对齐</td></tr>
<tr><td>脚注</td><td>注入为 Word 页面底部真脚注（4号 仿宋）</td></tr>
<tr><td>版记</td><td>仿宋 12pt，顶部单线，行距 24 磅，置于末页末尾</td></tr>
</table>

<h4>完整示例</h4>
<pre class="syn-example">@header ××集团有限公司文件
@docno ××集团发〔2026〕8号
@secret 密级★1年
@redline
@title 关于印发《公文管理办法》的通知
@to 各部门、各下属单位：
现将《公文管理办法》印发给你们，请认真贯彻执行。
# 总体要求
## 适用范围
本办法适用于集团各部门及下属单位。
## 施行时间
本办法自2026年9月1日起施行[^1]。
@attach [公文管理办法][修订说明]
@sign ××集团有限公司
@date 2026年8月25日
@stamp
@note （此件公开发布）
@cc 办公室，人力资源部
@print ××集团有限公司办公室 2026年8月25日印发
[^1]: 原《公文管理办法（试行）》同时废止。</pre>
`;

/* ---------------- 启动 ---------------- */
if (typeof document !== 'undefined') {
  document.addEventListener('DOMContentLoaded', init);
}

// Node 环境导出解析器与宽度判定（供前后端一致性测试使用）
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { parse: GWMD.parse, isWideChar, validate: GWMD.validate };
}
