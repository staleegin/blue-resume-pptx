// 单页 A4 简历 PPT 生成器
// 用法: node build_resume.js data.json out.pptx
const fs = require('fs'), path = require('path');
const pptxgen = require('pptxgenjs');
const JSZip = require(require.resolve('jszip', { paths: [path.dirname(require.resolve('pptxgenjs'))] }));
const React = require('react'), RDS = require('react-dom/server'), sharp = require('sharp');
const fa = require('react-icons/fa');

let D = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const OUT = process.argv[3] || 'resume.pptx';
const BASE = path.dirname(path.resolve(process.argv[2]));

// ---------- 常量 ----------
const P = v => v / 72;
const W = 595.25, H = 841.85, BOTTOM = H - 28;
const THEME = (D.theme || '2563EB').replace('#', '').toUpperCase();
const tint = (hex, a) => [0, 2, 4].map(i => Math.round(255 - (255 - parseInt(hex.substr(i, 2), 16)) * a).toString(16).padStart(2, '0')).join('').toUpperCase();
const CARD = tint(THEME, 0.1);
const CJK = 'SimHei', LAT = 'Arial', YH = 'Microsoft YaHei', GREY = '595959', BLACK = '000000';
const L_X = 25.6, DATE_R = 380.4, BUL_X = 30.1, BUL_IND = 13.5, BUL_W = 359, PARA_X = 24.6, PARA_W = 356;
const R_X = 403, R_W = 174, R_IN = 411.9, R_TXT_W = 140;
const HEAD_X = 32.5;   // 顶栏里姓名、求职意向、分隔线、简介的左边距
// 自动微调档位：正文字号 / 标题字号 / 行距 / 间距系数
const LEVELS = [
  { fs: 9, ts: 9.9, lh: 15.1, g: 1 },
  { fs: 9, ts: 9.9, lh: 14.2, g: 0.85 },
  { fs: 9, ts: 9.9, lh: 13.4, g: 0.7 },
  { fs: 8.5, ts: 9.4, lh: 12.6, g: 0.7 },
  { fs: 8, ts: 9, lh: 11.8, g: 0.6 },
];

// ---------- 语言模式：英文简历用 Arial、正文不加粗 ----------
const strip = t => String(t || '').replace(/\*\*/g, '');
const ALLTXT = JSON.stringify(D);
const EN = D.lang ? D.lang === 'en' : (ALLTXT.match(/[一-鿿]/g) || []).length < ALLTXT.length * 0.03;
const BODY = EN ? LAT : CJK;   // 正文字体
const BB = !EN;                // 正文是否加粗（中文版式全加粗，英文只加粗标题和重点）

// ---------- 文字宽度估算 ----------
// 黑体：汉字 1em、半角 0.5em；Arial 用实际字宽表（千分之一 em，ASCII 32-126）
const ARIAL = [278,278,355,556,556,889,667,191,333,333,389,584,278,333,278,278,556,556,556,556,556,556,556,556,556,556,278,278,584,584,584,556,1015,
  667,667,722,722,667,611,778,722,278,500,667,556,833,722,778,667,778,722,667,611,722,667,944,667,667,611,278,278,278,469,556,333,
  556,556,500,556,556,278,556,556,222,222,500,222,833,556,556,556,556,333,500,278,556,500,722,500,500,500,334,260,334,584];
function cw(ch, size, arial, bold) {
  const c = ch.codePointAt(0);
  if (c >= 0x80) return size;
  if (!arial) return size * 0.5;
  return size * (ARIAL[c - 32] || 556) / 1000 * (bold ? 1.08 : 1);
}
// arial：半角字符是否用 Arial 渲染；默认跟随语言模式
const textW = (t, size, arial = EN, bold = BB) => [...strip(t)].reduce((w, ch) => w + cw(ch, size, arial, bold), 0);
// 断行：英文按单词、中文按字，模拟 PowerPoint 的换行
function wrap(t, size, width, arial = EN, bold = BB) {
  const toks = strip(t).match(/[\x21-\x7E]+|\s+|[^\s]/gu) || [];
  const lim = width * (arial ? 0.97 : 1) + 0.5, lines = [];
  let cur = '', w = 0;
  for (const tok of toks) {
    const tw = textW(tok, size, arial, bold);
    if (/^\s+$/.test(tok)) { if (cur) { cur += tok; w += tw; } continue; }
    if (w + tw <= lim) { cur += tok; w += tw; continue; }
    if (cur.trim()) { lines.push(cur.trimEnd()); cur = ''; w = 0; }
    if (tw <= lim) { cur = tok; w = tw; continue; }
    for (const ch of tok) { const c = cw(ch, size, arial, bold); if (w + c > lim && cur) { lines.push(cur); cur = ''; w = 0; } cur += ch; w += c; }
  }
  if (cur.trim() || !lines.length) lines.push(cur.trimEnd());
  return lines;
}
const nLines = (t, size, width, arial, bold) => wrap(t, size, width, arial, bold).length;
// 标签排行：每个标签不拆开，用「 · 」连接；单个标签超宽时才按普通文字断行
function tagLines(tags, size, width) {
  const SEP = ' · ', lines = []; let cur = '';
  for (const t of tags.map(x => String(x).trim()).filter(Boolean)) {
    const cand = cur ? cur + SEP + t : t;
    if (textW(cand, size, true, BB) <= width) { cur = cand; continue; }
    if (cur) lines.push(cur);
    if (textW(t, size, true, BB) <= width) cur = t;
    else { const w = wrap(t, size, width, true, BB); lines.push(...w.slice(0, -1)); cur = w[w.length - 1]; }
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

// ---------- 富文本：**加粗** -> 黑色；数字成果自动加粗 ----------
// 数字 + 单位（中文量词、%、万/K/M/million、percent、x 倍）；货币符号开头的金额。全部用非捕获组
const NUM_RE = /(?:超|近|逾|约)?[$¥￥€£]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?\s*(?:%|％|\+)?\s*(?:万|千|百|亿|million|billion|[KMBkw](?![A-Za-z]))?(?:余)?\s*(?:人次|人|项|个月|个|次|篇|场|倍|名|元|单|秒|分钟|小时|天|件|款|家|所|条|份|位|台|套|门|周|月|年|percent|x(?![A-Za-z]))?(?:%|％|\+)?|(?:超|近|逾)?[百千万]余?人次/g;
function rich(text, color, auto) {
  const runs = [];
  String(text || '').split(/(\*\*.+?\*\*)/).forEach(seg => {
    if (!seg) return;
    if (seg.startsWith('**') && seg.endsWith('**')) return runs.push({ text: seg.slice(2, -2), options: { color: BLACK, bold: true } });
    if (!auto) return runs.push({ text: seg, options: { color, bold: BB } });
    let last = 0;
    seg.replace(NUM_RE, (m, off) => {
      const k = m.trim();
      // 纯数字（型号、编号）、年份「2026 年」、月份「9 月」不加粗
      if (/^[\d.,]+$/.test(k) || /^\d{4}\s*年$/.test(k) || /^\d{1,2}\s*月$/.test(k)) return m;
      if (off > last) runs.push({ text: seg.slice(last, off), options: { color, bold: BB } });
      runs.push({ text: m, options: { color: BLACK, bold: true } }); last = off + m.length; return m;
    });
    if (last < seg.length) runs.push({ text: seg.slice(last), options: { color, bold: BB } });
  });
  return runs;
}
// 拆分中英文，英文用 Arial
function mixRuns(t, o) {
  const out = []; let cur = '', lat = null;
  for (const ch of t) {
    const isLat = ch === ' ' ? lat : ch.codePointAt(0) < 0x80;
    if (lat !== null && isLat !== lat) { out.push({ text: cur, options: { ...o, fontFace: lat ? LAT : CJK } }); cur = ''; }
    cur += ch; lat = isLat;
  }
  if (cur) out.push({ text: cur, options: { ...o, fontFace: lat ? LAT : CJK } });
  return out;
}

// ---------- 图标 / 照片 ----------
const ICONS = { phone: 'FaMobileAlt', email: 'FaEnvelope', location: 'FaMapMarkerAlt', wechat: 'FaWeixin', web: 'FaGlobe', birthday: 'FaBirthdayCake', github: 'FaGithub', user: 'FaUser', linkedin: 'FaLinkedin', link: 'FaLink', portfolio: 'FaPalette' };
// 透明 1x1 图片：盖在联系方式胶囊上承载超链接，整颗胶囊可点，文字不变色、不加下划线
const CLEAR = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
// 链接安全：只放行 http/https/mailto/tel，并做 XML 转义（pptxgenjs 写 rels 时不转义，带 & 的链接会让文件损坏）
function safeUrl(u) {
  u = String(u || '').trim();
  if (!/^(https?:\/\/|mailto:|tel:)/i.test(u) || /[\s<>"]/.test(u)) return null;
  return u.replace(/&(?!amp;)/g, '&amp;').replace(/'/g, '&apos;');
}
function contactLink(c) { const u = rawLink(c); return u ? safeUrl(u) : null; }
function rawLink(c) {
  if (c.link) return c.link;
  const t = String(c.text || '').trim();
  if (c.type === 'email' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)) return 'mailto:' + t;
  if (/^(web|github|linkedin|link|portfolio)$/.test(c.type) && /^(https?:\/\/)?[\w-]+(\.[\w-]+)+(\/\S*)?$/.test(t)) return /^https?:/.test(t) ? t : 'https://' + t;
  return null;
}
async function iconPng(name) {
  const svg = RDS.renderToStaticMarkup(React.createElement(fa[name], { color: '#FFFFFF', size: 256 }));
  return 'image/png;base64,' + (await sharp(Buffer.from(svg)).png().toBuffer()).toString('base64');
}
async function circlePhoto(p) {
  const img = sharp(p).rotate(); const m = await img.metadata();
  const s = Math.min(m.width, m.height);
  const sq = await img.extract({ left: Math.floor((m.width - s) / 2), top: Math.floor((m.height - s) / 4), width: s, height: s }).resize(600, 600).toBuffer();
  const mask = Buffer.from('<svg width="600" height="600"><circle cx="300" cy="300" r="300" fill="#fff"/></svg>');
  const out = await sharp(sq).composite([{ input: mask, blend: 'dest-in' }]).png().toBuffer();
  return 'image/png;base64,' + out.toString('base64');
}

// ---------- 排版（S 为空时只计算高度）----------
function layout(S, cfg, assets) {
  const { fs, ts, lh, g } = cfg;
  const T = (x, top, w, size, text, o = {}) => {
    if (!S) return;
    const l = o.lh || size * 1.25, n = o.lines || 1;
    S.addText(o.runs || text, {
      x: P(x), y: P(o.boxY ?? (top - size * 0.12 - (l - size) / 2)), w: P(w), h: P(o.boxH ?? l * n), fontFace: o.font || BODY, fontSize: size,
      color: o.color || BLACK, bold: o.bold ?? true, align: o.align || 'left', valign: o.valign || 'top', margin: 0, lineSpacing: l,
      isTextBox: true, fit: 'none', wrap: true, objectName: o.name || strip(text).slice(0, 20),
    });
  };
  const RR = (x, y, w, h, color, r, o = {}) => S && S.addShape('roundRect', { x: P(x), y: P(y), w: P(w), h: P(h), fill: { color, transparency: o.tr }, line: { type: 'none' }, rectRadius: P(r), objectName: o.name });
  const bar = (x, y) => RR(x, y, 8.5, 2.85, THEME, 1.4);

  // ===== 头部 =====
  const sumLines = D.summary ? nLines(D.summary, 8.1, 521) : 0;
  const headBottom = D.summary ? 128.9 + (sumLines - 1) * 12 + 22.9 : 112;
  RR(20.2, 20.3, 555, headBottom - 20.3, THEME, 10, { name: '顶部背景' });
  T(HEAD_X, 33.4, 300, 17.1, D.name, { color: 'FFFFFF', font: EN ? LAT : YH, name: '姓名' });
  if (D.title) T(HEAD_X, 58.0, 450, 9, (D.titleLabel ?? (EN ? '' : '求职意向：')) + D.title, { color: 'FFFFFF', name: '求职意向' });
  let px = 30;
  (D.contacts || []).forEach((c, i) => {
    const lat = !/[^\x00-\x7F]/.test(c.text), icon = ICONS[c.type];
    const tx = icon ? 18.1 : 8, w = tx + textW(c.text, 8, lat, true) + 8.5;
    RR(px, 81.5, w, 17, 'FFFFFF', 8.5, { tr: 70, name: '联系方式底框' + (i + 1) });
    if (S && icon) S.addImage({ data: assets.icons[c.type], x: P(px + 6.3), y: P(81.5 + (17 - 8.5) / 2), w: P(8.5), h: P(8.5) });
    T(px + tx, 0, w - tx, lat ? 8 : 8.1, c.text, { color: 'FFFFFF', font: lat ? LAT : CJK, boxY: 81.5, boxH: 17, valign: 'middle', name: '联系方式' + (i + 1) });
    const url = contactLink(c);
    if (S && url) S.addImage({ data: CLEAR, x: P(px), y: P(81.5), w: P(w), h: P(17), hyperlink: { url, tooltip: c.text }, objectName: '链接' + (i + 1) });
    px += w + 4;
  });
  if (D.summary) {
    if (S) S.addShape('rect', { x: P(HEAD_X), y: P(112.8), w: P(568.5 - HEAD_X), h: P(1), fill: { color: 'FFFFFF', transparency: 70 }, line: { type: 'none' }, objectName: '分隔线' });
    T(HEAD_X, 128.9, 526, 8.1, D.summary, { color: 'FFFFFF', lh: 12, lines: sumLines, bold: BB, name: '个人简介' });
  }
  if (S && assets.photo) S.addImage({ data: assets.photo, x: P(493.2), y: P(35.0), w: P(62.4), h: P(62.4), objectName: '照片' });

  // ===== 左栏 =====
  const top0 = headBottom + 18;
  let y = top0;
  const auto = D.autoBold !== false;
  (D.left || []).forEach((sec, si) => {
    bar(L_X, y + 2.9);
    T(36, y, 200, ts, sec.title, { name: sec.title });
    y += 21 * Math.max(g, 0.8);
    (sec.entries || []).forEach((e, ei) => {
      // 没有标题、时间和副标题的条目（纯列表/段落）：直接接在板块标题下面，不留空行
      if (!e.name && !e.date && !e.sub && !e.subRight) y -= ei === 0 ? 21 * Math.max(g, 0.8) : 11.7 * g + 18.3 * Math.max(g, 0.75) - 4;
      const dw = e.date ? textW(e.date, fs, EN, true) + 4 : 0;
      const nameW = DATE_R - L_X - dw - 6, nameN = e.name ? nLines(e.name, ts, nameW, EN, true) : 1;
      if (e.name) T(L_X, y, nameW, ts, e.name, { name: (e.name || '条目').slice(0, 20), lh: ts * 1.25, lines: nameN });
      if (e.date) T(DATE_R - 120, y - 1.5, 120, fs, e.date, { align: 'right', name: '时间' });
      if (e.subRight) T(DATE_R - 120, y + 9.3, 120, fs, e.subRight, { align: 'right', name: '右侧副信息' });
      let cur;
      const ny = y + (nameN - 1) * ts * 1.25;   // 标题多行时整体下移
      if (e.sub) {
        const subW = e.subRight ? DATE_R - L_X - textW(e.subRight, fs, EN, true) - 10 : DATE_R - L_X, subN = nLines(e.sub, fs, subW);
        const sy = ny + 12.5 * Math.max(g, 0.85);
        T(L_X, sy, subW, fs, e.sub, { color: GREY, bold: BB, lh: fs * 1.3, lines: subN, name: '副标题' });
        cur = sy + (subN - 1) * fs * 1.3 + 16.7 * Math.max(g, 0.75);
      } else cur = ny + 18.3 * Math.max(g, 0.75);
      let any = false;
      // details：成绩/主修课程/技术栈等明细行，每条一行，冒号前的标签用黑色
      for (const dl of e.details || []) {
        const n = nLines(dl, fs, PARA_W);
        const mm = String(dl).match(/^([^：:]{1,12}[：:]\s?)([\s\S]*)$/);
        const runs = mm ? [{ text: mm[1], options: { color: BLACK, bold: true } }, ...rich(mm[2], GREY, auto)] : rich(dl, GREY, auto);
        T(PARA_X, cur, PARA_W, fs, dl, { runs, color: GREY, bold: BB, lh, lines: n, name: '明细' });
        cur += n * lh; any = true;
      }
      if (e.bullets && e.bullets.length) {
        const counts = e.bullets.map(b => nLines(b, fs, BUL_W - BUL_IND));
        const total = counts.reduce((a, b) => a + b, 0);
        if (S) {
          const paras = [];
          e.bullets.forEach((b, i) => {
            const runs = rich(b, GREY, auto);
            runs.forEach((r, j) => { if (j === runs.length - 1 && i < e.bullets.length - 1) r.options.breakLine = true; paras.push(r); });
          });
          S.addText(paras, { x: P(BUL_X), y: P(cur - fs * 0.12 - (lh - fs) / 2), w: P(BUL_W), h: P(lh * total), fontFace: BODY, fontSize: fs, color: GREY, bold: BB, valign: 'top', margin: 0, lineSpacing: lh, isTextBox: true, fit: 'none', objectName: '列表_' + (e.name || '').slice(0, 16) });
        }
        cur += total * lh; any = true;
      }
      if (e.text) {
        const n = nLines(e.text, fs, PARA_W);
        T(PARA_X, cur, PARA_W, fs, e.text, { runs: rich(e.text, GREY, auto), color: GREY, bold: BB, lh, lines: n, name: '描述' });
        cur += n * lh; any = true;
      }
      if (!any) cur -= lh * 0.3;
      y = cur + 11.7 * g;
    });
    y += 1.6 * g;
  });
  const leftBottom = y - 13.3 * g;

  // ===== 右栏卡片 =====
  let ry = top0 - 1.2;
  (D.right || []).forEach(card => {
    const c0 = ry;
    let it = c0 + 28.3 * Math.max(g, 0.9), lastTop = it;
    // tags：标签式卡片，没有小标题和三个点，标签整体不拆开
    const items = card.tags ? [{ name: '', tags: card.tags }] : (card.items || []);
    const ops = [];
    items.forEach(item => {
      // 按宽度断行（英文按单词），断好的行用换行写入，PowerPoint 不会再自行折行
      const lines = item.tags ? tagLines(item.tags, 8, R_TXT_W)
        : String(item.content || '').split('\n').flatMap(l => wrap(l, 8, R_TXT_W, true, BB));
      const dW = item.date ? textW(item.date, 7.5, true, true) + 6 : 0;   // 卡片小项右侧日期
      const nameN = item.name ? nLines(item.name, ts, 150 - dW, EN, true) : 0, shift = Math.max(nameN - 1, 0) * ts * 1.25;
      const t0 = item.name ? it + 25.3 + shift : it;
      ops.push({ it, t0, item, lines, nameN, shift, dW });
      lastTop = t0 + (lines.length - 1) * 9.6;
      it = lastTop + 18.3 * g;
    });
    const bottom = items.length ? lastTop + 15.3 : c0 + 30;
    if (S) {
      RR(R_X, c0, R_W, bottom - c0, CARD, 6, { name: (card.title || '') + '卡片' });
      bar(410.8, c0 + 12.7);
      T(422.9, c0 + 10.6, 140, ts, card.title, { name: card.title });
      ops.forEach(({ it, t0, item, lines, nameN, shift, dW }) => {
        if (item.name) {
          T(R_IN, it, 150 - dW, ts, item.name, { name: item.name, lh: ts * 1.25, lines: nameN });
          if (item.date) T(R_X + R_W - 8 - 70, it + 0.8, 70, 7.5, item.date, { color: GREY, align: 'right', font: LAT, name: '时间' });
          [412.0, 423.0, 434.2].forEach(x => bar(x, it + 15.9 + shift));
        }
        const runs = [];
        lines.forEach((l, i) => { const r = mixRuns(l, { color: GREY }); if (i < lines.length - 1) r[r.length - 1].options.breakLine = true; runs.push(...r); });
        S.addText(runs, { x: P(R_IN), y: P(t0 - 8 * 0.12 - 0.8), w: P(R_TXT_W + 15), h: P(9.6 * lines.length), fontSize: 8, bold: BB, valign: 'top', margin: 0, lineSpacing: 9.6, isTextBox: true, fit: 'none', objectName: item.name + '内容' });
      });
    }
    ry = bottom + 12;
  });
  const rightBottom = ry - 12;
  return Math.max(leftBottom, rightBottom);
}

// ---------- XML 后处理：每段加项目符号、去掉多余 pPr ----------
function postXml(x) {
  return x.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, sp => {
    const isBul = /name="列表_/.test(sp);
    const ln = (sp.match(/<a:lnSpc>[\s\S]*?<\/a:lnSpc>/) || [''])[0];
    const BUL = `<a:pPr marL="171450" indent="-171450">${ln}<a:buSzPct val="100000"/><a:buFont typeface="Arial"/><a:buChar char="&#x2022;"/></a:pPr>`;
    return sp.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, body) => {
      const first = body.match(/^(<a:pPr[^>]*\/>|<a:pPr[^>]*>[\s\S]*?<\/a:pPr>)/);
      const keep = first ? first[0] : '';
      const rest = body.slice(keep.length).replace(/<a:pPr[^>]*\/>|<a:pPr[^>]*>[\s\S]*?<\/a:pPr>/g, '');
      return '<a:p>' + (isBul ? BUL : keep) + rest + '</a:p>';
    });
  });
}

(async () => {
  const assets = { icons: {}, photo: null };
  for (const c of D.contacts || []) if (ICONS[c.type] && !assets.icons[c.type]) assets.icons[c.type] = await iconPng(ICONS[c.type]);
  if (D.photo) {
    const pp = path.resolve(BASE, D.photo);
    if (fs.existsSync(pp)) assets.photo = await circlePhoto(pp); else console.warn('照片不存在：' + pp);
  }
  // 选最松的、能放下的档位；都放不下就用最紧一档并记录超出量
  const fit = () => {
    const lv = LEVELS.findIndex(c => layout(null, c, assets) <= BOTTOM);
    if (lv >= 0) return { level: lv, overflow: 0 };
    return { level: LEVELS.length - 1, overflow: layout(null, LEVELS[LEVELS.length - 1], assets) - BOTTOM };
  };
  let best = fit();
  // 左右栏平衡：原始版式放不下时，把左栏的短板块（教育、证书、荣誉、语言等）挪成右栏卡片
  const moved = [];
  const MOVABLE = /教育|学历|证书|资格|荣誉|奖|语言|兴趣|爱好|社团|技能|education|certif|licen|award|honou?r|language|interest|affiliation|skill|activit/i;
  const toCard = sec => ({ title: sec.title, items: (sec.entries || []).map(e => ({
    name: e.name || '', date: e.name ? e.date : undefined, content: [e.name ? '' : e.date, [e.sub, e.subRight].filter(Boolean).join(' | '), ...(e.details || []), e.text, ...(e.bullets || [])].filter(Boolean).join('\n') })) });
  if (best.level > 0 && D.balance !== false) {
    const size = sec => JSON.stringify(sec).length;
    const cands = (D.left || []).filter(s => s.side === 'right' || (s.side !== 'left' && MOVABLE.test(s.title))).sort((a, b) => size(a) - size(b));
    for (const sec of cands) {
      const prev = D;
      D = { ...D, left: D.left.filter(s => s !== sec), right: [...(D.right || []), toCard(sec)] };
      const r = fit();
      if (r.level < best.level || (r.level === best.level && r.overflow < best.overflow - 1)) { best = r; moved.push(sec.title); }
      else D = prev;
      if (best.level === 0) break;
    }
  }
  const { level, overflow } = best;
  const cfg = LEVELS[level];
  const warns = [];
  if (D.summary && nLines(D.summary, 8.1, 521) > 3) warns.push('个人简介超过 3 行，顶栏会被撑高，建议压缩到 2～3 句');

  const pres = new pptxgen();
  pres.defineLayout({ name: 'A4P', width: P(W), height: P(H) });
  pres.layout = 'A4P';
  pres.title = (D.name || '') + ' 个人简历';
  const S = pres.addSlide();
  S.background = { color: 'FFFFFF' };
  const bottom = layout(S, cfg, assets);

  const zip = await JSZip.loadAsync(await pres.write({ outputType: 'nodebuffer' }));
  const f = 'ppt/slides/slide1.xml';
  zip.file(f, postXml(await zip.file(f).async('string')));
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(JSON.stringify({ out: OUT, level, fontSize: cfg.fs, lineHeight: cfg.lh, bottom: Math.round(bottom), pageBottom: BOTTOM, overflowPt: Math.round(overflow), lang: EN ? 'en' : 'zh', movedToRight: moved }));
  warns.forEach(w => console.log('提示：' + w));
  if (overflow > 0) console.log(`内容超出一页约 ${Math.round(overflow)}pt（约 ${Math.ceil(overflow / cfg.lh)} 行），需要删减。`);
})();
