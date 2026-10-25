/*!
 * resume_core.js —— 单页 A4 简历 PPT 的排版核心（Node 和浏览器通用）
 * https://github.com/staleegin/blue-resume-pptx
 *
 * Node:    const core = require('./resume_core.js')
 * 浏览器:  <script src="resume_core.js"></script>  →  window.ResumeCore
 *
 * core.generate({ PptxGenJS, JSZip, data, photo, outputType })  → { file, info }
 *    photo：已裁成正方形的图片 data（'image/jpeg;base64,...'），可省略
 *    outputType：'nodebuffer'（Node）或 'blob' / 'arraybuffer'（浏览器）
 * core.preview(data, { photo })  → { items, info }   用同一套排版算出所有元素，供网页画预览
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ResumeCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
'use strict';

function createEngine(data, assets) {
let D = JSON.parse(JSON.stringify(data || {}));
assets = assets || {};
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
// 联系方式图标（Font Awesome Free，CC BY 4.0；白色 96px PNG，内嵌以免依赖网络和 react-icons）
const ICON_DATA = {
  phone: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAApVBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////Mid1sAAAANnRSTlMA+O3WmTDzLwIMNu7jJC7YOD1C15MxwUUm9JL5wAFEJQuRm0PO8pD3vs0ON73rIZzkzzm/PpqWWsoUAAAACXBIWXMAAAsTAAALEwEAmpwYAAABOElEQVRo3u3aSXOCQBQE4EF2RdS4xX1fkrhknf//03JIApdM2fPIHFLVfW7ehyhgwSj1e4br1kjDGXU/5som9wNtnekCn//e0II0OvD+i+ZrXWuCwFELM8DmD7U4Gwi4yoEAAlpyoAsBkRyIIKCoP+QzpD+Lt8UmdkCM/q5DIZCgQCIE8FOfAAECBAgQIECAAAECBAgQIECAAAECBAj8CeD8yS/87LovBLYh9BmS/osQEIQAgVvAYV9fLp9Wnisgy78KceYGOOQ/jdhzAuzLytgJUC8rbSdAWlZS18Dj/zxEq7LSqwTsDBt7xQ0onBgqO1XphXX2LYTPpoYPAWvjeeqN22n62ptUfOU+l1/r7rB7+Ek6fwr+SbicHS/cUJ2aaD689ESpxZv9/GNT2WQT+BbLHyI/MH2/n3mtaBdgTd8QAAAAAElFTkSuQmCC',
  email: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB3VBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////49NrfAAAAnnRSTlMACfr5/iSAAe5vRPOR/BD9Be1CBLOZwYixJTTeI9cT+x3CA9YmEcqr+Hc+Q6rYFDjdvj1m7EzxcBjhf0Eix+sIrfANVTbboiECRQfvDPQwFrvfwOTPU5s5GuPOpOdHzdIOUfYyuFa9PKPmhCcs0fUSmLzgHOXqyytPVJc1nDobXoIgpumHsE7UtZbZv10fhsmLr03y03EPtFIzmlziwwYVdiYAAAAJcEhZcwAACxMAAAsTAQCanBgAAALxSURBVGje7ZnnWxpBEMaXfjSR3hQUItKMiC0ae01sicYSTe+999577/3+1mBk96jHnew8j0luPvFwc785luV9ZxaEpJBCCinWSmyeb5hkqcZk9Xcdwa9zsCAx417hRxkWKJjkn+cH47OsWZUq0MQChiP1/bKgEUM3YQvYUQNsgWqkhy2gRyxw/AMFjLB8I5qGLZBAD2ELvEYTs5D8pQmEKlxwfFfFspxqtHKgL1i9Ke040WMQ/EucpaEt5+jza7dnmrJVbaKLVxoqc3z/PVVja7uW31m8CtHjz6o4rmefJv1K9owW/7osjdQcGkIM+6gOF/tKZZmYOczzPWWZZTX9uYDfqaktn/+mC9NueLFc+yMaspuMtHaPRqvI8AMXWaaN28rhP9iFOXvrsw2noxFfWW9bPX/rBkz5lsh1NHkEZX828eEn2oOaFQUss76CbN3h1fCHPfj+ur7CntxxBGcMHBXP555vKFHM9P3qAE66I7Jf2q/lVljO01U0EQk8sSSG/+QkEeaD/G1LTyvOHDMoBfNHsDYg3c5SfZHSMIaTH7cIw/cP4jsq81W/QOO1243zLwiarbpHcf69bmGd3YeX+I5ApKRdG53EWQbHBbeOcbLlGr38/MUozpQZxPSm00Q5Psd5+xIfzjvwXFzzq/iEBVbTXnSZFO0kSRsW3V1zwuUpskyLtwQIJE/7ftFC7v9S6LrNJ0TikTDzaD6dezGsxcsT4DUp/gEkVIUf8sqenOFuSqDNlphwxq+SXXg8wyXCd4k2tJrLG6GUb63kzOR+WgdMO8hv3epUlj2j/aohPZSq8zLDhDq5ruq2g8YQ2P+j2DHTXAulKTN+phB+wE5vjD2vy+dbvFTn5HeqbHxXH+1B3GRb4PBTI3KISf+jU1fV21sVc576v84qpAJSgb+gAPCxZhv0wWwQzcMWOIt0sAVSvdVhSP5MSttHe1jYPyhQ0gzGT674k/sFDL+Js9mYPUj596AP2i3S34dSSCHFmonfS8GeRXnvHuoAAAAASUVORK5CYII=',
  location: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACN1BMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ZmOKXAAAAvHRSTlMAAQMFFwf5/fz+9xuu9ghzBvAE+x1QAu8MFvjZFe3xiUKp7AlRdTfCEuU65CDuEFfdez2AHsPSinCzNc1MLddW3HpbqNGTdLzBRYisK5J4fT9EJatyd1g5fMSGJM4K0w+Vdjhdb7uwY6QjyGaPlDITuL3hYOaEIsdlRusnzPXW+jZaO+B+X8Zkp+rLSo3QsU/0MNURVJfawF6CxaaHSYzPTi/Ult6/46VD6Its8nG0M9gUuZq+XOJhurIcNIPCva0AAAAJcEhZcwAACxMAAAsTAQCanBgAAAPFSURBVGjetVllc9tAEJVjxxCyXYcccpiZuYE2TcNMDWOZmZmZmZkZ78d1pp1OM/bu6UB+n9++N5JOewuKwgaf8rmS2eLcLKMxK7d41jX3y0fRED9Ly2KIG+w/pv21UU+YmjEREObvX4Kk5cOn0wgFJwrGpOT1SdFEBdFWvbh+/ihhwPtOQfmgcQthgqXVV0Q/cyVhRsUyfv3OGMIBeyyvfqIf4YIxh09/lYlwwrSdRz/ETLhhWsiuHxtBBGDcwarfFUOEYF/Bpu9bQQSxm+1/GCfCaGXKDxZxA0uHuv6CRUQCo+qZL4lI4Y1q/o+XM4gPVzHoJZKwqtyPL/BQc2PPpFOvd05aUyg/+mb6LTqFBgbejvpPc7oCUWId1WAGC4tzKyGCuzHmepq+P/LslgKdO1XXiyRcE62aKeVJlDmIwyuKQRkc0gaz82B2AOUvtsPvH6HrmkF6Fl5VlsPnB32pwfC9+g41OAbyh/BHdoEB31B+Cfh/ReEGBvDUNaH8sxD9Gu1cF/F9ZTBT99AMGqCIeyj9EERPphmkQxGPUTqYqqNoBgYoYidKN0J0G7VAgCIiULpZGwMzSg/U5hXtQelhEL2fZnAQbHpQeg33HTgBRUSi9G0QPYVmcBOKGETp9eAnc1I+AXgl3EL5BWDucuEGd8GAAbytgdN1MMYvhNP1fdTgA1yWduuQCycAvmBT8Uc+A1+CtTC7HWYPUw7FYeTSBzu8PuTSz6MY7ML6r1rPsmUA6xOpPXMKVk41u33pwgCM+Zla2V1GK0K/EsO88/8Sb6ND6O1HLqVPLWpIN9hshvSkQUoXfVVlFLZJtnyvU+ugIuX096rO8kLkDCrVu8CPMvpLGNrYKoe4/r4Wlka5SdxgC9sooVpUP5Jxytkh2Oub8lnHLZ/EDB4xz4vGMkT007LZJ1KVIgbneGZq9fz6y7mGgtlfefUzUvnGjks5T5LlDu/g9AqfwWLuye8IV1qtCec2UPo5ZqeOZJHp+BC7wQ2x7cRJVv1hm9gCoYpxAOzXogjCymZQKrxiCd3AdIuFii9x/KPV9cMKZdZQB9QN+uT2aMfV9J9JLuqyN2qb4zzxnLpsMcUq0lhNMzitwbLUtxjXv5igxTr2ghHTj3igzcK3DTOY0GhfHfoE1i/SbCueCa6l7Jna7dyPQAbXtdzqx3FNkAVw6alHHbdOUwOPMsZyStEYa2TLFNUF/MP5+tUjmhsoa+f90I79ihfQzjbvkPihG//pH/XxioGy9fVf/fNdipeQ+KeadIQoXkPy27CwsnKukN8qqy88IW8DWgAAAABJRU5ErkJggg==',
  wechat: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACl1BMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////921Z3zAAAA3HRSTlMA8hox7QT+/P0C+rIGKKffRM73DR0IjgpNvibPR8L7+eHxVY80fUETZusu9jKrARkJvQOttxeiB0+2giSAqRSYQvgrFVao26Oey3nXyK5ayUiwv+DZEQ7aIxznii115mWJUNZuDPSXFrs6IAsQQw9SM7lcw4WWLM0YX+9e2KQnmwVsr6XQ0+5d6oiE5CqQuEDEPRI1V37wb9zRtUzSmTdbIkZqnWmsjbGSMNVzeJzAyoxt8y9TuljefONiqotxtJWaOJ/iYUVKhnTdzHAfS4Nnx3fGSejlJaE7s2DUc/cwgAAAAAlwSFlzAAALEwAACxMBAJqcGAAABH5JREFUaN7tWWVbWzEUTimlt0ChUGC4DtmGDZiwYWPI2AZDxgZzd3dnBnN3d3d3d3fPj1kF1ptcS9LyZU/fT3Cek/dtcpOTc04AcMIJJ/57zEncEBq9x/DIT+Nwas+x97xOt4FN0EZ6dAo0Oox9RFl2FBQifPpAP0fQnxqlglLgzhfYu1pbMqE8Iod72kHvl+4GFeE9LZmRXjdcBYnQtjMTf85USIo2gxn4Wz6BFFjYgpY/LgpSoVdXOv6hXSAlerjQ8G/yhdToPZqcP9gVMqCnPyl/Bx/IhPakAgHouC73h9wsDRfQuXplrb6AmlYQblAOGXXN8vXUrTD+gAlm82FUsx2RQBgyaHZjrDmA8q9rdPZCrMUk/IEo0/VGs38E3+rWtGU6ouF1FoFAECowpcneD1mMf+uJeicp87fG1nqr6Ay4HNEZwHDl49YfE/DpYLWXYd9Y9BtA2F1RYJggGOvN5kR8Fx237KIluLePEv9B4fnRZgyJrdAKz0Fh1ovXwltUKehFsx1i7XKv9x+i35UXJLxVEOjEwP5tXXkf4jh3h5adK43TmcYZ87p3qyxaGRXREFmSPWOo9JEeQ0n/o8Y06Punj/g3GjZjirhAL7oIvQUAzdZK8dwjJEHsHg2hoHeP1QBdWV9ph0FLPJUChRxa/TSde4UZF+XhAunkV2QN0G1zV/Jyq8PWaRQpf98UUL2bxHFvMCIwmJC/qhak/CFzdd0vF0ylcrlEULuTdLKX43gCaWQJ12/QZxf5dlCpeQqPSUb8Asm7ac5LP941kUfyi2rBDroTH6SzJe3eyu6rQQ1t6nfGNoXbis7bq0ESYshQj0iMQY7EEf2ksdn8CKJK+SegWaAkEANaItGnrWX+0TzbPgvTNr7TQ8m8RYh8UIz8b7COK7FZrOVn2jyeky8vgneT5/cB8WhyP8A6zMN2v6WKRM5Ym4DxhqzAKrBCLJPo2mCzTLZY8pFLop53FtSy1Z8BFGLHeqJpzLiN/DhtzvDUWCDnl4qhclEyHuA7mcu8m47uW99jFzM4zGs+PyTN5ST5vwKjG2RBABJVCyKk/D6TBkQcHujFEDpSwm892MQmMB272wb0FvdLAl/YBJbjt6exUNSvAmxgE6gUFoOiq7QRqNkEhLXDc/GKG2xmE+gvEBAPGduTwRomgUCcf7FEOb4WVLDwc9W4QGhT6Go/OXZplc0zASSwCGQKVqiukd6a18QbLsWcCJv9qr7kDXDRMgh0xPk15qJJW7xILEVeSs9fJGgfGkz0CxdJdJPoBa4KSHK5ZXpHlREQhglbgjF6mULoCuVXUFH3DSlruWnUPcPFPWn4TzJ0JVOqKG6CNJa+av5KUv6pqWydYf0twt/vDxgxjqi3l2tHgz71nHIletS+9n+5Ui4+0973ixZn5XusC+x/Ipk0ULZYC3bAK4xu4kz8jtPuerDe+tdLxzwkJa99mhvkbeneuvcozTpkbu88CxGPcvZgfLzLeN7M5kea5pLTnE97njvmwbnN+3q4edWyZn+gdL7ROuHEf4a/6whiPQ08kqkAAAAASUVORK5CYII=',
  web: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACdlBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8rISQAAAA0XRSTlMA+y3O/Qn29wED8PX+ye9xHPoCCivpzAYdNvkfYzDPbTPigKXLcvyiNe34HiRzdnCj6zHbFxsgqyrzL+gYkya7qJVFtAdmGj7nTJk3Ro50eZ3q0PR91LWBSCnyLhmfVifFFdGmERAEWVoPamxHa/GUMnpbf+aEiUssb6eNbpLVoe5OkTR344uQFLmarFiWFijXYMdEyqDTUg5BI55Qv2IFOXxUqeyKrlF1E9wiCLw7fuWDIXjfXuSCJUkLU95nqq9XOHtcpJc9zWUNt8FAur4/OsZo1BYAAAAJcEhZcwAACxMAAAsTAQCanBgAAAVpSURBVGje7Vr5X1NHEH/BkASEkCAQQI4iQgBBRUE5xCIqiqIWsVpBRaCFel9tvWq11VZrq622VuvZ+77v+77v/Y+azOTBe7uzb19I8vn4A/NTdr47M2/ZnWNn0bSbirZcLCu7eClh6uf9nM9CVHWmLq5qp1VEftzRxSKU3Rlh5eXGqv3O3/6sbMefHzvYCPn3Is/nvVZwVwzqnwhWMXYOfz9Yzgzk2I/cxYy5L8wfo/qsQWdI1XkP2upmJmpErZ5fQ7/Lf1o/BvWTX3SCpiUwCqxmHE0IANAGA+ehaDcj5WAj6lmI4x+YQKsQuRVHPcWZ0ejvTdfVPIcnySUamIgf/aw+Pn/AtnrPa35dajNyFjCCdiGWo4/zM2zqnzQ4quRV4NRXUQbceH6HRzkzk+3o900dlWgeAtYRRtIK3P+uUU7RHLX+5bUGFS8Dq85FG5jYD/BDBlZ2hUr/LceMKpKAV8YkdBDgJCMrNU/x/anG2UXILJIZyEH8ASPPe9z+97M16NFMSrjCe5ndNVyupOQfkRu4Dz/LzKydLNP/xYB55hvoFe/JDTwMgcpTwkURyWn1tHLiZ4BdyJjqb3Qbxy2lDZzgpb8E9nQrAzUw5W6eTfr0bgc3y/m6MZrR9Cg6ShrH9jcR8TOHF16NgcNtZcCVLB7UMKWnCAZEb1qGmZhZUie5CYzdL+SXRmHOOeKQC7ROz5wcdX/AGfhWlG0AYK21AUxIvSIw16y/KU0MZZiillobuAc3UNyotCSTgbmi6FRMBUxBG2DaBBFoNepf7xQnbMRqRWWgA6YdFQFnlsHATEJyJSBTVAbehWkrCSSojdM4xY0WkfHyCGDFTElfw8RdFFT1FmBLSLl99tyAsT0WE98BbBWJFQPWojawwmKp0wF7jMQeB6xVbeAonTUN8ehJEmsD7G21gbUwsZPEnoE9dpDYblmUFLIjVRvpFf2NEFRNy1WD3Gy1gdth4nwaDNdg/9LQNJArURuYjSmXBv+Tbg+bB3LNagMl2AWgwRdC0FYawvu1V20gG2/sNFgQgmpoCAsel9pAKpZPNLgtBD1PQ1g5lasNuDHtM6mnJdxAwv9ECd/kYRqqi9sxpaM1y42Lo4U7EZcTGSrC8SBAB7tZcQl2Q/Jw/ZT6emOsr7eT2ICs8A3Rd4AdUhu4ABP3khjeBj8ksa3xSJk1Fkl/jd2kj52vdST2vUXZchKwbWoDi+UtK1dgvCgcp/iR/BK4T3VK0V9mEMigrWtsh8rAfkmzgruJfyO7iLfbu4g/reol9MbWSsjsFhdQaN4FoobPst8MaRCBV/iXmh4yYWva+9YGpsCkAnH9b/IHSTwHp+00pPAacZqpu3Yp6eTtwV5LbUA4IWJLTWvy89vUb7cpeIo/5Y5CytsyeOlNdtuam3h2Ge3PG8muoI3G7DJ7jVktmStSPsKG8Ceq1rL2lZlbNEkWkubUUrWRsjnexDXHLZ5BZpnb75+q2vsHiIRv/URx3CtWhfLyLvJAYXrE8yoeWcxrKKSPl7m9a1ph6nJVaqjIFmqXs7JHolMAHzawuvLUyafe4NIvLRLemQx0GMChpdE9c4VOa5BvidST4cK9gfcyew914cCXb846dN3WgtjIgvNn2M/RDZt1qb9gnEs9lvoA+mXk2bMhmiogc8dOFPsbx3tEAycQ+QNHO3dkRllo+EqdhotIQPCFnBuG115nqW8MtUxfMHy5ysFw08c9YPR8hoEqnO3Lg31jLJeqr7j18lzrMF3lHJ9req50X6mO5V9Brv/jjaz+R0NC8g8jr73y2vUtsVZ9Pt05t4/4eHOb7va5cS0wzy6Al3F3S3/CathLVzMyrv5+c9XV/wMznaVgIelZVAAAAABJRU5ErkJggg==',
  birthday: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABOFBMVEVMaXH////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////10QvYAAAAZ3RSTlMAC+wJBMYl0QfzaT/tgAjY9ZlGb8ruydMN3NImBj5HMVTP+P2+PVVmj9c3fx3Cx/rgAkXQbvQW/l6CxemrbS/EZ84PlcgiTHCjKfwy1sFfiAzVSdSBJC739lYsZEqcwESRm0NICppcxcDC7wAAAAlwSFlzAAALEwAACxMBAJqcGAAAAgFJREFUaN7tmVdXwkAQhRWE0FEUkW5X7IqIomLvvYC9l/n//0A8ARKSnRyOs0+euY/37sxHstkNZ9PSIlZPdFiKj+oAYE+Gj2oX4HBIgo9qHwCiLrqPaae9UgDzZB+Tfft3PMzMEn1MPYOgKkTzMS1sVMfDKMnHFOiqjYcuio+qsz4elik+qm6tYJriowprBTaKj8qtFfRR/GZu0RbFb2aSxyg+KpdSG79C8nFNVCsmN2m+xTWsTrndDk8b1Wf9VaATxWcAAxjAAAYwgAEMYMD/Bcy9BXNAVi7z9SJsHz8BaXp2mvsnwyBRYdPZUVxqfwBfyQAYBMk6McwvSNd7A+BJPqCjARCUD8hg5xOy5EZXrywxgAEMwACK3+tsbXUmx/vMFZElb7ySea8VcXb+WUqOZy0BsZRug802trh61LJUzND+UstCCg7oSOi9Nv+AfnN0WWSNdWd5MaB/0fia622vtbixyPpPjdndrQgQTJtf1MVjtcdx0ZyVr9RMVHcUNAJ8Do9deNQ/UrD5CusB4dnkyJoNq7N7HD7s68F6AS1zPeAtVZz4p+iVzlQv/Mic3Ufwm1K8ULOLsnX/YYup663/7bN8HL4t2tcesCYeWjzLvyaw/h8RfPkYl5Z+SYYMK1sRfIxCNoDY7wZQ2Tj8gqy+qWSRjLdrBjCAAQxgwH8C/ABsSnrVS04x3AAAAABJRU5ErkJggg==',
  github: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACslBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////+D/333AAAA5XRSTlMA/vcBE+1J/PkCEv0FZhvuAxp3CB9ZfKIMeLo5KfSALwm24U0hMx3nBPb7MRyRNHvi6PPMRSC58n9eZ3onu910FDzcyM0ywm/WnX6IB+8rsdXFq/iWFWKGyapsLoMqTDAZP8tULURfvmkRz64LYansKI/ScA6zVp7mhCIs9Zg223ngwerAoeSMoIHEcQ/YibKw09ojTmBqjuVTk1o+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',
  user: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABpFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9bMWvtAAAAi3RSTlMAEwWP9BXt/vzzjPgk8fvrb/r5JfL128tZbRiRuUIEa/bXmceJ1hcHFhsBU87YVxnKhSPsza5RdVY3f2AiJ1C3mOWN0LG2wD/pBu4LTv1840NnLrT3FD2AWm6QpMhmiiiz3L3ho2qODbJ0ElWdwSFkAogmDHNUeCBEaM+W2Xebvx/oyYtsld2+n+IclCRQlQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAq1JREFUaN7tmGdPG0EQhs/9zsZgG1dscMHGNBd6SeiEFjqEAOm9915Jb/Ong5QvrH0+b5mViOTn653eR9q7md1ZRaGlaTcaCy55vUvBWDTcpCBj2gsBQTFiQoz/sx+AMqyWDqR4s0cDXQJnsxj5YzNQkbdPxfO7NDDAPyiaf6IeDFHfiOVvQFUiQuujVheo4/z5dW6gQLvEm784A1R0mzkFL4GSO3z5HVZagVbgEuwDNdNc/S1AL7DydL49YCDOIQixCL6w5xccLAJHG7NgF5jIMwuibAILsyDGJjjFLPjBJggyC+xsggZmgY9N4Dp+AulLdEX2R5b+m0ovtDCbYJhZ0Ca72SlFFkErx34QYRHc4NkyrZK3TMVCL3BKPrb4OSeqD7SC57yjTTddfr9Z8uF3jP94PUhzfF8WGRAuVK/hR5JHqA3RIW3cb5TvXlaEaTT4l/rrMAblFo+/0iC+iDTrF27qFLV1GvNGxPTtPLE/OFrjd7EvXNrzlomc3eez5yYs+XalRo1jgO2ecyvZlzj85z+nGys0kfTC4eNEX/Ly1xc2xnRP6GhdOeaGNktf2Rx6QLxSjNA7Tk66yjvD1EBPeKTXls3aekfCPQNTOiPCu1dU8c+c9cCJN326en7zGgjwe75KfOeoCkKoq52Grf8+CHPd4BBjvg0IzLZUXJ9fgMK5xxUEo4DEqn7+vIolWFnXy7+1Bmhc06sHJyDyRKc/1GMKvJ/KBClA5XvZmcGFK9g+UyLwADJxkUtMjsnZ5sAWrJBrNAzodMkrgn9cJARb+IJZQvAeX5AkBA34ggVCsI0vSBACkEBNUBP8bwINP99NCHL4ggwh+IgvmCQEr/EFzeSeeRU7f65k03+I3LDtZVdVO3bU/B2dqfcAL/+n/lXbeiqDUA9aJnX0+/4F9ltf91/nTXoAAAAASUVORK5CYII=',
  linkedin: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABblBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8c2ar3AAAAeXRSTlMAmvQB7vFY75c1BgXZa/bH+f74/cMHBMxqCCxv1vrGywwbXmOw8xAVGslIzmzyebgCofwZpMhm7IooCUxRmTcYvVt/wqPmIkvwDZMSVZgXfsGiIacmaUpucxG2VNo63yDFpkSHJcpozwuRuoHEQ4aqLrRS9xQ4HbGVIkpHRgAAAAlwSFlzAAALEwAACxMBAJqcGAAAAdhJREFUaN7t2tVTQlEQBnDSAGkUFGzKQMLu7u7u7o773zuMILuIDzrfeXA8+/jtzvnNMNzLPfegUsmSBavC1RqNAiqNxWTIXL/yVIGWq5Gvb6hTwFVuZIBJgdcJA8rwwBUD/HhAzwBFQEngnwP6w55A8ZLbKwq4CH6kRbM2IUCk6DOvFwHoJknDJwAYoI07AUAJ62zBAQe/kfvggFXLOvv4jyjIOoN4IIc2Kh7xgIV+RmsiLrTqdD6iEwHYH1LxrlfMvcg2XZkIi02twm7XjuaYekMjf3D+AqDRlf4csOlI2ZNhmGTxj7WbqtoTD+kdJVMTOz8CcmkjLxnmk6wgMdRQQZJo7TEYiGQ8/qu0nlIk8Bb9ukWqteOAWNZNWBUMeC3Mvs27RQEv3+wjF6wg4NtaFA2siAbOrFBgfeZL5IQB2v4+s6KYnV08vkEB4/ep4QaWV6OAofT0Nc0PQMA2mZ6jjSMQME+m40+kEQABw3R8lL4TwgAhdj156JfLAQGeGXBJx80QoJMBbXRcDwE2GXCOB3IY0C0aUEtAAhKQgAQk8E8B+TJEAr8Dwvj1x0QfNToZ4MYDLfy414Vev4afEqhC4APl8t7MTbqh3gI7cvfvLRvlnxhk4eoduYJqUCldnKoAAAAASUVORK5CYII=',
  link: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACcFBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////8h3Fl2AAAAz3RSTlMA8QoJaQMC/gzCAesN7y+s1AT7qGruxA5bmLr8vhdPgK1Zhi7nI2tlkzvlB9XawJEQ+B8F8jOIxVMcsVTXhxmwjNtmlBNW8I6yUNa3+iGNMPk1eAZtCzT9Q5Ueri2EWskbqexMj9Kz9rgYPOEiJwh0EnmdfqJk9JJzEbZ9giCmq07zltl3WF0+JKpsr5BxtFL3dntc4tHkzChVyokrV6SFyIpRMjd6vZ6jS/Uxwcanlxa7nDpeP2NE6WhJKnIVmzka44Gl6Ejt0w+5n2HQy4tL4xJkAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFZ0lEQVRo3rVa90MTSRSeEBNCAgQCGMALLWABRJoiRUBFBA8FBVRERM/zPHvv7e4sZ29nb9d77733Ov/SbWY3knkzO7uT3byfkjff++bNzNs3b2YXIQuyuvi/toxSB67/JaOtfdVqZK+sLJmNaZk9K2wbu7v5DubJzXkuW/gLpmI9mdptnX7zo1gkf5ZZ5J+ZhsVS/5wVetcz2FjyMuPmrziGzcgTSXHye6qwOUmJr4ekR7BZuRjPLHlSsHnJS6j/EclNqP+KOMoS6r8i+Qnmx/gp6/OTk3VqZcgdCudm8R7v4y6L/v9WmPoQ0T/nCxYw05r/w9UUqDqLQYxa4U8u9AKYd1cyBN2zwD+Og9wBUSP28iOvH8BaTPD3f2iaH6HgYwDYasjv3c3jn6gHXwKQ9w07+EnCf0UGO2noXSP+ZyfI+K8IWIUiA/7gj1L+K7KHBl836GCRpP8ILabR5WL+r3sk/UfoCMjZYvRRaX70B40PCMELA5Lzo0iZzAgWSfuvhB1tcUNYY5XDDs5qLVteq2nR2XKzaYt/RB0sh/yTtYZ0khB2OnlGT9ImflEH8wH/O1qpk+5T/7/QwdpU19A2M0Qd3KSxy3ppfowXNDI2I8Cp2wL+VJAlslR12DemqtwObDbsBVEREnSwFDjzKdGGrsXqxtPr4L4IbDJEM1QMCgh1Bd6jtR9RJrNgWBSKOljCy4sLlwGK4tgQZfbkW2zsd2/bqP0sobHziXIcUyC2Cvgb2P03X1EPqb/30eAr3NDFeHfUdhLDjwt09vfp5M9kGvwNTxkJlO9V2xV9TFOlW6d++JK3Yi8S5SV2gzioGueyLV/p1SeHyN+rNPpnFTSXHUI6aficaejSrX8eJ/8/AXuTVsgNMDn8O6J3vgXUabW69ZUakR8Dg+kacC2sA3IqiL5LXPnG1m/riKYjmZ9L0QX+OeAsrZxG8ztjz6fr64huAW3RdyQKPgM6GOKkliL6bOA6R7W+TJR5mJvtlMe5j5dx6i7HqKq20gMA9wPZRNkNp+JCFA5CKfkA0f4ecw730PwnAVObem0Aj0WBtfziB79CtMO6/JsdwGCvmoV3wiEEBrS0XErr1euV83r8iL3/WU70vWx2matarKG1O4jykB5/AfuQF3N38Ihc4hU0z8dmWobfxd6P1Wib7RR2CPtIwzZezC/l86NmNrvsibY9YNraeZuRmlvQUOR4z/CjO4L67Qe4/HgO0U+jlYc1+Mb9b7rZ+0+GPzumdRVonKBmsCZau1+09ZbADkqo5tfpxjfUhxYMrFnUAby/baIH6WyIbdz0LbeoTBfdP8M7yQ8AYHvlWKMvzD2I9SSZr344lUzjw6zq0zxdDKL3XdEMtdPYy0EW0jFK89c1YIkKtw1cePIwW0lU5kS3tPdhWEwRdZDBC3NGclvW52/Rfg9A/mteUQelEs6oGYdJH8IZQk/T4BVG/JNgcYp7xDagRAhK+4//QjIjqJXmrzktNinlFz+m+UFmMYyiz2T5N3kMOgB58YxofTn8AcOwA09y2qCU//iqYVjDhH9Fxn981G3YQSMw8QUl/M/vN3GrCPcDv9c0f4rHBD979HyVd8URPz8KM3vyLjgG92EL/Aj9y5hm0bfeB5qwFX40j/Na4MTYFcHgiRwe//hU8+/2eG8+HWtOvb0hs7X372Od2Jr/irwk//JEwv+IHEwwPzpdL8cvNT9662yj/7zznr3+RyKpKKH+R95jmX2PWOVB8UmF3xT/OWfc77ozzazDeVdCvwZwnLT4vUGZ+HuGrlobvsg4rktf+atN35SMcukbCtzILrk30gLvbwtvIXul9f7d4XXlnTjtxnX/jNshK1T/A5cWLGcTiJ0aAAAAAElFTkSuQmCC',
  portfolio: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACDVBMVEVMaXH///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////9xxQJjAAAArnRSTlMA+S7+4zg2+AwvkAv3k5Iwke70CuT9NzmuDi1LAsrrtOpDJN8y4YQNxQFdYsnt8hQa/K2qZvZqP1gIh6NChQTN3DzHbzES+iHvK7EWwF6CYwYQ3sToBc5XGR7BJt3x9YGAWykXQBVP2G7S50fsiigJjxMYf+YDqEbwslB0VeB+y9VzfemMz/PUclOWNHd8v6/TcQ+awmUcrLCgpFbQe8i4jtZfEVmrmz5Ia2iZjWeFplepAAAACXBIWXMAAAsTAAALEwEAmpwYAAADjklEQVRo3rVa918aMRQPoB6CCxAFFUFRW/feu1pn1Vqt27pX995777333s3fWKooVrwkLyHfH+/zve83l7u8e3kvCDFhKmk+vO2c22TF2GoacdgMJd8+oUAh+WelG/tD9+HrZ7u4etzYAlaH4szMFVG3ZGZhGvKdSbzymztbMBOuZGznkH9YnYOZEVSUAJRP7+jFIPTVw16tA0Ohq8xjH36Dgjlw/DSj/sEtmA86w1YW/ZRUzI3aZLp+h4IFkB1L07+uI9weZggxGmOiSEMw7yLrdxK/du0SKTSIQNKUkfTHSPph2mVaKOkZrF3q+g+I82vwESOI0SleTT9DRzQI8TGDiUTzLZXvs5j8heh9VD2Z6Vr3ax2IxswGGyjU8xZ//QTq+mWeIg9u+hucoi6iKB85nBo1/F60lr6AldBlcqSGSu6dWTNBWSy/Fa9DZDQD2fa/QSZTpFEigvX62HANE7lqtX73JWjMofNrV/9F0+Axh87PWJU/jHLEHCo/tVzlAVhjDpVfspJftXAvKCL/2vLlLt6YQ+OXei87VG+wAw3W8Au9SZBAzCHzm5ZSpXGBmEPhzy9eXRCJOWT+u8X9hVjMIfIVo+dykWDMIfJnPQaVWCJ+eQzcMg3mEJqRqY+tdjQt1cCzmH/INfjDFjD58RsVyDWwoU1yDTaibXIN6lCiXINEtEOugRlZ5RrkoyahjIhhinJEMiI6XiOXQEbEgJdoTiAjYsArNCTwu6dj1I4K+TZJjHiDUIVMA7dnr/ZR4hSZ/2UDIQIZEU1/caNWruHPiMhIbVzSeiuQEZGGP270Dva9SEakUjW6f3Sie2W2vwtp6Xou1LeTK2nlIgG7II6hSPecW950kqnK2M+rHzTJWALP4Rx/K2sh9hGfQT9zpbdVx6PvBJSq7/IYpAAMjnHoO0DV9h64wR2QQRLcoArWkGgDGzTDDIbBWQS0+3QIajAANKjpAxrEQLtOjcBJKgL3tRqgWxco0gtBBsoU2GEfbDd1Ed78awdtpw5z9Bf3miAOPE3S0qsAgzSeDmww4BnucfV4tTsBW2AuNNdJNkB5Q4wGJ3g74TUVbP/oPfy9/PhsFoNpgdMCg066vuuI0HGH2TM0g8uCByos1eQcfncuEsXgfjMhbb8RiEMnB86qFntuB+hci6WsYL0OZ/ETFDjoy7641hRDnw2jAGNyosKWFW3ShJlGHr8oeQq59S93tkBme65lnQAAAABJRU5ErkJggg==',
};
// 简洁单栏版式用的深灰图标（#595959）
const ICON_DARK = {
  phone: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAApVBMVEVMaXFZWVlYWFhZWVlaWlpZWVlZWVlWVlZ/f39VVVVYWFhYWFhXV1dYWFhYWFhbW1tcXFxZWVlYWFhYWFhZWVlZWVlXV1dYWFhZWVlZWVlYWFhYWFhaWlpYWFhYWFhZWVlYWFhYWFhcXFwAAABXV1dYWFhYWFhYWFhYWFhbW1tYWFhYWFhZWVlZWVlZWVlcXFxaWlpZWVlYWFhZWVlYWFhZWVlZWVna3J8PAAAANnRSTlMA7Zk2MPP4LwIM7uM9Ltg4JNb0wJL5Q5slkfKQRPe+zkLXCwEmkzHBRQ6cz+Q5vyE+mjfN673D32sSAAAACXBIWXMAAAsTAAALEwEAmpwYAAABOUlEQVRo3u3aSXOCQBQE4BFZRVTcd6PGJdFEs83//2k5JIFLpux5ZA6p6j4370MUsGCU+j3HUaur4XTXL3Nlk/uDtk5vjM+/NrQgjSa8/6L5Wtc6ILDUwhyw+UctzhACHuWABwEtObCGgEAOBBBQ1N+yLdLfRqtiEzsgQn/XvhCIUSAWAvipT4AAAQIECBAgQIAAAQIECBAgQIAAAQIE/gRw/uQXfna9FwIrH/oM8f4kBAQhQOAWsJjVN5vXSegKSLOvQpS6ARbZTyMKnQCzstJ3AtTLStsJkJSVxDXw8D8P0aSs7CoBU8PGYXED8geGylRVemGdfgv+2dTIIWBkPE/DfjtJnnaDiq/c5/Jr3R12D/+Qzu+BfxIu744XbqhmTTQfXnqi1PjZfv6yo2wy9HKL5Q9B7pm+30+kOWgXz9xUnwAAAABJRU5ErkJggg==',
  email: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAB3VBMVEVMaXFVVVVZWVlZWVlZWVlcXFxZWVkAAABYWFhZWVlZWVlaWlpZWVlYWFhfX19YWFhmZmZZWVk/Pz9ZWVlYWFhYWFhZWVlYWFhYWFhYWFhYWFhYWFhYWFhXV1dZWVlaWlpVVVVdXV1XV1dYWFhZWVlXV1dYWFhYWFhYWFhaWlpXV1daWlpZWVlYWFhZWVlXV1dZWVlbW1tZWVlYWFhZWVlaWlpYWFhaWlpZWVlYWFhYWFhaWlpYWFhiYmJZWVlVVVVZWVlYWFhYWFhaWlpYWFhZWVlZWVl/f39cXFxZWVlYWFhVVVVYWFhYWFhcXFxYWFhISEhYWFhYWFhfX19YWFhaWlpZWVlYWFhaWlpZWVlYWFhXV1dYWFhVVVVZWVlbW1tYWFhZWVlYWFhZWVlZWVlYWFhYWFhYWFhXV1dbW1tZWVlYWFhYWFhYWFhZWVlYWFhZWVlaWlpbW1tYWFhWVlZZWVlZWVlYWFhYWFhYWFhVVVVZWVlZWVlYWFhZWVlYWFhWVlZYWFhZWVlZWVlZWVlaWlpZWVlYWFhYWFhZWVlYWFhYWFhVVVVZWVlZWVlaWlpZWVlaWlpYWFhYWFhZWVlWVlZZWVlZWVlYWFhZWVlYWFhZWVlXV1dZWVlZWVkVGvw0AAAAnnRSTlMACfnzgCRvAe7+kUT8+hD9Be0EsTTeJUKzmcGIykOrEQMTPar4I9jdvncdPhTX+ybCONbvNkH0IqK7wFXPDd8Y5JsaMONTOQIhzuwM8XAWfwfH6witZuHbTPBFICsPvA5U5Zzqy4Lph10nl7BOtV6Wpk8cyTLU2a/ycRK/hlyLwyykTdO0H5rSUeK4VhvnR1LNM+aE9jW9PJij4NE69S/VufIAAAAJcEhZcwAACxMAAAsTAQCanBgAAALySURBVGje7ZnnWxpBEMZXEO7o0psoVUCiUTRqbImxxBp7S++999577z33twYjuwcIx53sPI9Jbj7xcHO/OZblfWcWhOSQQw451kps76lp4ahGS91jHcGv93Ag0e5Y5idYDijY4J/nB+NznFmVKhDjAMOT+n450PCh/bAF3KgGtkAdUsIWUCIOOP6BAnpYvh51wxaIoiOwBW6ikU5I/tAIQkYrHN9qXJJTRm0A+oLLt6YdJ/Eegn+AtzS0bTd9fsOWTFO2lJvo4jXaihzfv0/V2Eb3ruwsHlXS43eqeK7tKZN+VXaGFv94WRrJvGhGLHeiGhd7Q2WZ2H7MC5zi2CU1fTaA36ltKJ1/oRXTPtmxXNfHGbKb9LR2D6NWZPiBlSzThs2l8A+/xpw9VdmGM9eIr6xzrZ6/cROmvIvmOpohjrI/m/SoJ9qDZhV5LLPKSLbu/Gr48zZ8f/VMfk+e+4Izur5J5/PP1xwtZPr15X6cdFpiv/RRza+wQaCriBEJPDYkhX/pLhHm58Jty9QCzpzQakTz27A2IN2rYn2RRjuBk08Oi8MvduA7Klaqfp7G660D558XNVtNDuL8HZPiOrurR/Ed/nhRu9Z7ibN0NIluHUNkyzXahfl9CZxZppXSm3YT5dgVEuxLAjjv4mVpza/iGhZYZqzgMinGSJI6LLm75oXLVmCZ+n6KEEiB9n2fk9z/MN91V0CMxCNx5jH7K/diWI2Xxy9oUsIDSGUEP+TOJznDXa9Imy0y4TSdI7vwQ4ZLhA8RbVgwlzZCaW5byJnJ2bQOmO6R37rFqyl5RrtTS3oo1fhBlq0c57uqlx4aQ+DirULHTP3DlKbM0PV8+C43vTH2u24l32mnOidfUWXjW2doD+Im1wCP720zQEz6D7y6yPR0xOf98X+dVcgF5AJ/QQHgY81R6IPZJOqBLfAV6WALpHqrz5D89pS2D05xsH9QoKAZjB9c9ifHDRh+jLdZnztJ+fegTLqd8t+Hcsghx5qJ3wOUnkWsXwrMAAAAAElFTkSuQmCC',
  location: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACN1BMVEVMaXEAAABVVVVmZmZYWFhZWVlISEhYWFhZWVlZWVlYWFhVVVVYWFhZWVlVVVVfX19ZWVlZWVlYWFg/Pz9ZWVlZWVl/f39ZWVlXV1dVVVVcXFxZWVlVVVVZWVlYWFhZWVlXV1dZWVlVVVVVVVVYWFhfX19ZWVlXV1dXV1dVVVVYWFhYWFhYWFhYWFhZWVlYWFhZWVlYWFhZWVlZWVlZWVlXV1dYWFhYWFhYWFhYWFhWVlZYWFhYWFhVVVVZWVlZWVlZWVlZWVlZWVlYWFhZWVlZWVlYWFhZWVlZWVlZWVlaWlpYWFhYWFhZWVlZWVlaWlpZWVlYWFhYWFhXV1dMTExZWVlZWVlZWVlZWVlYWFhaWlpaWlpbW1tZWVlYWFhcXFxZWVlYWFhYWFhXV1dYWFhZWVlYWFhaWlpYWFhaWlpZWVlYWFhaWlpYWFhXV1dZWVlYWFhYWFhZWVlZWVlYWFhZWVlZWVlYWFhZWVlZWVlYWFhaWlpbW1tZWVlYWFhXV1dYWFhZWVlYWFhZWVlZWVlYWFhaWlpaWlpYWFhaWlpZWVldXV1WVlZZWVlYWFhXV1dYWFhZWVlYWFhZWVlZWVlYWFhYWFhZWVlZWVlYWFhZWVlZWVlZWVlZWVlXV1dZWVlYWFhZWVlZWVlaWlpYWFhYWFhXV1dZWVlYWFhYWFhYWFhZWVlYWFhZWVlaWlpZWVlYWFhYWFhYWFhYWFhWVlZYWFhZWVlZWVlZWVlYWFhbW1tYWFhZWVly+JeGAAAAvHRSTlMAAQMFF/kH/fz+9xtzrgYI9vDvBPjZAu0dDBb7FVDxiSDlHgnuEORXPRLdw0Lse1GAN6l1wjrSinCzNawrD7ySeH0/fCWrxHJYOUSVdobOd9PXVnQKzdx6W5NMLTiowSTRRYhdu2+wY41gNvQi1Ul+VNpk6sDLSoLQsYc7J5fPQ05elsWm3k8w4xGMEzLU8iNxv9il6L5ci2yktBS5mkbiuGHIM+aEWo/HZeuUzL1m4frgX8Yvp/XWujQcsp/7UGoAAAAJcEhZcwAACxMAAAsTAQCanBgAAAPHSURBVGjetVllUxtRFF1IQhIIgaRoQgju7lrcvWixtmjd3d3d3d1d34/rTDudMuTet0825/O558zuvr3viqKwwbdstnimKC3BYEhIK5pxz/7wVTTEz5LSaDIPtl/TAdqox09OGQkI0/erQdLykdOxhIJNOZVS8rr8cKKCcItOXD+lmzDgYqugfFCbmTDBXOUnoh/zkDCjfwm/fms04YAtg1c/259wwZDMp//WSDhhfMGj7zIRbhgXsetnhBABGO6y6tdHEyHYzrHp+/UTQbxm+x/aiDCqmPKDWdzAXKuuv2ApkUC3eubLJ1I4r5r/M+UMMiNVDLqIJCwq9+NlPNQ03jTq0Okco5YCyo/+mH6LTqKBwTuj/tMc7mCUOEQ1mMLC8uaVEBENGPM+TT8AeXZzjs98qk8XknCNtGqmhCdRJiMO1ykGpXBINczOgtmBlL/YBr9/hO5TDtIT8KqyDD4/6EuNgO/VR6jBRpDfgT+yGwx4gPKLwf8rCjfQg6duAuXvgOibaec6ne8rg5m6iWbQC0XsRekfIHouzaAQijiC0sFUHUUz0EMRL1G6AaJbqQUCFBGC0k3aGJhQerA2r+grSrdD9BaawS6w6UHpA9x3YDsUEYrSn0P0AprBISiiEaWPgZ/MQfkE4JVwGOXngLnLjRucBAPq8LYGTtcRGD8OTtf7UYNbcFna4INcOIHwBZuKP/J2+BIchNk1MLuTcijWIZc+2OE1I5d+FsVgN9Z/DXqWLXVYn0jtmQuwcqp83peOC8SYV6iV3Xq0IvQv1s85/9fwNtpFbz/SKH1qem+h3mrVF+Y3UrroDSqjsBWy5fuQWgcVKqf/TXWW55IzGFbvAm/L6C9kaGNHnOL67ytYGuUJcYMnbKOEHlH9UMYpZ61gr29MYR233BEzOMg8L6pMEtGPTWSfSA2LGBzlmamN8evf4BoKJt7j1U9K5Rs7LuM8Sea1vIPTPXwGi7knv31caXUgkttAaeGYnTpzRabjHewG28S2E1tY9TutYguEEcYBsH+FIggLm0GJ8Iol7A3TLRYmvsQJCFfXt8fJrKGeqhs0y+3RvqjpX5Jc1CW+0zbHeWIfddlizFCk8YxmsFyDZalfEa6/Jl6LdewxA6Yf8lmbhW81ZtCu0b467AKsn67ZVjwGXEvZYrTbuX+EDA5oudXP45ogC+DEcY867pSmBh5ljHmrojFeyZYpqgv4M3P1e/o0N1BOz/mhnSsVL6CGbd4h8UOP/9P/5OsVA2X12b/6q+oVLyH7TzXpdCleQ+5Nu720jCvkN4wqLzwXzIsLAAAAAElFTkSuQmCC',
  wechat: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACl1BMVEVMaXFYWFhYWFhYWFhZWVk/Pz9ZWVlZWVlYWFh/f39YWFhYWFhVVVVZWVlYWFhZWVlaWlpZWVlYWFhiYmJYWFhZWVlXV1dMTExfX19ZWVlYWFhXV1dZWVlZWVlZWVlZWVlZWVlaWlpYWFhYWFhZWVlZWVlaWlpdXV1aWlpYWFhYWFhWVlZZWVlZWVkAAABbW1tZWVlZWVlYWFhVVVVZWVlaWlpVVVVISEhYWFhZWVlYWFhcXFxZWVlZWVlZWVlYWFhZWVlYWFhYWFhVVVVYWFhYWFhZWVlZWVlYWFhZWVlYWFhYWFhYWFhZWVlZWVlXV1dYWFhYWFhZWVlZWVlYWFhaWlpbW1tYWFhXV1dbW1tXV1dYWFhVVVVYWFhYWFhZWVlXV1dfX19YWFhYWFhYWFhVVVVZWVlcXFxZWVlaWlpaWlpYWFhYWFhcXFxZWVlaWlpZWVlXV1dZWVlZWVlYWFhYWFhWVlZYWFhZWVlVVVVYWFhZWVlYWFhbW1tYWFhmZmZZWVlYWFhYWFhZWVlZWVlZWVlYWFhXV1dZWVlYWFhYWFhZWVlYWFhbW1tYWFhXV1dYWFhXV1dVVVVXV1dWVlZZWVlZWVlZWVlZWVlZWVlYWFhYWFhYWFhaWlpZWVlYWFhYWFhYWFhYWFhZWVlZWVlZWVlYWFhYWFhZWVlYWFhYWFhaWlpYWFhZWVlYWFhZWVlYWFhZWVlZWVlZWVlYWFhYWFhYWFhZWVlZWVlZWVlYWFhXV1dZWVlZWVlYWFhYWFhZWVlaWlpbW1tZWVlWVlZZWVlZWVlYWFhXV1dYWFhYWFhYWFhaWlpZWVlZWVlYWFhYWFhaWlpXV1dZWVlYWFhZWVlZWVlYWFhaWlpYWFhaWlpZWVlZWVn7/I9sAAAA3HRSTlMA8hox7QT+/P0C+rIGKKffRM73Db5NHQoIjs8mR8L7+eFV8TSPfUETZusuMvarARm3ooIJtk8DB629FySAqRSY+EIrFVbbqKOey3nXyciuWkiwv9ngEQ7aIxwgbg/0u5dDEFzDigy5FudSM+ZlC3UtiTpQ1paFLO9eGNjNXyebBWykr6XQ0+5d6oiE5JAqxD24QBJXNX7wb9zRtayNIp3Vc5zAabGSym143nww44xi86pTulhxlZ+LtJrSRuJhmTdbTDhqL0qGRXTdzHAfg2dLx3dJ6MblJaE7s2DU4C809wAAAAlwSFlzAAALEwAACxMBAJqcGAAABH9JREFUaN7tWWVbWzEUTimlt0ChUGC4DdsYssEY25DBkDHBBhvM3d2dGczd3d3d3d09P2YVWG9yLUnLlz19P8F5Tt63yU1OzjkBwAknnPjvMTdxU2jUI8PjAI3DqT2nXvU60AY2QRvp0c7f6DD2kWXZ0VCI8OGTAhxBf2awCkqBO1xg72ptTYfyiOzvaQd9QKYbVIT3sGRGel1/FSRCx85M/DmjICnaTGDgb3kXUmBJC1r+uGhIhZ5d6findYGU6OFCw7/BF1Kj92hy/hBXyIBegaT8HXwgE9qSCgSh47rcnHK6NFxA5+qVtfo4alpBuEE5ZNRly9dTt8L4g4aazU9QzU5EAmHIoPmNseYhyr+20dkLsRaT8PujTFcazYERfKtb05Zpj4bXWQQCfqjAuCb7GGQx/q0n6p2kzN8aW+stojPgckRnAMOVj9sITMCng9Vehn1j0W8AYXdFgemCYKw3mxPxXbTXsosW494+Svz7hOdHmzoltkIrPAdFWc9fC29RpaAXxXaItcu93r2P+lBekPBWQaAdA/u3teV9iOPcRVp2rjROZxpnzOverbJwZXREQ2RJ9sxp0kd6ECX9jxrToO+fPuLfaPrMceICPeki9FYANFsqxXOP4ASxezSYgt49VgN0ZX2lHcYv9lQKFHJo9dN07hVmXJiHC2SSX5E1QLfNXcnLbRG2ToNJ+fumgOrdJI63QxCBCYT8VbUg5Q+Zq+sDuWAqlcslgtqdpJO9EMcTyCBLuH6DPrvIt4NKzVO4TjLiF0jeTXNexvCuiTySX1QLdtCdeD+dLWn3VnZfDWpoU7+DtimcV3TeXg2SEEOqemRiDHIk9usHTM3mRxBVyj8BzTwlgRjQEok+HS3zj+LZ7luYtvGdbknmLULkg2Lkf4N1XInNYi0/M2bznHx5EbybPL8PiEeT+8nWYR62+y1NJHLG2gSMp2QFVoEVYplE1wabZazFko9cEvW8s6CWrf4MoAg71kNMY/qt58dpc4anxgI5v1QMlYuS8QDfyVz6jUx03/ruOZfKYV5z+CFpASfJ/wUY3SALgpCoWhAh5feVNCDi8EAvhtCBEn4bwQY2geHY3Ta5t7hfEvjMJrAcvz2NRaJ+FWATm0ClsBgUXaX1QM0mIKwdnolX3GAzm8AIgYB4yNieDNYwCfjj/HUS5fg6UMHCz1XjAqFNoavt2NilVTbPBJDAIpAuWKFFjfTWvCbecCLmaNj8V/Ulb4CLlkGgPc6vMRdN2uKFYinyUnr+QkH70GCiX7JQoptEL3BJQJLLLdM7qoyAMEzYEozRyxRCTym/goq6b0hZyw2j7hnW9aLhP8bQlUyporgJMlj6qvkrSflHpbF1hvVnCX9/IGBEP6LeXq4dDfq0Q8qV6BH72v/lSrn4DHvfL1qclO+xzrP/iWTAJNliLcQBrzC6ITPwO06769pG618vHfOQlLzuXq6ft6V7696jNOuOub3zIlg8ytmDifEuE3kzmxNpmktOcz7tee6YDRc07+vh5lXLmv2B0vlG64QT/xn+AolHYj3FTDdlAAAAAElFTkSuQmCC',
  web: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACdlBMVEVMaXFZWVlZWVlaWlpYWFhVVVVZWVlYWFgAAABVVVVZWVlZWVlZWVlYWFhYWFhYWFhbW1t/f39MTExYWFhYWFhYWFhYWFhVVVVZWVlXV1dZWVlaWlpaWlpYWFhYWFhaWlpZWVlZWVlaWlpZWVlZWVlZWVlZWVlZWVlZWVlWVlZZWVlVVVVcXFxYWFhYWFhXV1dZWVlZWVlZWVlbW1tYWFhVVVVYWFhYWFhZWVlYWFhYWFhWVlZVVVVYWFhXV1dYWFhYWFhZWVlZWVlYWFhISEhYWFhaWlpaWlpYWFhZWVlZWVlZWVlZWVlYWFhXV1dYWFhZWVlYWFhYWFhYWFhYWFhXV1dYWFhYWFhZWVlXV1dYWFhaWlpbW1tZWVlbW1tZWVlYWFhVVVVZWVlZWVlfX19aWlo/Pz9YWFhVVVVXV1dZWVlZWVlYWFhYWFhYWFhYWFhZWVlYWFhYWFhYWFhYWFhZWVlWVlZYWFhZWVlWVlZaWlpYWFhZWVlZWVlYWFhZWVlZWVlYWFhYWFhYWFhZWVlYWFhZWVlZWVlZWVlZWVlYWFhYWFhZWVlYWFhcXFxaWlpYWFhaWlpZWVlYWFhZWVlYWFhZWVlaWlpaWlpbW1tXV1dYWFhZWVlYWFhZWVlmZmZZWVlYWFhYWFhZWVlaWlpXV1dZWVlZWVlZWVlYWFhaWlpZWVlZWVlZWVlZWVlYWFhZWVlXV1ddXV1YWFhZWVlYWFhYWFhcXFxZWVlYWFhYWFhZWVlYWFhZWVlZWVlcXFxfX19ZWVlZWVlbW1tYWFhZWVlXV1dZWVliYmJZWVlYWFhZWVlYWFhXV1dYWFhZWVlXV1dZWVnyZNh8AAAA0XRSTlMA+84t/Qn29wED/vD1ye9xHAIK+ivpzAY2HfkfY89tMOKAM6XLcvyi7TX4HiRzdiCr8+gqcBvrMaPbFy8YuyaTlai0RQcaZj70nerQfbUpgdRI8i6fRpk353R5TBmOJ8VWFabREBEEWQ9aamyNbtWhku5ONOORMpCLLHdrFLnxmkd/5oSUS3pbiW+nrFiWFmDKRKDXKMfTUkEOI2K/nlAFOXxUvCJJfuWDgjt4317k3iVXE6pTr1whZ+yKe1GprgsIddw4pJc9tw26vs1lQME/Ou2dZEYAAAAJcEhZcwAACxMAAAsTAQCanBgAAAVpSURBVGje7Vr5X1NHEH+RkASQEAwkgFyKEEBQAUEFK1JEKV4tiBeKAi1YtdXW1tajXvWo1qO1tda7933f933f+x81meHBe7uzb19I8vn4A/NTdr47M2/ZnWNn0bSbirZcLC29eClh6uf/nMfCVHmmMa5qp1cM/7ijgw1TdvcwKyc3Vu13/fVPeRv+fNvBRsh/Ank+79X8O2NQ/3SokrFz+Pv+MmYgxzHkLmbMfX7BGNVn9TvDqi540FY7M1ENavX8Gf5d9tOKMaif/LUTNC2BUXAp42hCEIBWGDhPRbsZKSdrUM+tOP6SCbQekVtw1FmYGY3+rnRdzXN4klyigYn40c/q4wt7bKv3fO/XpTYhZyEj6ChiAX2cl2FT/6T+USVfAae+kjLgxvM7NMqZm2xHv2/aqETtILB2MpI24v53jHIKZqr1r2kwqPgCWI0u2sDEPoAfNLCyK1T6p95nVJEEvFImoZMAJxlZaTmK708zzi5AZoHMQADxe4087y7738+WoUczKeEKH2V213C5nJJ/WG7gMfwsM7Nhskz/twPmmT+gV7wqN/AABCpPMRdFJKfV08KJnwF2EWOqv9FtHLeENrCXl/4I2DOsDFTDlLt5NunTux3cLOcHxmhG0yPoKKkc299ExM8AL7wUA4fbyoArWTyoEUpPEQyI3rQaMzGzpG5yExi7R8gvNcKcc8QhF2ilnjk5av+UM/CxKFsHwHJrA5iQukRglll/U6oYyjBFrbU2cAQ3UNyo1CSTgVmi6DRMBUxBG2DaBBFoMepf4RQnrMJqRWWgB6btEAFnlsHAXEJyHSBTVAZegWnrCCSkjdM4xY0WkfFyJ2CFTEmfwcSjFFT5HWBLSLnD9tyAseMWE98AbD2JFQLWrDaw0WKpMwB7nMROA9aiNrCDzpqGePQUibUC9qLawHKY2E1iz8AeO0hstyxKCtmRqo30iv5GGKqi5apAbp7awO0wcQENRmqwf2loOsgVqw3Mw5RLg/9Jt4fNB7latYFi7ALQ4AthaCsN4f3aqzaQjTd2GswPQ9U0hAWPS20gDcsnGtwWhp6nIaycytQG3Jj2mdTTEm4g4X+ihG/yEA01xu2Y0tGa5cbF0SKdiMuJDBWReBCkg92cuAS7QXm4flJ9vTHW14dIbEBW+IbpfcBOqQ2ch4knSAxvg2+S2NZ4pMxqi6S/zG7Sx87XShL70aJseQiwbWoDi+UtK1dwvCgcp/iR/BJ4WHVK0V9mE0i/rWtsj8rAMUmzgruJfyi7iLfZu4g/oeoldMXWSshsFxdQZN4FoobPst8MqROBb/iXmk4yYWvaa9YGpsCkfHH9L/EHSTwHB+00pPAacZCpu3Yp6eTtwV5LbUA4IWJLTWvy89vUZ7cpuJ8/5Y4iytsyeOnNdtuam3l2Ke3Pq8iuoI3G7Gp7jVktmStS3sKG8Duq1rL2rplbMEkWkmY2ULWRsjnexDXHLZ5B5pjb76+r2vt7iIRv/USxyytWhfLybviBwvSI51U8spjXUEQfL3N717TCtDWq1FCRLdQuZ2WPRPsB3m5gdeSok0+9waU/XyS8MxloO4CDa6N75gqf1hDfEqknw4V7A+9l9h7qIoEvz5x16LqtGbGRBefNtp+j6zbpUn/DOJd6LPUB9MvIs2ddNFVA5r4DKPY7jo+LBvYi8iuODuzLjLLQ8JU4DReRoOALgRuG115niW8MtUxvKHK5CmC46eUeMDrfw0AVyfZlod4xlktV19x6ea71mK5yjpc1PVe6r1XF8q8gV/7wDq/+E0NC8g8hr6386pUtsVZ9Pt05D434eG2r7va5cS0wzy6El3F3c1/CathL1zMyrv92c9XV/wNUSqVg60qaiwAAAABJRU5ErkJggg==',
  birthday: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABOFBMVEVMaXFYWFhZWVlVVVU/Pz9ZWVlcXFxZWVlISEhZWVlZWVlZWVlYWFhZWVlYWFhfX19ZWVlYWFhYWFhYWFhZWVlXV1dYWFhZWVliYmJZWVlaWlpYWFhVVVVXV1dYWFhYWFhYWFhZWVlYWFhZWVlYWFhXV1daWlpYWFhVVVVYWFhYWFhZWVlYWFhYWFhZWVlZWVlZWVlZWVlYWFh/f39YWFhZWVlZWVlXV1dcXFxYWFhYWFhYWFhZWVlYWFhaWlpZWVlYWFhYWFhWVlZYWFhaWlpZWVlXV1dZWVlYWFhZWVlaWlpYWFhXV1dYWFhWVlZZWVlYWFhYWFhVVVVYWFhYWFhcXFxYWFhZWVlZWVlYWFhYWFhZWVlXV1dYWFhZWVlYWFhaWlpYWFhWVlZZWVlYWFhMTExZWVlZWVlBSnDSAAAAZ3RSTlMAxmkJBCUL0QfzP+3sgNgI9cruydNGmW8N3D7SBiZUz/34vkcxPVVuD/SC0Olt/l7Fq8QClWfOHRbXN3+Px2bC+uAvRSLIKfxwo0zVSYEy1C73DMFfJIjWZJzASkObkUhEXCyaVgr2DtIqnAAAAAlwSFlzAAALEwAACxMBAJqcGAAAAgFJREFUaN7tmVdXwkAQhQWE0FEEFRAQO2LBLnax94K99/n//0A4ARKSnRyOs0+euY/37sxHstkNZ9PSIlZ3bEiKj2oXYF+Gj+oQYG9Qgo/qAABibrqPabutXADDZB+TY6cyHjZniD6m7j5QFaX5mBbnq+NhheRjCrXXxkM7xUfVVR8PYxQfVYdWsEHxUYW1Aj/FR+XRCtIUv5lbNEvxm5nkJYqPyq3Uxk+RfFwL1Yr1EZpvcQ3jqx6Pzd5K9Vl/FehE8RnAAAYwgAEMYAADGPB/AVvPwSyQlS18vgnbJ0sgTQ9Oc/94GCQqbDo7SkrtD+C7NgD6QLJKhvkF6XpqANzLB3Q2AILyAQXsfEKWPOjqlSUGMIABGEAJeJ0ulzO+nDZXREa9yXLmPVHE2eX3dXw5ZwlI9Os22Fxji7NHLetPGNofa1lUwQGdPXqvNdCr3xzdFllj3c2EGJCZNL7mUm21FhcWWebKmH39iADBAfOLunir9rgtmrPzMzUT1Z0GjQCfze4QHvWv5f2+/FxIeDa5Nu3H6hx2mw/7ejCXR8vcd3hLFSf+KXoNFKoXfmrOPiL4TSkeqdnRuXX/IYupS9X/9lk+Du8W7WsPWBMPLZ5NvPRg/V8j+PIxLi39kowaVrYi+BiFbACJygZQ3jgCgqy+qeSQjLdrBjCAAQxgwH8C/AJwq3rVb0vwPgAAAABJRU5ErkJggg==',
  github: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACslBMVEVMaXFZWVlYWFgAAABdXV1ZWVlXV1dZWVlZWVl/f39VVVVYWFhmZmZaWlpVVVVYWFhVVVVYWFhaWlpfX19aWlpYWFhYWFhZWVlZWVlXV1dZWVlZWVlVVVVYWFhZWVlWVlZVVVVZWVlZWVlZWVlcXFxaWlpXV1c/Pz9YWFhbW1tZWVlZWVlZWVlZWVlZWVlZWVlYWFhZWVlZWVlYWFhYWFhXV1dZWVlYWFhYWFhZWVlZWVlZWVlbW1tYWFhYWFhXV1dZWVlZWVlZWVlYWFhYWFhZWVlYWFhZWVlZWVlYWFhZWVlZWVlYWFhYWFhWVlZYWFhZWVlYWFhZWVlZWVlZWVlZWVlISEhVVVVZWVlZWVlYWFhZWVlZWVlbW1taWlpaWlpbW1tZWVlYWFhZWVlaWlpaWlpYWFhYWFhZWVlaWlpYWFhZWVlcXFxZWVlZWVlVVVVaWlpbW1tZWVlYWFhYWFhZWVlYWFhZWVlYWFhYWFhZWVlYWFhWVlZYWFhYWFhYWFhYWFhYWFhYWFhZWVlZWVlYWFhYWFhZWVlYWFhYWFhYWFhYWFhYWFhZWVlYWFhYWFhZWVlZWVlYWFhXV1dYWFhaWlpZWVlZWVlZWVlZWVlYWFhXV1daWlpYWFhZWVlZWVlYWFhYWFhZWVlZWVlXV1dYWFhVVVVXV1dZWVlaWlpYWFhZWVlYWFhXV1dYWFhZWVlYWFhZWVlYWFhZWVlYWFhYWFhaWlpYWFhYWFhYWFhXV1diYmJZWVlZWVlYWFhYWFhaWlpYWFhZWVlXV1dZWVlYWFhYWFhZWVlZWVlYWFhZWVlYWFhZWVlXV1dcXFxZWVlYWFhWVlZYWFhZWVlXV1dYWFhZWVlbW1tZWVlYWFhaWlpYWFhYWFhYWFhYWFhaWlpZWVlVVVVcXFxYWFhYWFhZWVmPMFkQAAAA5XRSTlMA/vcBE+1J/PkCEv0FZhvuAxp3CB9ZfKJ4Kbo5DPSALwm24U0hMx0ENByRe+Ln9vsx6PPMRSC58n9eZ3onu910FDzc7yud1X6xlsWrYskyqvguhmzIzQcVwm+I1oMqMEwZP1TLRC2+X2kRz64LYTYPIg7qwKHkgYzEcaDYLOzScLNWnqko5oSPmNt54MH1sLKJ09ojTmBq5VOOk1o+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',
  user: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABpFBMVEVMaXFdXV1mZmZZWVlYWFhZWVlVVVVZWVlZWVlZWVlZWVlZWVlcXFxZWVlZWVlYWFhYWFhYWFhZWVlYWFhZWVlYWFhZWVlYWFhYWFhZWVlVVVVZWVlZWVk/Pz9XV1dZWVlYWFhZWVlYWFhYWFhcXFxYWFhYWFhVVVVISEhYWFhZWVkAAABYWFhbW1tZWVlZWVlYWFhZWVlYWFhaWlpaWlpZWVlYWFhZWVlZWVlYWFhZWVlYWFhbW1tXV1dYWFhZWVlYWFhYWFhYWFhXV1dXV1dYWFhYWFhZWVlYWFhZWVlYWFhZWVlYWFhZWVlYWFhYWFhYWFhZWVlZWVlVVVVcXFxZWVlYWFhZWVlYWFhYWFhXV1dXV1dZWVlYWFhYWFhZWVlaWlpYWFhYWFhZWVliYmJYWFhYWFhYWFhZWVlVVVV/f39aWlpZWVlYWFhYWFhZWVlYWFhcXFxYWFhZWVlZWVlYWFhaWlpYWFhVVVVaWlpYWFhZWVlZWVlXV1dZWVlZWVlYWFhZWVlXV1daWlpZWVlZWVlYWFhZWVlZWVlYWFhbW1tZWVkbvEK5AAAAi3RSTlMAEwWP9O0V/vzzjPgk+SXy8ev7+m9Zy23b9RiRuQRXU9jOQmsW15kbB8f2ARcZidbKt41gIuXA0LHptu4nQ04//XzjPSMu92eFtOwUUYBWN3/NrgYLdZhQbpBaIJ1zVGRVaM94DZabydkSAkS/ld3oviGfi2ykd4oMH7PiyHQo3LK9Jmbho8FqjogcUD0cNwAAAAlwSFlzAAALEwAACxMBAJqcGAAAAq1JREFUaN7tmGdTIkEQhpe8CyJIFEFEEUSUoHemM+dwnvGMl3POOXo59Z8+q+6LA8syoafKq+L5ulvvUzW73TM9ikJL03401jjs9w83xqLBJgUZ024ICIphE2L8n70IlGG1dCDFmwMa6BI5k8HIHxmAilz9Ip7fooEBnmXR/BM+MES9IpY/D1UJC62PWl2g7vDn13mBAu0jb/7mAFDRauYUXANKPvDld1hpBVqOS7AH1PRy9bcIvcDK0/l2gYEkhyDEIrjFnp9zsAgc3cyCfWAiziyIsgkszIIYm+AUs+Ahm6CRWWBjEzQwC9xsAtfxE0hfogPZH1n6byq90IJsghVmQbfsZqcUWQSdHPtBmEXwk2fLtEreMhULvcAp+dji4ZyoXtIKLvCONq10+etmyYffEf7j9TLN8X1MZED4Wr2GL0keoeZFh7Qdj1G+d0wRpt7gX1qvwxiUmwOeSoP4JtKsn/ulU9TWXswbEdPz98T+4OhM3sG+cGmPW3oSNrfbluixxNuVGjWOAfbLzsXUUvbwn3+Ur6/QRPJDh4+zS6m3zy7aGdMDoaN15ejfmi19ZXbrB/FKMUzvOLntKu8MfRsLwdUZeyZjn1kNLmz06YwIN85Rxb92+oATf/5s9fy2KRDg3XiV+K41FYRQJ7sMW/95EOa7wSHG/AkQGGyuuD5PAYUXdysI1gCJSf38cRVLMDGql399CtD4plcPTkDkt05/8GEK/G/KBGlA5UHZmcGFK5g7XSIIADJJkUtMjsnZ7sAWTJBrtALotMgrgn+8IgSL+IJBQnAbX5AiBA34giFCMIcvyBICkEBNUBP8bwINP99LCBL4ggIhuIkv2CYE9/AFbeSe+Rk7v79k03+C3LBtZVdV0zbU/Gmdqfc+Xv5j/au20XQBoR60Qvro9/0LAwZf9wGaJzQAAAAASUVORK5CYII=',
  linkedin: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAABblBMVEVMaXFZWVlYWFhVVVVYWFhYWFhZWVlYWFhZWVlWVlYAAABmZmZZWVlZWVlZWVlYWFhZWVlYWFhYWFhYWFhZWVlISEg/Pz9ZWVlYWFhfX19ZWVlYWFhZWVlYWFhYWFhYWFhaWlpWVlZZWVlZWVlVVVVVVVVfX19YWFhVVVVYWFhZWVlYWFhZWVlZWVlYWFhYWFh/f39ZWVlYWFhbW1tXV1dYWFhaWlpYWFhZWVlYWFhZWVlYWFhZWVlZWVlaWlpYWFhYWFhYWFhVVVVZWVlYWFhXV1diYmJZWVlVVVVZWVlZWVlZWVlaWlpYWFhYWFhZWVlYWFhZWVlYWFhVVVVaWlpYWFhZWVlcXFxYWFhZWVlYWFhYWFhZWVlYWFhYWFhYWFhZWVlYWFhaWlpZWVlYWFhcXFxYWFhZWVlXV1dZWVlYWFhaWlpZWVlYWFhZWVlYWFhaWlpbW1tYWFhZWVlXV1dYWFhZWVlYWFhXV1dZWVku2URHAAAAeXRSTlMAmvQG7vFY75c1AQXZ+f79+MNrx/YHBMuwEF4a88lI8mMszmwMGwjMFfpqxm/WeaEC/LgZIG4ic35UotppSlWHymgJts9DDd8YxaYlEYHEkaq6LhJE94YhpLTsihRRmTfIf1Io5gtLvSZbk2bCmKMXTDjB8DqnsZUdtFrL+QAAAAlwSFlzAAALEwAACxMBAJqcGAAAAdhJREFUaN7t2tdPAkEQBnCkqSgdVBAQ7CLYe++9995773r/vSGAzCA8YL59MOw8fjPZX0K44/YWlUqWLFg5271qBVRqu02TvH7VmAKtmla+vqZcAVdxAQNsCryGGVCPBxYY4MEDOgYoAkoCWQ7oBg9d1VfBsChguTaaFnWZhAAtRT95hQhAu0oaJQKAY9pYEQAEWKcXDpj5jbwEDhgNrDOD/4hqWecOD+hpo3IdD9jpZ9Qp4kLzJ/IOrQjAchKP28Ji7kWm+6pIWG3bEXa7Nh815oyq5Q/OfwDU2rrMAZOWlCUWlpGsObr2lvs08pA+Ehi4mcwIyKWNvFiYT7LCyFBDJUmafENgoCXp8V9lCNUhgc+m31sknwUHNKbchLlhwIcz9TZvCQV8pdlHbhpBQNp6Fw30iwbOjVBg7vFXZIUBhutLh6I4rPs83kABE9vx4QaW+1HAQ2J6nOZrIOCNTL/QxiwI6CHTzXuk4QIBZ3T8gr4TwgCl7HoK0S+XGQI8M2CejjsgwBMDdum4DgJ0M2ARD+gZcCAayJGABCQgAQlIIEsB+TJEAn8DyvDrv4o+arQyIIgHbvlxbw16fS8/JVCVgg+Ui6eSN+maCjvsyN0z3Vcg/8QgC1ff/RVqUE/ocDsAAAAASUVORK5CYII=',
  link: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACcFBMVEVMaXFYWFhMTExVVVVZWVlVVVV/f39ZWVlVVVVZWVkAAABYWFhiYmJYWFhWVlZYWFhZWVk/Pz9YWFhYWFhZWVlZWVlZWVlbW1tZWVlYWFhZWVlZWVlYWFhaWlpZWVlYWFhYWFhYWFhZWVlYWFhYWFhYWFhZWVlYWFhaWlpfX19mZmZZWVlZWVlYWFhXV1dYWFhaWlpZWVlaWlpISEhYWFhYWFhZWVlZWVlYWFhZWVlbW1tYWFhYWFhYWFhbW1tYWFhZWVlYWFhZWVlYWFhVVVVZWVlYWFhXV1dZWVlYWFhYWFhaWlpdXV1YWFhcXFxYWFhVVVVcXFxZWVlWVlZYWFhaWlpYWFhZWVlZWVlZWVlZWVlZWVlaWlpYWFhYWFhXV1dVVVVXV1dYWFhaWlpZWVlYWFhZWVlZWVlZWVlYWFhXV1dZWVlZWVlbW1tYWFhZWVlYWFhVVVVZWVlZWVlaWlpZWVlYWFhZWVlYWFhZWVlYWFhYWFhVVVVYWFhYWFhYWFhZWVlaWlpZWVlYWFhaWlpYWFhZWVlYWFhZWVlYWFhXV1dZWVlaWlpZWVlfX19ZWVlZWVlZWVlYWFhaWlpcXFxZWVlZWVlYWFhaWlpZWVlYWFhYWFhXV1dZWVlVVVVXV1dYWFhYWFhYWFhYWFhZWVlZWVlZWVlYWFhYWFhYWFhYWFhWVlZZWVlYWFhaWlpZWVlaWlpZWVlZWVlYWFhZWVlYWFhYWFhZWVlYWFhZWVlbW1tYWFhcXFxZWVlYWFhZWVlYWFhYWFhZWVlZWVlYWFhVVVVZWVlYWFhYWFhXV1dYWFhZWVlZWVlZWVlZWVlZWVmH6k3iAAAAz3RSTlMA8QoJaQMC/gzCAe8N6y+s1ATuxPuoag66vvxbmE+AF61Zhi7V2uXAOxAFkfjyI2sf5zMHZZPFU4ixHFSH1xmwjNu3jR75bUN4NP0wE5UhVgYLlDWyZvrwjlDWri3JhFobIPQinXN+omSCXZK2J059lhimqxHzqtmvWJBxEvd2XGx3tOw+0nuz4rh0qVKPCPY84XlMJOTRzFUoyitXiQ9Ju5zpaJdeP5sa44EyckhjOUSl6J/tpIXTirkqURZhN8ieS3q9MRWjwcY6p/XQy4soUNhhAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFaElEQVRo3rVa90MTSRSeEBNCIAQwEE5KgCAIAgeiKCDFU7AACoogIti7nmc5u556p+fZe9frvffee51/6TazG8m8mZ3dyW7eT8mb733zZubtmzezi5AFWVX8X1tGnQNXf5TR1r5yFbJXVpTMwrTMmhm2jd3dfAfz5OYily38BROxnkwcsk6/9jEski9rLPJPT8ViqX7KCr3rCWws+Zlx89fvxmbk2aQ4+T3l2JykxNdD0iPYrJyIZ5Y8Kdi85CfU/4jkJdR/RRw1CfVfkdwE82P8uPX5ycresiLkDoXzsnmP9z6XRf8/KUx7iOhd+hILmG7N/+EKClSRzSBGrPAnF3oBzLshGYLuWuAfx0FegahRe/mR1w9gLSb4e18wzY9Q8GkAbDXk927k8T+qB58HkPcMO3hLwn9F+gZp6H0j/mcmyPivCFiFIgP+4JtS/iuyiQbfMOhgtqT/CC2h0WVi/q8GJP1HaDXI2WL0Lml+9IDGB4TghQHJ+VGkRmYEs6X9V8KOtrgurLHKYAfHtJbtP1e16Gy5ObTFP6IO5kL+aVpDOkkI65w8o+doE7+og2WA/3et1En3qf+/72ZtKqpomxmiDm7S2AVdND/Gy5sYm1Hg1C0BfxrIEtmqOuwbU5UeBjY714CoCAk6mA+ceY1oQ6djdePpdXCfADYZohkqBgWEugJ/0doXKZOZMCwKRR3M4+XFhQsARXFsiDJ78m029ofOXNJ+ltDYZUQ5jikQWwX8Dez+m6uoe9TfW2nwUW7oYrwxajuV4ccFOvv7FPJnGg1+h6eMBMr7qu2cDqap1K1TP7zMW7EviPIcu0FsU43z2Jbn9eqTHeTvKRr9oQpazA4hnTR8wDT069Y/58n/V8HepBVynUwOf53onX8DdWqtbn2lRuSfwGCKBrwK64CseqLvF1e+sfXbSaLpTubnUnSIfw44Risn0/zO2PPpxUqiW05bdKyOgveCDno4qaWIPhu49lOtbxNlPuZmO+Vx7uBlnMojMaryC/QAwP1ADlEOwak4FIWDUEreQ7Sfx5zDPTT/AcDUpl4bwGNR4Cq/+MHfEe2wLv9aBzBYo2bhdXAIgU4tLdfRevV65aAeP2Lvf+YSfRebXRarFmdp7RWi3KHHX8A+5MXcHTwi53gFzWexmZbhd7H3Y1XaZjuJHcJW0nCGF/Pz+fyomc0um6JtnzJt7bzNSM0tqCdyvGf40R1B/fYGXH68lOgn08rNGvzS+h/d7P0nw58T07oSNE5QM1gjrV0v2npLYAclVPMPdONv6kMLBtYs6gDe3zbSg3Q2xDZe/pVbVKaL7p/hneQ3AHC4dKzRF+YexAaSzFc/nEqm6WFW9WmeLgHR+4dohtpp7JEgC+keofkrG7BEhdsGLjx5mAskKrOiW9orMCwmiTrI4IU5I3ktF3O3a787If9pr6iDOgln1IzDpA/hDKEnafAcI/6psDjFA2IbUCIEpf3HHyOZEdRK81ddE5vU8Ysf0/wgsxhG0Xuy/Jc9Bh2AvLhXtL4c/oBh2IEnObVPyn98yjCsYcI/KuM/3uU27KAJmPiCEv7n9pq4VYT7gd9rmj/FY4KfPXr+xLviiJ8fhZk9eQMcg3uzBX6E/mVMs+lb7z2N2Ao/WsR5LXB87Iqg73gWj398mvl3e7w3n46zW37Zmdna9e3uQWzNf0W+ln95IuF/RLYlmB9dq5bjl5ofvXW20X/eec9e/yORVJRQ/yPvscy+Ryz3oPik3m+Kf78z7nfdmWbW4aAroV8DOA5Y/N6gRvw9Q3+tDV9k7NOlL33Xpm9KRrj0DQVuZJfcHW2B97eFt5G90nrv/vDJskGcev2Gf8atkBWq/wGmwCxnJZp1EQAAAABJRU5ErkJggg==',
  portfolio: 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAACDVBMVEVMaXFZWVlYWFhZWVlYWFhbW1tZWVlZWVlVVVVWVlZYWFhcXFxYWFhZWVlYWFhaWlpZWVlYWFhYWFhMTExZWVlYWFhZWVlYWFhZWVlbW1t/f39YWFhaWlpYWFhZWVlYWFhZWVlXV1dcXFxZWVlXV1diYmJZWVlYWFhYWFhWVlZYWFhZWVlZWVlYWFgAAABZWVlYWFhZWVlYWFhYWFhZWVlaWlpZWVlZWVlZWVlYWFhfX19ZWVlYWFhYWFg/Pz9VVVVYWFhYWFhZWVlfX19ZWVlYWFhXV1dYWFhmZmZaWlpVVVVZWVlcXFxZWVlYWFhYWFhcXFxZWVlYWFhVVVVZWVlYWFhZWVlYWFhbW1tZWVlYWFhXV1dYWFhYWFhZWVlYWFhZWVlXV1dZWVlXV1dVVVVYWFhaWlpYWFhYWFhYWFhVVVVYWFhYWFhZWVlZWVlaWlpYWFhVVVVYWFhZWVlYWFhYWFhVVVVYWFhZWVlVVVVdXV1ZWVlZWVlZWVlZWVlYWFhYWFhZWVlaWlpZWVlYWFhYWFhZWVlXV1dZWVlXV1dZWVlYWFhZWVlYWFhZWVlYWFhZWVlYWFhZWVlZWVlYWFhYWFhbW1tZWVlYWFhZWVlZWVlYWFhZWVlYWFhYWFhYWFhZWVlaWlpYWFhaWlpYWFhZWVlYWFhZWVlZWVlYWFhYWFhYWFhZWVlYWFhZWVl0buXFAAAArnRSTlMA+S7+4zg2+AwvkAv3kpMwke70CuT9OTeuDgJLLcq06+pDJN9dDcViyTLy7RSEAeEa/K2q9mZqP1iHCKPvKwQewIKxEF7eV8QFYxLoIc5ChRbNxwbcMTz6GW/dJsHx9YGAKVtAFRdP2G7SD9VzfstV6QnPfZY0GHyMAxPz1HJTr3G/d9PsippG53RHfyjmj7Ko4PBQrLAcoGXC0KR7VlmbqxFIPrjIX47WaGuNZ5k9Mdc9AAAACXBIWXMAAAsTAAALEwEAmpwYAAADjklEQVRo3rVa9V8bMRRPe0BLi7Wl0AIthbLh7j5kMIZvuAy3ubu7u7v78jeug0IZ5ZK8pPn+eJ/vfb+53OXdy3tBiAkzSfMR9XvdJivGVtOQw2Yo+fQBBQrJPyrd2B/K93dv7eLqCWMLWB06Z2qWiLolNQ3TkO9M4pXfWleFmXAwZQuH/M7ODMyMoOI4oHxOYysGob0P9modGAqlMo99+IM6zIGH+xj1j27CfFAM21n003MxN7qT6fqNOiyAzBia/hmFcHuYIcRojI0iDcF8k6xfR/zatUuk0CACSVNG0h8j6Ydpl2mhpGewtqjrnyfOr8FHDCdGp3g1/RSFaBDiYwYTieYjKt9nEfkL0fuoejLTte7XOhCNmQ02UKi7LP76cdT1yzxFHhz2N7hBXURRPnIENWr4vWgtfQHrQpfJkRoquXV6zQSlsfxWvA6R0Qxk2/8GqUyRRhcerNfHRGiYyB2r9Rv2Q2MOnd+9+i+aDY85dH7KqvxhlCPmUPm5FSoPwBpzqPySlfyqintBEfmnly+38MYcGr/Ue9mheoMdaLCGX+hNggRiDpnfv5QqjQvEHAp/fvHqgkjMIfM/L+4vxGIOka8zei4XC8YcIn/WY1CJJeKPx8At02AOoWmZ+thqR1NSDTyL+adcg99sAZMfv1CBXAMb2ibXYCPaLNegByXKNUhEO+QamJFVrkE+6hfKiBimKEMkI6LjJXIJZEQMeIXmBDIiBrxBIwK/ezpG7aiQb5PEiNcIlcs0cHv2al8lTpH5XzYQIpAR0fQXN2oVGv6MiIzc3iWtLwIZEWn440bvYL+JZEQqVaNL9ycbVmb7o5CW0ry7r4tcSasQCdgFCQxFuifc8qZHTFXGJl79oAnGEngG5/irWQuxV/kMmpgrvdUKj74TUKq+wGOQDjC4xqHvAFXbm+EGx0AGSXCDDlhDoh5s0AYzGAZnEdDu0y2owQDQoKYdaBAL7Tr1AiepGNzXGoRuXaDIKQQZ6GbADgdgu6lD8OZfF2g7dYejv3jbBHHgaZKWngIYZPN0YIMBz3COq8ervQ7YAnOhrUeyAcobYTS4x9sJryln+0df4e/lx2eyGEwJnBaoddL1XSeEjjvMPqAZHBc8UGHpJOfwl7OQKGrvmglp+9lAHDo5uUe12HMxQOdaLGUF63U4i56jwEFf9t61phj6eBgFGBOT5ba0aJMmzDT04lnJU8itfwGRUUBmQPW0NAAAAABJRU5ErkJggg==',
};
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
// ---------- 排版（S 为空时只计算高度）----------
function layout(S, cfg, assets) {
  const { fs, ts, lh, g } = cfg;
  const T = (x, top, w, size, text, o = {}) => {
    if (!S) return;
    const l = o.lh || size * 1.25, n = o.lines || 1;
    S.addText(o.runs || text, {
      x: P(x), y: P(o.boxY ?? (top - size * 0.12 - (l - size) / 2)), w: P(w), h: P(o.boxH ?? l * n), fontFace: o.font || BODY, fontSize: size,
      color: o.color || BLACK, bold: o.bold ?? true, align: o.align || 'left', valign: o.valign || 'top', margin: 0, lineSpacing: l,
      isTextBox: true, fit: 'none', wrap: o.wrap ?? true, objectName: o.name || strip(text).slice(0, 20),
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
    const lat = !/[^\x00-\x7F]/.test(c.text), icon = ICON_DATA[c.type];
    const tx = icon ? 18.1 : 8, w = tx + textW(c.text, 8, lat, true) * 1.06 + 10;   // 留余量，避免字体差异把文字挤成两行
    RR(px, 81.5, w, 17, 'FFFFFF', 8.5, { tr: 70, name: '联系方式底框' + (i + 1) });
    if (S && icon) S.addImage({ data: icon, x: P(px + 6.3), y: P(81.5 + (17 - 8.5) / 2), w: P(8.5), h: P(8.5) });
    T(px + tx, 0, w - tx, lat ? 8 : 8.1, c.text, { color: 'FFFFFF', font: lat ? LAT : CJK, boxY: 81.5, boxH: 17, valign: 'middle', wrap: false, name: '联系方式' + (i + 1) });
    const url = contactLink(c);
    if (S && url) S.addImage({ data: CLEAR, x: P(px), y: P(81.5), w: P(w), h: P(17), hyperlink: { url, tooltip: c.text }, objectName: '链接' + (i + 1) });
    px += w + 4;
  });
  if (D.summary) {
    if (S) S.addShape('rect', { x: P(HEAD_X), y: P(112.8), w: P(568.5 - HEAD_X), h: P(1), fill: { color: 'FFFFFF', transparency: 70 }, line: { type: 'none' }, objectName: '分隔线' });
    T(HEAD_X, 128.9, 526, 8.1, D.summary, { color: 'FFFFFF', lh: 12, lines: sumLines, bold: BB, name: '个人简介' });
  }
  if (S && assets.photo) S.addImage({ data: assets.photo, x: P(493.2), y: P(35.0), w: P(62.4), h: P(62.4), rounding: true, objectName: '照片' });

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




// ---------- 简洁单栏版式（layout: "simple"）：白底、无色块，右栏卡片排进单栏 ----------
const S_X0 = 51, S_X1 = W - 51, S_CW = S_X1 - S_X0, S_TOP = 40, S_BOTTOM = H - 36;
// 单栏一页能装更多内容：前 4 档比原始版式更宽松（内容少时自动放大），之后与蓝色版相同
const LEVELS_S = [
  { fs: 10.5, ts: 11.5, lh: 18, g: 1.3 },
  { fs: 10, ts: 11, lh: 17, g: 1.2 },
  { fs: 10, ts: 11, lh: 16, g: 1.1 },
  { fs: 9.5, ts: 10.5, lh: 15.5, g: 1.05 },
  ...LEVELS,
];
const S_BASE = 4;   // LEVELS_S[S_BASE] 就是原始版式（level 0），更宽松的档位 level 为负数
const INFO_RE = /个人信息|基本信息|基本资料|个人资料|基本情况|personal/i;
const splitParts = t => String(t || '').split(/\n|\s+[·•｜|]\s+|[｜|]/).map(x => x.trim()).filter(Boolean);
// 把数据整理成单栏：个人信息并入页眉，其余右栏卡片排在左栏板块后面
function simpleParts(merge) {
  const own = Array.isArray(D.info);
  const info = own ? D.info.map(String).filter(x => x.trim()) : [];
  const secs = (D.left || []).map(s => ({ title: s.title, kind: 'entries', entries: s.entries || [] }));
  for (const c of D.right || []) {
    if (!own && INFO_RE.test(c.title || '')) {
      if (c.tags) info.push(...c.tags.map(String).filter(x => x.trim()));
      else (c.items || []).forEach(it => {
        const parts = splitParts(it.content);
        if (it.name && !INFO_RE.test(it.name) && parts.length) parts[0] = it.name + (EN ? ': ' : '：') + parts[0];
        info.push(...parts);
      });
      continue;
    }
    if (c.tags) secs.push({ title: c.title, kind: 'tags', tags: c.tags, card: true });
    else secs.push({ title: c.title, kind: 'rows', items: c.items || [], card: true });
  }
  // 排不下时：把右栏来的卡片合成一个「其他信息」板块，每张卡片一行
  if (merge) {
    const cards = secs.filter(x => x.card);
    if (cards.length >= 2) {
      const sep = EN ? '; ' : '；';
      const items = cards.map(c => ({
        name: c.title,
        content: c.kind === 'tags'
          ? c.tags.map(String).filter(x => x.trim()).join(EN ? ', ' : '、')
          : c.items.filter(it => String(it.name || '').trim() || String(it.content || '').trim()).map(it => {
              const body = splitParts(it.content).join(EN ? ', ' : '，');
              const head = String(it.name || '') + (it.date ? (EN ? ' (' + it.date + ')' : '（' + it.date + '）') : '');
              return head && body ? head + (EN ? ': ' : '：') + body : head || body;
            }).join(sep),
      }));
      return { info, secs: [...secs.filter(x => !x.card), { title: EN ? 'Additional Information' : '其他信息', kind: 'rows', items }] };
    }
  }
  return { info, secs };
}

function layoutSimple(S, cfg, assets, merge) {
  const { fs, ts, lh, g } = cfg;
  const k = fs / 9;   // 条目内部间距随字号缩放
  const gi = Math.min(Math.max(g, 0.75), 1);   // 条目内部只收紧、不放大；放大只作用于板块和条目之间
  const auto = D.autoBold !== false;
  const LINE = tint(THEME, 0.35);
  const T = (x, top, w, size, text, o = {}) => {
    if (!S) return;
    const l = o.lh || size * 1.25, n = o.lines || 1;
    S.addText(o.runs || text, {
      x: P(x), y: P(o.boxY ?? (top - size * 0.12 - (l - size) / 2)), w: P(w), h: P(o.boxH ?? l * n), fontFace: o.font || BODY, fontSize: size,
      color: o.color || BLACK, bold: o.bold ?? true, align: o.align || 'left', valign: o.valign || 'top', margin: 0, lineSpacing: l,
      isTextBox: true, fit: 'none', wrap: o.wrap ?? true, objectName: o.name || strip(text).slice(0, 20),
    });
  };
  const RR = (x, y, w, h, color, r, o = {}) => S && S.addShape(r ? 'roundRect' : 'rect', { x: P(x), y: P(y), w: P(w), h: P(h), fill: { color }, line: { type: 'none' }, rectRadius: r ? P(r) : undefined, objectName: o.name });
  const { info, secs } = simpleParts(merge);
  const hasPhoto = !!assets.photo, PH = 72;
  const HW = S_CW - (hasPhoto ? PH + 12 : 0);   // 页眉文字区宽度（给照片让位）

  // ===== 页眉 =====
  let y = S_TOP;
  T(S_X0, y, HW, 22, D.name, { font: EN ? LAT : YH, name: '姓名' });
  y += 22 * 1.25 + 6;
  if (D.title) {
    T(S_X0, y, HW, 10.5, (D.titleLabel ?? (EN ? '' : '求职意向：')) + D.title, { color: '333333', bold: BB, name: '求职意向' });
    y += 10.5 * 1.25 + 7;
  }
  const contacts = (D.contacts || []).filter(c => String(c.text || '').trim());
  if (contacts.length) {
    let cx = S_X0, row = y;
    contacts.forEach((c, i) => {
      const text = String(c.text), lat = !/[^\x00-\x7F]/.test(text), icon = ICON_DARK[c.type];
      const iw = icon ? 13 : 0, tw = textW(text, 9, lat, BB) * 1.06 + 2, w = iw + tw;
      if (cx > S_X0 && cx + w > S_X0 + HW) { cx = S_X0; row += 16; }
      if (S && icon) S.addImage({ data: icon, x: P(cx), y: P(row + 0.6), w: P(9), h: P(9) });
      T(cx + iw, 0, tw, lat ? 9 : 9, text, { color: '333333', bold: BB, font: lat ? LAT : CJK, boxY: row - 2, boxH: 14, valign: 'middle', wrap: false, name: '联系方式' + (i + 1) });
      const url = contactLink(c);
      if (S && url) S.addImage({ data: CLEAR, x: P(cx), y: P(row - 2), w: P(w), h: P(14), hyperlink: { url, tooltip: text }, objectName: '链接' + (i + 1) });
      cx += w + 16;
    });
    y = row + 16;
  }
  if (info.length) {
    const t = info.join('  |  '), n = nLines(t, 9, HW);
    T(S_X0, y, HW, 9, t, { color: GREY, bold: BB, lh: 14, lines: n, name: '个人信息' });
    y += n * 14 + 2;
  }
  if (S && hasPhoto) S.addImage({ data: assets.photo, x: P(S_X1 - PH), y: P(S_TOP - 4), w: P(PH), h: P(PH), rounding: true, objectName: '照片' });
  if (hasPhoto) y = Math.max(y, S_TOP - 4 + PH + 4);
  if (D.summary) {
    y += 6;
    const n = nLines(D.summary, fs, S_CW);
    T(S_X0, y, S_CW, fs, D.summary, { runs: rich(D.summary, GREY, auto), color: GREY, bold: BB, lh, lines: n, name: '个人简介' });
    y += n * lh;
  }
  y += 10 * g;

  // ===== 板块 =====
  const gap = 6 * g;
  secs.forEach(sec => {
    const th = ts + 1;
    RR(S_X0, y - th * 0.05, 3, th * 1.1, THEME, 1.2);            // 小竖条
    T(S_X0 + 8, y, S_CW - 8, th, sec.title, { name: sec.title });
    RR(S_X0, y + th * 1.25 + 2, S_CW, 0.75, LINE, 0, { name: '分隔线' });   // 细线
    y += th * 1.25 + 2 + 7 * Math.max(g, 0.8);

    if (sec.kind === 'entries') {
      sec.entries.forEach((e, ei) => {
        const bare = !e.name && !e.date && !e.sub && !e.subRight;
        const dw = e.date ? textW(e.date, fs, EN, true) + 4 : 0;
        const nameW = S_CW - dw - 8, nameN = e.name ? nLines(e.name, ts, nameW, EN, true) : 1;
        let cur;
        if (bare) cur = y;
        else {
          if (e.name) T(S_X0, y, nameW, ts, e.name, { name: String(e.name).slice(0, 20), lh: ts * 1.25, lines: nameN });
          if (e.date) T(S_X1 - 140, y - 1.5 * k, 140, fs, e.date, { align: 'right', name: '时间' });
          if (e.subRight) T(S_X1 - 140, y + 9.3 * k, 140, fs, e.subRight, { align: 'right', name: '右侧副信息' });
          const ny = y + (nameN - 1) * ts * 1.25;
          if (e.sub) {
            const subW = e.subRight ? S_CW - textW(e.subRight, fs, EN, true) - 10 : S_CW, subN = nLines(e.sub, fs, subW);
            const sy = ny + 12.5 * k * gi;
            T(S_X0, sy, subW, fs, e.sub, { color: GREY, bold: BB, lh: fs * 1.3, lines: subN, name: '副标题' });
            cur = sy + (subN - 1) * fs * 1.3 + 16.7 * k * gi;
          } else cur = ny + 18.3 * k * gi;
        }
        for (const dl of e.details || []) {
          const n = nLines(dl, fs, S_CW);
          const mm = String(dl).match(/^([^：:]{1,12}[：:]\s?)([\s\S]*)$/);
          const runs = mm ? [{ text: mm[1], options: { color: BLACK, bold: true } }, ...rich(mm[2], GREY, auto)] : rich(dl, GREY, auto);
          T(S_X0 - 1, cur, S_CW, fs, dl, { runs, color: GREY, bold: BB, lh, lines: n, name: '明细' });
          cur += n * lh;
        }
        if (e.bullets && e.bullets.length) {
          const BX = S_X0 + 4.5, BW = S_CW - 4.5;
          const total = e.bullets.reduce((a, b) => a + nLines(b, fs, BW - BUL_IND), 0);
          if (S) {
            const paras = [];
            e.bullets.forEach((b, i) => {
              const runs = rich(b, GREY, auto);
              runs.forEach((r, j) => { if (j === runs.length - 1 && i < e.bullets.length - 1) r.options.breakLine = true; paras.push(r); });
            });
            S.addText(paras, { x: P(BX), y: P(cur - fs * 0.12 - (lh - fs) / 2), w: P(BW), h: P(lh * total), fontFace: BODY, fontSize: fs, color: GREY, bold: BB, valign: 'top', margin: 0, lineSpacing: lh, isTextBox: true, fit: 'none', objectName: '列表_' + String(e.name || '').slice(0, 16) });
          }
          cur += total * lh;
        }
        if (e.text) {
          const n = nLines(e.text, fs, S_CW);
          T(S_X0 - 1, cur, S_CW, fs, e.text, { runs: rich(e.text, GREY, auto), color: GREY, bold: BB, lh, lines: n, name: '描述' });
          cur += n * lh;
        }
        y = cur + (ei < sec.entries.length - 1 ? 7 * k * g : 0);
      });
    } else if (sec.kind === 'rows') {
      // 小项：每项一行——名称（黑）｜内容（灰）｜日期（右对齐）
      const items = sec.items.filter(it => String(it.name || '').trim() || String(it.content || '').trim());
      const named = items.some(it => String(it.name || '').trim());
      const nameW = named ? Math.min(Math.max(...items.map(it => textW(it.name || '', fs, EN, true))) + 16, S_CW * 0.34) : 0;
      const dW = Math.max(0, ...items.map(it => it.date ? textW(it.date, fs, EN, true) + 12 : 0));
      const cX = S_X0 + nameW, cW = S_CW - nameW - dW;
      items.forEach((it, i) => {
        const content = splitParts(String(it.content || '').replace(/\n/g, '\n')).join(EN ? '; ' : '；');
        const nN = named && it.name ? nLines(it.name, fs, nameW - 8, EN, true) : 1;
        const cN = content ? nLines(content, fs, cW) : 1;
        if (named && it.name) T(S_X0, y, nameW - 8, fs, it.name, { lh, lines: nN, name: String(it.name).slice(0, 20) });
        if (content) T(cX, y, cW, fs, content, { runs: rich(content, GREY, auto), color: GREY, bold: BB, lh, lines: cN, name: '内容' });
        if (it.date) T(S_X1 - 120, y, 120, fs, it.date, { align: 'right', lh, name: '时间' });
        y += Math.max(nN, cN) * lh + (i < items.length - 1 ? 2 * g : 0);
      });
    } else {
      const t = (sec.tags || []).map(String).filter(x => x.trim()).join(EN ? ', ' : '、');
      const n = nLines(t, fs, S_CW);
      T(S_X0 - 1, y, S_CW, fs, t, { runs: rich(t, GREY, auto), color: GREY, bold: BB, lh, lines: n, name: '标签' });
      y += n * lh;
    }
    y += gap + 4 * g;
  });
  return y - gap - 4 * g + lh * 0.25;
}

// ---------- 规划：选档位 + 左右栏平衡 ----------
function plan() {
  if (D.layout === 'simple') {
    const fitS = merge => {
      const lv0 = LEVELS_S.findIndex(c => layoutSimple(null, c, assets, merge) <= S_BOTTOM);
      const lv = lv0 >= 0 ? lv0 : LEVELS_S.length - 1;
      return { lv, merge, overflow: lv0 >= 0 ? 0 : layoutSimple(null, LEVELS_S[lv], assets, merge) - S_BOTTOM };
    };
    let best = fitS(false);
    // 比原始版式还挤时，试试把右栏卡片合并成「其他信息」，能放得更宽松就采用
    if (best.lv > S_BASE && D.balance !== false) {
      const m = fitS(true);
      if (m.lv < best.lv || (m.lv === best.lv && m.overflow < best.overflow - 1)) best = m;
    }
    const merged = best.merge ? simpleParts(false).secs.filter(x => x.card).map(x => x.title) : [];
    return { level: best.lv - S_BASE, overflow: best.overflow, moved: merged, merge: best.merge, cfg: LEVELS_S[best.lv], warns: [], pageBottom: S_BOTTOM };
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

  return { level, overflow, moved, cfg, warns };
}

return {
  plan,
  draw: (S, cfg, merge) => (D.layout === 'simple' ? layoutSimple(S, cfg, assets, merge) : layout(S, cfg, assets)),
  get data() { return D; },
  EN, W, H, BOTTOM,
};
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

function summary(p, bottom, EN, BOTTOM) {
  return { level: p.level, fontSize: p.cfg.fs, lineHeight: p.cfg.lh, bottom: Math.round(bottom), pageBottom: p.pageBottom ?? BOTTOM,
           overflowPt: Math.round(p.overflow), lang: EN ? 'en' : 'zh', movedToRight: p.moved, warnings: p.warns };
}

async function generate({ PptxGenJS, JSZip, data, photo, outputType }) {
  const eng = createEngine(data, { photo });
  const p = eng.plan();
  const pres = new PptxGenJS();
  pres.defineLayout({ name: 'A4P', width: eng.W / 72, height: eng.H / 72 });
  pres.layout = 'A4P';
  pres.title = (eng.data.name || '') + ' 个人简历';
  const S = pres.addSlide();
  S.background = { color: 'FFFFFF' };
  const bottom = eng.draw(S, p.cfg, p.merge);
  const zip = await JSZip.loadAsync(await pres.write({ outputType: 'arraybuffer' }));
  const f = 'ppt/slides/slide1.xml';
  zip.file(f, postXml(await zip.file(f).async('string')));
  const file = await zip.generateAsync({ type: outputType || 'blob', compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  return { file, info: summary(p, bottom, eng.EN, eng.BOTTOM) };
}

// 预览：用一个「记录器」代替幻灯片，收集每个元素的位置、文字和样式（单位：pt）
function preview(data, opts) {
  const eng = createEngine(data, { photo: (opts || {}).photo });
  const p = eng.plan();
  const items = [];
  const pt = v => v * 72;
  const box = o => ({ x: pt(o.x), y: pt(o.y), w: pt(o.w), h: pt(o.h) });
  const S = {
    addText(text, o) { items.push(Object.assign({ type: 'text', text, o }, box(o))); },
    addShape(shape, o) { items.push(Object.assign({ type: 'shape', shape, o }, box(o))); },
    addImage(o) { items.push(Object.assign({ type: 'image', o }, box(o))); },
  };
  const bottom = eng.draw(S, p.cfg, p.merge);
  return { items, info: summary(p, bottom, eng.EN, eng.BOTTOM), W: eng.W, H: eng.H };
}

return { generate, preview, postXml, createEngine };
});
