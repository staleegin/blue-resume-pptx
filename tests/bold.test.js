// 自动加粗规则单测：node tests/bold.test.js
const path = require('path');
const src = require('fs').readFileSync(path.join(__dirname, '../skill/blue-resume-pptx/scripts/build_resume.js'), 'utf8');
const re = eval(src.match(/const NUM_RE = (\/.*\/g);/)[1]);
const skip = k => /^[\d.,]+$/.test(k) || /^\d{4}\s*年$/.test(k) || /^\d{1,2}\s*月$/.test(k);
const bolded = t => { const out = []; t.replace(re, m => { const k = m.trim(); if (!skip(k)) out.push(k); return m; }); return out.join(' | '); };
const cases = [
  ['月均面试约 90 人，9 月试岗留岗率 56%', '90 人 | 56%'],
  ['2026 年 7–10 月约 300 条，3 年经验', '300 条 | 3 年'],
  ['Increased revenue by $2.5M (35%) and 3 million users', '$2.5M | 35% | 3 million'],
  ['Windows 10 and 6-axis robot, 1,200 clients, grew 3x, cut costs 20 percent', '3x | 20 percent'],
  ['超百人次，近 4 个月，STM32 芯片，带教 15 人团队', '超百人次 | 4 个月 | 15 人'],
  ['创办 6 个月推出 5 款应用，近 4 个月复盘 300 条', '6 个月 | 5 款 | 4 个月 | 300 条'],
];
let fail = 0;
for (const [t, exp] of cases) {
  const got = bolded(t);
  if (got === exp) console.log('PASS', t);
  else { fail++; console.log('FAIL', t, '\n  期望:', exp, '\n  实际:', got); }
}
process.exit(fail ? 1 : 0);
