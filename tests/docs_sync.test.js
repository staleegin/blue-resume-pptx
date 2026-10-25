// 检查网站用到的副本和源文件一致：node tests/docs_sync.test.js（不一致时运行 tools/sync_docs.sh）
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..');
const pairs = [
  ['skill/blue-resume-pptx/scripts/resume_core.js', 'docs/lib/resume_core.js'],
  ['resume_pptx.py', 'docs/resume_pptx.py'],
  ...['zh_campus.json', 'zh_campus.pptx', 'zh_campus.pdf', 'en_marketing.json', 'en_marketing.pptx', 'en_marketing.pdf',
      'preview_zh_campus.png', 'preview_en_marketing.png', 'avatar.png'].map(f => ['examples/' + f, 'docs/examples/' + f]),
];
let bad = 0;
for (const [a, b] of pairs) {
  const same = fs.readFileSync(path.join(root, a)).equals(fs.readFileSync(path.join(root, b)));
  if (!same) { bad++; console.log('不一致:', a, '≠', b); }
}
console.log(bad ? `${bad} 个文件需要同步，请运行 tools/sync_docs.sh` : `docs/ 副本全部一致（${pairs.length} 个）`);
process.exit(bad ? 1 : 0);
