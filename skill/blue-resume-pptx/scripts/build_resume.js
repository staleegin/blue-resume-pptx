// 单页 A4 简历 PPT 生成器（命令行）
// 用法: node build_resume.js data.json out.pptx
// 排版逻辑全部在同目录的 resume_core.js 里（网页生成器用的也是它）
const fs = require('fs'), path = require('path');
const PptxGenJS = require('pptxgenjs');
const JSZip = require(require.resolve('jszip', { paths: [path.dirname(require.resolve('pptxgenjs'))] }));
const core = require('./resume_core.js');

// 头像：裁成正方形（左右居中、上下偏上取景，尽量保住头顶），由 PowerPoint 的圆形显示
async function squarePhoto(file) {
  const sharp = require('sharp');
  const img = sharp(file).rotate();
  const m = await img.metadata();
  const w = m.orientation >= 5 ? m.height : m.width, h = m.orientation >= 5 ? m.width : m.height;
  const s = Math.min(w, h);
  const buf = await sharp(file).rotate()
    .extract({ left: Math.floor((w - s) / 2), top: Math.floor((h - s) / 4), width: s, height: s })
    .resize(600, 600).flatten({ background: '#ffffff' }).jpeg({ quality: 90 }).toBuffer();
  return 'image/jpeg;base64,' + buf.toString('base64');
}

(async () => {
  const src = process.argv[2];
  if (!src) { console.log('用法: node build_resume.js data.json [out.pptx]'); process.exit(2); }
  const OUT = process.argv[3] || 'resume.pptx';
  const data = JSON.parse(fs.readFileSync(src, 'utf8'));
  let photo = null;
  if (data.photo) {
    const pp = path.resolve(path.dirname(path.resolve(src)), data.photo);
    if (fs.existsSync(pp)) photo = await squarePhoto(pp); else console.warn('照片不存在：' + pp);
  }
  const { file, info } = await core.generate({ PptxGenJS, JSZip, data, photo, outputType: 'nodebuffer' });
  fs.writeFileSync(OUT, file);
  const { warnings, ...rest } = info;
  console.log(JSON.stringify({ out: OUT, ...rest }));
  (warnings || []).forEach(w => console.log('提示：' + w));
  if (info.overflowPt > 0) console.log(`内容超出一页约 ${info.overflowPt}pt（约 ${Math.ceil(info.overflowPt / info.lineHeight)} 行），需要删减。`);
})();
