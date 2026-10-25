/* 简历生成器：表单录入 → 用 ResumeCore 实时排版预览 → 浏览器内生成 .pptx（不上传任何数据） */
(function () {
  'use strict';

  // ---------- 小工具 ----------
  var $ = function (s, r) { return (r || document).querySelector(s); };
  function h(tag, props) {
    var el = document.createElement(tag);
    props = props || {};
    Object.keys(props).forEach(function (k) {
      var v = props[k];
      if (v == null) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.slice(0, 2) === 'on') el.addEventListener(k.slice(2), v);
      else if (k === 'value') el.value = v;
      else el.setAttribute(k, v);
    });
    (function add(list) {   // 子元素可以是节点、文字或任意层嵌套的数组
      list.forEach(function (x) {
        if (x == null || x === false) return;
        if (Array.isArray(x)) add(x);
        else el.appendChild(typeof x === 'string' ? document.createTextNode(x) : x);
      });
    })(Array.prototype.slice.call(arguments, 2));
    return el;
  }
  var str = function (v) { return v == null ? '' : String(v); };
  var lines = function (v) { return str(v).split('\n').map(function (s) { return s.trim(); }).filter(Boolean); };
  var tagsOf = function (v) { return str(v).split(/[\n,，、;；]/).map(function (s) { return s.trim(); }).filter(Boolean); };
  var store = {
    get: function (k) { try { return window.localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { if (v == null) window.localStorage.removeItem(k); else window.localStorage.setItem(k, v); } catch (e) { /* 存不下就算了 */ } }
  };
  var toastTimer;
  function toast(msg, action) {
    var t = $('#toast');
    t.replaceChildren(document.createTextNode(msg));
    if (action) {
      t.style.pointerEvents = 'auto';
      t.appendChild(h('button', { class: 'btn btn-small', type: 'button', style: 'margin-left:12px;min-height:26px', text: action.label, onclick: function () { action.run(); t.classList.remove('show'); } }));
    } else t.style.pointerEvents = 'none';
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.classList.remove('show'); }, action ? 6000 : 2400);
  }
  function copy(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    var ta = h('textarea', { style: 'position:fixed;opacity:0' });
    ta.value = text; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); } finally { ta.remove(); }
    return Promise.resolve();
  }

  // ---------- 数据：只保留认识的字段，类型不对就修正（导入的 JSON 不可信） ----------
  var CONTACT_TYPES = [
    ['phone', '电话'], ['email', '邮箱'], ['location', '城市'], ['wechat', '微信'], ['github', 'GitHub'],
    ['linkedin', 'LinkedIn'], ['portfolio', '作品集'], ['web', '网站'], ['link', '链接'], ['birthday', '生日'], ['user', '其他']
  ];
  var LINKABLE = { portfolio: 1, web: 1, github: 1, linkedin: 1, link: 1 };
  function normEntry(e) {
    e = e && typeof e === 'object' ? e : {};
    var o = {};
    ['name', 'date', 'sub', 'subRight', 'text'].forEach(function (k) { if (str(e[k]).trim()) o[k] = str(e[k]); });
    ['details', 'bullets'].forEach(function (k) { var a = Array.isArray(e[k]) ? e[k].map(str).filter(function (s) { return s.trim(); }) : []; if (a.length) o[k] = a; });
    return o;
  }
  function norm(d) {
    d = d && typeof d === 'object' && !Array.isArray(d) ? d : {};
    var o = {
      name: str(d.name), title: str(d.title), summary: str(d.summary),
      contacts: (Array.isArray(d.contacts) ? d.contacts : []).map(function (c) {
        c = c && typeof c === 'object' ? c : {};
        var x = { type: str(c.type) || 'user', text: str(c.text) };
        if (str(c.link).trim()) x.link = str(c.link).trim();
        return x;
      }),
      left: (Array.isArray(d.left) ? d.left : []).map(function (s) {
        s = s && typeof s === 'object' ? s : {};
        var x = { title: str(s.title), entries: (Array.isArray(s.entries) ? s.entries : []).map(normEntry) };
        if (s.side === 'left' || s.side === 'right') x.side = s.side;
        return x;
      }),
      right: (Array.isArray(d.right) ? d.right : []).map(function (c) {
        c = c && typeof c === 'object' ? c : {};
        if (Array.isArray(c.tags)) return { title: str(c.title), tags: c.tags.map(str).filter(function (t) { return t.trim(); }) };
        return { title: str(c.title), items: (Array.isArray(c.items) ? c.items : []).map(function (i) {
          i = i && typeof i === 'object' ? i : {};
          var x = { name: str(i.name), content: str(i.content) };
          if (str(i.date).trim()) x.date = str(i.date);
          return x;
        }) };
      })
    };
    if (/^#?[0-9A-Fa-f]{6}$/.test(str(d.theme))) o.theme = str(d.theme).replace('#', '').toUpperCase();
    if (d.titleLabel != null) o.titleLabel = str(d.titleLabel);
    if (d.lang === 'zh' || d.lang === 'en') o.lang = d.lang;
    if (d.layout === 'simple') o.layout = 'simple';
    if (Array.isArray(d.info)) o.info = d.info.map(str).filter(function (t) { return t.trim(); });
    return o;
  }
  // 交给排版核心前去掉空项，避免空白条目占位
  function clean(d) {
    var o = JSON.parse(JSON.stringify(d));
    o.contacts = o.contacts.filter(function (c) { return c.text.trim(); });
    o.left = o.left.map(function (s) { return Object.assign({}, s, { entries: s.entries.map(normEntry).filter(function (e) { return Object.keys(e).length; }) }); })
      .filter(function (s) { return s.title.trim() || s.entries.length; });
    o.right = o.right.map(function (c) {
      if (c.tags) return c;
      return Object.assign({}, c, { items: c.items.filter(function (i) { return i.name.trim() || i.content.trim(); }) });
    }).filter(function (c) { return c.title.trim() || (c.tags ? c.tags.length : c.items.length); });
    ['name', 'title', 'summary'].forEach(function (k) { if (!o[k].trim()) delete o[k]; });
    return o;
  }

  var BLANK = { name: '', title: '', summary: '', contacts: [{ type: 'phone', text: '' }, { type: 'email', text: '' }, { type: 'location', text: '' }],
    left: [{ title: '教育经历', entries: [{}] }, { title: '工作经历', entries: [{}] }, { title: '项目经历', entries: [{}] }],
    right: [{ title: '技能', tags: [] }] };

  var state = norm(BLANK);
  var photo = null;   // 'data:image/jpeg;base64,...'

  // ---------- 头像：裁成正方形（左右居中、上下偏上取景） ----------
  function squarePhoto(src) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () {
        var w = img.naturalWidth, hh = img.naturalHeight, s = Math.min(w, hh);
        var cv = document.createElement('canvas'); cv.width = cv.height = 600;
        var ctx = cv.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, 600, 600);
        ctx.drawImage(img, (w - s) / 2, (hh - s) / 4, s, s, 0, 0, 600, 600);
        resolve(cv.toDataURL('image/jpeg', 0.9));
      };
      img.onerror = function () { reject(new Error('图片读取失败')); };
      img.src = src;
    });
  }

  // ---------- 保存与预览 ----------
  var timer;
  function changed() {
    clearTimeout(timer);
    timer = setTimeout(function () {
      store.set('brp:data', JSON.stringify(state));
      store.set('brp:photo', photo);
      refresh();
    }, 160);
  }
  function refresh() {
    var st = $('#status');
    try {
      var r = window.ResumeCore.preview(clean(state), { photo: photo ? photo.replace(/^data:/, '') : null });
      window.ResumePreview.render($('#preview'), r);
      var i = r.info, msg, cls = '';
      if (i.overflowPt > 0) { msg = '超出一页约 ' + Math.ceil(i.overflowPt / i.lineHeight) + ' 行，已压到最小字号，请删减一些内容'; cls = 'bad'; }
      else if (i.level > 0) { msg = '内容较多，已自动收紧到第 ' + i.level + ' 档'; cls = 'warn'; }
      else if (i.level < 0) msg = '内容较少，已自动放大到 ' + i.fontSize + ' 号字';
      else msg = '一页放得下，未压缩';
      if (i.movedToRight.length) msg += state.layout === 'simple'
        ? '；「' + i.movedToRight.join('」「') + '」已合并到「其他信息」'
        : '；「' + i.movedToRight.join('」「') + '」已挪到右栏';
      (i.warnings || []).forEach(function (w) { msg += '；' + w; cls = cls || 'warn'; });
      st.textContent = msg; st.className = 'status ' + cls;
    } catch (e) {
      st.textContent = '预览出错：' + e.message; st.className = 'status bad';
    }
  }

  // ---------- 表单 ----------
  function field(label, value, onInput, opt) {
    opt = opt || {};
    var input = opt.multi
      ? h('textarea', { rows: opt.rows || 3, placeholder: opt.ph || null, value: value })
      : h('input', { type: 'text', placeholder: opt.ph || null, value: value });
    input.addEventListener('input', function () { onInput(input.value); changed(); });
    return h('label', { class: 'field' + (opt.cls ? ' ' + opt.cls : '') }, h('span', { text: label }), input);
  }
  function tool(label, title, fn) {
    return h('button', { class: 'btn btn-small btn-quiet', type: 'button', title: title, 'aria-label': title, text: label, onclick: function () { fn(); render(); changed(); } });
  }
  function move(arr, i, d) { var j = i + d; if (j < 0 || j >= arr.length) return; var t = arr[i]; arr[i] = arr[j]; arr[j] = t; }
  function listTools(arr, i, name) {
    return [tool('↑', name + '上移', function () { move(arr, i, -1); }), tool('↓', name + '下移', function () { move(arr, i, 1); }),
      tool('删除', '删除' + name, function () { arr.splice(i, 1); })];
  }

  var THEMES = [['2563EB', '蓝'], ['1E3A8A', '深蓝'], ['0F766E', '墨绿'], ['9F1239', '酒红'], ['6D28D9', '紫'], ['EA580C', '橙'], ['334155', '深灰']];
  function basic() {
    var cur = state.theme || '2563EB';
    var custom = h('input', { type: 'color', value: '#' + cur, 'aria-label': '自定义主色', title: '自定义主色', style: 'width:36px;height:30px;border:0;padding:0;background:none' });
    custom.addEventListener('input', function () { state.theme = custom.value.slice(1).toUpperCase(); render(); changed(); });
    var file = h('input', { type: 'file', accept: 'image/*', hidden: '' });
    file.addEventListener('change', function () {
      var f = file.files && file.files[0];
      if (!f) return;
      if (!/^image\//.test(f.type)) { toast('请选择图片文件'); return; }
      if (f.size > 15 * 1024 * 1024) { toast('图片太大，请选 15MB 以内的'); return; }
      var url = URL.createObjectURL(f);
      squarePhoto(url).then(function (d) { photo = d; render(); changed(); })
        .catch(function () { toast('这张图片读取失败，换一张试试'); })
        .then(function () { URL.revokeObjectURL(url); });
    });
    return h('fieldset', null, h('legend', { text: '基本信息' }),
      h('div', { class: 'grid2' },
        field('姓名', state.name, function (v) { state.name = v; }, { ph: '王小明' }),
        field('求职意向', state.title, function (v) { state.title = v; }, { ph: '前端开发工程师' })),
      field('个人简介（一两句，可不填）', state.summary, function (v) { state.summary = v; }, { multi: true, rows: 2 }),
      h('div', { class: 'field' }, h('span', { text: '版式' }),
        h('div', { class: 'seg', role: 'group', 'aria-label': '版式' },
          [['', '蓝色顶栏'], ['simple', '简洁单栏']].map(function (l) {
            return h('button', { type: 'button', text: l[1], 'aria-pressed': String((state.layout || '') === l[0]),
              onclick: function () { if (l[0]) state.layout = l[0]; else delete state.layout; render(); changed(); } });
          }))),
      h('div', { class: 'field' }, h('span', { text: state.layout === 'simple' ? '主色（用在板块标题的小竖条）' : '主色' }),
        h('div', { class: 'swatches', role: 'group', 'aria-label': '主色' },
          THEMES.map(function (t) {
            return h('button', { class: 'swatch', type: 'button', title: t[1], 'aria-label': t[1], 'aria-pressed': String(cur === t[0]), style: 'background:#' + t[0],
              onclick: function () { state.theme = t[0]; render(); changed(); } });
          }), custom)),
      h('div', { class: 'field' }, h('span', { text: state.layout === 'simple' ? '头像（可不传，放在右上角）' : '头像（可不传，没有时右上角留空）' }),
        h('div', { class: 'photo-row' },
          photo ? h('img', { class: 'photo-thumb', src: photo, alt: '当前头像' }) : h('span', { class: 'photo-thumb', 'aria-hidden': 'true' }),
          h('label', { class: 'btn btn-small' }, photo ? '换一张' : '上传照片', file),
          photo ? h('button', { class: 'btn btn-small btn-quiet', type: 'button', text: '移除', onclick: function () { photo = null; render(); changed(); } }) : null)));
  }

  function contacts() {
    return h('fieldset', null, h('legend', { text: '联系方式' }),
      h('p', { class: 'hint', text: '邮箱和网址会变成可点击的链接。' }),
      state.contacts.map(function (c, i) {
        var sel = h('select', { 'aria-label': '类型' }, CONTACT_TYPES.map(function (t) { return h('option', { value: t[0], text: t[1] }); }));
        sel.value = c.type;
        sel.addEventListener('change', function () { c.type = sel.value; render(); changed(); });
        var txt = h('input', { type: 'text', value: c.text, 'aria-label': '内容', placeholder: '显示的文字' });
        txt.addEventListener('input', function () { c.text = txt.value; changed(); });
        var link = LINKABLE[c.type] ? h('input', { type: 'url', value: c.link || '', 'aria-label': '链接', placeholder: '链接（可不填）' }) : h('span');
        if (LINKABLE[c.type]) link.addEventListener('input', function () { if (link.value.trim()) c.link = link.value.trim(); else delete c.link; changed(); });
        return h('div', { class: 'row' }, sel, txt, link, tool('删除', '删除这条联系方式', function () { state.contacts.splice(i, 1); }));
      }),
      h('button', { class: 'btn btn-small', type: 'button', text: '添加联系方式', onclick: function () { state.contacts.push({ type: 'email', text: '' }); render(); } }));
  }

  function entryForm(sec, e, i) {
    var set = function (k) { return function (v) { if (v.trim()) e[k] = v; else delete e[k]; }; };
    var setArr = function (k) { return function (v) { var a = lines(v); if (a.length) e[k] = a; else delete e[k]; }; };
    return h('div', { class: 'block' },
      h('div', { class: 'grid2' },
        field('名称', str(e.name), set('name'), { ph: '学校 / 公司 · 职位 / 项目名' }),
        field('时间', str(e.date), set('date'), { ph: '2024.09 - 至今' }),
        field('副标题', str(e.sub), set('sub'), { ph: '专业 / 部门（可不填）' }),
        field('右侧信息', str(e.subRight), set('subRight'), { ph: '学历等（可不填）' })),
      field('明细行（每行一条，如「技术栈：React」）', (e.details || []).join('\n'), setArr('details'), { multi: true, rows: 2 }),
      field('要点（每行一条，用 **重点** 加粗）', (e.bullets || []).join('\n'), setArr('bullets'), { multi: true, rows: 4 }),
      field('段落（不带项目符号，可不填）', str(e.text), set('text'), { multi: true, rows: 2 }),
      h('div', { class: 'block-head', style: 'margin:0;justify-content:flex-end' }, listTools(sec.entries, i, '这条')));
  }
  function leftSections() {
    return h('fieldset', null, h('legend', { text: '左栏板块' }),
      h('p', { class: 'hint', text: '教育、工作、项目、校园经历等，名称和数量都可以改。数字（如 90 人、56%）会自动加粗。' }),
      state.left.map(function (sec, i) {
        var title = h('input', { type: 'text', value: sec.title, 'aria-label': '板块名称', placeholder: '板块名称' });
        title.addEventListener('input', function () { sec.title = title.value; changed(); });
        return h('div', { class: 'block' },
          h('div', { class: 'block-head' }, title, listTools(state.left, i, '板块')),
          sec.entries.map(function (e, j) { return entryForm(sec, e, j); }),
          h('button', { class: 'btn btn-small', type: 'button', text: '添加一条', onclick: function () { sec.entries.push({}); render(); } }));
      }),
      h('button', { class: 'btn btn-small', type: 'button', text: '添加板块', onclick: function () { state.left.push({ title: '新板块', entries: [{}] }); render(); changed(); } }));
  }

  function rightCards() {
    return h('fieldset', null, h('legend', { text: '右栏卡片' }),
      h('p', { class: 'hint', text: '技能、证书、荣誉、语言等。「标签」适合一串关键词，「小项」适合带标题和日期的内容。' }),
      state.right.map(function (card, i) {
        var title = h('input', { type: 'text', value: card.title, 'aria-label': '卡片名称', placeholder: '卡片名称' });
        title.addEventListener('input', function () { card.title = title.value; changed(); });
        var isTags = !!card.tags;
        var seg = h('div', { class: 'seg', role: 'group', 'aria-label': '卡片样式' },
          h('button', { type: 'button', 'aria-pressed': String(isTags), text: '标签', onclick: function () {
            if (!card.tags) { card.tags = (card.items || []).map(function (x) { return x.name; }).filter(Boolean); delete card.items; render(); changed(); } } }),
          h('button', { type: 'button', 'aria-pressed': String(!isTags), text: '小项', onclick: function () {
            if (card.tags) { card.items = card.tags.map(function (t) { return { name: t, content: '' }; }); if (!card.items.length) card.items.push({ name: '', content: '' }); delete card.tags; render(); changed(); } } }));
        var body;
        if (isTags) {
          body = field('关键词（用逗号、顿号或换行分隔）', card.tags.join('、'), function (v) { card.tags = tagsOf(v); }, { multi: true, rows: 3 });
        } else {
          body = [card.items.map(function (it, j) {
            return h('div', { class: 'block' },
              h('div', { class: 'grid2' },
                field('小标题', it.name, function (v) { it.name = v; }, { ph: '如 CET-6' }),
                field('日期', str(it.date), function (v) { if (v.trim()) it.date = v; else delete it.date; }, { ph: '可不填' })),
              field('内容（可换行）', it.content, function (v) { it.content = v; }, { multi: true, rows: 2 }),
              h('div', { class: 'block-head', style: 'margin:0;justify-content:flex-end' }, listTools(card.items, j, '小项')));
          }), h('button', { class: 'btn btn-small', type: 'button', text: '添加小项', onclick: function () { card.items.push({ name: '', content: '' }); render(); } })];
        }
        return h('div', { class: 'block' }, h('div', { class: 'block-head' }, title, seg), body,
          h('div', { class: 'block-head', style: 'margin:8px 0 0;justify-content:flex-end' }, listTools(state.right, i, '卡片')));
      }),
      h('button', { class: 'btn btn-small', type: 'button', text: '添加卡片', onclick: function () { state.right.push({ title: '新卡片', tags: [] }); render(); changed(); } }));
  }

  function render() {
    var form = $('#form');
    var y = window.scrollY;
    form.replaceChildren(basic(), contacts(), leftSections(), rightCards());
    window.scrollTo(0, y);
  }

  // ---------- 示例、导入、导出、下载 ----------
  function loadExample(name) {
    return fetch('examples/' + name + '.json').then(function (r) { return r.json(); }).then(function (d) {
      state = norm(d);
      var p = d.photo ? squarePhoto('examples/' + d.photo).catch(function () { return null; }) : Promise.resolve(null);
      return p.then(function (ph) { photo = ph; render(); changed(); });
    });
  }

  var AI_PROMPT = [
    '请把我下面提供的简历材料整理成一份 JSON，用于生成单页 A4 的 PPT 简历。要求：',
    '1. 不要编造任何经历、数字、证书；缺少的关键信息先问我，我回答后再输出。',
    '2. 最终只输出 JSON（放在代码块里），不要输出其他格式。',
    '3. 内容控制在一页：每段经历 3～4 条要点，每条一两句；重点短语用 **…** 包起来，每条最多一处。',
    '格式如下（不需要的字段可以省略）：',
    '{',
    '  "name": "姓名", "title": "求职意向", "summary": "一两句个人简介",',
    '  "contacts": [{"type": "phone", "text": "手机号"}, {"type": "email", "text": "邮箱"}, {"type": "location", "text": "城市"}],',
    '  "left": [',
    '    {"title": "教育经历", "entries": [{"name": "学校", "date": "2021.09 - 2024.06", "sub": "专业", "subRight": "学历", "details": ["主修课程：……"]}]},',
    '    {"title": "工作经历", "entries": [{"name": "公司 · 职位", "date": "2024.07 - 至今", "bullets": ["要点，**重点**"]}]}',
    '  ],',
    '  "right": [',
    '    {"title": "技能", "tags": ["关键词1", "关键词2"]},',
    '    {"title": "证书", "items": [{"name": "证书名", "date": "2024.06", "content": "说明"}]}',
    '  ]',
    '}',
    'contacts 的 type 可选：phone、email、location、wechat、github、linkedin、portfolio、web、link。',
    '我的材料如下：',
    ''
  ].join('\n');

  function parseLoose(text) {
    // 允许直接粘贴 AI 的整段回复：取第一个 { 到最后一个 } 之间的内容
    var t = str(text).trim();
    var a = t.indexOf('{'), b = t.lastIndexOf('}');
    if (a < 0 || b <= a) throw new Error('没有找到 JSON 内容');
    return JSON.parse(t.slice(a, b + 1));
  }

  function bind() {
    var dImp = $('#dlg-import'), dExp = $('#dlg-export');
    $('#btn-import').addEventListener('click', function () { $('#import-error').textContent = ''; dImp.showModal(); });
    $('#btn-do-import').addEventListener('click', function () {
      try {
        state = norm(parseLoose($('#import-text').value));
        render(); changed(); dImp.close(); toast('已导入');
      } catch (e) { $('#import-error').textContent = '导入失败：' + (e.message || e) + '。请检查粘贴的内容是不是完整的 JSON。'; }
    });
    $('#import-file').addEventListener('change', function (ev) {
      var f = ev.target.files && ev.target.files[0];
      if (!f) return;
      f.text().then(function (t) { $('#import-text').value = t; $('#import-error').textContent = ''; });
      ev.target.value = '';
    });
    $('#btn-copy-prompt').addEventListener('click', function () { copy(AI_PROMPT).then(function () { toast('提示语已复制，粘贴给 AI 后在末尾补上你的材料'); }); });
    document.querySelectorAll('[data-example]').forEach(function (b) {
      b.addEventListener('click', function () { loadExample(b.getAttribute('data-example')).then(function () { dImp.close(); toast('已载入示例'); }); });
    });

    $('#btn-export').addEventListener('click', function () { $('#export-text').value = JSON.stringify(clean(state), null, 2); dExp.showModal(); });
    $('#btn-copy-json').addEventListener('click', function () { copy($('#export-text').value).then(function () { toast('已复制'); }); });
    $('#btn-save-json').addEventListener('click', function () {
      saveBlob(new Blob([$('#export-text').value], { type: 'application/json' }), 'data.json');
    });

    $('#btn-clear').addEventListener('click', function () {
      var prev = { state: JSON.parse(JSON.stringify(state)), photo: photo };
      state = norm(BLANK); photo = null; render(); changed();
      toast('已清空', { label: '撤销', run: function () { state = prev.state; photo = prev.photo; render(); changed(); } });
    });

    $('#btn-download').addEventListener('click', function () {
      var btn = this;
      btn.disabled = true; btn.textContent = '正在生成…';
      window.ResumeCore.generate({ PptxGenJS: window.PptxGenJS, JSZip: window.JSZip, data: clean(state),
        photo: photo ? photo.replace(/^data:/, '') : null, outputType: 'blob' })
        .then(function (r) {
          var name = (state.name.trim() || '我的') .replace(/[\\/:*?"<>|]/g, '');
          saveBlob(r.file, name + '简历.pptx');
          toast(r.info.overflowPt > 0 ? '已下载，但内容超出一页，建议删减后再下载' : '已下载，用 PowerPoint 或 WPS 打开后可另存为 PDF');
        })
        .catch(function (e) { toast('生成失败：' + e.message); })
        .then(function () { btn.disabled = false; btn.textContent = '下载 PPT'; });
    });

    document.querySelectorAll('.mobile-tabs button').forEach(function (b) {
      b.addEventListener('click', function () {
        $('.app').setAttribute('data-tab', b.getAttribute('data-tab'));
        document.querySelectorAll('.mobile-tabs button').forEach(function (x) { x.setAttribute('aria-pressed', String(x === b)); });
        if (b.getAttribute('data-tab') === 'preview') refresh();
      });
    });
  }
  function saveBlob(blob, name) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
  }

  // ---------- 启动：有保存的草稿就用草稿，否则载入中文示例 ----------
  bind();
  var saved = store.get('brp:data');
  if (saved) {
    try { state = norm(JSON.parse(saved)); } catch (e) { state = norm(BLANK); }
    var sp = store.get('brp:photo');
    photo = sp && /^data:image\/(jpeg|png);base64,/.test(sp) ? sp : null;
    render(); refresh();
  } else {
    loadExample('zh_campus').catch(function () { render(); refresh(); });
  }
})();
