/* 把 ResumeCore.preview() 算出的元素画成网页里的一张 A4 纸（1pt = 1px，再整体缩放） */
(function () {
  'use strict';
  var FONTS = {
    'SimHei': 'SimHei, "Heiti SC", "PingFang SC", "Microsoft YaHei", "Noto Sans SC", sans-serif',
    'Microsoft YaHei': '"Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", "Noto Sans SC", sans-serif',
    'Arial': 'Arial, Helvetica, "Liberation Sans", sans-serif'
  };
  var font = function (f) { return FONTS[f] || FONTS.SimHei; };
  var hex = function (c) { return /^[0-9A-Fa-f]{6}$/.test(c || '') ? '#' + c : '#000'; };
  var dataUrl = function (d) { return /^data:/.test(d) ? d : 'data:' + d; };

  function paragraphs(text, o) {
    // 把 pptxgenjs 的 run 列表按 breakLine 切成段落
    if (!Array.isArray(text)) return [[{ text: String(text == null ? '' : text), options: {} }]];
    var out = [[]];
    text.forEach(function (r) {
      out[out.length - 1].push(r);
      if (r.options && r.options.breakLine) out.push([]);
    });
    return out;
  }

  function renderText(it) {
    var o = it.o, el = document.createElement('div');
    el.className = 'el tx' + (o.valign === 'middle' ? ' mid' : '') + (/^列表_/.test(o.objectName || '') ? ' bul' : '');
    el.style.fontSize = o.fontSize + 'px';
    el.style.lineHeight = (o.lineSpacing || o.fontSize * 1.2) + 'px';
    el.style.textAlign = o.align === 'right' ? 'right' : 'left';
    paragraphs(it.text, o).forEach(function (runs) {
      var p = document.createElement('p');
      if (!runs.length) p.textContent = '​';
      runs.forEach(function (r) {
        var ro = r.options || {}, s = document.createElement('span');
        s.textContent = r.text;
        s.style.color = hex(ro.color || o.color);
        s.style.fontWeight = (ro.bold != null ? ro.bold : o.bold) ? '700' : '400';
        s.style.fontFamily = font(ro.fontFace || o.fontFace);
        p.appendChild(s);
      });
      el.appendChild(p);
    });
    return el;
  }

  function renderShape(it) {
    var o = it.o, el = document.createElement('div'), f = o.fill || {};
    el.className = 'el';
    var c = hex(f.color), a = f.transparency ? 1 - f.transparency / 100 : 1;
    el.style.background = c;
    el.style.opacity = a;
    if (it.shape === 'roundRect') el.style.borderRadius = Math.min((o.rectRadius || 0) * 72, it.w / 2, it.h / 2) + 'px';
    return el;
  }

  function renderImage(it) {
    if (it.o.hyperlink) return null;   // 链接用的透明遮罩，预览里不需要
    var el = document.createElement('img');
    el.className = 'el';
    el.alt = it.o.objectName === '照片' ? '头像' : '';
    el.src = dataUrl(it.o.data);
    el.style.objectFit = 'cover';
    if (it.o.rounding) el.style.borderRadius = '50%';
    return el;
  }

  window.ResumePreview = {
    render: function (holder, result) {
      var page = document.createElement('div');
      page.className = 'page';
      page.setAttribute('role', 'img');
      page.setAttribute('aria-label', '简历预览');
      result.items.forEach(function (it) {
        var el = it.type === 'text' ? renderText(it) : it.type === 'shape' ? renderShape(it) : renderImage(it);
        if (!el) return;
        el.style.left = it.x + 'px'; el.style.top = it.y + 'px';
        el.style.width = it.w + 'px'; el.style.height = it.h + 'px';
        page.appendChild(el);
      });
      if (result.info.overflowPt > 0) {
        var line = document.createElement('div');
        line.className = 'overflow-line';
        line.style.top = result.info.pageBottom + 'px';
        page.appendChild(line);
      }
      holder.replaceChildren(page);
      var fit = function () {
        var k = holder.clientWidth / result.W;
        page.style.transform = 'scale(' + k + ')';
        holder.style.height = result.H * k + 'px';
      };
      fit();
      if (!holder._ro && window.ResizeObserver) {
        holder._ro = new ResizeObserver(function () { var p = holder.querySelector('.page'); if (p) { var k = holder.clientWidth / result.W; p.style.transform = 'scale(' + k + ')'; holder.style.height = result.H * k + 'px'; } });
        holder._ro.observe(holder);
      }
    }
  };
})();
