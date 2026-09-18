/* ============================================================================
   Canvas Course Content Exporter — self-contained bookmarklet source

   No server, no libraries, no access token. Runs on the Canvas origin under
   the signed-in user's session and writes a folder tree of .docx and .csv
   files to a folder the user picks (Chrome/Edge), or a single .zip (others).

   This is the readable source. Minify it and prefix with "javascript:" to
   get the bookmarklet; install-exporter.html already contains that version.
   ========================================================================== */
(function () {
  'use strict';

  if (window.__cce) { alert('The exporter is already open in this tab.'); return; }
  window.__cce = true;

  var API = location.origin + '/api/v1';
  var CONCURRENCY = 4;

  /* ======================================================================
     1. CRC-32 and a minimal ZIP writer

     A .docx file IS a zip archive. Rather than ship JSZip we write stored
     (uncompressed) entries by hand: local header, then the bytes, then a
     central directory, then the end-of-central-directory record. Word does
     not care that the entries are uncompressed.
     ====================================================================== */

  var CRC_TABLE = (function () {
    var table = [], c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  var ENC = new TextEncoder();
  function bytes(x) { return typeof x === 'string' ? ENC.encode(x) : x; }

  function Zip() { this.entries = []; }

  Zip.prototype.add = function (name, data) {
    this.entries.push({ name: name, data: bytes(data) });
  };

  Zip.prototype.blob = function () {
    var now = new Date();
    var dosTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xFFFF;
    var dosDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xFFFF;

    var parts = [], central = [], offset = 0;

    for (var i = 0; i < this.entries.length; i++) {
      var e = this.entries[i];
      var nameBytes = ENC.encode(e.name);
      var crc = crc32(e.data);
      var size = e.data.length;

      var local = new Uint8Array(30 + nameBytes.length);
      var lv = new DataView(local.buffer);
      lv.setUint32(0, 0x04034B50, true);   // local file header signature
      lv.setUint16(4, 20, true);           // version needed
      lv.setUint16(6, 0x0800, true);       // flags: names are UTF-8
      lv.setUint16(8, 0, true);            // method 0 = stored
      lv.setUint16(10, dosTime, true);
      lv.setUint16(12, dosDate, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);        // compressed size
      lv.setUint32(22, size, true);        // uncompressed size
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);           // extra field length
      local.set(nameBytes, 30);

      parts.push(local, e.data);

      var cd = new Uint8Array(46 + nameBytes.length);
      var cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014B50, true);   // central directory signature
      cv.setUint16(4, 20, true);           // version made by
      cv.setUint16(6, 20, true);           // version needed
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, dosTime, true);
      cv.setUint16(14, dosDate, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint32(42, offset, true);      // offset of local header
      cd.set(nameBytes, 46);
      central.push(cd);

      offset += local.length + size;
    }

    var cdSize = 0;
    for (var j = 0; j < central.length; j++) cdSize += central[j].length;

    var end = new Uint8Array(22);
    var ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054B50, true);
    ev.setUint16(8, this.entries.length, true);
    ev.setUint16(10, this.entries.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);

    return new Blob(parts.concat(central, [end]), { type: 'application/zip' });
  };

  /* ======================================================================
     2. WordprocessingML

     A minimal .docx needs [Content_Types].xml, _rels/.rels,
     word/document.xml, and word/_rels/document.xml.rels for hyperlinks.
     Headings use direct formatting plus <w:outlineLvl>, which is what
     actually drives Word's navigation pane — so no styles.xml is needed.
     ====================================================================== */

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  }

  // Element order inside <w:rPr> and <w:pPr> is fixed by the schema.
  // Word tolerates a wrong order; other readers do not.
  function run(text, o) {
    o = o || {};
    var rPr = '';
    if (o.bold) rPr += '<w:b/>';
    if (o.italic) rPr += '<w:i/>';
    if (o.color) rPr += '<w:color w:val="' + o.color + '"/>';
    if (o.size) rPr += '<w:sz w:val="' + o.size + '"/>';
    if (o.underline) rPr += '<w:u w:val="single"/>';
    if (o.mono) rPr = '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>' + rPr;
    return '<w:r>' + (rPr ? '<w:rPr>' + rPr + '</w:rPr>' : '') +
      '<w:t xml:space="preserve">' + esc(text) + '</w:t></w:r>';
  }

  function lineBreak() { return '<w:r><w:br/></w:r>'; }

  function para(runsXml, o) {
    o = o || {};
    var pPr = '<w:spacing w:after="' + (o.after == null ? 120 : o.after) + '"' +
      (o.before ? ' w:before="' + o.before + '"' : '') + '/>';
    if (o.indent) pPr += '<w:ind w:left="' + o.indent + '"/>';
    if (o.align) pPr += '<w:jc w:val="' + o.align + '"/>';
    if (o.outline != null) pPr += '<w:outlineLvl w:val="' + o.outline + '"/>';
    return '<w:p><w:pPr>' + pPr + '</w:pPr>' + (runsXml || '') + '</w:p>';
  }

  var HEADING_SIZE = [36, 30, 26, 24, 22, 22]; // half-points

  function heading(text, level, opts) {
    opts = opts || {};
    return para(run(text, { bold: true, size: HEADING_SIZE[level], color: '1F3864' }), {
      outline: level, before: level === 0 ? 0 : 240, after: 120, align: opts.align
    });
  }

  var CELL_BORDERS = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map(function (s) { return '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="999999"/>'; })
    .join('');

  function table(rows) {
    var xml = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>' +
      CELL_BORDERS + '</w:tblBorders></w:tblPr>';
    for (var r = 0; r < rows.length; r++) {
      xml += '<w:tr>';
      for (var c = 0; c < rows[r].length; c++) {
        // every <w:tc> must contain at least one paragraph
        xml += '<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>' +
          (rows[r][c] || para('')) + '</w:tc>';
      }
      xml += '</w:tr>';
    }
    // a table must be followed by a paragraph or Word repairs the file
    return xml + '</w:tbl>' + para('');
  }

  var CONTENT_TYPES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '</Types>';

  var ROOT_RELS =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>';

  // Build the finished .docx bytes from body XML plus any hyperlink targets.
  function docxBytes(bodyXml, links) {
    var docRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      links.map(function (l) {
        return '<Relationship Id="' + l.id +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' +
          esc(l.url) + '" TargetMode="External"/>';
      }).join('') +
      '</Relationships>';

    var document_xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>' +
      bodyXml +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"' +
      ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>';

    var z = new Zip();
    z.add('[Content_Types].xml', CONTENT_TYPES);
    z.add('_rels/.rels', ROOT_RELS);
    z.add('word/document.xml', document_xml);
    z.add('word/_rels/document.xml.rels', docRels);
    return z.blob();
  }

  /* ======================================================================
     3. Canvas HTML -> WordprocessingML
     ====================================================================== */

  function parseHtml(html) {
    return new DOMParser()
      .parseFromString('<div id="r">' + (html || '') + '</div>', 'text/html')
      .querySelector('#r');
  }

  function plainText(html) {
    if (!html) return '';
    return (parseHtml(html).textContent || '').replace(/[ \t]+/g, ' ').trim();
  }

  var BLOCKS = { P: 1, DIV: 1, H1: 1, H2: 1, H3: 1, H4: 1, H5: 1, H6: 1, LI: 1, UL: 1, OL: 1, TABLE: 1, BLOCKQUOTE: 1, PRE: 1, SECTION: 1, ARTICLE: 1, HEADER: 1, FOOTER: 1, FIGURE: 1, HR: 1 };

  function hasBlockChild(el) {
    for (var i = 0; i < el.children.length; i++) if (BLOCKS[el.children[i].tagName]) return true;
    return false;
  }

  // Inline runs for one block, ignoring nested block structure.
  function inlineRuns(node, ctx) {
    var out = '';
    (function walk(n, fmt) {
      for (var i = 0; i < n.childNodes.length; i++) {
        var c = n.childNodes[i];

        if (c.nodeType === 3) {
          var t = c.nodeValue.replace(/\s+/g, ' ');
          if (t.trim()) out += run(t, fmt);
          continue;
        }
        if (c.nodeType !== 1) continue;

        var tag = c.tagName;
        if (tag === 'SCRIPT' || tag === 'STYLE') continue;
        if (tag === 'BR') { out += lineBreak(); continue; }

        if (tag === 'IMG') {
          out += run('[image: ' + (c.getAttribute('alt') || c.getAttribute('src') || '') + ']',
            { italic: true, color: '777777' });
          continue;
        }

        if (tag === 'A' && c.getAttribute('href')) {
          var label = (c.textContent || '').replace(/\s+/g, ' ').trim();
          if (label) {
            var href;
            try { href = new URL(c.getAttribute('href'), location.origin).href; }
            catch (e) { href = c.getAttribute('href'); }
            var id = 'rL' + (ctx.links.length + 1);
            ctx.links.push({ id: id, url: href });
            out += '<w:hyperlink r:id="' + id + '">' +
              run(label, { color: '0563C1', underline: true }) + '</w:hyperlink>';
          }
          continue;
        }

        walk(c, {
          bold: fmt.bold || tag === 'B' || tag === 'STRONG' || tag === 'TH',
          italic: fmt.italic || tag === 'I' || tag === 'EM',
          mono: fmt.mono || tag === 'CODE' || tag === 'PRE' || tag === 'KBD'
        });
      }
    })(node, { bold: false, italic: false, mono: false });
    return out;
  }

  var HEADING_LEVEL = { H1: 1, H2: 1, H3: 2, H4: 3, H5: 4, H6: 5 };

  function cellXml(cell, ctx) {
    var out = '';
    if (hasBlockChild(cell)) out = walkBlocks(cell, ctx, 0, null);
    else { var r = inlineRuns(cell, ctx); if (r) out = para(r, { after: 40 }); }
    return out || para('');
  }

  function tableXml(el, ctx) {
    var trs = el.querySelectorAll('tr');
    if (!trs.length) return '';
    var rows = [];
    for (var i = 0; i < trs.length; i++) {
      var cells = [], kids = trs[i].children;
      for (var j = 0; j < kids.length; j++) {
        if (kids[j].tagName === 'TD' || kids[j].tagName === 'TH') cells.push(cellXml(kids[j], ctx));
      }
      if (cells.length) rows.push(cells);
    }
    return rows.length ? table(rows) : '';
  }

  /* Only emit a paragraph for a block that actually holds text. A container
     that holds other blocks gets recursed into instead — otherwise
     <div><p>Hello</p></div> writes "Hello" twice. */
  function walkBlocks(node, ctx, level, list) {
    var out = '';
    var counter = 1;

    for (var i = 0; i < node.children.length; i++) {
      var el = node.children[i];
      var tag = el.tagName;

      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'IFRAME') continue;

      if (tag === 'HR') { out += para(run('\u2014\u2014\u2014\u2014\u2014\u2014\u2014\u2014')); continue; }

      if (tag === 'TABLE') { out += tableXml(el, ctx); continue; }

      if (tag === 'UL' || tag === 'OL') {
        out += walkBlocks(el, ctx, level, tag === 'OL' ? 'ordered' : 'bullet');
        continue;
      }

      if (tag === 'LI') {
        var clone = el.cloneNode(true);
        var nested = clone.querySelectorAll('ul, ol');
        for (var n = 0; n < nested.length; n++) nested[n].parentNode.removeChild(nested[n]);

        var marker = list === 'ordered' ? (counter++) + '. ' : '\u2022  ';
        var liRuns = inlineRuns(clone, ctx);
        if (liRuns) out += para(run(marker) + liRuns, { indent: 360 + level * 360, after: 60 });

        for (var m = 0; m < el.children.length; m++) {
          var sub = el.children[m];
          if (sub.tagName === 'UL' || sub.tagName === 'OL') {
            out += walkBlocks(sub, ctx, level + 1, sub.tagName === 'OL' ? 'ordered' : 'bullet');
          }
        }
        continue;
      }

      if (HEADING_LEVEL[tag]) {
        var text = (el.textContent || '').replace(/\s+/g, ' ').trim();
        if (text) out += heading(text, HEADING_LEVEL[tag]);
        continue;
      }

      if (hasBlockChild(el)) { out += walkBlocks(el, ctx, level, list); continue; }

      var runs = inlineRuns(el, ctx);
      if (runs) out += para(runs, { indent: tag === 'BLOCKQUOTE' ? 720 : 0 });
    }

    return out;
  }

  function htmlDocx(html, title) {
    var ctx = { links: [] };
    var body = heading(title, 0, { align: 'center' });
    if (html && String(html).trim()) {
      var inner = walkBlocks(parseHtml(html), ctx, 0, null);
      body += inner || para(run(plainText(html)));
    } else {
      body += para(run('[No content or body text provided.]', { italic: true }));
    }
    return docxBytes(body, ctx.links);
  }

  /* ======================================================================
     4. Quiz, bank and rubric documents
     ====================================================================== */

  function weightOf(a) { var w = parseFloat(a.weight); return isNaN(w) ? 0 : w; }

  function questionDocx(questions, title, withAnswers, isBank) {
    var body = heading((isBank ? 'Question Bank: ' : '') + title + (withAnswers ? ' (Answer Key)' : ''),
      0, { align: 'center' });

    if (!withAnswers) {
      body += para(run('Name: _________________________\t\tDate: _____________'), { after: 360 });
    }
    body += heading('Questions', 1);

    questions.forEach(function (q, idx) {
      var qText = plainText(q.question_text || '');
      var answers = Array.isArray(q.answers) ? q.answers : [];
      if (!qText && !answers.length) return;

      var pts = (q.points_possible == null) ? 0 : q.points_possible;
      body += heading((idx + 1) + '. ' + (q.question_name || 'Question ' + (idx + 1)) + ' (' + pts + ' pts)', 2);
      body += para(run(qText), { after: 160 });

      var type = q.question_type || '';

      if (type === 'multiple_choice_question' || type === 'multiple_answers_question' || type === 'true_false_question') {
        answers.forEach(function (a) {
          var text = plainText(a.text || '') || plainText(a.html || '');
          if (withAnswers && weightOf(a) > 0) {
            body += para(run('[X] ', { bold: true }) + run(text + '  <-- CORRECT ANSWER', { bold: true }),
              { indent: 360, after: 80 });
          } else {
            body += para(run('[  ] ' + text), { indent: 360, after: 80 });
          }
        });

      } else if (type === 'essay_question') {
        body += withAnswers
          ? para(run('[ Essay question: student responses will vary. ]', { italic: true }), { after: 240 })
          : para(lineBreak() + lineBreak() + lineBreak() + lineBreak() + lineBreak(), { after: 240 });

      } else if (type === 'short_answer_question' || type === 'numerical_question') {
        if (withAnswers) {
          var correct = answers
            .filter(function (a) { return weightOf(a) > 0 || a.text != null; })
            .map(function (a) { return plainText(String(a.text == null ? '' : a.text)); })
            .filter(Boolean);
          body += para(run('Correct answer(s): ' + correct.join(', '), { bold: true }), { after: 240 });
        } else {
          body += para(run('Answer: _____________________________________'), { after: 240 });
        }

      } else if (type === 'matching_question') {
        answers.forEach(function (a) {
          var left = plainText(a.text || '');
          body += para(run(withAnswers
            ? '_____  ' + left + '  ==> Correct match: ' + plainText(a.match_text || '')
            : '_____  ' + left), { indent: 360, after: 80 });
        });
        var matches = Array.isArray(q.matches) ? q.matches : [];
        if (!withAnswers && matches.length) {
          body += para(run('Possible choices:', { bold: true }), { before: 160 });
          var choices = matches.map(function (m) { return plainText(m.text || ''); }).filter(Boolean).sort();
          choices.filter(function (c, i, arr) { return arr.indexOf(c) === i; }).forEach(function (c) {
            body += para(run('- ' + c), { indent: 720, after: 40 });
          });
        }

      } else {
        if (withAnswers && answers.length) {
          var txt = answers.filter(function (a) { return weightOf(a) > 0; })
            .map(function (a) { return '[' + plainText(a.text || '') + ']'; }).join(' ');
          body += para(run('Answers: ' + txt), { after: 240 });
        } else {
          body += para(run('_____________________________________'), { after: 240 });
        }
      }
    });

    return docxBytes(body, []);
  }

  function rubricDocx(rubric) {
    var body = heading(rubric.title || 'Unnamed Rubric', 0, { align: 'center' });
    var criteria = Array.isArray(rubric.data) ? rubric.data : [];

    if (!criteria.length) {
      body += para(run('[This rubric contains no criteria rows.]', { italic: true }));
      return docxBytes(body, []);
    }

    var rows = [[
      para(run('Criterion', { bold: true })),
      para(run('Ratings', { bold: true })),
      para(run('Pts', { bold: true }))
    ]];

    criteria.forEach(function (item) {
      var left = para(run(item.description || 'No description', { bold: true }), { after: 40 });
      if (item.long_description) left += para(run(plainText(item.long_description)), { after: 40 });

      var mid = '';
      (item.ratings || []).forEach(function (r) {
        var s = '\u2022 ' + (r.points == null ? 0 : r.points) + ' pts - ' + (r.description || 'No rating name');
        if (r.long_description) s += ' (' + plainText(r.long_description) + ')';
        mid += para(run(s), { after: 40 });
      });

      rows.push([left, mid || para(''), para(run((item.points == null ? 0 : item.points) + ' pts'))]);
    });

    return docxBytes(body + table(rows), []);
  }

  /* ======================================================================
     5. CSV
     ====================================================================== */

  function csvCell(v) {
    var s = (v == null) ? '' : String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function csvRow(arr) { return arr.map(csvCell).join(',') + '\r\n'; }

  function rubricCsv(rubric) {
    var criteria = Array.isArray(rubric.data) ? rubric.data : [];
    if (!criteria.length) return null;

    var max = 0;
    criteria.forEach(function (c) { max = Math.max(max, (c.ratings || []).length); });

    var header = ['Rubric Name', 'Criteria Name', 'Criteria Description', 'Criteria Enable Range'];
    for (var i = 0; i < max; i++) header.push('Rating Name', 'Rating Description', 'Rating Points');

    var out = '\uFEFF' + csvRow(header);
    criteria.forEach(function (item) {
      var row = [rubric.title || 'Unnamed Rubric', item.description || '',
        plainText(item.long_description || ''), String(!!item.criterion_use_range)];
      var ratings = item.ratings || [];
      for (var j = 0; j < max; j++) {
        var r = ratings[j];
        if (r) row.push(r.description || '', plainText(r.long_description || ''), r.points == null ? 0 : r.points);
        else row.push('', '', '');
      }
      out += csvRow(row);
    });
    return out;
  }

  /* ======================================================================
     6. Canvas API
     ====================================================================== */

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function nextLink(header) {
    if (!header) return null;
    var parts = header.split(',');
    for (var i = 0; i < parts.length; i++) {
      var m = parts[i].match(/<([^>]+)>\s*;\s*rel="next"/);
      if (m) return m[1];
    }
    return null;
  }

  async function apiFetch(url) {
    for (var attempt = 0; attempt < 5; attempt++) {
      var res = await fetch(url, {
        credentials: 'same-origin',
        headers: {
          'Accept': 'application/json+canvas-string-ids, application/json',
          'X-Requested-With': 'XMLHttpRequest'
        }
      });

      if (res.status === 403) {
        var t = await res.text();
        if (/rate limit/i.test(t)) { await sleep(2000 * (attempt + 1)); continue; }
        throw new Error('403 Forbidden — ' + url);
      }
      if (res.status === 401) throw new Error('Your Canvas session expired. Reload the page and sign in again.');
      if (!res.ok) throw new Error(res.status + ' — ' + url);

      if (parseFloat(res.headers.get('X-Rate-Limit-Remaining') || '999') < 150) await sleep(1500);

      var text = await res.text();
      // Canvas prefixes session-authenticated JSON with while(1); as anti-hijacking
      return {
        json: JSON.parse(text.replace(/^\s*while\s*\(1\)\s*;?\s*/, '')),
        next: nextLink(res.headers.get('Link'))
      };
    }
    throw new Error('Rate limited repeatedly on ' + url);
  }

  async function getAll(path) {
    var url = /^https?:/.test(path) ? path
      : API + path + (/[?&]per_page=/.test(path) ? '' : (path.indexOf('?') === -1 ? '?' : '&') + 'per_page=100');
    var out = [];
    while (url) {
      var r = await apiFetch(url);
      out = out.concat(Array.isArray(r.json) ? r.json : [r.json]);
      url = r.next;
    }
    return out;
  }

  async function getOne(path) {
    return (await apiFetch(/^https?:/.test(path) ? path : API + path)).json;
  }

  async function pool(items, worker, limit) {
    var i = 0, runners = [];
    for (var k = 0; k < Math.min(limit, items.length); k++) {
      runners.push((async function () {
        while (i < items.length) { var idx = i++; await worker(items[idx], idx); }
      })());
    }
    await Promise.all(runners);
  }

  /* ======================================================================
     7. Output target: real folder tree, or a zip if the browser lacks the
        File System Access API.
     ====================================================================== */

  function sanitize(name) {
    var s = String(name == null ? '' : name)
      .replace(/[\\/*?:"<>|\u0000-\u001f]/g, '').replace(/\s+/g, ' ').trim().slice(0, 80);
    return s || 'Untitled';
  }

  var usedPaths = Object.create(null);
  function uniquePath(dir, base, ext) {
    var path = dir + '/' + base + ext, n = 2;
    while (usedPaths[path.toLowerCase()]) { path = dir + '/' + base + ' (' + n + ')' + ext; n++; }
    usedPaths[path.toLowerCase()] = true;
    return path;
  }

  function FolderTarget(rootHandle) {
    this.root = rootHandle;
    this.dirs = Object.create(null);
  }
  FolderTarget.prototype.dir = async function (path) {
    if (!path) return this.root;
    if (this.dirs[path]) return this.dirs[path];
    var parts = path.split('/'), h = this.root;
    for (var i = 0; i < parts.length; i++) h = await h.getDirectoryHandle(parts[i], { create: true });
    this.dirs[path] = h;
    return h;
  };
  FolderTarget.prototype.write = async function (path, data) {
    var parts = path.split('/');
    var name = parts.pop();
    var dir = await this.dir(parts.join('/'));
    var fh = await dir.getFileHandle(name, { create: true });
    var w = await fh.createWritable();
    await w.write(data);
    await w.close();
  };
  FolderTarget.prototype.finish = async function () { };

  function ZipTarget(name) { this.zip = new Zip(); this.name = name; this.pending = []; }
  ZipTarget.prototype.write = async function (path, data) {
    var buf = (data instanceof Blob) ? new Uint8Array(await data.arrayBuffer()) : bytes(data);
    this.zip.add(path, buf);
  };
  ZipTarget.prototype.finish = async function () {
    var url = URL.createObjectURL(this.zip.blob());
    var a = document.createElement('a');
    a.href = url; a.download = this.name + '.zip';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 600000);
  };

  /* ======================================================================
     8. Interface
     ====================================================================== */

  var host = document.createElement('div');
  var shadow = host.attachShadow({ mode: 'open' });
  shadow.innerHTML =
    '<style>' +
    ':host{all:initial}' +
    '.s{position:fixed;inset:0;background:rgba(15,22,33,.55);z-index:2147483647;display:flex;align-items:center;justify-content:center;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}' +
    '.p{background:#fff;color:#16202c;width:540px;max-width:94vw;max-height:90vh;overflow:auto;border-radius:6px;padding:24px 26px;box-shadow:0 18px 48px rgba(0,0,0,.35)}' +
    'h1{font-size:19px;margin:0 0 6px}' +
    '.d{font-size:13px;color:#5b6774;margin:0 0 18px;line-height:1.5}' +
    'label.f{display:block;font-size:13px;font-weight:700;margin-bottom:6px}' +
    'input[type=text]{width:100%;box-sizing:border-box;padding:9px 10px;font-size:15px;border:1px solid #a9b0b8;border-radius:4px}' +
    '.h{font-size:12px;color:#5b6774;margin-top:5px;line-height:1.4}' +
    'fieldset{border:1px solid #e2e5e9;border-radius:4px;margin:18px 0 0;padding:12px 14px}' +
    'legend{font-size:12px;font-weight:700;color:#5b6774;padding:0 6px}' +
    '.c{display:grid;grid-template-columns:1fr 1fr;gap:6px 14px}' +
    '.c label{font-size:13px;display:flex;align-items:center;gap:7px}' +
    '.w{background:#fdf5e6;border:1px solid #e8c26a;border-radius:4px;padding:10px 12px;font-size:12px;line-height:1.5;margin-top:12px;color:#6b4d09}' +
    '.a{display:flex;gap:10px;margin-top:20px}' +
    'button{font:inherit;font-size:15px;font-weight:700;padding:11px 18px;border-radius:4px;border:none;cursor:pointer}' +
    '#go{background:#0a5c8a;color:#fff;flex:1}#go:disabled{background:#8a959f;cursor:not-allowed}' +
    '#x{background:#fff;border:1px solid #a9b0b8;color:#16202c}' +
    'button:focus-visible{outline:3px solid #1f7ab8;outline-offset:2px}' +
    '#log{display:none;margin-top:18px;background:#10161e;color:#a9e6b0;border-radius:4px;padding:12px 14px;height:200px;overflow-y:auto;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12.5px;line-height:1.6;white-space:pre-wrap}' +
    '#log .e{color:#ff9c8a}#log .k{color:#8fd4ff}' +
    '</style>' +
    '<div class="s"><div class="p" role="dialog" aria-modal="true">' +
    '<h1>Export course content</h1>' +
    '<p class="d">Saves this course as Word documents and spreadsheets. You will be asked to choose a folder — pick your Desktop and the export writes a new folder there. Keep this tab open until it finishes.</p>' +
    '<label class="f" for="cid">Course ID</label><input type="text" id="cid" inputmode="numeric" autocomplete="off">' +
    '<p class="h">Filled in from the page you are on. Change it to export a different course you can access.</p>' +
    '<fieldset><legend>Include</legend><div class="c">' +
    '<label><input type="checkbox" data-k="assignments" checked> Assignments</label>' +
    '<label><input type="checkbox" data-k="quizzes" checked> Quizzes</label>' +
    '<label><input type="checkbox" data-k="banks" checked> Question banks</label>' +
    '<label><input type="checkbox" data-k="rubrics" checked> Rubrics</label>' +
    '<label><input type="checkbox" data-k="discussions" checked> Discussions</label>' +
    '<label><input type="checkbox" data-k="announcements" checked> Announcements</label>' +
    '<label><input type="checkbox" data-k="pages" checked> Pages</label>' +
    '<label><input type="checkbox" data-k="syllabus" checked> Syllabus</label>' +
    '<label><input type="checkbox" data-k="gradebook"> Gradebook</label>' +
    '</div>' +
    '<div class="w">Gradebook creates a spreadsheet of named students and their scores. Leave it unchecked unless you need it, and keep the file wherever your institution requires grades to be kept.</div>' +
    '</fieldset>' +
    '<div class="a"><button id="go">Choose folder and export</button><button id="x">Close</button></div>' +
    '<div id="log" aria-live="polite"></div>' +
    '</div></div>';
  document.body.appendChild(host);

  var $ = function (s) { return shadow.querySelector(s); };
  var logEl = $('#log'), goBtn = $('#go'), cidInput = $('#cid');

  var m = location.pathname.match(/\/courses\/(\d+)/);
  cidInput.value = m ? m[1] : '';

  function log(msg, cls) {
    logEl.style.display = 'block';
    var d = document.createElement('div');
    if (cls) d.className = cls;
    d.textContent = msg;
    logEl.appendChild(d);
    logEl.scrollTop = logEl.scrollHeight;
  }

  var running = false;
  $('#x').addEventListener('click', function () {
    if (running && !confirm('The export is still running. Close anyway?')) return;
    window.__cce = false;
    host.remove();
  });
  window.addEventListener('beforeunload', function (e) {
    if (running) { e.preventDefault(); e.returnValue = ''; }
  });

  goBtn.addEventListener('click', async function () {
    var courseId = cidInput.value.trim();
    if (!/^\d+$/.test(courseId)) { alert('Enter a numeric course ID.'); return; }

    var want = {};
    shadow.querySelectorAll('input[type=checkbox]').forEach(function (cb) { want[cb.dataset.k] = cb.checked; });

    // The folder picker must be the first thing this click does. Any await
    // before it spends the user gesture and the call is rejected.
    var dirHandle = null;
    if (window.showDirectoryPicker) {
      try {
        dirHandle = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'desktop' });
      } catch (e) {
        return; // user cancelled the picker
      }
    }

    goBtn.disabled = true;
    goBtn.textContent = 'Exporting…';
    running = true;

    try {
      await runExport(courseId, want, dirHandle);
    } catch (err) {
      log('Stopped: ' + err.message, 'e');
    }
    running = false;
    goBtn.disabled = false;
    goBtn.textContent = 'Choose folder and export';
  });

  /* ======================================================================
     9. The export
     ====================================================================== */

  async function runExport(courseId, want, dirHandle) {
    log('Connecting to ' + location.host + '…');
    var course = await getOne('/courses/' + courseId + '?include[]=syllabus_body');
    log('Course: ' + course.name, 'k');

    var folderName = sanitize('Canvas_Backup_' + course.name + '_' + courseId);
    var target;
    if (dirHandle) {
      var root = await dirHandle.getDirectoryHandle(folderName, { create: true });
      target = new FolderTarget(root);
      log('Writing to ' + folderName + '/');
    } else {
      target = new ZipTarget(folderName);
      log('This browser cannot write folders directly, so the export will arrive as one zip file.');
    }

    var assignments = [];

    if (want.assignments || want.gradebook || want.quizzes) {
      assignments = await getAll('/courses/' + courseId + '/assignments');
    }

    if (want.assignments) {
      log('Assignments: ' + assignments.length);
      await pool(assignments, async function (a) {
        await target.write(uniquePath('Assignments', sanitize(a.name), '.docx'), htmlDocx(a.description, a.name));
      }, CONCURRENCY);
    }

    if (want.quizzes) {
      try {
        var quizzes = await getAll('/courses/' + courseId + '/quizzes');
        log('Quizzes: ' + quizzes.length);
        await pool(quizzes, async function (q) {
          var questions;
          try { questions = await getAll('/courses/' + courseId + '/quizzes/' + q.id + '/questions'); }
          catch (e) { log('  Skipped "' + q.title + '": ' + e.message, 'e'); return; }
          var base = sanitize(q.title);
          await target.write(uniquePath('Quizzes', base + '_Printable_Quiz', '.docx'), questionDocx(questions, q.title, false, false));
          await target.write(uniquePath('Quizzes', base + '_Answer_Key', '.docx'), questionDocx(questions, q.title, true, false));
        }, 2);

        var newQ = assignments.filter(function (a) { return a.is_quiz_lti_assignment === true; });
        if (newQ.length) {
          log('  ' + newQ.length + ' New Quizzes assessment(s) cannot be exported. See the note in Quizzes/.', 'e');
          await target.write('Quizzes/READ_ME_New_Quizzes.txt',
            'These assessments are built in New Quizzes. Canvas does not expose their questions to this export:\r\n\r\n' +
            newQ.map(function (a) { return '- ' + a.name; }).join('\r\n') +
            '\r\n\r\nOpen each one in Canvas and use its own export or print option to keep a copy.\r\n');
        }
      } catch (e) { log('  Quizzes failed: ' + e.message, 'e'); }
    }

    if (want.banks) {
      try {
        var banks;
        try { banks = await getAll('/courses/' + courseId + '/question_banks'); }
        catch (e) { banks = await getAll('/question_banks?context_type=Course&context_id=' + courseId); }
        log('Question banks: ' + banks.length);
        await pool(banks, async function (b) {
          var qs = await getAll('/question_banks/' + b.id + '/questions');
          if (!qs.length) return;
          var title = b.title || 'Unnamed Bank', base = sanitize(title);
          await target.write(uniquePath('Question_Banks', base + '_Printable_Bank', '.docx'), questionDocx(qs, title, false, true));
          await target.write(uniquePath('Question_Banks', base + '_Answer_Key', '.docx'), questionDocx(qs, title, true, true));
        }, 2);
      } catch (e) { log('  Question banks failed: ' + e.message, 'e'); }
    }

    if (want.rubrics) {
      try {
        var rubrics = await getAll('/courses/' + courseId + '/rubrics');
        log('Rubrics: ' + rubrics.length);
        await pool(rubrics, async function (r) {
          var full = await getOne('/courses/' + courseId + '/rubrics/' + r.id);
          var base = sanitize(full.title || 'Rubric_' + r.id);
          await target.write(uniquePath('Rubrics', base, '.docx'), rubricDocx(full));
          var csv = rubricCsv(full);
          if (csv) await target.write(uniquePath('Rubrics', base + '_Importable', '.csv'), csv);
        }, CONCURRENCY);
      } catch (e) { log('  Rubrics failed: ' + e.message, 'e'); }
    }

    if (want.discussions) {
      try {
        var topics = (await getAll('/courses/' + courseId + '/discussion_topics'))
          .filter(function (t) { return !t.is_announcement; });
        log('Discussions: ' + topics.length);
        await pool(topics, async function (t) {
          await target.write(uniquePath('Discussions', sanitize(t.title), '.docx'), htmlDocx(t.message, t.title));
        }, CONCURRENCY);
      } catch (e) { log('  Discussions failed: ' + e.message, 'e'); }
    }

    if (want.announcements) {
      try {
        var anns = await getAll('/courses/' + courseId + '/discussion_topics?only_announcements=true');
        log('Announcements: ' + anns.length);
        await pool(anns, async function (a) {
          await target.write(uniquePath('Announcements', sanitize(a.title), '.docx'), htmlDocx(a.message, a.title));
        }, CONCURRENCY);
      } catch (e) { log('  Announcements failed: ' + e.message, 'e'); }
    }

    if (want.pages) {
      try {
        var pages = await getAll('/courses/' + courseId + '/pages');
        log('Pages: ' + pages.length);
        await pool(pages, async function (p) {
          var full = await getOne('/courses/' + courseId + '/pages/' + encodeURIComponent(p.url));
          await target.write(uniquePath('Pages', sanitize(full.title), '.docx'), htmlDocx(full.body, full.title));
        }, CONCURRENCY);
      } catch (e) { log('  Pages failed: ' + e.message, 'e'); }
    }

    if (want.syllabus) {
      log('Syllabus');
      await target.write('Syllabus/Syllabus.docx', htmlDocx(course.syllabus_body, course.name + ' Syllabus'));
    }

    if (want.gradebook) {
      try {
        log('Gradebook');
        var students = await getAll('/courses/' + courseId + '/users?enrollment_type[]=student&enrollment_state[]=active');
        var nameById = {};
        students.forEach(function (s) { nameById[String(s.id)] = s.sortable_name || s.name || ''; });

        var header = ['Student', 'Canvas User ID'], colById = {};
        assignments.forEach(function (a) {
          var col = a.name + ' (Max: ' + (a.points_possible == null ? 0 : a.points_possible) + ')';
          colById[String(a.id)] = col;
          header.push(col);
        });

        // One bulk call set, not one call per student.
        var grouped = await getAll('/courses/' + courseId + '/students/submissions?student_ids[]=all&grouped=true&per_page=50');

        var csv = '\uFEFF' + csvRow(header);
        grouped.forEach(function (g) {
          var uid = String(g.user_id), byCol = {};
          (g.submissions || []).forEach(function (s) {
            var col = colById[String(s.assignment_id)];
            if (col) byCol[col] = (s.score == null) ? '' : s.score;
          });
          var row = [nameById[uid] || '(not an active student)', uid];
          for (var c = 2; c < header.length; c++) row.push(byCol[header[c]] == null ? '' : byCol[header[c]]);
          csv += csvRow(row);
        });
        await target.write('Gradebook/Gradebook_' + courseId + '.csv', csv);
      } catch (e) { log('  Gradebook failed: ' + e.message, 'e'); }
    }

    await target.finish();
    log('Done.', 'k');
  }
})();
