// app.js — 화면. 파일 내용은 이 브라우저 안에서만 다룬다 (서버로 보내지 않는다).
// 네트워크로 가져오는 것은 같은 사이트의 rules/<파일 이름>.json 하나뿐이다.
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var state = { name: null, table: null, rules: null, result: null, twice: null };

  function el(tag, attrs, kids) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (k === 'text') e.textContent = attrs[k]; else e.setAttribute(k, attrs[k]);
    });
    (kids || []).forEach(function (c) { e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); });
    return e;
  }
  function clear(n) { while (n.firstChild) n.removeChild(n.firstChild); }
  function show(v) { return v === null || v === undefined || v === '' ? '(빈칸)' : String(v); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  // ---- 엑셀 · CSV 읽기 ----------------------------------------------------
  // 첫 시트만 읽는다. 날짜 칸은 YYYY-MM-DD 글자로, 빈 글자는 빈칸(null)으로.
  function readTable(file, asText) {
    return file.arrayBuffer().then(function (buf) {
      var wb;
      if (/\.csv$/i.test(file.name)) {
        var text = new TextDecoder('utf-8').decode(buf);
        if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
        wb = XLSX.read(text, { type: 'string', raw: !!asText, cellNF: true });
      } else wb = XLSX.read(buf, { type: 'array', cellNF: true });
      var ws = wb.Sheets[wb.SheetNames[0]];
      if (!ws || !ws['!ref']) throw new Error('읽을 내용이 없습니다');
      var rg = XLSX.utils.decode_range(ws['!ref']), grid = [], lines = [];
      for (var r = rg.s.r; r <= rg.e.r; r++) {
        var row = [], any = false;
        for (var c = rg.s.c; c <= rg.e.c; c++) {
          var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })], v = null;
          if (cell) {
            if (cell.t === 'n' && cell.z && XLSX.SSF.is_date(cell.z)) {
              var p = XLSX.SSF.parse_date_code(cell.v);
              v = p.y + '-' + pad(p.m) + '-' + pad(p.d);
            } else if (cell.t === 'd') {
              v = cell.v.getFullYear() + '-' + pad(cell.v.getMonth() + 1) + '-' + pad(cell.v.getDate());
            } else if (cell.t === 'n') v = cell.v;
            else if (cell.t === 's') v = Cleaner.isBlank(cell.v) ? null : cell.v;
            else if (cell.t === 'b') v = cell.v ? 'TRUE' : 'FALSE';
          }
          if (v !== null) any = true;
          row.push(v);
        }
        if (any || r === rg.s.r) { grid.push(row); lines.push(r + 1); }
      }
      var headers = grid[0].map(function (h, i) { return Cleaner.isBlank(h) ? '열' + (i + 1) : String(h).trim(); });
      return { headers: headers, rows: grid.slice(1), lines: lines.slice(1), sheetCount: wb.SheetNames.length, sheetName: wb.SheetNames[0] };
    });
  }

  // ---- 1. 파일 넣기 ---------------------------------------------------------
  function baseName(n) { return n.replace(/\.[^.]+$/, ''); }

  function onFile(file) {
    if (!file) return;
    state = { name: baseName(file.name), table: null, rules: null, result: null, twice: null };
    ['s-result', 's-compare'].forEach(function (i) { $(i).hidden = true; });
    readTable(file, false).then(function (t) {
      state.table = t;
      var info = $('info'); clear(info); info.hidden = false;
      info.appendChild(el('p', {}, [el('strong', { text: file.name }), ' — ' + t.rows.length + '줄 × ' + t.headers.length + '칸 (머리글 제외)']));
      info.appendChild(el('p', { class: 'note', text: '칸: ' + t.headers.join(' · ') }));
      if (t.sheetCount > 1) info.appendChild(el('div', { class: 'warn', text: '시트가 ' + t.sheetCount + '개입니다. 첫 시트(' + t.sheetName + ')만 읽었습니다.' }));
      loadRules();
    }).catch(function (e) {
      var info = $('info'); clear(info); info.hidden = false;
      info.appendChild(el('div', { class: 'warn bad', text: '파일을 읽지 못했습니다: ' + e.message }));
    });
  }

  // ---- 2. 처방 ---------------------------------------------------------------
  function loadRules() {
    $('s-rules').hidden = false;
    clear($('rules-body')); $('rules-pick').hidden = true; $('run').hidden = true;
    fetch('rules/' + encodeURIComponent(state.name) + '.json', { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('없음'); return r.json(); })
      .then(function (j) { acceptRules(j, 'rules/' + state.name + '.json'); })
      .catch(function () { noRules('이 파일에 맞는 처방이 없습니다'); });
  }
  function noRules(msg) {
    state.rules = null; $('run').hidden = true;
    clear($('rules-body'));
    $('rules-body').appendChild(el('div', { class: 'warn bad', text: msg }));
    $('rules-pick').hidden = false;
  }
  function acceptRules(j, from) {
    if (!j || !Array.isArray(j['처방']) || !j['표시칸']) { noRules('이 파일에 맞는 처방이 없습니다 (처방 파일 모양이 아닙니다)'); return; }
    var used = 0, missing = {};
    j['처방'].forEach(function (r) {
      var cols = Cleaner.ruleColumns(r), hit = cols.filter(function (c) { return state.table.headers.indexOf(c) >= 0; });
      used += hit.length;
      cols.forEach(function (c) { if (state.table.headers.indexOf(c) < 0) missing[c] = 1; });
    });
    if (!used) { noRules('이 파일에 맞는 처방이 없습니다 (처방이 가리키는 칸이 이 파일에 없습니다)'); return; }
    state.rules = j;
    var body = $('rules-body'); clear(body);
    body.appendChild(el('p', { class: 'note', text: from + ' — 처방 ' + j['처방'].length + '개' }));
    if (Object.keys(missing).length)
      body.appendChild(el('div', { class: 'warn', text: '이 파일에 없는 칸: ' + Object.keys(missing).join(', ') + ' — 이 칸을 쓰는 처방은 건너뜁니다.' }));
    j['처방'].forEach(function (r) {
      body.appendChild(el('div', { class: 'rule' }, [
        el('strong', { text: r['하는일'] }),
        el('small', { text: '이유: ' + r.why + '  ·  정한 사람: ' + r.owner })
      ]));
    });
    $('run').hidden = false;
  }

  // ---- 3. 정리 ---------------------------------------------------------------
  function table(headers, rows) {
    var t = el('table'), tr = el('tr');
    headers.forEach(function (h) { tr.appendChild(el('th', { text: h })); });
    t.appendChild(el('thead', {}, [tr]));
    var tb = el('tbody');
    rows.forEach(function (r) {
      var row = el('tr');
      r.forEach(function (c) { row.appendChild(el('td', { text: c })); });
      tb.appendChild(row);
    });
    t.appendChild(tb);
    return el('div', { class: 'tablewrap' }, [t]);
  }
  function card(num, label, cls) { return el('div', { class: 'card' }, [el('b', { text: String(num), class: cls || '' }), el('span', { text: label })]); }
  function section(title, node) { var d = el('details'); d.open = true; d.appendChild(el('summary', { text: title })); d.appendChild(node); return d; }

  function runClean() {
    var tw = Cleaner.runTwice(state.table, state.rules['처방'], state.rules['표시칸']);
    state.twice = tw; state.result = tw.first;
    var r = tw.first, ruleById = {};
    state.rules['처방'].forEach(function (p) { ruleById[p.id] = p; });

    $('s-result').hidden = false; $('s-compare').hidden = false;
    var sm = $('summary'); clear(sm);
    sm.appendChild(card(r.before + ' → ' + r.after, '행 (전 → 후)'));
    sm.appendChild(card(r.deleted.length, '지운 행'));
    sm.appendChild(card(r.changes.length, '고친 값'));
    sm.appendChild(card(r.flaggedRows, '표시한 행'));
    sm.appendChild(card(tw.same ? '같음' : '다름', '두 번 돌려도 같은 결과', tw.same ? 'ok' : 'bad'));

    var dl = $('downloads'); clear(dl);
    dl.appendChild(button('정리본 내려받기 (CSV)', function () { download(state.name + '_정리본.csv', cleanedCsv(r)); }));
    dl.appendChild(button('바꾼 기록 내려받기 (CSV)', function () { download(state.name + '_바꾼기록.csv', logCsv(r, ruleById)); }));
    dl.appendChild(el('p', { class: 'note', text: '엑셀에서 한글이 깨지지 않도록 UTF-8(BOM) 형식으로 저장됩니다.' }));

    var d = $('details'); clear(d);
    if (r.skipped.length) d.appendChild(el('div', { class: 'warn', text: '건너뛴 처방: ' + r.skipped.map(function (s) { return s.id + ' (없는 칸: ' + s.missing.join(', ') + ')'; }).join(' / ') }));
    if (!tw.same) d.appendChild(el('div', { class: 'warn bad', text: '정리본을 다시 정리하면 결과가 달라집니다. 두 번째에 바뀐 값 ' + tw.second.changes.length + '개, 지운 행 ' + tw.second.deleted.length + '개.' }));

    d.appendChild(section('처방별 건수', table(['처방', '고친·표시한 칸', '해당 행', '지운 행'],
      state.rules['처방'].map(function (p) {
        var s = r.stats[p.id];
        return [p['하는일'] + (s.skipped ? ' (건너뜀)' : ''), s.cells, s.rows, s.deleted];
      }))));
    d.appendChild(section('고친 값 (' + r.changes.length + ')', table(['줄', '칸', '전', '후', '처방'],
      r.changes.map(function (c) { return [c.line, c.col, show(c.before), show(c.after), c.rule]; }))));
    d.appendChild(section('지운 행 (' + r.deleted.length + ')', table(['줄', '같은 줄(남김)', '처방'].concat(r.headers),
      r.deleted.map(function (x) { return [x.line, x.sameAs, x.rule].concat(x.values.map(show)); }))));
    d.appendChild(section('표시한 행 (' + r.flaggedRows + ')', table(['줄', '칸', '까닭', '처방'],
      r.flagged.map(function (f) { return [f.line, f.cols.join(', '), f.reason, f.rule]; }))));
    d.appendChild(section('정리본 미리보기', table(['줄'].concat(r.outHeaders),
      r.rows.map(function (row, i) { return [r.lines[i]].concat(row.map(show), [r.marks[i]]); }))));
    $('compare-out').textContent = '';
    $('s-result').scrollIntoView({ behavior: 'smooth' });
  }
  function button(text, fn) { var b = el('button', { type: 'button', text: text }); b.addEventListener('click', fn); return b; }

  // ---- 내려받기 ----------------------------------------------------------------
  function csvCell(v) {
    var s = v === null || v === undefined ? '' : String(v);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function toCsv(rows) { return rows.map(function (r) { return r.map(csvCell).join(','); }).join('\r\n') + '\r\n'; }
  function cleanedCsv(r) {
    return toCsv([r.outHeaders].concat(r.rows.map(function (row, i) { return row.concat([r.marks[i]]); })));
  }
  function logCsv(r, ruleById) {
    var why = function (id) { return ruleById[id] ? ruleById[id].why : ''; };
    var out = [['종류', '줄', '칸', '전', '후', '처방', '이유']];
    r.changes.forEach(function (c) { out.push(['고침', c.line, c.col, c.before, c.after, c.rule, why(c.rule)]); });
    r.deleted.forEach(function (d) { out.push(['지움', d.line, '', '같은 줄: ' + d.sameAs, '', d.rule, why(d.rule)]); });
    r.flagged.forEach(function (f) { out.push(['표시', f.line, f.cols.join(', '), '', f.reason, f.rule, why(f.rule)]); });
    return toCsv(out);
  }
  function download(name, text) {
    var blob = new Blob(['﻿' + text], { type: 'text/csv;charset=utf-8' });
    var a = el('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  // ---- 4. 정답과 대조 ------------------------------------------------------------
  function onAnswer(file) {
    if (!file || !state.result) return;
    var out = $('compare-out'); clear(out);
    readTable(file, true).then(function (ans) {
      var c = Cleaner.compare(state.result, ans, state.rules['표시칸']);
      var cards = el('div', { class: 'cards' });
      cards.appendChild(card(c.answerRows, '정답 행'));
      cards.appendChild(card(c.myRows, '내 정리본 행'));
      cards.appendChild(card(c.rowOk + ' / ' + c.matched, '일치한 행 / 짝지은 행', c.rowOk === c.matched && !c.missing.length && !c.extra.length ? 'ok' : 'bad'));
      cards.appendChild(card(c.diffs.length, '어긋난 칸', c.diffs.length ? 'bad' : 'ok'));
      out.appendChild(cards);
      out.appendChild(el('p', { class: 'note', text: '짝짓기: ' + (c.byKey ? '«' + c.keyCol + '» 값이 같은 행끼리' : '위에서부터 순서대로') + ' · 비교한 칸: ' + c.common.join(', ') }));
      if (c.onlyMine.length || c.onlyAns.length) out.appendChild(el('div', { class: 'warn', text: '한쪽에만 있는 칸 — 내 정리본: ' + (c.onlyMine.join(', ') || '없음') + ' / 정답: ' + (c.onlyAns.join(', ') || '없음') }));
      if (c.missing.length) out.appendChild(el('div', { class: 'warn bad', text: '정답에만 있는 행(정답 파일 줄): ' + c.missing.join(', ') }));
      if (c.extra.length) out.appendChild(el('div', { class: 'warn bad', text: '내 정리본에만 있는 행(원본 줄): ' + c.extra.join(', ') }));
      out.appendChild(section('어긋난 칸 (' + c.diffs.length + ')', table(['내 원본 줄', '정답 줄', '칸', '내 값', '정답 값'],
        c.diffs.map(function (d) { return [d.myLine, d.ansLine, d.col, show(d.mine), show(d.answer)]; }))));
    }).catch(function (e) { out.appendChild(el('div', { class: 'warn bad', text: '정답 파일을 읽지 못했습니다: ' + e.message })); });
  }

  // ---- 이벤트 ---------------------------------------------------------------------
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('over'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove('over'); }); });
  drop.addEventListener('drop', function (e) { onFile(e.dataTransfer.files[0]); });
  drop.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
  $('file').addEventListener('change', function (e) { onFile(e.target.files[0]); e.target.value = ''; });
  $('run').addEventListener('click', runClean);
  $('rules-file').addEventListener('change', function (e) {
    var f = e.target.files[0]; if (!f) return;
    f.text().then(function (t) { acceptRules(JSON.parse(t), f.name); }).catch(function () { noRules('이 파일에 맞는 처방이 없습니다 (처방 파일을 읽지 못했습니다)'); });
    e.target.value = '';
  });
  $('answer-file').addEventListener('change', function (e) { onAnswer(e.target.files[0]); e.target.value = ''; });
})();
