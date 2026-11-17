// clean.js — 처방(rules/*.json)을 표에 적용한다. 화면과 무관한 계산만 한다.
// 처방의 내용(어떤 값을 어떻게 바꿀지)은 모두 처방 파일에서 읽는다. 여기에는 «처방을 읽는 방법»만 있다.
(function (root) {
  'use strict';

  var FIX_KINDS = ['정식이름', '바꿈표', '날짜', '글자떼기'];
  var FLAG_KINDS = ['빈칸표시', '큰값표시', '칸관계표시'];

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }
  function asList(x) { return Array.isArray(x) ? x : (x === undefined ? [] : [x]); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function ruleColumns(r) {
    if (r['종류'] === '큰값표시') return Object.keys(r['한도'] || {});
    if (r['종류'] === '칸관계표시') return [r['만약']['칸'], r['그리고']['칸']];
    return asList(r['칸']);
  }

  // ---- 고치기 ------------------------------------------------------------
  function fixFormalName(v, rule) {
    if (typeof v !== 'string') return v;
    var s = v.trim();
    var opt = rule['같은것으로보는차이'] || {};
    var tails = asList(opt['끝글자']);
    function key(x) {
      if (opt['안의공백']) x = x.replace(/\s+/g, '');
      if (opt['끝숫자']) x = x.replace(/\d+$/, '');
      for (var i = 0; i < tails.length; i++) {
        if (x.length > tails[i].length && x.slice(-tails[i].length) === tails[i]) { x = x.slice(0, -tails[i].length); break; }
      }
      return x;
    }
    var names = asList(rule['정식이름']), k = key(s);
    for (var j = 0; j < names.length; j++) if (key(names[j]) === k) return names[j];
    return s;
  }

  function fixTable(v, rule) {
    if (typeof v !== 'string') return v;
    var t = rule['바꿈표'] || {}, s = v.trim();
    return Object.prototype.hasOwnProperty.call(t, s) ? t[s] : v;
  }

  // 날짜: {value, flag} — flag 가 있으면 고치지 않는다
  function fixDate(v, rule) {
    if (isBlank(v)) return { value: v };
    if (typeof v !== 'string') return { value: v, flag: '알 수 없는 날짜 표기' };
    var s = v.trim(), m, y, mo, d;
    if ((m = /^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/.exec(s)) ||
        (m = /^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일$/.exec(s)) ||
        (m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) {
      y = +m[1]; mo = +m[2]; d = +m[3];
      var dt = new Date(Date.UTC(y, mo - 1, d));
      if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d)
        return { value: v, flag: '없는 날짜' };
      return { value: y + '-' + pad(mo) + '-' + pad(d) };
    }
    if (/^\d{1,2}[-\/.]\d{1,2}[-\/.]\d{1,2}$/.test(s)) return { value: v, flag: rule['표시까닭'] || '연도가 두 자리' };
    return { value: v, flag: '알 수 없는 날짜 표기' };
  }

  function fixStrip(v, rule) {
    if (typeof v !== 'string') return v;
    var s = v.trim(), t = rule['바꿈표'] || {};
    if (Object.prototype.hasOwnProperty.call(t, s)) return t[s];
    asList(rule['뗄글자']).forEach(function (tok) { s = s.split(tok).join(''); });
    s = s.trim();
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : v;
  }

  // ---- 한 번 돌리기 --------------------------------------------------------
  // table: {headers:[], rows:[[]], lines:[]}  (lines = 파일의 줄 번호)
  function run(table, rules, markCol) {
    var headers = table.headers.slice();
    var rows = table.rows.map(function (r) { return r.slice(); });
    var lines = table.lines.slice();
    var mi = headers.indexOf(markCol);
    if (mi >= 0) { headers.splice(mi, 1); rows.forEach(function (r) { r.splice(mi, 1); }); }

    var skipped = [], active = [];
    rules.forEach(function (r) {
      var miss = ruleColumns(r).filter(function (c) { return headers.indexOf(c) < 0; });
      if (miss.length) skipped.push({ id: r.id, missing: miss }); else active.push(r);
    });

    var changes = [], flags = rows.map(function () { return []; });
    var ci = function (c) { return headers.indexOf(c); };

    function setCell(i, col, after, rule) {
      var j = ci(col), before = rows[i][j];
      if (before === after) return;
      rows[i][j] = after;
      changes.push({ line: lines[i], col: col, before: before, after: after, rule: rule.id });
    }

    // 1) 고치는 처방 — 처방 파일에 적힌 순서대로
    active.filter(function (r) { return FIX_KINDS.indexOf(r['종류']) >= 0; }).forEach(function (rule) {
      asList(rule['칸']).forEach(function (col) {
        for (var i = 0; i < rows.length; i++) {
          var v = rows[i][ci(col)];
          if (isBlank(v)) continue;
          if (rule['종류'] === '정식이름') setCell(i, col, fixFormalName(v, rule), rule);
          else if (rule['종류'] === '바꿈표') setCell(i, col, fixTable(v, rule), rule);
          else if (rule['종류'] === '글자떼기') setCell(i, col, fixStrip(v, rule), rule);
          else if (rule['종류'] === '날짜') {
            var r = fixDate(v, rule);
            if (r.flag) flags[i].push({ cols: [col], reason: r.flag, rule: rule.id });
            else setCell(i, col, r.value, rule);
          }
        }
      });
    });

    // 2) 표시하는 처방 — 고친 뒤의 값을 본다
    active.filter(function (r) { return FLAG_KINDS.indexOf(r['종류']) >= 0; }).forEach(function (rule) {
      for (var i = 0; i < rows.length; i++) {
        if (rule['종류'] === '빈칸표시') {
          asList(rule['칸']).forEach(function (col) {
            if (isBlank(rows[i][ci(col)])) flags[i].push({ cols: [col], reason: rule['표시까닭'], rule: rule.id });
          });
        } else if (rule['종류'] === '큰값표시') {
          Object.keys(rule['한도']).forEach(function (col) {
            var v = rows[i][ci(col)];
            if (typeof v === 'number' && v > rule['한도'][col])
              flags[i].push({ cols: [col], reason: rule['표시까닭'], rule: rule.id });
          });
        } else if (rule['종류'] === '칸관계표시') {
          var a = rows[i][ci(rule['만약']['칸'])], b = rows[i][ci(rule['그리고']['칸'])];
          if (typeof a === 'number' && typeof b === 'number' &&
              a === rule['만약']['같다'] && b > rule['그리고']['크다'])
            flags[i].push({ cols: [rule['만약']['칸'], rule['그리고']['칸']], reason: rule['표시까닭'], rule: rule.id });
        }
      }
    });

    // 3) 중복 — 마지막에, 모양을 다 맞춘 뒤
    var deleted = [], keep = rows.map(function () { return true; });
    var dupRule = active.filter(function (r) { return r['종류'] === '중복'; })[0];
    if (dupRule) {
      var seen = {};
      for (var i = 0; i < rows.length; i++) {
        var k = JSON.stringify(rows[i]);
        if (Object.prototype.hasOwnProperty.call(seen, k)) {
          keep[i] = false;
          deleted.push({ line: lines[i], sameAs: lines[seen[k]], rule: dupRule.id, values: rows[i].slice() });
        } else seen[k] = i;
      }
    }

    var kept = {};
    lines.forEach(function (l, i) { if (keep[i]) kept[l] = true; });
    var outRows = [], outLines = [], outFlags = [];
    rows.forEach(function (r, i) { if (keep[i]) { outRows.push(r); outLines.push(lines[i]); outFlags.push(flags[i]); } });
    changes = changes.filter(function (c) { return kept[c.line]; });

    var flagged = [];
    outFlags.forEach(function (f, i) {
      f.forEach(function (x) { flagged.push({ line: outLines[i], cols: x.cols, reason: x.reason, rule: x.rule }); });
    });
    var marks = outFlags.map(function (f) {
      return f.map(function (x) { return x.cols.join('·') + ': ' + x.reason; }).join(' / ');
    });
    var flaggedRowCount = outFlags.filter(function (f) { return f.length; }).length;

    var stats = {};
    rules.forEach(function (r) { stats[r.id] = { cells: 0, rows: 0, deleted: 0, skipped: false }; });
    skipped.forEach(function (s) { stats[s.id].skipped = true; });
    var rowSets = {};
    changes.forEach(function (c) { stats[c.rule].cells++; (rowSets[c.rule] = rowSets[c.rule] || {})[c.line] = 1; });
    flagged.forEach(function (c) { stats[c.rule].cells++; (rowSets[c.rule] = rowSets[c.rule] || {})[c.line] = 1; });
    Object.keys(rowSets).forEach(function (id) { stats[id].rows = Object.keys(rowSets[id]).length; });
    deleted.forEach(function (d) { stats[d.rule].deleted++; stats[d.rule].rows++; });

    return {
      headers: headers, outHeaders: headers.concat([markCol]),
      rows: outRows, lines: outLines, marks: marks,
      changes: changes, deleted: deleted, flagged: flagged,
      before: table.rows.length, after: outRows.length,
      flaggedRows: flaggedRowCount, skipped: skipped, stats: stats
    };
  }

  // 정리본을 다시 정리해 본다 — 같아야 한다
  function runTwice(table, rules, markCol) {
    var r1 = run(table, rules, markCol);
    var r2 = run({ headers: r1.headers, rows: r1.rows, lines: r1.lines }, rules, markCol);
    var same = JSON.stringify(r1.rows) === JSON.stringify(r2.rows) &&
               JSON.stringify(r1.marks) === JSON.stringify(r2.marks) &&
               r2.changes.length === 0 && r2.deleted.length === 0;
    return { first: r1, second: r2, same: same };
  }

  // ---- 정답과 대조 ----------------------------------------------------------
  function cellText(v) {
    if (isBlank(v)) return '';
    if (typeof v === 'number') return String(v);
    var s = String(v).trim();
    return /^-?\d+(\.\d+)?$/.test(s) ? String(Number(s)) : s;
  }
  // mine: run() 의 결과(outHeaders, rows, marks, lines) · answer: {headers, rows, lines}
  function compare(mine, answer, markCol) {
    var mh = mine.outHeaders, ah = answer.headers;
    var common = mh.filter(function (h) { return ah.indexOf(h) >= 0; });
    var onlyMine = mh.filter(function (h) { return ah.indexOf(h) < 0; });
    var onlyAns = ah.filter(function (h) { return mh.indexOf(h) < 0; });
    var mRows = mine.rows.map(function (r, i) { return r.concat([mine.marks[i]]); });

    var keyCol = common[0], byKey = false;
    var uniq = function (rows, hs) {
      var j = hs.indexOf(keyCol), s = {};
      for (var i = 0; i < rows.length; i++) { var k = cellText(rows[i][j]); if (k === '' || s[k]) return false; s[k] = 1; }
      return true;
    };
    if (keyCol && uniq(mRows, mh) && uniq(answer.rows, ah)) byKey = true;

    var pairs = [], missing = [], extra = [];
    if (byKey) {
      var am = {}, aj = ah.indexOf(keyCol), mj = mh.indexOf(keyCol);
      answer.rows.forEach(function (r, i) { am[cellText(r[aj])] = i; });
      var used = {};
      mRows.forEach(function (r, i) {
        var k = cellText(r[mj]);
        if (Object.prototype.hasOwnProperty.call(am, k)) { pairs.push([i, am[k]]); used[am[k]] = 1; }
        else extra.push(mine.lines[i]);
      });
      answer.rows.forEach(function (r, i) { if (!used[i]) missing.push(answer.lines[i]); });
    } else {
      var n = Math.min(mRows.length, answer.rows.length);
      for (var i = 0; i < n; i++) pairs.push([i, i]);
      for (var a = n; a < answer.rows.length; a++) missing.push(answer.lines[a]);
      for (var b = n; b < mRows.length; b++) extra.push(mine.lines[b]);
    }

    var diffs = [], rowOk = 0;
    pairs.forEach(function (p) {
      var bad = 0;
      common.forEach(function (h) {
        var mv = cellText(mRows[p[0]][mh.indexOf(h)]), av = cellText(answer.rows[p[1]][ah.indexOf(h)]);
        var differ;
        if (h === markCol) differ = (mv === '') !== (av === '');   // 표시 문구는 보지 않고, 표시했는지만 본다
        else differ = mv !== av;
        if (differ) { bad++; diffs.push({ myLine: mine.lines[p[0]], ansLine: answer.lines[p[1]], col: h, mine: mv, answer: av }); }
      });
      if (!bad) rowOk++;
    });
    return {
      answerRows: answer.rows.length, myRows: mine.rows.length, matched: pairs.length, rowOk: rowOk,
      diffs: diffs, missing: missing, extra: extra, byKey: byKey, keyCol: keyCol,
      common: common, onlyMine: onlyMine, onlyAns: onlyAns
    };
  }

  root.Cleaner = { run: run, runTwice: runTwice, compare: compare, isBlank: isBlank, ruleColumns: ruleColumns };
})(typeof window !== 'undefined' ? window : this);
