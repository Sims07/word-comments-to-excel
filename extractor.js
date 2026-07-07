/*
 * FDL Extractor v1.5
 * Transforme un PDF "Liste des marques de révision et commentaires" (Word) en fichier Excel de Fiche De Lecture.
 * Usage : coller ce script entier dans la console développeur (F12) de n'importe quelle page, puis Entrée.
 * Optimisé pour supporter les relecteurs sans entité/lieu entre parenthèses et les références complexes.
 */
(function () {
  'use strict';

  // Nettoyage si le panneau existe déjà (relance du script)
  var old = document.getElementById('fdl-extractor-panel');
  if (old) old.remove();
  var oldStyle = document.getElementById('fdl-extractor-style');
  if (oldStyle) oldStyle.remove();

  var PDFJS_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  var PDFJS_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  var TESSERACT_URL = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';

  var COLUMNS = [
    { key: 'refFdl', label: 'RefFDL', editable: false, width: '52px' },
    { key: 'relecteur', label: 'Relecteur', editable: true, width: '170px' },
    { key: 'dateOuverture', label: 'Date ouverture remarque', editable: true, width: '112px', type: 'date' },
    { key: 'priorite', label: 'Priorité', editable: true, width: '58px', type: 'datalist', options: ['P0', 'P1', 'P2', 'P3'] },
    { key: 'jira', label: 'JIRA', editable: true, width: '70px' },
    { key: 'boFo', label: 'BO / FO', editable: true, width: '64px', type: 'select', options: ['BO', 'FO', 'NA'] },
    { key: 'version', label: 'Version', editable: true, width: '58px' },
    { key: 'pageParagraphe', label: 'Page, Paragraphe', editable: true, width: '92px' },
    { key: 'contenu', label: 'Contenu - Si applicable', editable: true, width: '100px', type: 'select', options: ['', 'Esthétique', 'Question'] },
    { key: 'remarque', label: 'Remarque', editable: true, width: 'auto', type: 'textarea' }
  ];

  var SUSPECT_RE = /[\u3040-\u30FF\u3400-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF]/;
  
  // REGEX V1.5 : Rend l'auteur, le lieu () et la date optionnels pour éviter les fusions de lignes
  var TITLE_RE = /^Page\s+(\d+)\s*:\s*Comment[ée]e?\s*\[([\s\S]+?)\](?:\s+([\s\S]+?))?(?:\s*\(([^)]+)\))?\s*(\d{1,2}\/\d{1,2}\/\d{2,4}.*|$)/i;
  var DATE_RE = /(\d{1,2}\/\d{1,2}\/\d{2,4})(?:[^0-9]{1,3}(\d{1,2}:\d{2}(?::\d{2})?))?/;
  var PRIORITY_TAG_RE = /#\s*([0-9])\b/;
  var VERSION_RE = /\bV(\d{1,2})\b/i;

  var state = { rows: [], fileName: '', ocrMode: false, filter: '' };

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = src;
      s.onload = resolve;
      s.onerror = function () { reject(new Error('Échec de chargement : ' + src)); };
      document.head.appendChild(s);
    });
  }

  function ensureLibs() {
    var p = Promise.resolve();
    if (!window.pdfjsLib) {
      p = p.then(function () { return loadScript(PDFJS_URL); })
           .then(function () { window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL; });
    }
    if (!window.XLSX) {
      p = p.then(function () { return loadScript(XLSX_URL); });
    }
    return p;
  }

  function ensureTesseract() {
    if (window.Tesseract) return Promise.resolve();
    return loadScript(TESSERACT_URL);
  }

  function extractLinesText(pdfDoc) {
    var allLines = [];
    var pageCount = pdfDoc.numPages;
    var chain = Promise.resolve();
    var _loop = function (p) {
      chain = chain.then(function () { return pdfDoc.getPage(p); })
        .then(function (page) { return page.getTextContent(); })
        .then(function (content) {
          var items = content.items.map(function (it) {
            return { str: it.str, x: it.transform[4], y: it.transform[5] };
          });
          items.sort(function (a, b) { return (b.y - a.y) || (a.x - b.x); });
          var curY = null, curLine = [];
          items.forEach(function (it) {
            if (curY === null || Math.abs(it.y - curY) > 2) {
              if (curLine.length) allLines.push(curLine.map(function (i) { return i.str; }).join(''));
              curLine = [it];
              curY = it.y;
            } else {
              curLine.push(it);
            }
          });
          if (curLine.length) allLines.push(curLine.map(function (i) { return i.str; }).join(''));
          allLines.push('\f');
        });
    };
    for (var p = 1; p <= pageCount; p++) _loop(p);
    return chain.then(function () { return allLines; });
  }

  function renderPageToCanvas(page, scale) {
    var viewport = page.getViewport({ scale: scale || 2.5 });
    var canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    var ctx = canvas.getContext('2d');
    return page.render({ canvasContext: ctx, viewport: viewport }).promise.then(function () { return canvas; });
  }

  function extractLinesOCR(pdfDoc, onProgress) {
    var allLines = [];
    var pageCount = pdfDoc.numPages;
    var worker;
    return ensureTesseract()
      .then(function () { return window.Tesseract.createWorker('fra'); })
      .then(function (w) {
        worker = w;
        var chain = Promise.resolve();
        var _loop = function (p) {
          chain = chain.then(function () {
            if (onProgress) onProgress(p, pageCount);
            return pdfDoc.getPage(p);
          })
            .then(function (page) { return renderPageToCanvas(page, 2.5); })
            .then(function (canvas) { return worker.recognize(canvas); })
            .then(function (result) {
              var text = (result && result.data && result.data.text) || '';
              text.split('\n').forEach(function (line) {
                if (line.trim()) allLines.push(line.trim());
              });
              allLines.push('\f');
            });
        };
        for (var p = 1; p <= pageCount; p++) _loop(p);
        return chain;
      })
      .then(function () { return worker.terminate(); })
      .then(function () { return allLines; });
  }

  function parseDateToken(str) {
    var m = str && str.match(DATE_RE);
    if (!m) return null;
    var parts = m[1].split('/');
    var d = parseInt(parts[0], 10), mo = parseInt(parts[1], 10) - 1, y = parseInt(parts[2], 10);
    if (y < 100) y += 2000;
    var date = new Date(y, mo, d);
    return isNaN(date.getTime()) ? null : date;
  }

  function parseComments(lines, fileName, detectedVersion) {
    var comments = [];
    var current = null;
    lines.forEach(function (raw) {
      var line = (raw || '').trim();
      if (!line || line === '\f') return;
      var m = line.match(TITLE_RE);
      if (m) {
        if (current) comments.push(current);
        
        var auteurBrut = (m[3] || '').trim();
        var horodatage = (m[5] || '').trim();
        
        // Sécurité si l'auteur et la date sont collés sans espace (Ex: Valentin Petit04/06/2025)
        var dateMatch = auteurBrut.match(DATE_RE);
        if (dateMatch && !horodatage) {
          horodatage = auteurBrut.substring(dateMatch.index);
          auteurBrut = auteurBrut.substring(0, dateMatch.index).trim();
        }

        current = {
          page: m[1],
          refInterne: m[2].trim(),
          auteur: auteurBrut || 'Anonyme',
          lieu: (m[4] || '').trim(),
          horodatageBrut: horodatage,
          texteLignes: []
        };
      } else if (current) {
        current.texteLignes.push(line);
      }
    });
    if (current) comments.push(current);

    comments.forEach(function (c) {
      if (!parseDateToken(c.horodatageBrut) && c.texteLignes.length && parseDateToken(c.texteLignes[0])) {
        c.horodatageBrut = c.texteLignes.shift();
      }
    });

    var boFo = 'NA';
    if (/\bBO\b/i.test(fileName)) boFo = 'BO';
    else if (/\bFO\b/i.test(fileName)) boFo = 'FO';

    return comments.map(function (c, idx) {
      var d = parseDateToken(c.horodatageBrut);
      var remarque = c.texteLignes.join(' ').replace(/\s+/g, ' ').trim();
      var priorite = '';
      
      var pm = remarque.match(PRIORITY_TAG_RE);
      if (pm) {
        priorite = 'P' + pm[1];
        remarque = remarque.replace(pm[0], '').replace(/\s{2,}/g, ' ').trim();
      }

      return {
        refFdl: String(idx + 1),
        relecteur: c.auteur,
        dateOuverture: d ? d.toISOString().slice(0, 10) : '',
        priorite: priorite,
        jira: '',
        boFo: boFo,
        version: detectedVersion,
        pageParagraphe: c.page,
        contenu: '',
        remarque: remarque
      };
    });
  }

  function exportXlsx() {
    var header = COLUMNS.map(function (c) { return c.label; });
    var data = [header].concat(state.rows.map(function (r) {
      return COLUMNS.map(function (c) { return r[c.key] || ''; });
    }));
    var ws = window.XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [
      { wch: 8 }, { wch: 20 }, { wch: 14 }, { wch: 8 }, { wch: 12 },
      { wch: 8 }, { wch: 8 }, { wch: 16 }, { wch: 16 }, { wch: 60 }
    ];
    var wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, 'Détail');
    var base = (state.fileName || 'commentaires').replace(/\.pdf$/i, '');
    window.XLSX.writeFile(wb, base + '_FDL.xlsx');
  }

  function injectStyles() {
    var style = document.createElement('style');
    style.id = 'fdl-extractor-style';
    style.textContent = [
      ':root{--fdl-blue:#2A61B1;--fdl-blue-dark:#1c4685;--fdl-green:#00966D;--fdl-green-dark:#007d5a;',
      '--fdl-bg:#F5F8FC;--fdl-border:#D7E1EF;--fdl-text:#1A2B3C;--fdl-warn-bg:#FFF6DA;--fdl-warn-border:#F0C75E;',
      '--fdl-zebra:#F7FAFD;--fdl-hover:#EAF2FC;}',
      '#fdl-extractor-panel{position:fixed;top:16px;right:16px;left:16px;bottom:16px;background:var(--fdl-bg);',
      'border:1px solid var(--fdl-border);border-radius:16px;box-shadow:0 12px 40px rgba(20,50,90,.28);z-index:2147483647;',
      'display:flex;flex-direction:column;font-family:"Segoe UI","Helvetica Neue",Arial,sans-serif;font-size:13px;color:var(--fdl-text);overflow:hidden;}',
      '#fdl-extractor-panel .fdl-header{display:flex;align-items:center;gap:12px;padding:14px 18px;',
      'background:linear-gradient(135deg,var(--fdl-blue),var(--fdl-blue-dark));color:#fff;flex-wrap:wrap;}',
      '#fdl-extractor-panel .fdl-header .fdl-badge{width:26px;height:26px;border-radius:50%;background:var(--fdl-green);',
      'flex:none;display:inline-block;box-shadow:inset 0 0 0 3px rgba(255,255,255,.5);}',
      '#fdl-extractor-panel .fdl-header h2{font-size:15px;margin:0;flex:1;font-weight:600;letter-spacing:.2px;min-width:180px;}',
      '#fdl-extractor-panel .fdl-header label.fdl-ocr-toggle{display:flex;align-items:center;gap:6px;font-size:12px;',
      'background:rgba(255,255,255,.14);padding:5px 10px;border-radius:20px;cursor:pointer;white-space:nowrap;}',
      '#fdl-extractor-panel .fdl-upload-btn{display:flex;align-items:center;gap:7px;background:#fff;',
      'color:var(--fdl-blue-dark);border:none;border-radius:20px;padding:7px 14px;font-size:12px;font-weight:600;',
      'font-family:inherit;cursor:pointer;box-shadow:0 1px 2px rgba(20,50,90,.25);transition:filter .1s,transform .05s;white-space:nowrap;}',
      '#fdl-extractor-panel .fdl-upload-btn:hover{filter:brightness(.96);}',
      '#fdl-extractor-panel .fdl-upload-btn:active{transform:scale(.97);}',
      '#fdl-extractor-panel .fdl-upload-btn svg{color:var(--fdl-green);flex:none;}',
      '#fdl-extractor-panel .fdl-upload-btn.fdl-has-file{background:rgba(255,255,255,.16);color:#fff;}',
      '#fdl-extractor-panel .fdl-upload-btn.fdl-has-file svg{color:#fff;}',
      '#fdl-extractor-panel .fdl-upload-btn span{max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}',
      '#fdl-extractor-panel .fdl-stats{display:flex;gap:16px;align-items:center;padding:7px 18px;background:#fff;',
      'border-bottom:1px solid var(--fdl-border);font-size:12px;color:var(--fdl-blue-dark);}',
      '#fdl-extractor-panel .fdl-stats b{color:var(--fdl-text);}',
      '#fdl-extractor-panel .fdl-stats .fdl-search{margin-left:auto;display:flex;align-items:center;gap:6px;}',
      '#fdl-extractor-panel .fdl-stats .fdl-search input{border:1px solid var(--fdl-border);border-radius:16px;',
      'padding:4px 12px;font-size:12px;width:180px;}',
      '#fdl-extractor-panel .fdl-stats .fdl-chip{background:var(--fdl-bg);border:1px solid var(--fdl-border);',
      'border-radius:12px;padding:2px 9px;}',
      '#fdl-extractor-panel .fdl-stats .fdl-chip.fdl-chip-warn{background:var(--fdl-warn-bg);border-color:var(--fdl-warn-border);color:#6b5300;}',
      '#fdl-extractor-panel .fdl-warning{display:none;background:var(--fdl-warn-bg);border-bottom:1px solid var(--fdl-warn-border);',
      'padding:8px 18px;font-size:12px;color:#6b5300;}',
      '#fdl-extractor-panel .fdl-warning.show{display:block;}',
      '#fdl-extractor-panel .fdl-body{flex:1;overflow:auto;padding:14px 18px;}',
      '#fdl-extractor-panel table{border-collapse:separate;border-spacing:0;width:100%;background:#fff;',
      'table-layout:fixed;border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(20,50,90,.12);}',
      '#fdl-extractor-panel th,#fdl-extractor-panel td{border-bottom:1px solid var(--fdl-border);',
      'border-right:1px solid var(--fdl-border);padding:6px 8px;vertical-align:top;overflow:hidden;}',
      '#fdl-extractor-panel th{background:#EAF0F9;color:var(--fdl-blue-dark);position:sticky;top:0;z-index:2;',
      'text-align:left;font-weight:600;font-size:12px;}',
      '#fdl-extractor-panel th:first-child,#fdl-extractor-panel td:first-child{position:sticky;left:0;z-index:1;',
      'background:#EAF0F9;box-shadow:2px 0 3px rgba(20,50,90,.06);text-align:center;}',
      '#fdl-extractor-panel td:first-child{background:inherit;font-weight:600;color:var(--fdl-blue-dark);}',
      '#fdl-extractor-panel tbody tr:nth-child(even){background:var(--fdl-zebra);}',
      '#fdl-extractor-panel tbody tr:hover{background:var(--fdl-hover);}',
      '#fdl-extractor-panel tbody tr:nth-child(even) td:first-child{background:var(--fdl-zebra);}',
      '#fdl-extractor-panel tbody tr:hover td:first-child{background:var(--fdl-hover);}',
      '#fdl-extractor-panel tbody tr.fdl-page-start td{border-top:2px solid var(--fdl-blue);}',
      '#fdl-extractor-panel tbody tr.fdl-hidden{display:none;}',
      '#fdl-extractor-panel td.fdl-suspect{background:#FFF1E0 !important;box-shadow:inset 0 0 0 1px #E8A33D;}',
      '#fdl-extractor-panel input,#fdl-extractor-panel select,#fdl-extractor-panel textarea{width:100%;',
      'box-sizing:border-box;font-size:12px;font-family:inherit;border:1px solid var(--fdl-border);border-radius:5px;',
      'padding:4px 6px;background:#fff;color:var(--fdl-text);}',
      '#fdl-extractor-panel input:focus,#fdl-extractor-panel select:focus,#fdl-extractor-panel textarea:focus{',
      'outline:none;border-color:var(--fdl-blue);box-shadow:0 0 0 2px rgba(42,97,177,.18);position:relative;z-index:3;}',
      '#fdl-extractor-panel textarea{min-height:44px;resize:vertical;}',
      '#fdl-extractor-panel textarea:focus{min-height:110px;}',
      '#fdl-extractor-panel .fdl-footer{padding:12px 18px;border-top:1px solid var(--fdl-border);display:flex;',
      'gap:10px;align-items:center;background:#fff;}',
      '#fdl-extractor-panel button{cursor:pointer;padding:7px 14px;border-radius:20px;border:1px solid var(--fdl-border);',
      'background:#fff;color:var(--fdl-text);font-weight:600;font-size:12px;transition:filter .1s;}',
      '#fdl-extractor-panel button:hover{filter:brightness(.97);}',
      '#fdl-extractor-panel button.primary{background:var(--fdl-green);color:#fff;border-color:var(--fdl-green-dark);}',
      '#fdl-extractor-panel button.primary:disabled{background:#B9C6D6;border-color:#B9C6D6;cursor:not-allowed;}',
      '#fdl-extractor-panel .fdl-close{margin-left:auto;background:rgba(255,255,255,.16);border:none;color:#fff;',
      'font-size:15px;cursor:pointer;border-radius:50%;width:26px;height:26px;line-height:26px;padding:0;text-align:center;}',
      '#fdl-extractor-panel .fdl-status{color:#fff;opacity:.9;font-size:12px;white-space:nowrap;}',
      '#fdl-extractor-panel .fdl-del{color:#B3261E;background:none;border:none;font-size:15px;cursor:pointer;padding:0;}'
    ].join('');
    document.head.appendChild(style);
  }

  function buildPanel() {
    var panel = document.createElement('div');
    panel.id = 'fdl-extractor-panel';
    panel.innerHTML =
      '<div class="fdl-header">' +
        '<span class="fdl-badge" aria-hidden="true"></span>' +
        '<h2>FDL Extractor — PDF commentaires → Excel</h2>' +
        '<label class="fdl-ocr-toggle"><input type="checkbox" id="fdl-ocr-toggle" /> Mode OCR (police mal encodée)</label>' +
        '<button type="button" class="fdl-upload-btn" id="fdl-upload-btn">' +
          '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">' +
            '<path d="M12 4v11m0-11 4 4m-4-4-4 4M5 17v2a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-2" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>' +
          '</svg>' +
          '<span id="fdl-upload-label">Choisir un PDF</span>' +
        '</button>' +
        '<input type="file" id="fdl-file-input" accept="application/pdf" style="display:none" />' +
        '<span class="fdl-status" id="fdl-status"></span>' +
        '<button class="fdl-close" id="fdl-close" title="Fermer">✕</button>' +
      '</div>' +
      '<div class="fdl-stats" id="fdl-stats"></div>' +
      '<div class="fdl-warning" id="fdl-warning"></div>' +
      '<div class="fdl-body"><div id="fdl-table-wrap"></div></div>' +
      '<div class="fdl-footer">' +
        '<button id="fdl-add-row">+ Ajouter une ligne</button>' +
        '<span style="flex:1"></span>' +
        '<button id="fdl-export" class="primary" disabled>Exporter en Excel</button>' +
      '</div>';
    document.body.appendChild(panel);

    document.getElementById('fdl-close').onclick = function () { panel.remove(); document.getElementById('fdl-extractor-style').remove(); };
    document.getElementById('fdl-file-input').onchange = onFileSelected;
    document.getElementById('fdl-upload-btn').onclick = function () {
      document.getElementById('fdl-file-input').click();
    };
    document.getElementById('fdl-ocr-toggle').onchange = function (e) { state.ocrMode = e.target.checked; };
    document.getElementById('fdl-add-row').onclick = function () {
      state.rows.push(emptyRow());
      renderTable();
    };
    document.getElementById('fdl-export').onclick = exportXlsx;

    return panel;
  }

  function emptyRow() {
    return { refFdl: '', relecteur: '', dateOuverture: '', priorite: '', jira: '', boFo: 'NA', version: '', pageParagraphe: '', contenu: '', remarque: '' };
  }

  function renumber() {
    state.rows.forEach(function (r, i) { r.refFdl = String(i + 1); });
  }

  function isSuspect(value) { return SUSPECT_RE.test(value || ''); }

  function pageOf(row) { return String(row.pageParagraphe || '').split(',')[0].trim(); }

  function renderStats() {
    var total = state.rows.length;
    var pages = state.rows.map(pageOf).filter(Boolean).map(Number).filter(function (n) { return !isNaN(n); });
    var minP = pages.length ? Math.min.apply(null, pages) : null;
    var maxP = pages.length ? Math.max.apply(null, pages) : null;
    var sansPriorite = state.rows.filter(function (r) { return !r.priorite; }).length;
    var suspectCount = state.rows.filter(function (r) { return isSuspect(r.relecteur) || isSuspect(r.remarque); }).length;

    var el = document.getElementById('fdl-stats');
    if (!total) { el.innerHTML = ''; return; }
    el.innerHTML =
      '<span><b>' + total + '</b> commentaire(s)</span>' +
      (minP !== null ? '<span class="fdl-chip">Pages ' + minP + '–' + maxP + '</span>' : '') +
      '<span class="fdl-chip">' + sansPriorite + ' sans priorité</span>' +
      (suspectCount ? '<span class="fdl-chip fdl-chip-warn">' + suspectCount + ' à vérifier (encodage)</span>' : '') +
      '<span class="fdl-search"><input type="text" id="fdl-search-input" placeholder="Filtrer (page, relecteur, texte)…" value="' + (state.filter || '').replace(/"/g, '&quot;') + '" /></span>';

    document.getElementById('fdl-search-input').addEventListener('input', function (e) {
      state.filter = e.target.value;
      applyFilter();
    });
  }

  function applyFilter() {
    var q = (state.filter || '').toLowerCase().trim();
    var rows = document.querySelectorAll('#fdl-table-wrap tbody tr');
    rows.forEach(function (tr) {
      var idx = parseInt(tr.getAttribute('data-idx'), 10);
      var r = state.rows[idx];
      var hay = [r.relecteur, r.remarque, r.pageParagraphe, r.jira, r.priorite].join(' ').toLowerCase();
      tr.classList.toggle('fdl-hidden', q.length > 0 && hay.indexOf(q) === -1);
    });
  }

  function updateWarningBanner() {
    var count = state.rows.filter(function (r) { return isSuspect(r.relecteur) || isSuspect(r.remarque); }).length;
    var el = document.getElementById('fdl-warning');
    if (count > 0) {
      el.textContent = '⚠ ' + count + ' ligne(s) contiennent des caractères probablement mal encodés (police PDF non standard, surlignées en orange). Corrigez-les manuellement, ou cochez "Mode OCR" et relancez l\'extraction sur ce PDF pour contourner le problème.';
      el.classList.add('show');
    } else {
      el.classList.remove('show');
    }
  }

  function renderTable() {
    renumber();
    var wrap = document.getElementById('fdl-table-wrap');
    var html = '<datalist id="fdl-priorite-options">' +
      COLUMNS.filter(function (c) { return c.key === 'priorite'; })[0].options.map(function (o) { return '<option value="' + o + '">'; }).join('') +
      '</datalist>';
    html += '<table><colgroup>' +
      COLUMNS.map(function (c) { return '<col style="width:' + (c.width === 'auto' ? 'auto' : c.width) + '">'; }).join('') +
      '<col style="width:30px"></colgroup>';
    html += '<thead><tr>';
    COLUMNS.forEach(function (c) { html += '<th>' + c.label + '</th>'; });
    html += '<th></th></tr></thead><tbody>';
    var prevPage = null;
    state.rows.forEach(function (r, i) {
      var pg = pageOf(r);
      var pageStart = prevPage !== null && pg && pg !== prevPage;
      prevPage = pg || prevPage;
      html += '<tr data-idx="' + i + '"' + (pageStart ? ' class="fdl-page-start"' : '') + '>';
      COLUMNS.forEach(function (c) {
        var suspect = (c.key === 'relecteur' || c.key === 'remarque') && isSuspect(r[c.key]);
        html += '<td' + (suspect ? ' class="fdl-suspect" title="Caractères suspects : probable problème d\'encodage de police dans le PDF"' : '') + '>' + renderCell(c, r[c.key], i) + '</td>';
      });
      html += '<td><button class="fdl-del" data-idx="' + i + '" title="Supprimer la ligne">🗑</button></td></tr>';
    });
    html += '</tbody></table>';
    wrap.innerHTML = html;

    wrap.querySelectorAll('[data-key]').forEach(function (el) {
      el.addEventListener('input', function () {
        var idx = parseInt(el.closest('tr').getAttribute('data-idx'), 10);
        state.rows[idx][el.getAttribute('data-key')] = el.value;
      });
      el.addEventListener('change', function () {
        var idx = parseInt(el.closest('tr').getAttribute('data-idx'), 10);
        state.rows[idx][el.getAttribute('data-key')] = el.value;
        renderTable();
      });
    });
    wrap.querySelectorAll('.fdl-del').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var idx = parseInt(btn.getAttribute('data-idx'), 10);
        state.rows.splice(idx, 1);
        renderTable();
      });
    });
    document.getElementById('fdl-export').disabled = state.rows.length === 0;
    renderStats();
    updateWarningBanner();
    applyFilter();
  }

  function renderCell(col, value, idx) {
    var v = value == null ? '' : String(value).replace(/"/g, '&quot;');
    if (!col.editable) {
      return '<span>' + v + '</span>';
    }
    if (col.type === 'select') {
      var opts = col.options.map(function (o) {
        return '<option value="' + o + '"' + (o === value ? ' selected' : '') + '>' + (o || '—') + '</option>';
      }).join('');
      return '<select data-key="' + col.key + '">' + opts + '</select>';
    }
    if (col.type === 'datalist') {
      return '<input type="text" list="fdl-priorite-options" data-key="' + col.key + '" value="' + v + '" />';
    }
    if (col.type === 'textarea') {
      return '<textarea data-key="' + col.key + '" title="' + v + '">' + v + '</textarea>';
    }
    var type = col.type === 'date' ? 'date' : 'text';
    return '<input type="' + type + '" data-key="' + col.key + '" value="' + v + '" title="' + v + '" />';
  }

  function setStatus(msg) {
    document.getElementById('fdl-status').textContent = msg;
  }

  function onFileSelected(ev) {
    var file = ev.target.files[0];
    if (!file) return;
    state.fileName = file.name;
    
    var versionMatch = file.name.match(VERSION_RE);
    var detectedVersion = versionMatch ? versionMatch[1] : '';

    var uploadBtn = document.getElementById('fdl-upload-btn');
    document.getElementById('fdl-upload-label').textContent = file.name;
    uploadBtn.classList.add('fdl-has-file');
    setStatus('Chargement des librairies…');
    ensureLibs()
      .then(function () {
        setStatus('Lecture du PDF…');
        return file.arrayBuffer();
      })
      .then(function (buf) { return window.pdfjsLib.getDocument({ data: buf }).promise; })
      .then(function (pdfDoc) {
        if (state.ocrMode) {
          return extractLinesOCR(pdfDoc, function (p, total) {
            setStatus('OCR page ' + p + '/' + total + '…');
          });
        }
        setStatus('Extraction du texte (' + pdfDoc.numPages + ' pages)…');
        return extractLinesText(pdfDoc);
      })
      .then(function (lines) {
        var comments = parseComments(lines, state.fileName, detectedVersion);
        state.rows = comments;
        renderTable();
        setStatus(comments.length + ' commentaire(s) extrait(s)' + (state.ocrMode ? ' (OCR)' : '') + '.');
      })
      .catch(function (err) {
        console.error(err);
        setStatus('Erreur : ' + err.message);
      });
  }

  injectStyles();
  buildPanel();
  renderTable();
})();
