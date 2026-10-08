(function(){
  'use strict';
  var $ = function(id){ return document.getElementById(id); };
  var rows = [];
  var headers = [];
  var data = [];               // {square, address, cell, qty}
  var layout = {};             // { square: [ {cell, level, square, address, qty, origCell, moved} x N ] }
  var levelMap = { 1: [], 2: [] };
  var sortKey = 'cell', sortDir = 1;
  var viewMode = 'was';
  var paintMode = 0;
  var totalCells = 54;
  var cellRefs = [];
  var LEVEL_LOAD = { high: 'Ходовой', mid: 'Средний', low: 'Низкий' };

  /* ---------- XLSX ---------- */
  var xlsxReady = null;
  function loadXLSX(){
    if (window.XLSX) return Promise.resolve();
    if (xlsxReady) return xlsxReady;
    var urls = [
      'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
      'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
    ];
    xlsxReady = new Promise(function(resolve, reject){
      (function next(i){
        if (i >= urls.length) { xlsxReady = null; return reject(new Error('Не удалось загрузить XLSX.')); }
        var s = document.createElement('script');
        s.src = urls[i];
        s.onload = function(){ window.XLSX ? resolve() : next(i + 1); };
        s.onerror = function(){ next(i + 1); };
        document.head.appendChild(s);
      })(0);
    });
    return xlsxReady;
  }
  loadXLSX().catch(function(){});

  /* ---------- Утилиты ---------- */
  function msg(text, type){ var m = $('msg'); m.textContent = text || ''; m.className = text ? type : ''; }
  function esc(s){ return String(s).replace(/[&<>"']/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function cellsNeeded(maxCell){
    var groups = Math.max(6, Math.ceil((maxCell || 54) / 9));
    return groups * 9;
  }

  /* ---------- Сетка ----------
     Нумерация СВЕРХУ ВНИЗ: 1 4 7 / 2 5 8 / 3 6 9 в первой группе.
  */
  function buildGrid(maxCell){
    var total = cellsNeeded(maxCell);
    totalCells = total;
    var groups = total / 9;

    var grid = $('grid');
    grid.innerHTML = '';
    cellRefs = [];

    for (var g = 0; g < groups; g++) {
      var group = document.createElement('div');
      group.className = 'group';

      for (var row = 0; row < 3; row++) {
        for (var col = 0; col < 3; col++) {
          var num = g * 9 + col * 3 + row + 1;
          var c = document.createElement('div');
          c.className = 'cell empty';
          c.dataset.num = num;
          c.innerHTML = '<span class="n">' + num + '</span><div class="q">0</div><div class="a">Пусто</div><span class="mv"></span>';
          c.addEventListener('click', onCellClick);
          group.appendChild(c);
          cellRefs[num - 1] = c;
        }
      }
      grid.appendChild(group);
    }
  }

  function onCellClick(e){
    if (!paintMode) return;
    var el = e.currentTarget;
    var num = +el.dataset.num;
    var cur = (levelMap[1].indexOf(num) !== -1) ? 1 : (levelMap[2].indexOf(num) !== -1 ? 2 : 0);

    if (cur === paintMode) {
      levelMap[1] = levelMap[1].filter(function(n){ return n !== num; });
      levelMap[2] = levelMap[2].filter(function(n){ return n !== num; });
    } else {
      if (paintMode === 1) levelMap[2] = levelMap[2].filter(function(n){ return n !== num; });
      if (paintMode === 2) levelMap[1] = levelMap[1].filter(function(n){ return n !== num; });
      if (levelMap[paintMode].indexOf(num) === -1) levelMap[paintMode].push(num);
      levelMap[paintMode].sort(function(a, b){ return a - b; });
    }
    rebuildLayout();
    render();
  }

  /* ---------- Загрузка ---------- */
  function handleFile(file){
    if (!file) return;
    msg('Читаю файл…', 'ok');
    $('btnExportOne').disabled = true;
    $('btnExportAll').disabled = true;
    $('btnRecalc').disabled = true;
    loadXLSX().then(function(){
      return new Promise(function(res, rej){
        var r = new FileReader();
        r.onload = function(e){ res(e.target.result); };
        r.onerror = function(){ rej(new Error('Не удалось прочитать файл')); };
        r.readAsArrayBuffer(file);
      });
    }).then(function(buf){
      var wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      var ws = wb.Sheets[wb.SheetNames[0]];
      var json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
      json = json.filter(function(r){ return r.some(function(v){ return String(v).trim() !== ''; }); });
      if (json.length < 2) throw new Error('В файле нет данных.');
      rows = json;
      headers = json[0].map(function(h){ return String(h).trim(); });
      fillMapping(detect(headers));
      applyMapping();
    }).catch(function(err){ msg(err.message || String(err), 'err'); });
  }

  function detect(h){
    var low = h.map(function(x){ return x.toLowerCase(); });
    function find(words){
      for (var i = 0; i < low.length; i++)
        for (var j = 0; j < words.length; j++)
          if (low[i].indexOf(words[j]) !== -1) return i;
      return -1;
    }
    return {
      addr: find(['адрес','пвз','офис','address','pvz']),
      sq:   find(['квадрат','место','square']),
      cell: find(['ячейк','cell']),
      qty:  find(['короб','кол','qty','count','шт'])
    };
  }

  function fillMapping(m){
    ['cAddr','cSq','cCell','cQty'].forEach(function(id){
      var s = $(id);
      s.innerHTML = '<option value="-1">— не выбрано —</option>' +
        headers.map(function(h, i){ return '<option value="' + i + '">' + esc(h || ('Колонка ' + (i + 1))) + '</option>'; }).join('');
    });
    $('cAddr').value = m.addr;
    $('cSq').value   = m.sq;
    $('cCell').value = m.cell;
    $('cQty').value  = m.qty;
  }

  function applyMapping(){
    var a = +$('cAddr').value, s = +$('cSq').value, c = +$('cCell').value, q = +$('cQty').value;
    if (a < 0 || s < 0 || c < 0 || q < 0) {
      $('app').style.display = 'block';
      $('mapBox').open = true;
      msg('Не удалось определить колонки автоматически. Выберите их вручную.', 'err');
      return;
    }
    var agg = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var sq = String(r[s] == null ? '' : r[s]).trim();
      var ad = String(r[a] == null ? '' : r[a]).trim();
      var cn = parseInt(String(r[c] == null ? '' : r[c]).trim(), 10);
      if (!sq || !ad) continue;
      if (!(cn >= 1)) continue;
      var n = parseFloat(String(r[q]).replace(/\s/g, '').replace(',', '.'));
      if (!isFinite(n)) n = 0;

      var key = sq + '\u0000' + cn + '\u0000' + ad;
      if (agg[key]) agg[key].qty += n;
      else agg[key] = { square: sq, address: ad, cell: cn, qty: n };
    }
    data = Object.keys(agg).map(function(k){ return agg[k]; });
    if (!data.length) { msg('Не найдено ни одной строки с адресом и ячейкой.', 'err'); return; }

    var maxCell = 54;
    data.forEach(function(d){ if (d.cell > maxCell) maxCell = d.cell; });
    buildGrid(maxCell);

    var squares = Array.from(new Set(data.map(function(d){ return d.square; })))
      .sort(function(x, y){ return x.localeCompare(y, 'ru', { numeric: true }); });
    var sel = $('sq'), prev = sel.value;
    sel.innerHTML = squares.map(function(v){ return '<option>' + esc(v) + '</option>'; }).join('');
    if (squares.indexOf(prev) !== -1) sel.value = prev;

    $('app').style.display = 'block';
    $('btnRecalc').disabled = false;
    $('btnReset').disabled = false;
    msg('Загружено: ' + data.length + ' строк, квадратов: ' + squares.length + '. Разметьте уровни и нажмите «Пересчитать».', 'ok');

    levelMap = { 1: [], 2: [] };
    rebuildLayout();
    render();
  }

  function recalcAll(){
    rebuildLayout();
    render();
    msg('Пересчёт: 1-й ур. — ' + levelMap[1].length + ' яч., 2-й — ' + levelMap[2].length + ' яч.', 'ok');
    $('btnExportOne').disabled = false;
    $('btnExportAll').disabled = false;
  }

  /* ---------- ЯДРО: перестановка внутри квадрата ---------- */
  function applyLevelsToSquare(square){
    var list = data.filter(function(d){ return d.square === square; })
      .map(function(d){ return { square: d.square, address: d.address, origCell: d.cell, qty: d.qty }; });
    if (!list.length) return [];

    var byQtyDesc = list.slice().sort(function(a, b){
      if (b.qty !== a.qty) return b.qty - a.qty;
      return a.address.localeCompare(b.address, 'ru');
    });

    var l1 = levelMap[1].slice();
    var l2 = levelMap[2].slice();

    var strongCount = Math.min(l1.length, byQtyDesc.length);
    var strongOffices = byQtyDesc.slice(0, strongCount);

    var taken = {};
    strongOffices.forEach(function(o){ taken[o.address + '|' + o.origCell] = true; });

    var rest = byQtyDesc.filter(function(o){ return !taken[o.address + '|' + o.origCell]; })
      .slice()
      .sort(function(a, b){
        if (a.qty !== b.qty) return a.qty - b.qty;
        return a.address.localeCompare(b.address, 'ru');
      });
    var weakCount = Math.min(l2.length, rest.length);
    var weakOffices = rest.slice(0, weakCount);

    var placement = {};
    list.forEach(function(o){ placement[o.origCell] = o; });

    function swapCells(cellA, cellB){
      if (cellA === cellB) return;
      var a = placement[cellA];
      var b = placement[cellB];
      if (b) placement[cellA] = b; else delete placement[cellA];
      if (a) placement[cellB] = a; else delete placement[cellB];
    }

    for (var i = 0; i < strongOffices.length; i++) {
      var strong = strongOffices[i];
      var target = l1[i];
      var curCell = null;
      for (var k in placement) { if (placement[k] === strong) { curCell = +k; break; } }
      if (curCell === null) curCell = strong.origCell;
      swapCells(curCell, target);
    }

    for (var j = 0; j < weakOffices.length; j++) {
      var weak = weakOffices[j];
      var target2 = l2[j];
      var curCell2 = null;
      for (var k2 in placement) { if (placement[k2] === weak) { curCell2 = +k2; break; } }
      if (curCell2 === null) curCell2 = weak.origCell;
      swapCells(curCell2, target2);
    }

    var maxCell = 54;
    Object.keys(placement).forEach(function(k){
      var n = +k;
      if (n > maxCell) maxCell = n;
    });
    if (l1.length) maxCell = Math.max(maxCell, l1[l1.length - 1]);
    if (l2.length) maxCell = Math.max(maxCell, l2[l2.length - 1]);

    var result = [];
    for (var c = 1; c <= maxCell; c++) {
      var off = placement[c];
      var lvl = 0;
      if (l1.indexOf(c) !== -1) lvl = 1;
      else if (l2.indexOf(c) !== -1) lvl = 2;
      result.push({
        cell: c,
        level: lvl,
        square: square,
        address: off ? off.address : null,
        qty: off ? off.qty : 0,
        origCell: off ? off.origCell : null,
        moved: off ? (off.origCell !== c) : false
      });
    }
    return result;
  }

  function rebuildLayout(){
    layout = {};
    var squares = Array.from(new Set(data.map(function(d){ return d.square; })));
    squares.forEach(function(sq){ layout[sq] = applyLevelsToSquare(sq); });
  }

  function getLevel(cellNum){
    if (levelMap[1].indexOf(cellNum) !== -1) return 1;
    if (levelMap[2].indexOf(cellNum) !== -1) return 2;
    return 0;
  }

  /* ---------- Отрисовка ---------- */
  function render(){
    var sq = $('sq').value;
    var cells = layout[sq] || [];

    var maxCell = 54;
    data.forEach(function(d){ if (d.square === sq && d.cell > maxCell) maxCell = d.cell; });
    var needed = cellsNeeded(maxCell);
    if (needed !== totalCells) buildGrid(maxCell);

    var byCell = {};
    cells.forEach(function(c){ byCell[c.cell] = c; });

    var maxQty = 1;
    cells.forEach(function(c){ if (c.qty > maxQty) maxQty = c.qty; });
    var tH = 0.7, tM = 0.3;

    for (var n = 1; n <= totalCells; n++) {
      var c = cellRefs[n - 1];
      if (!c) continue;
      var d = byCell[n];

      if (viewMode === 'was') {
        var wasOffice = null;
        cells.forEach(function(x){ if (x.origCell === n) wasOffice = x; });
        d = wasOffice ? { cell: n, level: 0, address: wasOffice.address, qty: wasOffice.qty, moved: false, origCell: n } : null;
      }

      c.className = 'cell';
      if (paintMode === 1) c.classList.add('paint-1');
      else if (paintMode === 2) c.classList.add('paint-2');

      var lvl = getLevel(n);
      if (lvl === 1) c.classList.add('lvl-1');
      else if (lvl === 2) c.classList.add('lvl-2');

      if (d && d.address) {
        var ratio = d.qty / maxQty;
        var load = ratio >= tH ? 'high' : (ratio >= tM ? 'mid' : 'low');
        c.classList.add(load);
        c.querySelector('.q').textContent = d.qty;
        c.querySelector('.a').textContent = d.address;
        c.querySelector('.mv').textContent = (viewMode === 'now' && d.moved) ? '← ' + d.origCell : '';
        c.title = 'Ячейка ' + n +
          (viewMode === 'now' ? ' (уровень ' + lvl + ')' : '') +
          '\n' + d.address + '\n' + d.qty + ' кор.' +
          (viewMode === 'now' && d.moved ? '\nПриехал с ячейки ' + d.origCell : '');
      } else {
        c.classList.add('empty');
        c.querySelector('.q').textContent = '0';
        c.querySelector('.a').textContent = 'Пусто';
        c.querySelector('.mv').textContent = '';
        c.removeAttribute('title');
      }
    }

    var sum = 0, officesCount = 0;
    cells.forEach(function(c){
      if (c.address) { sum += c.qty; officesCount++; }
    });
    $('sSq').textContent = sq || '—';
    $('sCount').textContent = officesCount;
    $('sSum').textContent = Math.round(sum * 100) / 100;
    $('sL1').textContent = levelMap[1].length;
    $('sL2').textContent = levelMap[2].length;

    renderTable(cells);
  }

  function renderTable(cells){
    var term = $('q').value.trim().toLowerCase();
    var view = cells.filter(function(c){ return c.address; })
      .filter(function(c){ return !term || c.address.toLowerCase().indexOf(term) !== -1; });

    view.sort(function(a, b){
      var x = a[sortKey], y = b[sortKey];
      if (typeof x === 'string') return sortDir * x.localeCompare(y, 'ru', { numeric: true });
      return sortDir * (x - y);
    });

    $('tb').innerHTML = view.length ? view.map(function(c){
      var lvClass = 'lvl' + (c.level || 0);
      var lvText = c.level ? (c.level + '-й') : '—';
      var moved = c.moved ? ' <span class="badge low">переезд</span>' : '';
      return '<tr>' +
        '<td class="num">' + c.cell + '</td>' +
        '<td class="num">' + (c.origCell || '—') + '</td>' +
        '<td class="num"><span class="badge ' + lvClass + '">' + lvText + '</span></td>' +
        '<td>' + esc(c.address) + moved + '</td>' +
        '<td class="num">' + c.qty + '</td>' +
        '</tr>';
    }).join('') : '<tr><td colspan="5" style="color:var(--muted)">Ничего не найдено</td></tr>';
  }

  function resetLevels(){
    levelMap = { 1: [], 2: [] };
    rebuildLayout();
    render();
    msg('Уровни сброшены.', 'ok');
  }

  /* ---------- Экспорт: текущий КС ---------- */
  function sheetFromLayout(square){
    var cells = layout[square] || [];
    var offices = cells.filter(function(c){ return c.address; });
    var sum = offices.reduce(function(s, c){ return s + c.qty; }, 0);

    var aoa = [];
    aoa.push(['Квадрат', square]);
    aoa.push([]);
    aoa.push(['Итог по квадрату']);
    aoa.push(['Офисов', offices.length]);
    aoa.push(['Коробок всего', Math.round(sum * 100) / 100]);
    aoa.push(['Ячеек 1-го ур.', levelMap[1].length]);
    aoa.push(['Ячеек 2-го ур.', levelMap[2].length]);
    aoa.push([]);
    aoa.push(['Ячейка', 'Уровень', 'Была ячейка', 'Адрес ПВЗ', 'Коробок', 'Переезд']);
    offices.sort(function(a,b){ return a.cell - b.cell; });
    offices.forEach(function(c){
      aoa.push([c.cell, c.level || '', c.origCell, c.address, c.qty, c.moved ? 'да' : '']);
    });
    return aoa;
  }

  function exportCurrent(){
    if (!data.length) return;
    var sq = $('sq').value;
    if (!sq) return;
    loadXLSX().then(function(){
      var wb = XLSX.utils.book_new();
      var aoa = sheetFromLayout(sq);
      var ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [{wch:10},{wch:10},{wch:12},{wch:50},{wch:10},{wch:10}];
      var sheetName = ('Квадрат ' + sq).slice(0, 31).replace(/[\\\/\?\*\[\]:]/g, '_');
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      var safeName = String(sq).replace(/[^\wа-яё\-]+/gi, '_').slice(0, 40) || 'square';
      XLSX.writeFile(wb, 'ks_' + safeName + '.xlsx');
      msg('Скачан отчёт по квадрату: ' + sq, 'ok');
    }).catch(function(err){ msg(err.message, 'err'); });
  }

  /* ---------- Экспорт: все КС на ОДНОМ ЛИСТЕ ---------- */
  function exportAll(){
    if (!data.length) return;
    var answer = prompt(
      'Как применить уровни?\n\n' +
      '1 — применить уровни ко ВСЕМ КС (каждый квадрат пересчитан отдельно)\n' +
      '2 — применить уровни ТОЛЬКО к выделенному КС, остальные — как в исходном файле\n\n' +
      'Введите 1 или 2 (Enter — применить ко всем):', '1');
    if (answer === null) return;
    var applyAll = (answer.trim() !== '2');

    var selected = $('sq').value;
    var oldLayout = layout;
    var newLayout = {};
    var squares = Array.from(new Set(data.map(function(d){ return d.square; })))
      .sort(function(a, b){ return a.localeCompare(b, 'ru', { numeric: true }); });

    squares.forEach(function(sq){
      if (applyAll) newLayout[sq] = applyLevelsToSquare(sq);
      else newLayout[sq] = (sq === selected) ? applyLevelsToSquare(sq) : (oldLayout[sq] || []);
    });

    loadXLSX().then(function(){
      var wb = XLSX.utils.book_new();

      // Один лист: сначала сводка, потом общая таблица по всем квадратам
      var aoa = [];

      // ---- Сводка ----
      aoa.push(['СВОДКА ПО КВАДРАТАМ']);
      aoa.push(['Квадрат', 'Офисов', 'Коробок всего', 'Ячеек 1-го ур.', 'Ячеек 2-го ур.']);
      squares.forEach(function(sq){
        var cells = newLayout[sq] || [];
        var offices = cells.filter(function(c){ return c.address; });
        var sum = offices.reduce(function(s, c){ return s + c.qty; }, 0);
        aoa.push([sq, offices.length, Math.round(sum * 100) / 100, levelMap[1].length, levelMap[2].length]);
      });

      // Пустые строки-разделители
      aoa.push([]);
      aoa.push([]);

      // ---- Общая таблица по всем квадратам ----
      aoa.push(['ДЕТАЛИЗАЦИЯ ПО ВСЕМ КВАДРАТАМ']);
      aoa.push(['Квадрат', 'Ячейка', 'Уровень', 'Была ячейка', 'Адрес ПВЗ', 'Коробок', 'Переезд']);

      squares.forEach(function(sq){
        var cells = newLayout[sq] || [];
        var offices = cells.filter(function(c){ return c.address; });
        offices.sort(function(a,b){ return a.cell - b.cell; });
        offices.forEach(function(c){
          aoa.push([
            sq,
            c.cell,
            c.level || '',
            c.origCell,
            c.address,
            c.qty,
            c.moved ? 'да' : ''
          ]);
        });
      });

      var ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [
        {wch:28}, // Квадрат
        {wch:10}, // Ячейка
        {wch:10}, // Уровень
        {wch:12}, // Была ячейка
        {wch:50}, // Адрес
        {wch:10}, // Коробок
        {wch:10}  // Переезд
      ];
      XLSX.utils.book_append_sheet(wb, ws, 'Все КС');

      XLSX.writeFile(wb, 'ks_all.xlsx');
      msg('Скачан отчёт по всем КС (' + (applyAll ? 'уровни применены ко всем' : 'уровни применены только к ' + selected) + ').', 'ok');
    }).catch(function(err){ msg(err.message, 'err'); });
  }

  function downloadTemplate(){
    loadXLSX().then(function(){
      var ws = XLSX.utils.aoa_to_sheet([
        ['Адрес', 'Квадрат', 'Ячейка', 'Кол-во коробок'],
        ['г. Москва, ул. Ленина, 1', 'A1', 1, 120],
        ['г. Москва, ул. Мира, 5', 'A1', 2, 45],
        ['г. Москва, пр-т Победы, 10', 'A1', 3, 80]
      ]);
      ws['!cols'] = [{wch:40},{wch:10},{wch:10},{wch:16}];
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Данные');
      XLSX.writeFile(wb, 'template.xlsx');
    }).catch(function(err){ msg(err.message, 'err'); });
  }

  /* ---------- События ---------- */
  var fileInput = $('file'), drop = $('drop');
  function pick(){ fileInput.click(); }
  $('btnOpen').addEventListener('click', pick);
  drop.addEventListener('click', pick);
  drop.addEventListener('keydown', function(e){ if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); } });
  fileInput.addEventListener('change', function(){
    var f = fileInput.files && fileInput.files[0];
    fileInput.value = '';
    handleFile(f);
  });
  ['dragenter','dragover'].forEach(function(ev){
    drop.addEventListener(ev, function(e){ e.preventDefault(); drop.classList.add('over'); });
  });
  ['dragleave','drop'].forEach(function(ev){
    drop.addEventListener(ev, function(e){ e.preventDefault(); drop.classList.remove('over'); });
  });
  drop.addEventListener('drop', function(e){ handleFile(e.dataTransfer.files && e.dataTransfer.files[0]); });
  window.addEventListener('dragover', function(e){ e.preventDefault(); });
  window.addEventListener('drop', function(e){ e.preventDefault(); });

  $('btnRecalc').addEventListener('click', recalcAll);
  $('btnReset').addEventListener('click', resetLevels);
  $('btnExportOne').addEventListener('click', exportCurrent);
  $('btnExportAll').addEventListener('click', exportAll);
  $('btnTpl').addEventListener('click', downloadTemplate);
  $('sq').addEventListener('change', render);
  $('q').addEventListener('input', function(){ renderTable(layout[$('sq').value] || []); });
  ['cAddr','cSq','cCell','cQty'].forEach(function(id){ $(id).addEventListener('change', applyMapping); });

  document.querySelectorAll('#paintModes .seg').forEach(function(b){
    b.addEventListener('click', function(){
      document.querySelectorAll('#paintModes .seg').forEach(function(x){ x.classList.remove('active'); });
      b.classList.add('active');
      paintMode = +b.dataset.mode;
      render();
    });
  });

  document.querySelectorAll('#viewModes .seg').forEach(function(b){
    b.addEventListener('click', function(){
      document.querySelectorAll('#viewModes .seg').forEach(function(x){ x.classList.remove('active'); });
      b.classList.add('active');
      viewMode = b.dataset.view;
      render();
    });
  });

  document.querySelectorAll('th[data-k]').forEach(function(th){
    th.addEventListener('click', function(){
      var k = th.getAttribute('data-k');
      if (sortKey === k) sortDir = -sortDir; else { sortKey = k; sortDir = (k === 'address' || k === 'cell' || k === 'origCell') ? 1 : -1; }
      renderTable(layout[$('sq').value] || []);
    });
  });
})();
