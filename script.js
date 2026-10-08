(function(){
  'use strict';
  var $ = function(id){ return document.getElementById(id); };
  var rows = [];               // сырые строки листа
  var headers = [];
  var data = [];               // {square, address, cell, qty} — исходные данные из файла
  var layout = {};             // { square: [ {square, address, cell, origCell, qty, level} x N ] }
  var levelMap = { 1: [], 2: [] };
  var sortKey = 'cell', sortDir = 1;
  var viewMode = 'was';        // 'was' | 'now'
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
  function parseCellInput(str, max){
    var out = [];
    String(str).split(/[,\s]+/).forEach(function(part){
      if (!part) return;
      var m = part.match(/^(\d+)\s*-\s*(\d+)$/);
      if (m) {
        var a = +m[1], b = +m[2];
        if (a > b) { var t = a; a = b; b = t; }
        for (var i = a; i <= b; i++) if (i >= 1 && i <= max) out.push(i);
      } else {
        var n = parseInt(part, 10);
        if (n >= 1 && n <= max) out.push(n);
      }
    });
    return Array.from(new Set(out)).sort(function(x,y){ return x-y; });
  }

  /* ---------- Сетка: динамическая, минимум 54 (6 групп по 9) ---------- */
  var cellRefs = [];      // [cell] по порядку номеров 1..N
  var totalCells = 54;    // может увеличиться, если офисов > 54

  function buildGrid(n){
    if (!n) n = 54;
    // Округляем до кратного 9 (полный столбец 3×3)
    var groups = Math.ceil(n / 9);
    totalCells = groups * 9;

    var grid = $('grid');
    grid.innerHTML = '';
    cellRefs = [];

    for (var g = 0; g < groups; g++) {
      var group = document.createElement('div');
      group.className = 'group';
      var localCells = [];
      for (var col = 0; col < 3; col++) {
        for (var row = 0; row < 3; row++) {
          var num = (g * 9) + (col * 3) + row + 1;
          var c = document.createElement('div');
          c.className = 'cell empty';
          c.dataset.num = num;
          c.innerHTML = '<span class="n">' + num + '</span><div class="q">0</div><div class="a">Пусто</div><span class="mv"></span>';
          group.appendChild(c);
          localCells.push(c);
        }
      }
      grid.appendChild(group);
      cellRefs = cellRefs.concat(localCells);
    }
  }
  buildGrid(54);

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
      if (!(cn >= 1)) continue; // ячейка должна быть >=1
      var n = parseFloat(String(r[q]).replace(/\s/g, '').replace(',', '.'));
      if (!isFinite(n)) n = 0;

      var key = sq + '\u0000' + cn + '\u0000' + ad;
      if (agg[key]) agg[key].qty += n;
      else agg[key] = { square: sq, address: ad, cell: cn, qty: n };
    }
    data = Object.keys(agg).map(function(k){ return agg[k]; });
    if (!data.length) { msg('Не найдено ни одной строки с адресом и ячейкой.', 'err'); return; }

    // Максимальный номер ячейки — для расширения сетки
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
    msg('Загружено: ' + data.length + ' строк, квадратов: ' + squares.length + '. Нажмите «Пересчитать».', 'ok');

    levelMap = { 1: [], 2: [] };
    layout = {};
    rebuildLayout();
    render();
  }

  /* ---------- Пересчёт ---------- */
  function recalcFlow(){
    var s1 = prompt(
      'ШАГ 1/2 — Ячейки 1-го уровня (сюда попадут САМЫЕ СИЛЬНЫЕ офисы каждого квадрата).\n' +
      'Формат: 1-5, 8, 12-14\n\n' +
      'Оставьте пустым, если 1-й уровень не нужен.', '1-5');
    if (s1 === null) return;
    var l1 = s1.trim() ? parseCellInput(s1, 1e6) : [];

    var s2 = prompt(
      'ШАГ 2/2 — Ячейки 2-го уровня (сюда попадут САМЫЕ СЛАБЫЕ офисы каждого квадрата).\n' +
      'Уже занято 1-м: ' + (l1.join(',') || '—') + '\n\n' +
      'Оставьте пустым, если 2-й уровень не нужен.', '48-54');
    if (s2 === null) return;
    var l2 = s2.trim() ? parseCellInput(s2, 1e6).filter(function(n){ return l1.indexOf(n) === -1; }) : [];

    if (!l1.length && !l2.length) { msg('Ни один уровень не назначен.', 'err'); return; }
    levelMap = { 1: l1, 2: l2 };
    rebuildLayout();
    render();
    msg('Уровни: 1-й — ' + l1.length + ' яч., 2-й — ' + l2.length + ' яч.', 'ok');
    $('btnExportOne').disabled = false;
    $('btnExportAll').disabled = false;
  }

  /* ---------- Основная логика перестановки ----------
     Для каждого квадрата:
     - Сортируем офисы по qty убыв.
     - Сильные: топ-N по ячейкам 1-го уровня.
     - Слабые: топ-M по ячейкам 2-го уровня из оставшихся.
     - Остальные остаются на своих ячейках, но если их ячейку забрали —
       они переезжают на освободившуюся (обмен местами).
  */
  function applyLevelsToSquare(square){
    var list = data.filter(function(d){ return d.square === square; })
      .map(function(d){ return { square: d.square, address: d.address, origCell: d.cell, qty: d.qty }; });

    if (!list.length) return [];

    // Сортировка по убыванию коробок (для сильных) — стабильная
    var byQtyDesc = list.slice().sort(function(a, b){
      if (b.qty !== a.qty) return b.qty - a.qty;
      return a.address.localeCompare(b.address, 'ru');
    });

    // Ячейки уровней
    var l1 = levelMap[1].slice();
    var l2 = levelMap[2].slice();

    // Сформировать итоговую карту: cell -> office
    // Стартуем с исходного размещения
    var placement = {};       // cell -> office
    list.forEach(function(d){ placement[d.origCell] = d; });

    // Отмечаем, какие офисы уже переставлены
    var used = {}; // address -> true
    var strongOffices = [];
    var weakOffices = [];

    // Сильные: топ-N из byQtyDesc
    var strongCount = Math.min(l1.length, byQtyDesc.length);
    for (var i = 0; i < strongCount; i++) {
      strongOffices.push(byQtyDesc[i]);
      used[byQtyDesc[i].address] = true;
    }

    // Слабые: топ-M из byQtyDesc по возрастанию qty, исключая уже взятых
    var rest = byQtyDesc.filter(function(d){ return !used[d.address]; })
      .slice()
      .sort(function(a, b){
        if (a.qty !== b.qty) return a.qty - b.qty;
        return a.address.localeCompare(b.address, 'ru');
      });
    var weakCount = Math.min(l2.length, rest.length);
    for (var j = 0; j < weakCount; j++) {
      weakOffices.push(rest[j]);
      used[rest[j].address] = true;
    }

    // Раскладка "сильные -> ячейки 1-го уровня"
    // При этом тот офис, который сидел в этой ячейке, отправляется в origCell сильного.
    strongOffices.forEach(function(strong, idx){
      var targetCell = l1[idx];
      var displaced = placement[targetCell];   // офис, сидевший в целевой ячейке
      var strongOrigCell = strong.origCell;

      // Убираем displaced из его позиции (если он не тот же самый)
      if (displaced && displaced !== strong) {
        // Переезжает на origCell сильного
        delete placement[targetCell];
        placement[strongOrigCell] = displaced;
        displaced.newCell = strongOrigCell;
      } else if (!displaced) {
        // Ячейка была пустая
        // Сильный уезжает со своего origCell
        delete placement[strong.origCell];
      }
      placement[targetCell] = strong;
      strong.newCell = targetCell;
    });

    // Раскладка "слабые -> ячейки 2-го уровня"
    weakOffices.forEach(function(weak, idx){
      var targetCell = l2[idx];
      var displaced = placement[targetCell];
      var weakOrigCell = weak.origCell;

      if (displaced && displaced !== weak) {
        delete placement[targetCell];
        placement[weakOrigCell] = displaced;
        displaced.newCell = weakOrigCell;
      } else if (!displaced) {
        delete placement[weak.origCell];
      }
      placement[targetCell] = weak;
      weak.newCell = targetCell;
    });

    // Собираем результат: каждой ячейке либо офис, либо пусто
    // Возвращаем массив ячеек от 1 до N, где N = максимальная занятая ячейка
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
      var level = 0;
      if (l1.indexOf(c) !== -1) level = 1;
      else if (l2.indexOf(c) !== -1) level = 2;
      result.push({
        cell: c,
        level: level,
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
    squares.forEach(function(sq){
      layout[sq] = applyLevelsToSquare(sq);
    });
  }

  /* ---------- Отрисовка ---------- */
  function getLevel(cellNum){
    if (levelMap[1].indexOf(cellNum) !== -1) return 1;
    if (levelMap[2].indexOf(cellNum) !== -1) return 2;
    return 0;
  }

  function render(){
    var sq = $('sq').value;
    var cells = layout[sq] || [];

    // Собираем мапу: cell -> { address, qty, level, moved, origCell }
    var byCell = {};
    cells.forEach(function(c){ byCell[c.cell] = c; });

    // Определяем максимальный номер ячейки, чтобы расширить сетку при необходимости
    var maxCell = totalCells;
    cells.forEach(function(c){ if (c.cell > maxCell) maxCell = c.cell; });
    if (maxCell > totalCells) buildGrid(maxCell);

    // Цветовая шкала нагрузки: считаем максимум по всем офисам квадрата
    var maxQty = 1;
    cells.forEach(function(c){ if (c.qty > maxQty) maxQty = c.qty; });
    var tH = 0.7, tM = 0.3;

    for (var n = 1; n <= totalCells; n++) {
      var c = cellRefs[n - 1];
      if (!c) continue;
      var d = byCell[n];

      // В режиме "Было" показываем origCell-позицию, т.е. офис, который изначально сидел в этой ячейке
      if (viewMode === 'was') {
        // Найти офис, у которого origCell == n
        var wasOffice = null;
        cells.forEach(function(x){ if (x.origCell === n) wasOffice = x; });
        d = wasOffice ? { cell: n, level: 0, address: wasOffice.address, qty: wasOffice.qty, moved: false, origCell: n } : null;
      }

      c.className = 'cell';
      var lvl = getLevel(n);
      if (lvl === 1) c.classList.add('lvl-1');
      else if (lvl === 2) c.classList.add('lvl-2');

      if (d && d.address) {
        var ratio = d.qty / maxQty;
        var load = ratio >= tH ? 'high' : (ratio >= tM ? 'mid' : 'low');
        c.classList.add(load);
        c.querySelector('.q').textContent = d.qty;
        c.querySelector('.a').textContent = d.address;
        // В режиме "Стало" показываем, откуда приехал офис
        if (viewMode === 'now' && d.moved) {
          c.querySelector('.mv').textContent = '← ' + d.origCell;
        } else {
          c.querySelector('.mv').textContent = '';
        }
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

    // Статистика
    var sum = 0, l1Count = 0, l2Count = 0, officesCount = 0;
    cells.forEach(function(c){
      if (c.address) { sum += c.qty; officesCount++; }
      if (c.level === 1) l1Count++;
      else if (c.level === 2) l2Count++;
    });
    $('sSq').textContent = sq || '—';
    $('sCount').textContent = officesCount;
    $('sSum').textContent = Math.round(sum * 100) / 100;
    $('sL1').textContent = l1Count;
    $('sL2').textContent = l2Count;

    renderTable(cells);
  }

  function renderTable(cells){
    var term = $('q').value.trim().toLowerCase();
    // В таблице всегда показываем актуальное ("Стало"), но добавляем столбец "Была"
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

  /* ---------- Сброс уровней ---------- */
  function resetLevels(){
    levelMap = { 1: [], 2: [] };
    rebuildLayout();
    render();
    msg('Уровни сброшены.', 'ok');
  }

  /* ---------- Экспорт ---------- */
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
      aoa.push([
        c.cell,
        c.level || '',
        c.origCell,
        c.address,
        c.qty,
        c.moved ? 'да' : ''
      ]);
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
    // Сохраняем старый layout и строим новый по запросу
    var oldLayout = layout;
    var newLayout = {};
    var squares = Array.from(new Set(data.map(function(d){ return d.square; })));

    squares.forEach(function(sq){
      if (applyAll) {
        newLayout[sq] = applyLevelsToSquare(sq);
      } else {
        if (sq === selected) {
          newLayout[sq] = applyLevelsToSquare(sq);
        } else {
          // Квадрат без изменений: как в исходном файле
          newLayout[sq] = oldLayout[sq] || [];
        }
      }
    });

    loadXLSX().then(function(){
      var wb = XLSX.utils.book_new();

      // Лист 1: сводка
      var summary = [['Квадрат', 'Офисов', 'Коробок всего', 'Ячеек 1-го ур.', 'Ячеек 2-го ур.']];
      squares.forEach(function(sq){
        var cells = newLayout[sq] || [];
        var offices = cells.filter(function(c){ return c.address; });
        var sum = offices.reduce(function(s, c){ return s + c.qty; }, 0);
        var l1c = cells.filter(function(c){ return c.level === 1; }).length;
        var l2c = cells.filter(function(c){ return c.level === 2; }).length;
        summary.push([sq, offices.length, Math.round(sum * 100) / 100, l1c, l2c]);
      });
      var wsSum = XLSX.utils.aoa_to_sheet(summary);
      wsSum['!cols'] = [{wch:30},{wch:10},{wch:14},{wch:16},{wch:16}];
      XLSX.utils.book_append_sheet(wb, wsSum, 'Сводка');

      // Листы по квадратам
      squares.forEach(function(sq){
        var cells = newLayout[sq] || [];
        var offices = cells.filter(function(c){ return c.address; });
        offices.sort(function(a,b){ return a.cell - b.cell; });

        var aoa = [];
        aoa.push(['Ячейка', 'Уровень', 'Была ячейка', 'Адрес ПВЗ', 'Коробок', 'Переезд']);
        offices.forEach(function(c){
          aoa.push([c.cell, c.level || '', c.origCell, c.address, c.qty, c.moved ? 'да' : '']);
        });
        var ws = XLSX.utils.aoa_to_sheet(aoa);
        ws['!cols'] = [{wch:10},{wch:10},{wch:12},{wch:50},{wch:10},{wch:10}];
        var name = String(sq).slice(0, 28).replace(/[\\\/\?\*\[\]:]/g, '_') || 'Лист';
        // Гарантируем уникальность имени
        var base = name, k = 2;
        while (wb.SheetNames.indexOf(name) !== -1) { name = base + '_' + (k++); }
        XLSX.utils.book_append_sheet(wb, ws, name);
      });

      XLSX.writeFile(wb, 'ks_all.xlsx');
      msg('Скачан отчёт по всем КС (' + (applyAll ? 'уровни применены ко всем' : 'уровни применены только к ' + selected) + ').', 'ok');
    }).catch(function(err){ msg(err.message, 'err'); });
  }

  /* ---------- Шаблон ---------- */
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

  $('btnRecalc').addEventListener('click', recalcFlow);
  $('btnReset').addEventListener('click', resetLevels);
  $('btnExportOne').addEventListener('click', exportCurrent);
  $('btnExportAll').addEventListener('click', exportAll);
  $('btnTpl').addEventListener('click', downloadTemplate);
  $('sq').addEventListener('change', render);
  $('q').addEventListener('input', function(){ renderTable(layout[$('sq').value] || []); });
  ['cAddr','cSq','cCell','cQty'].forEach(function(id){ $(id).addEventListener('change', applyMapping); });

  // Переключатель "Было / Стало"
  $('tglWas').addEventListener('click', function(){
    viewMode = 'was';
    $('tglWas').classList.add('active');
    $('tglNow').classList.remove('active');
    render();
  });
  $('tglNow').addEventListener('click', function(){
    viewMode = 'now';
    $('tglNow').classList.add('active');
    $('tglWas').classList.remove('active');
    render();
  });

  document.querySelectorAll('th[data-k]').forEach(function(th){
    th.addEventListener('click', function(){
      var k = th.getAttribute('data-k');
      if (sortKey === k) sortDir = -sortDir; else { sortKey = k; sortDir = (k === 'address' || k === 'cell' || k === 'origCell') ? 1 : -1; }
      renderTable(layout[$('sq').value] || []);
    });
  });
})();
