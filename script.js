(function(){
  'use strict';
  var $ = function(id){ return document.getElementById(id); };
  var rows = [];          // сырые строки листа
  var headers = [];
  var data = [];          // {square,address,qty} — все ПВЗ по всем квадратам
  var sortKey = 'rank', sortDir = 1;
  var LEVEL = { high: 'Ходовой', mid: 'Средний', low: 'Низкий' };

  /* ---------- Библиотека XLSX с запасным источником ---------- */
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
        if (i >= urls.length) { xlsxReady = null; return reject(new Error('Не удалось загрузить библиотеку XLSX. Проверьте интернет или блокировщик рекламы.')); }
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

  /* ---------- Сообщения ---------- */
  function msg(text, type){
    var m = $('msg');
    m.textContent = text || '';
    m.className = text ? type : '';
  }
  function esc(s){
    return String(s).replace(/[&<>"']/g, function(c){
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];
    });
  }

  /* ---------- Сетка: 6 групп × (3 колонки × 3 строки) = 54 ячейки ---------- */
  var cellRefs = [];
  function buildGrid(){
    var grid = $('grid');
    grid.innerHTML = '';
    cellRefs = [];

    for (var g = 0; g < 6; g++) {
      var group = document.createElement('div');
      group.className = 'group';

      var localCells = [];
      for (var col = 0; col < 3; col++) {
        for (var row = 0; row < 3; row++) {
          var num = (g * 9) + (col * 3) + row + 1;
          var c = document.createElement('div');
          c.className = 'cell empty';
          c.dataset.num = num;
          c.innerHTML = '<span class="n">' + num + '</span><div class="q">0</div><div class="a">Пусто</div>';
          group.appendChild(c);
          localCells.push(c);
        }
      }
      grid.appendChild(group);
      cellRefs = cellRefs.concat(localCells);
    }
  }
  buildGrid();

  /* ---------- Загрузка файла ---------- */
  function handleFile(file){
    if (!file) return;
    msg('Читаю файл…', 'ok');
    $('btnExport').disabled = true;
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
      if (json.length < 2) throw new Error('В файле нет данных. Нужна строка заголовков и хотя бы одна строка ПВЗ.');
      rows = json;
      headers = json[0].map(function(h){ return String(h).trim(); });
      fillMapping(detect(headers));
      applyMapping();
    }).catch(function(err){
      msg(err.message || String(err), 'err');
    });
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
      qty:  find(['короб','кол','qty','count','шт'])
    };
  }

  function fillMapping(m){
    ['cAddr','cSq','cQty'].forEach(function(id){
      var s = $(id);
      s.innerHTML = '<option value="-1">— не выбрано —</option>' +
        headers.map(function(h, i){ return '<option value="' + i + '">' + esc(h || ('Колонка ' + (i + 1))) + '</option>'; }).join('');
    });
    $('cAddr').value = m.addr; $('cSq').value = m.sq; $('cQty').value = m.qty;
  }

  function applyMapping(){
    var a = +$('cAddr').value, s = +$('cSq').value, q = +$('cQty').value;
    if (a < 0 || s < 0 || q < 0) {
      $('app').style.display = 'block';
      $('mapBox').open = true;
      msg('Не удалось определить колонки автоматически. Выберите их вручную ниже: адрес, квадрат, коробок.', 'err');
      return;
    }
    var agg = {};
    for (var i = 1; i < rows.length; i++) {
      var r = rows[i];
      var sq = String(r[s] == null ? '' : r[s]).trim();
      var ad = String(r[a] == null ? '' : r[a]).trim();
      if (!sq || !ad) continue;
      var n = parseFloat(String(r[q]).replace(/\s/g, '').replace(',', '.'));
      if (!isFinite(n)) n = 0;
      var key = sq + '\u0000' + ad;
      if (agg[key]) agg[key].qty += n; else agg[key] = { square: sq, address: ad, qty: n };
    }
    data = Object.keys(agg).map(function(k){ return agg[k]; });
    if (!data.length) {
      msg('Не найдено ни одной строки с адресом и квадратом. Проверьте выбранные колонки.', 'err');
      return;
    }
    var squares = Array.from(new Set(data.map(function(d){ return d.square; })))
      .sort(function(x, y){ return x.localeCompare(y, 'ru', { numeric: true }); });
    var sel = $('sq'), prev = sel.value;
    sel.innerHTML = squares.map(function(v){ return '<option>' + esc(v) + '</option>'; }).join('');
    if (squares.indexOf(prev) !== -1) sel.value = prev;
    $('app').style.display = 'block';
    $('btnExport').disabled = false;
    msg('Загружено: ' + data.length + ' ПВЗ, квадратов: ' + squares.length, 'ok');
    render();
  }

  /* ---------- Расчёт — ТОЛЬКО ПО ВЫБРАННОМУ КВАДРАТУ ---------- */
  function thresholds(){
    var h = +$('tHigh').value / 100, m = +$('tMid').value / 100;
    if (m > h) m = h;
    return { high: h, mid: m };
  }
  function level(ratio, t){ return ratio >= t.high ? 'high' : (ratio >= t.mid ? 'mid' : 'low'); }

  // Возвращает список ПВЗ только для выбранного квадрата, с рассчитанными rank/pct/level
  function current(){
    var sq = $('sq').value;
    var list = data.filter(function(d){ return d.square === sq; })
      .sort(function(a, b){ return b.qty - a.qty || a.address.localeCompare(b.address, 'ru'); });

    // Максимум считаем ТОЛЬКО внутри этого квадрата
    var max = list.length ? (list[0].qty || 1) : 1;
    var t = thresholds();

    list.forEach(function(d, i){
      d.rank = i + 1;
      d.pct = Math.round(d.qty / max * 100);
      d.level = level(d.qty / max, t);
    });
    return list;
  }

  function render(){
    $('oHigh').textContent = $('tHigh').value + '%';
    $('oMid').textContent = $('tMid').value + '%';

    var list = current();          // список ТОЛЬКО по текущему квадрату

    // Заполняем 54 ячейки: первые N — ПВЗ текущего квадрата, остальные — пустые
    for (var i = 0; i < 54; i++) {
      var c = cellRefs[i], d = list[i];
      if (!c) continue;
      if (d) {
        c.className = 'cell ' + d.level;
        c.querySelector('.q').textContent = d.qty;
        c.querySelector('.a').textContent = d.address;
        c.title = 'Ячейка ' + (i + 1) + ': ' + d.address + ' — ' + d.qty + ' (' + LEVEL[d.level] + ')';
      } else {
        c.className = 'cell empty';
        c.querySelector('.q').textContent = '0';
        c.querySelector('.a').textContent = 'Пусто';
        c.removeAttribute('title');
      }
    }

    // Статистика — только по выбранному квадрату
    var sum = list.reduce(function(s, d){ return s + d.qty; }, 0);
    $('sSq').textContent = $('sq').value || '—';
    $('sCount').textContent = list.length;
    $('sSum').textContent = Math.round(sum * 100) / 100;
    $('sAvg').textContent = list.length ? Math.round(sum / list.length * 10) / 10 : 0;
    $('sOver').textContent = Math.max(0, list.length - 54);

    renderTable(list);
  }

  function renderTable(list){
    var term = $('q').value.trim().toLowerCase();
    var view = list.filter(function(d){ return !term || d.address.toLowerCase().indexOf(term) !== -1; });
    view.sort(function(a, b){
      var x = a[sortKey], y = b[sortKey];
      if (typeof x === 'string') return sortDir * x.localeCompare(y, 'ru', { numeric: true });
      return sortDir * (x - y);
    });
    $('tb').innerHTML = view.length ? view.map(function(d){
      return '<tr><td class="num">' + d.rank + '</td><td>' + esc(d.address) + '</td><td class="num">' + d.qty +
        '</td><td class="num">' + d.pct + '%</td><td><span class="badge ' + d.level + '">' + LEVEL[d.level] + '</span></td></tr>';
    }).join('') : '<tr><td colspan="5" style="color:var(--muted)">Ничего не найдено</td></tr>';
  }

  /* ---------- Экспорт: отчёт по выбранному квадрату + сводка по всем ---------- */
  function exportReport(){
    if (!data.length) return;
    var currentSquare = $('sq').value;
    if (!currentSquare) return;

    loadXLSX().then(function(){
      var t = thresholds();

      // Лист 1: детализация ТОЛЬКО по выбранному квадрату
      var detail = [['Квадрат', 'Ячейка №', 'Адрес ПВЗ', 'Коробок', '% от макс.', 'Уровень']];
      var list = data.filter(function(d){ return d.square === currentSquare; })
        .slice()
        .sort(function(a, b){ return b.qty - a.qty; });
      var max = list.length ? (list[0].qty || 1) : 1;
      list.forEach(function(d, i){
        var lv = level(d.qty / max, t);
        detail.push([currentSquare, i + 1, d.address, d.qty, Math.round(d.qty / max * 100) / 100, LEVEL[lv]]);
      });

      // Лист 2: сводка по ВСЕМ квадратам (для сравнения)
      var bySq = {};
      data.forEach(function(d){ (bySq[d.square] = bySq[d.square] || []).push(d); });
      var squares = Object.keys(bySq).sort(function(a, b){ return a.localeCompare(b, 'ru', { numeric: true }); });

      var summary = [['Квадрат', 'ПВЗ', 'Коробок всего', 'В среднем на ПВЗ', 'Ходовых', 'Средних', 'Низких']];
      squares.forEach(function(sq){
        var l = bySq[sq].slice().sort(function(a, b){ return b.qty - a.qty; });
        var m = l[0].qty || 1, sum = 0, cnt = { high: 0, mid: 0, low: 0 };
        l.forEach(function(d){
          var lv = level(d.qty / m, t);
          cnt[lv]++; sum += d.qty;
        });
        summary.push([sq, l.length, sum, Math.round(sum / l.length * 10) / 10, cnt.high, cnt.mid, cnt.low]);
      });

      var wb = XLSX.utils.book_new();
      var wDetail = XLSX.utils.aoa_to_sheet(detail);
      var wSummary = XLSX.utils.aoa_to_sheet(summary);

      wDetail['!cols'] = [{wch:14},{wch:10},{wch:50},{wch:10},{wch:12},{wch:12}];
      wSummary['!cols'] = [{wch:14},{wch:8},{wch:14},{wch:18},{wch:10},{wch:10},{wch:10}];

      // Формат процентов на листе детализации
      for (var r = 1; r < detail.length; r++) {
        var ref = XLSX.utils.encode_cell({ r: r, c: 4 });
        if (wDetail[ref]) wDetail[ref].z = '0%';
      }

      // Имя листа с квадратом (обрезаем до 31 символа — ограничение Excel)
      var sheetName = ('Квадрат ' + currentSquare).slice(0, 31).replace(/[\\\/\?\*\[\]:]/g, '_');

      XLSX.utils.book_append_sheet(wb, wDetail, sheetName);
      XLSX.utils.book_append_sheet(wb, wSummary, 'Сводка по всем');

      var safeName = String(currentSquare).replace(/[^\wа-яё\-]+/gi, '_').slice(0, 40) || 'square';
      XLSX.writeFile(wb, 'heatmap_' + safeName + '.xlsx');
      msg('Отчёт скачан по квадрату: ' + currentSquare, 'ok');
    }).catch(function(err){ msg(err.message, 'err'); });
  }

  function downloadTemplate(){
    loadXLSX().then(function(){
      var ws = XLSX.utils.aoa_to_sheet([
        ['Адрес', 'Квадрат', 'Кол-во коробок'],
        ['г. Москва, ул. Ленина, 1', 'A1', 120],
        ['г. Москва, ул. Мира, 5', 'A1', 45],
        ['г. Москва, пр-т Победы, 10', 'B2', 80]
      ]);
      ws['!cols'] = [{wch:36},{wch:10},{wch:16}];
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

  $('btnExport').addEventListener('click', exportReport);
  $('btnTpl').addEventListener('click', downloadTemplate);
  $('sq').addEventListener('change', render);       // смена квадрата -> полный пересчёт
  $('tHigh').addEventListener('input', render);
  $('tMid').addEventListener('input', render);
  $('q').addEventListener('input', function(){ renderTable(current()); });
  ['cAddr','cSq','cQty'].forEach(function(id){ $(id).addEventListener('change', applyMapping); });
  document.querySelectorAll('th[data-k]').forEach(function(th){
    th.addEventListener('click', function(){
      var k = th.getAttribute('data-k');
      if (sortKey === k) sortDir = -sortDir; else { sortKey = k; sortDir = (k === 'address' || k === 'rank') ? 1 : -1; }
      renderTable(current());
    });
  });
})();
