# ========== ПРОТОТИП САЙТА v1.4 (ТРЁХУРОВНЕВОЕ МЕНЮ + EXCEL) ==========
$outputPath = "C:\Прокст сайта"
$htmlFile = Join-Path -Path $outputPath -ChildPath "monitoring.html"

if (-not (Test-Path $outputPath)) { New-Item -ItemType Directory -Path $outputPath -Force | Out-Null }

$html = @"
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Мониторинг ПК v1.4</title>
    <style>
        :root {
            --bg: #f0f2f5; --card: #fff; --text: #1a1a2e; --text2: #666;
            --border: #e0e0e0; --primary: #1976D2; --phover: #1565C0;
            --success: #4CAF50; --slight: #d4edda; --warning: #ff9800;
            --wlight: #fff3cd; --danger: #f44336; --dlight: #f8d7da;
            --hbg: #1a1a2e; --radius: 10px; --shadow: 0 2px 4px rgba(0,0,0,0.08);
        }
        *{margin:0;padding:0;box-sizing:border-box}
        body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
        
        .header{background:var(--hbg);color:#fff;padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,.2)}
        .header-left{display:flex;align-items:center;gap:12px}
        .header-logo{font-size:20px;font-weight:700}
        .header-version{font-size:11px;opacity:.6;background:rgba(255,255,255,.1);padding:3px 8px;border-radius:12px}
        .header-right{display:flex;align-items:center;gap:20px}
        .header-time{font-size:14px;opacity:.8}
        .header-disk{background:rgba(255,255,255,.1);padding:6px 14px;border-radius:8px;font-size:13px;display:flex;align-items:center;gap:8px}
        .header-disk .disk-value{font-weight:600;color:#4CAF50}
        
        .main-container{max-width:100%;margin:0 auto;padding:16px 20px}
        .page-title{font-size:22px;font-weight:600;margin-bottom:4px}
        .page-subtitle{color:var(--text2);font-size:13px;margin-bottom:16px}
        
        /* ===== УРОВЕНЬ 1: ГЛАВНЫЕ КАТЕГОРИИ ===== */
        .main-cat-bar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
        .main-cat-btn{padding:12px 24px;border:2px solid var(--border);border-radius:var(--radius);cursor:pointer;font-size:15px;font-weight:500;background:var(--card);color:var(--text);transition:all .2s;display:flex;align-items:center;gap:8px;box-shadow:var(--shadow)}
        .main-cat-btn:hover{border-color:var(--primary);background:#e3f2fd}
        .main-cat-btn.active{background:var(--success);color:#fff;border-color:var(--success);box-shadow:0 2px 8px rgba(76,175,80,.3)}
        .main-cat-btn.excel-active{background:#E91E63;color:#fff;border-color:#E91E63}
        .main-cat-count{font-size:12px;opacity:.8;background:rgba(0,0,0,.1);padding:2px 8px;border-radius:10px;font-weight:600}
        .main-cat-btn.active .main-cat-count,.main-cat-btn.excel-active .main-cat-count{background:rgba(255,255,255,.2)}
        
        /* ===== УРОВЕНЬ 2: РЕЖИМЫ (скрываются для Excel) ===== */
        .mode-bar{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
        .mode-btn{padding:10px 20px;border:2px solid var(--border);border-radius:var(--radius);cursor:pointer;font-size:14px;font-weight:500;background:var(--card);color:var(--text);transition:all .2s;display:flex;align-items:center;gap:8px;box-shadow:var(--shadow)}
        .mode-btn:hover{border-color:var(--primary);background:#e3f2fd}
        .mode-btn.active{background:var(--primary);color:#fff;border-color:var(--primary)}
        
        .stats-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:10px;margin-bottom:20px}
        .stat-card{background:var(--card);padding:16px 12px;border-radius:var(--radius);box-shadow:var(--shadow);text-align:center;transition:transform .15s}
        .stat-card:hover{transform:translateY(-2px)}
        .stat-card .number{font-size:26px;font-weight:700;line-height:1;margin-bottom:4px}
        .stat-card .label{color:var(--text2);font-size:11px;text-transform:uppercase;letter-spacing:.5px}
        .stat-card.total .number{color:#2196F3}.stat-card.updated .number{color:#4CAF50}.stat-card.outdated .number{color:#ff9800}.stat-card.offline .number{color:#f44336}.stat-card.errors .number{color:#E91E63}.stat-card.lowdisk .number{color:#f44336}.stat-card.percent .number{color:#9C27B0}
        
        .section-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
        .section-title{font-size:16px;font-weight:600}
        .btn-refresh{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;background:var(--primary);color:#fff;transition:all .2s}
        .btn-refresh:hover{background:var(--phover)}
        
        .table-wrapper{overflow-x:auto;background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow)}
        table{width:100%;border-collapse:collapse;font-size:12px;min-width:1800px}
        th{background:var(--hbg);color:#fff;padding:10px 5px;text-align:center;font-weight:500;font-size:9px;text-transform:uppercase;letter-spacing:.3px;white-space:nowrap;cursor:pointer;user-select:none}
        th:hover{background:#2a2a4e}
        th.left{text-align:left}
        td{padding:6px 5px;border-bottom:1px solid #eee;text-align:center;font-size:11px;white-space:nowrap}
        td.left{text-align:left;font-weight:500}
        tr:hover td{background-color:#f8f9fa}
        tr:last-child td{border-bottom:none}
        
        .badge{padding:3px 8px;border-radius:12px;font-size:10px;font-weight:600;display:inline-block;white-space:nowrap}
        .badge-success{background:var(--slight);color:#155724}.badge-warning{background:var(--wlight);color:#856404}.badge-danger{background:var(--dlight);color:#721c24}.badge-info{background:#d1ecf1;color:#0c5460}.badge-excel{background:#fce4ec;color:#880e4f}
        
        .match-bar{display:inline-block;width:38px;height:4px;background:#e0e0e0;border-radius:2px;vertical-align:middle;margin-right:3px}
        .match-bar-fill{height:4px;border-radius:2px}.match-high{background:#4CAF50}.match-mid{background:#ff9800}.match-low{background:#f44336}
        
        /* ===== EXCEL TABS ===== */
        .excel-tabs{display:flex;gap:4px;margin-bottom:0;flex-wrap:wrap;padding:8px 12px;background:#f5f5f5;border-radius:var(--radius) var(--radius) 0 0}
        .excel-tab{padding:8px 16px;border:none;border-radius:6px 6px 0 0;cursor:pointer;font-size:12px;font-weight:500;background:#e0e0e0;color:var(--text);transition:all .2s}
        .excel-tab:hover{background:#d0d0d0}
        .excel-tab.active{background:var(--card);color:var(--primary);font-weight:600;box-shadow:0 -2px 4px rgba(0,0,0,.05)}
        
        .footer{text-align:center;color:#999;font-size:11px;padding:20px;margin-top:20px}
    </style>
</head>
<body>

<div class="header">
    <div class="header-left">
        <span class="header-logo">📊 Мониторинг ПК</span>
        <span class="header-version">v1.4</span>
    </div>
    <div class="header-right">
        <span class="header-time" id="currentTime">—</span>
        <div class="header-disk">
            <span>💾 Сервер:</span>
            <span class="disk-value" id="serverDiskValue">540 / 1000 GB</span>
        </div>
    </div>
</div>

<div class="main-container">
    <h1 class="page-title">Мониторинг ПК</h1>
    <p class="page-subtitle">Выберите категорию. Для Excel отображаются листы книги.</p>

    <!-- УРОВЕНЬ 1: ГЛАВНЫЕ КАТЕГОРИИ -->
    <div class="main-cat-bar">
        <button class="main-cat-btn active" onclick="selectMainCategory('server', this)">🖥 Сервер <span class="main-cat-count">5</span></button>
        <button class="main-cat-btn" onclick="selectMainCategory('kassa', this)">🏪 Касса <span class="main-cat-count">42</span></button>
        <button class="main-cat-btn" onclick="selectMainCategory('mark', this)">📦 Маркировка <span class="main-cat-count">3</span></button>
        <button class="main-cat-btn" onclick="selectMainCategory('office', this)">💼 Офис <span class="main-cat-count">0</span></button>
        <button class="main-cat-btn" onclick="selectMainCategory('excel', this)">📊 Excel Таблица <span class="main-cat-count">▼</span></button>
    </div>

    <!-- УРОВЕНЬ 2: РЕЖИМЫ (скрыты для Excel) -->
    <div class="mode-bar" id="modeBar">
        <button class="mode-btn active" onclick="selectMode('sync', this)">🔄 Синхронизация</button>
        <button class="mode-btn" onclick="selectMode('pc', this)">📋 Состояние ПК</button>
    </div>

    <!-- EXCEL ЛИСТЫ (только для Excel) -->
    <div class="excel-tabs" id="excelTabs" style="display:none;">
        <button class="excel-tab active" onclick="selectExcelSheet('sync', this)">🔄 Синхронизация</button>
        <button class="excel-tab" onclick="selectExcelSheet('pc', this)">📋 Состояние ПК</button>
        <button class="excel-tab" onclick="selectExcelSheet('disk', this)">💾 Диски</button>
        <button class="excel-tab" onclick="selectExcelSheet('errors', this)">⚠ Ошибки</button>
    </div>

    <!-- СТАТИСТИКА -->
    <div class="stats-grid">
        <div class="stat-card total"><div class="number" id="statTotal">50</div><div class="label">Всего</div></div>
        <div class="stat-card updated"><div class="number" id="statUpdated">45</div><div class="label">✅ Актуальны</div></div>
        <div class="stat-card outdated"><div class="number" id="statOutdated">3</div><div class="label">⚠ Устарела версия</div></div>
        <div class="stat-card offline"><div class="number" id="statOffline">2</div><div class="label">❌ Нет связи</div></div>
        <div class="stat-card errors"><div class="number" id="statErrors">1</div><div class="label">С ошибками</div></div>
        <div class="stat-card lowdisk"><div class="number" id="statLowDisk">2</div><div class="label">⚠ Мало места</div></div>
        <div class="stat-card percent"><div class="number" id="statPercent">90%</div><div class="label">Актуальных</div></div>
    </div>

    <!-- ЗАГОЛОВОК ТАБЛИЦЫ -->
    <div class="section-header">
        <h2 class="section-title" id="tableTitle">🔄 Синхронизация — Сервер</h2>
        <button class="btn-refresh" onclick="location.reload()">🔄 Обновить</button>
    </div>

    <!-- ТАБЛИЦА -->
    <div class="table-wrapper" id="tableContainer"></div>

    <div class="footer">
        C:\Прокст сайта\monitoring.html | Обновлено: <span id="updateTime">—</span>
    </div>
</div>

<script>
    var currentMainCat = 'server';
    var currentMode = 'sync';
    var currentExcelSheet = 'sync';
    var isExcelMode = false;

    var syncData = [
        {cat:'server',name:'SRV-DC01',ip:'192.168.1.10',status:'success',stText:'Актуален',ver:'20260703.0949',verOK:true,sync:'03.07 07:15',when:'5 мин',files:'72/72',match:100,diskFree:120,diskTotal:500,diskModel:'Samsung SSD 980 1TB',diskHealth:'Good',diskHealthPct:93,diskTemp:44,diskHours:23036,cpu:'45%',cpuTemp:52,ramUsed:12,ramTotal:32,uptime:15,err:0},
        {cat:'kassa',name:'KASSA-01',ip:'192.168.1.101',status:'success',stText:'Актуален',ver:'20260703.0949',verOK:true,sync:'03.07 07:15',when:'5 мин',files:'72/72',match:100,diskFree:45,diskTotal:128,diskModel:'Samsung SSD 980 1TB',diskHealth:'Good',diskHealthPct:93,diskTemp:38,diskHours:12000,cpu:'30%',cpuTemp:44,ramUsed:4,ramTotal:8,uptime:7,err:0},
        {cat:'kassa',name:'KASSA-05',ip:'192.168.1.105',status:'warning',stText:'Устарела',ver:'20260702.2200',verOK:false,sync:'02.07 12:00',when:'18 ч',files:'68/72',match:94.4,diskFree:12,diskTotal:128,diskModel:'WD Green 240GB',diskHealth:'Caution',diskHealthPct:67,diskTemp:42,diskHours:34000,cpu:'55%',cpuTemp:58,ramUsed:6,ramTotal:8,uptime:3,err:2},
        {cat:'mark',name:'MARK-01',ip:'192.168.2.10',status:'danger',stText:'Нет связи',ver:'—',verOK:false,sync:'—',when:'3 дн',files:'—',match:0,diskFree:0,diskTotal:0,diskModel:'—',diskHealth:'—',diskHealthPct:0,diskTemp:0,diskHours:0,cpu:'—',cpuTemp:0,ramUsed:0,ramTotal:0,uptime:0,err:0},
        {cat:'office',name:'OFFICE-03',ip:'192.168.3.15',status:'success',stText:'Актуален',ver:'20260703.0949',verOK:true,sync:'03.07 08:00',when:'12 мин',files:'72/72',match:100,diskFree:200,diskTotal:500,diskModel:'Kingston A400 480GB',diskHealth:'Good',diskHealthPct:85,diskTemp:35,diskHours:8000,cpu:'20%',cpuTemp:38,ramUsed:8,ramTotal:16,uptime:10,err:0}
    ];

    function updateClock() {
        var now = new Date();
        document.getElementById('currentTime').textContent = now.toLocaleString('ru-RU');
        document.getElementById('updateTime').textContent = now.toLocaleString('ru-RU');
    }
    updateClock(); setInterval(updateClock, 10000);

    function getMatchBar(pct) {
        var cls = pct >= 95 ? 'match-high' : (pct >= 70 ? 'match-mid' : 'match-low');
        return '<div class="match-bar"><div class="match-bar-fill '+cls+'" style="width:'+Math.max(5,pct)+'%"></div></div> '+pct+'%';
    }
    function getDiskBar(free, total) {
        if(total===0) return '—';
        var pct = Math.round((free/total)*100);
        var cls = pct>25?'match-high':(pct>10?'match-mid':'match-low');
        return '<div class="match-bar"><div class="match-bar-fill '+cls+'" style="width:'+Math.max(5,pct)+'%"></div></div> '+free+' / '+total+' GB';
    }
    function getRamBar(used, total) {
        if(total===0) return '—';
        var pct = Math.round((used/total)*100);
        var cls = pct>90?'match-low':(pct>60?'match-mid':'match-high');
        return '<div class="match-bar"><div class="match-bar-fill '+cls+'" style="width:'+Math.max(5,pct)+'%"></div></div> '+used+' / '+total+' GB';
    }

    function renderTable() {
        var container = document.getElementById('tableContainer');
        var filtered = syncData.filter(function(r){ return currentMainCat==='all' || r.cat===currentMainCat || isExcelMode; });
        if(isExcelMode) filtered = syncData;

        var html = '<table><thead><tr>';
        var mode = isExcelMode ? currentExcelSheet : currentMode;

        if(mode==='sync'){
            html += '<th>№</th><th class="left">Имя хоста</th><th>IP</th><th>Категория</th><th>Ver</th><th>Статус</th><th>Дата синхр.</th><th>Когда</th><th>Файлы</th><th>Актуал.</th><th>Диск</th><th>Модель диска</th><th>SMART</th><th>Ошибок</th>';
        } else if(mode==='pc'){
            html += '<th>№</th><th class="left">Имя хоста</th><th>IP</th><th>Категория</th><th>CPU</th><th>CPU t°</th><th>RAM</th><th>Диск</th><th>Модель диска</th><th>SMART</th><th>t° диска</th><th>Наработка</th><th>Uptime</th>';
        } else if(mode==='disk'){
            html += '<th>№</th><th class="left">Имя хоста</th><th>Категория</th><th>Диск</th><th>Модель диска</th><th>SMART</th><th>t° диска</th><th>Наработка</th>';
        } else if(mode==='errors'){
            html += '<th>№</th><th class="left">Имя хоста</th><th>Категория</th><th>Статус</th><th>Ошибок</th><th>Дата синхр.</th>';
        }
        html += '</tr></thead><tbody>';

        var updated=0, outdated=0, offline=0, errors=0;

        for(var i=0; i<filtered.length; i++){
            var r = filtered[i];
            html += '<tr>';

            if(mode==='sync'){
                html += '<td>'+(i+1)+'</td><td class="left">'+r.name+'</td><td>'+r.ip+'</td>';
                html += '<td><span class="badge '+(isExcelMode?'badge-excel':'badge-info')+'">'+(r.cat==='server'?'Сервер':r.cat==='kassa'?'Касса':r.cat==='mark'?'Маркировка':'Офис')+'</span></td>';
                var vc = r.verOK?'badge-success':'badge-warning';
                html += '<td><span class="badge '+vc+'">'+r.ver+'</span></td>';
                var sc = r.status==='success'?'badge-success':(r.status==='warning'?'badge-warning':'badge-danger');
                var si = r.status==='success'?'✅ ':(r.status==='warning'?'⚠ ':'❌ ');
                html += '<td><span class="badge '+sc+'">'+si+r.stText+'</span></td>';
                html += '<td>'+r.sync+'</td><td>'+r.when+'</td><td>'+r.files+'</td>';
                html += '<td>'+getMatchBar(r.match)+'</td>';
                html += '<td>'+getDiskBar(r.diskFree, r.diskTotal)+'</td>';
                html += '<td style="font-size:10px;">'+r.diskModel+'</td>';
                var hc = r.diskHealth==='Good'?'badge-success':(r.diskHealth==='Caution'?'badge-warning':'badge-danger');
                html += '<td><span class="badge '+hc+'">'+r.diskHealth+' ('+r.diskHealthPct+'%)</span></td>';
                html += '<td>'+(r.err>0?'<span class="badge badge-danger">'+r.err+'</span>':'0')+'</td>';
                if(r.status==='success') updated++;
                else if(r.status==='warning') outdated++;
                else if(r.status==='danger') offline++;
                if(r.err>0) errors++;
            } else if(mode==='pc'){
                html += '<td>'+(i+1)+'</td><td class="left">'+r.name+'</td><td>'+r.ip+'</td>';
                html += '<td><span class="badge '+(isExcelMode?'badge-excel':'badge-info')+'">'+(r.cat==='server'?'Сервер':r.cat==='kassa'?'Касса':r.cat==='mark'?'Маркировка':'Офис')+'</span></td>';
                html += '<td>'+r.cpu+'</td>';
                html += '<td style="color:'+(r.cpuTemp>60?'#f44336':r.cpuTemp>40?'#ff9800':'#4CAF50')+'">'+(r.cpuTemp>0?r.cpuTemp+'°C':'—')+'</td>';
                html += '<td>'+getRamBar(r.ramUsed, r.ramTotal)+'</td>';
                html += '<td>'+getDiskBar(r.diskFree, r.diskTotal)+'</td>';
                html += '<td style="font-size:10px;">'+r.diskModel+'</td>';
                var hc = r.diskHealth==='Good'?'badge-success':(r.diskHealth==='Caution'?'badge-warning':'badge-danger');
                html += '<td><span class="badge '+hc+'">'+r.diskHealth+' ('+r.diskHealthPct+'%)</span></td>';
                html += '<td style="color:'+(r.diskTemp>50?'#f44336':r.diskTemp>40?'#ff9800':'#4CAF50')+'">'+(r.diskTemp>0?r.diskTemp+'°C':'—')+'</td>';
                html += '<td>'+(r.diskHours>0?r.diskHours.toLocaleString()+' ч':'—')+'</td>';
                html += '<td>'+(r.uptime>0?r.uptime+' дн':'—')+'</td>';
            } else if(mode==='disk'){
                html += '<td>'+(i+1)+'</td><td class="left">'+r.name+'</td>';
                html += '<td><span class="badge '+(isExcelMode?'badge-excel':'badge-info')+'">'+(r.cat==='server'?'Сервер':r.cat==='kassa'?'Касса':r.cat==='mark'?'Маркировка':'Офис')+'</span></td>';
                html += '<td>'+getDiskBar(r.diskFree, r.diskTotal)+'</td>';
                html += '<td style="font-size:10px;">'+r.diskModel+'</td>';
                var hc = r.diskHealth==='Good'?'badge-success':(r.diskHealth==='Caution'?'badge-warning':'badge-danger');
                html += '<td><span class="badge '+hc+'">'+r.diskHealth+' ('+r.diskHealthPct+'%)</span></td>';
                html += '<td style="color:'+(r.diskTemp>50?'#f44336':r.diskTemp>40?'#ff9800':'#4CAF50')+'">'+(r.diskTemp>0?r.diskTemp+'°C':'—')+'</td>';
                html += '<td>'+(r.diskHours>0?r.diskHours.toLocaleString()+' ч':'—')+'</td>';
            } else if(mode==='errors'){
                html += '<td>'+(i+1)+'</td><td class="left">'+r.name+'</td>';
                html += '<td><span class="badge '+(isExcelMode?'badge-excel':'badge-info')+'">'+(r.cat==='server'?'Сервер':r.cat==='kassa'?'Касса':r.cat==='mark'?'Маркировка':'Офис')+'</span></td>';
                var sc = r.status==='success'?'badge-success':(r.status==='warning'?'badge-warning':'badge-danger');
                var si = r.status==='success'?'✅ ':(r.status==='warning'?'⚠ ':'❌ ');
                html += '<td><span class="badge '+sc+'">'+si+r.stText+'</span></td>';
                html += '<td>'+(r.err>0?'<span class="badge badge-danger">'+r.err+'</span>':'0')+'</td>';
                html += '<td>'+r.sync+'</td>';
                if(r.err>0) errors++;
            }
            html += '</tr>';
        }
        html += '</tbody></table>';
        container.innerHTML = html;

        var total = filtered.length;
        document.getElementById('statTotal').textContent = total;
        if(mode==='sync'){
            document.getElementById('statUpdated').textContent = updated;
            document.getElementById('statOutdated').textContent = outdated;
            document.getElementById('statOffline').textContent = offline;
            document.getElementById('statErrors').textContent = errors;
            document.getElementById('statPercent').textContent = total>0?Math.round((updated/total)*100)+'%':'—';
        }
    }

    function updateTitle(){
        var catName = currentMainCat==='server'?'Сервер':currentMainCat==='kassa'?'Касса':currentMainCat==='mark'?'Маркировка':currentMainCat==='office'?'Офис':'Excel';
        var modeName = isExcelMode ? (currentExcelSheet==='sync'?'Синхронизация':currentExcelSheet==='pc'?'Состояние ПК':currentExcelSheet==='disk'?'Диски':'Ошибки') : (currentMode==='sync'?'Синхронизация':'Состояние ПК');
        document.getElementById('tableTitle').textContent = (isExcelMode?'📊 ':'') + modeName + ' — ' + catName;
    }

    function selectMainCategory(cat, btn){
        currentMainCat = cat;
        isExcelMode = (cat === 'excel');
        
        document.querySelectorAll('.main-cat-btn').forEach(function(b){b.classList.remove('active','excel-active')});
        if(isExcelMode) btn.classList.add('excel-active');
        else btn.classList.add('active');
        
        document.getElementById('modeBar').style.display = isExcelMode ? 'none' : '';
        document.getElementById('excelTabs').style.display = isExcelMode ? '' : 'none';
        
        updateTitle();
        renderTable();
    }

    function selectMode(mode, btn){
        currentMode = mode;
        document.querySelectorAll('.mode-btn').forEach(function(b){b.classList.remove('active')});
        btn.classList.add('active');
        updateTitle();
        renderTable();
    }

    function selectExcelSheet(sheet, btn){
        currentExcelSheet = sheet;
        document.querySelectorAll('.excel-tab').forEach(function(b){b.classList.remove('active')});
        btn.classList.add('active');
        updateTitle();
        renderTable();
    }

    renderTable();
</script>

</body>
</html>
"@

try {
    $html | Set-Content $htmlFile -Encoding UTF8
    Write-Output "✅ Сайт создан: $htmlFile"
    Start-Process $htmlFile
} catch {
    Write-Output "❌ Ошибка: $_"
}