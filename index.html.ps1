$html = @"
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Мониторинг v7.5</title>
    <style>
        *{margin:0;padding:0;box-sizing:border-box}
        body{font-family:"Segoe UI",sans-serif;background:#f0f2f5;display:flex;flex-direction:column;height:100vh}
        .header{background:#1a1a2e;color:#fff;padding:0 20px;height:50px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
        .header-logo{font-size:18px;font-weight:700}.header-time{font-size:13px;opacity:.8}
        .main-tabs{display:flex;gap:4px;padding:10px 20px;background:#1a1a2e;flex-wrap:wrap;flex-shrink:0}
        .main-tab{padding:10px 20px;border:none;border-radius:8px 8px 0 0;cursor:pointer;font-size:14px;font-weight:500;background:#2a2a4e;color:#aaa;transition:.2s}
        .main-tab:hover{background:#3a3a5e;color:#fff}
        .main-tab.active{background:#f0f2f5;color:#1a1a2e;font-weight:600}
        .sub-tabs{display:flex;gap:4px;padding:8px 20px;background:#f0f2f5;border-bottom:1px solid #ddd;flex-wrap:wrap;flex-shrink:0}
        .sub-tab{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;background:#e8e8e8;color:#555;transition:.2s}
        .sub-tab:hover{background:#d0d0d0;color:#333}
        .sub-tab.active{background:#1976D2;color:#fff}
        .sub-tab.active.reklama{background:#E91E63;color:#fff}
        .frame-container{flex:1;overflow:hidden}
        iframe{width:100%;height:100%;border:none}
    </style>
</head>
<body>

<div class="header">
    <span class="header-logo">📊 Мониторинг</span>
    <span class="header-time" id="clock">—</span>
</div>

<div class="main-tabs">
    <button class="main-tab" onclick="selectMain('server')">🖥 Сервер</button>
    <button class="main-tab active" onclick="selectMain('kassa')">🏪 Касса</button>
    <button class="main-tab" onclick="selectMain('mark')">📦 Маркировка</button>
    <button class="main-tab" onclick="selectMain('reklama')">📢 Реклама</button>
    <button class="main-tab" onclick="selectMain('excel')">📊 Excel Таблица</button>
</div>

<div class="sub-tabs" id="subTabs"></div>

<div class="frame-container">
    <iframe id="contentFrame" src="summary.html"></iframe>
</div>

<script>
var currentMain = 'kassa';
var currentSub = 'sync';
var mainOrder = ['server','kassa','mark','reklama','excel'];

var subMenus = {
    'server': [
        {name:'📋 Характеристики', id:'specs', src:''},
        {name:'🔄 Синхронизация', id:'sync', src:'summary.html'}
    ],
    'kassa': [
        {name:'📋 Характеристики', id:'specs', src:''},
        {name:'🔄 Синхронизация', id:'sync', src:'summary.html'}
    ],
    'mark': [
        {name:'📋 Характеристики', id:'specs', src:''},
        {name:'🔄 Синхронизация', id:'sync', src:'summary.html'}
    ],
    'reklama': [
        {name:'📋 Характеристики', id:'specs', src:''},
        {name:'🔄 Синхронизация', id:'sync', src:'summary.html'},
        {name:'📢 Актуальность рекламы', id:'ads', src:''}
    ],
    'excel': [
        {name:'📊 Все листы', id:'sheets', src:'monitoring.html'}
    ]
};

function selectMain(main){
    currentMain = main;
    document.querySelectorAll('.main-tab').forEach(function(b){b.classList.remove('active')});
    document.querySelectorAll('.main-tab')[mainOrder.indexOf(main)].classList.add('active');
    renderSubTabs();
    var firstSub = subMenus[main][0];
    currentSub = firstSub.id;
    renderSubTabs();
    loadContent(firstSub);
}

function renderSubTabs(){
    var container = document.getElementById('subTabs');
    var subs = subMenus[currentMain] || [];
    container.innerHTML = '';
    subs.forEach(function(sub){
        var btn = document.createElement('button');
        btn.className = 'sub-tab';
        if(currentMain === 'reklama' && sub.id === 'ads') btn.classList.add('reklama');
        if(sub.id === currentSub) btn.classList.add('active');
        btn.textContent = sub.name;
        btn.onclick = function(){ 
            currentSub = sub.id;
            renderSubTabs();
            loadContent(sub);
        };
        container.appendChild(btn);
    });
}

function loadContent(sub){
    var frame = document.getElementById('contentFrame');
    if(sub.src){
        frame.src = sub.src;
        frame.style.display = '';
    } else {
        frame.style.display = 'none';
        // Показываем заглушку через frame (пустой about:blank с сообщением)
        frame.src = 'about:blank';
        frame.style.display = '';
        setTimeout(function(){
            var doc = frame.contentDocument || frame.contentWindow.document;
            doc.open();
            doc.write('<html><body style="display:flex;align-items:center;justify-content:center;height:100vh;font-family:Segoe UI;color:#999;font-size:18px;background:#fff;margin:0">📋 <b>'+sub.name+'</b> — раздел в разработке</body></html>');
            doc.close();
        }, 100);
    }
}

function updateClock(){document.getElementById('clock').textContent = new Date().toLocaleString('ru-RU')}
updateClock();setInterval(updateClock,10000);

// Старт — Касса → Синхронизация
selectMain('kassa');
</script>

</body>
</html>
"@

$html | Set-Content "C:\Прокст сайта\index.html" -Encoding UTF8
Write-Output "✅ index.html обновлён"
Start-Process "C:\Прокст сайта\index.html"