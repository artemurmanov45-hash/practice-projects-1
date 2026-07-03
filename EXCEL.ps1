# ========== ExcelSyn v6.1 (ТОЛЬКО EXCEL ТАБЛИЦА) ==========
$outputPath = "C:\Прокст сайта\Excel"
$htmlFile = Join-Path -Path $outputPath -ChildPath "monitoring.html"
$jsonOutputPath = Join-Path -Path $outputPath -ChildPath "data.json"
$configEncrypted = Join-Path -Path $outputPath -ChildPath "config.enc"
$configPlain = Join-Path -Path $outputPath -ChildPath "config.txt"
$csvTempPath = Join-Path -Path $env:TEMP -ChildPath "google_temp.csv"

$sheets = @(
    @{Name="Кассы ПК";         Gid=1906770819},
    @{Name="Инфо о аптеках";  Gid=2001303443},
    @{Name="Сервер ПК";       Gid=690415603},
    @{Name="Провайдеры";      Gid=662816165},
    @{Name="Выездные работы"; Gid=1919342825},
    @{Name="Маркировка ПК";   Gid=493644987},
    @{Name="СетьАптек";       Gid=687285541},
    @{Name="Терминалы";       Gid=413500129},
    @{Name="Принтеры";        Gid=1838643814},
    @{Name="Реклама ПК";      Gid=1069624634},
    @{Name="Стикеры";         Gid=138289406},
    @{Name="Справочник";      Gid=1887472056}
)

if (-not (Test-Path $outputPath)) { New-Item -ItemType Directory -Path $outputPath -Force | Out-Null }

function Get-GoogleSheetUrl {
    if (Test-Path $configEncrypted) {
        try {
            $secure = Import-Clixml -Path $configEncrypted -ErrorAction Stop
            $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
        } catch { return $null }
    }
    if (Test-Path $configPlain) {
        $url = (Get-Content $configPlain -Raw).Trim()
        if ($url -and $url -ne "ССЫЛКА_СЮДА") {
            $secure = ConvertTo-SecureString $url -AsPlainText -Force
            $secure | Export-Clixml -Path $configEncrypted -Force
            return $url
        }
    }
    return $null
}

if ($args[0] -eq "-encrypt") {
    if (Test-Path $configPlain) {
        $url = (Get-Content $configPlain -Raw).Trim()
        $secure = ConvertTo-SecureString $url -AsPlainText -Force
        $secure | Export-Clixml -Path $configEncrypted -Force
        Write-Output "OK: config.enc created"
    }
    exit 0
}

Write-Output "ExcelSyn v6.1 — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

$baseUrl = Get-GoogleSheetUrl
if (-not $baseUrl) { Write-Output "ERROR: нет config"; exit 1 }
$baseUrl = $baseUrl -replace '\?.*$', ''

$allSheets = @()
foreach ($sheet in $sheets) {
    $url = "$baseUrl`?gid=$($sheet.Gid)&single=true&output=csv"
    try {
        Invoke-WebRequest -Uri $url -OutFile $csvTempPath -UseBasicParsing -ErrorAction Stop
        $csvContent = Get-Content $csvTempPath -Raw -Encoding UTF8
        $lines = ($csvContent -split "`n") | Where-Object { $_ -notmatch '^\s*$' }
        if ($lines.Count -lt 2) { Write-Output "  $($sheet.Name): пусто"; continue }
        $headers = $lines[0] -split ','
        for ($h = 0; $h -lt $headers.Count; $h++) { if ([string]::IsNullOrWhiteSpace($headers[$h])) { $headers[$h] = "Col$($h+1)" } }
        $csvData = @()
        for ($l = 1; $l -lt $lines.Count; $l++) {
            $values = $lines[$l] -split ','
            $row = [ordered]@{}
            for ($v = 0; $v -lt [Math]::Min($headers.Count, $values.Count); $v++) { $row[$headers[$v]] = $values[$v] -replace '^"|"$', '' }
            $csvData += [PSCustomObject]$row
        }
        $allSheets += @{ SheetName = $sheet.Name; RowCount = $csvData.Count; Data = $csvData }
        Write-Output "  $($sheet.Name): $($csvData.Count) строк"
    } catch { Write-Output "  $($sheet.Name): ОШИБКА" }
}
Remove-Item $csvTempPath -Force -ErrorAction SilentlyContinue
$allSheets | ConvertTo-Json -Depth 5 -Compress | Set-Content $jsonOutputPath -Encoding UTF8
Write-Output "Сохранено: $($allSheets.Count) листов"

if ($allSheets.Count -eq 0) { Write-Output "Нет данных"; exit 1 }

$allDataJson = $allSheets | ConvertTo-Json -Depth 5 -Compress

$style = @'
<style>
*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',system-ui,sans-serif;background:#f0f2f5;color:#1a1a2e;min-height:100vh}
.header{background:#1a1a2e;color:#fff;padding:0 24px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.header-left{display:flex;align-items:center;gap:12px}.header-logo{font-size:20px;font-weight:700}.header-version{font-size:11px;opacity:.6;background:rgba(255,255,255,.1);padding:3px 8px;border-radius:12px}
.header-right{display:flex;align-items:center;gap:20px}.header-time{font-size:14px;opacity:.8}
.main-container{max-width:100%;margin:0 auto;padding:16px 20px}
.sub-tabs{display:flex;gap:4px;margin-bottom:12px;padding:10px 12px;background:#fff;border-radius:10px;box-shadow:0 2px 4px rgba(0,0,0,.08);overflow-x:auto;flex-wrap:wrap}
.sub-tab{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:500;background:#e8e8e8;color:#1a1a2e;transition:.2s;white-space:nowrap}
.sub-tab:hover{background:#d0d0d0}.sub-tab.active{background:#1976D2;color:#fff}
.sub-tab .tab-count{font-size:10px;opacity:.7;margin-left:4px}
.toggle-panel{display:flex;align-items:center;gap:8px;margin-bottom:12px;flex-wrap:wrap}
.toggle-btn{padding:6px 14px;border:1px solid #ccc;border-radius:6px;cursor:pointer;font-size:12px;background:#fff;transition:.2s}
.toggle-btn:hover{background:#e3f2fd}
.toggle-btn.active{background:#1976D2;color:#fff}
.toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin:12px 0;padding:10px 14px;background:#fff;border-radius:10px;box-shadow:0 2px 4px rgba(0,0,0,.08)}
.search-input{padding:8px 14px;border:2px solid #d0d0d0;border-radius:6px;font-size:13px;width:250px;outline:none}.search-input:focus{border-color:#1976D2}
.filter-select{padding:8px 14px;border:2px solid #d0d0d0;border-radius:6px;font-size:13px;background:#fff;cursor:pointer}
.result-count{color:#666;font-size:13px;margin-left:auto}
.btn-refresh{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;background:#1976D2;color:#fff}
.btn-reset-width{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:500;background:#666;color:#fff}
.table-outer{background:#fff;border-radius:10px;box-shadow:0 2px 4px rgba(0,0,0,.08);overflow:hidden}
.table-scroll{overflow-x:auto;overflow-y:auto;max-height:65vh;border-radius:10px}
.table-scroll::-webkit-scrollbar{height:12px;width:12px}.table-scroll::-webkit-scrollbar-track{background:#f1f1f1;border-radius:6px}.table-scroll::-webkit-scrollbar-thumb{background:#bbb;border-radius:6px;border:3px solid #f1f1f1}
table{width:max-content;min-width:100%;border-collapse:collapse;font-size:12px}
thead{position:sticky;top:0;z-index:20}
th{background:#1a1a2e;color:#fff;padding:10px 30px 10px 10px;text-align:left;font-weight:500;font-size:10px;text-transform:uppercase;white-space:nowrap;border:1px solid #3a3a5e;cursor:pointer;user-select:none;position:relative;overflow:visible}
th .resize-handle{position:absolute;right:0;top:0;bottom:0;width:8px;cursor:col-resize;background:transparent;z-index:30}
th .resize-handle:hover{background:rgba(255,255,255,0.4)}
th .resize-handle.active{background:rgba(255,255,255,0.6)}
th.drag-over{border-left:3px solid #4CAF50!important}
th.dragging{opacity:0.5;background:#4a4a6e}
th:hover{background:#2a2a4e}th .sort-icon{font-size:10px;margin-left:4px;opacity:.5}th.sorted .sort-icon{opacity:1}
th:first-child{text-align:center;width:40px;position:sticky;left:0;z-index:22;background:#222240;cursor:default;padding:10px}
td{padding:7px 10px;text-align:left;font-size:12px;white-space:nowrap;border:1px solid #d0d0d0;overflow:hidden;text-overflow:ellipsis}
td:first-child{text-align:center;font-weight:500;color:#666;background:#fafafa;position:sticky;left:0;z-index:5;border-right:2px solid #ccc}
tr:hover td{background-color:#e3f2fd!important}tr:hover td:first-child{background-color:#bbdefb!important}
tr:nth-child(even) td{background-color:#fafafa}tr:nth-child(even) td:first-child{background-color:#f0f0f0}
.placeholder{text-align:center;color:#999;padding:60px;font-size:16px}
.footer{text-align:center;color:#999;font-size:11px;padding:20px;margin-top:20px}
</style>
'@

$script = @'
<script>
var allData = __ALL_DATA__;
var currentSheet=-1,currentData=[],currentHeaders=[],sortCol=-1,sortDir='asc';
var columnWidths = {};
var hiddenColumns = {};
var dragSrcIndex = null;

function updateClock(){var n=new Date();document.getElementById('currentTime').textContent=n.toLocaleString('ru-RU');document.getElementById('updateTime').textContent=n.toLocaleString('ru-RU')}
updateClock();setInterval(updateClock,10000);

function renderSubTabs(){
    var t=document.getElementById('subTabs');t.innerHTML='';
    allData.forEach(function(s,i){var b=document.createElement('button');b.className='sub-tab';if(i===currentSheet)b.classList.add('active');b.innerHTML='📄 '+s.SheetName+'<span class="tab-count">('+s.RowCount+')</span>';b.onclick=function(){selectSheet(i)};t.appendChild(b)})
}

function selectSheet(i){
    currentSheet=i;renderSubTabs();var s=allData[i];
    currentData=s.Data||[];if(currentData.length===0){document.getElementById('tableContainer').innerHTML='<div class="placeholder">Нет данных</div>';return}
    currentHeaders=Object.keys(currentData[0]);
    var fs=document.getElementById('filterColumn');fs.innerHTML='<option value="">Все колонки</option>';
    currentHeaders.forEach(function(h){fs.innerHTML+='<option value="'+h+'">'+h+'</option>'});
    sortCol=-1;document.getElementById('searchInput').value='';columnWidths={};hiddenColumns={};
    renderToggleButtons();
    renderTable(currentData);
}

function renderToggleButtons(){
    var panel = document.getElementById('togglePanel');
    panel.innerHTML = '<span style="font-size:12px;color:#666;margin-right:4px;">Скрыть колонки:</span>';
    currentHeaders.forEach(function(h){
        var btn = document.createElement('button');
        btn.className = 'toggle-btn' + (hiddenColumns[h] ? ' active' : '');
        btn.textContent = h.length > 25 ? h.substring(0,23)+'…' : h;
        btn.title = h;
        btn.onclick = function(){ hiddenColumns[h]=!hiddenColumns[h]; btn.classList.toggle('active'); renderTable(currentData); };
        panel.appendChild(btn);
    });
}

function sortData(c){
    if(sortCol===c){sortDir=sortDir==='asc'?'desc':'asc'}else{sortCol=c;sortDir='asc'}
    var h=currentHeaders[c];
    currentData.sort(function(a,b){
        var va=(a[h]||'').toString().toLowerCase(),vb=(b[h]||'').toString().toLowerCase();
        var na=parseFloat(va),nb=parseFloat(vb);
        if(!isNaN(na)&&!isNaN(nb)&&va===na.toString()&&vb===nb.toString()){va=na;vb=nb}
        if(va<vb)return sortDir==='asc'?-1:1;
        if(va>vb)return sortDir==='asc'?1:-1;
        return 0;
    });
    renderTable(currentData);
}

function applyFilters(){
    if(currentData.length===0)return;
    var q=document.getElementById('searchInput').value.toLowerCase(),fc=document.getElementById('filterColumn').value;
    var f=currentData.filter(function(r){
        if(fc){return(r[fc]||'').toString().toLowerCase().indexOf(q)>=0}
        else{for(var k in r){if((r[k]||'').toString().toLowerCase().indexOf(q)>=0)return true}return false}
    });
    renderTable(f);
}

function resetWidths(){columnWidths={};renderTable(currentData)}

function renderTable(d){
    var visibleHeaders = currentHeaders.filter(function(h){ return !hiddenColumns[h]; });
    var h='<table><thead><tr><th style="width:40px">№</th>';
    visibleHeaders.forEach(function(hdr,i){
        var realIndex = currentHeaders.indexOf(hdr);
        var w = columnWidths[hdr] ? ' style="width:'+columnWidths[hdr]+'px"' : '';
        h+='<th'+w+' draggable="true" ondragstart="dragStart(event,'+realIndex+')" ondragover="dragOver(event)" ondrop="drop(event,'+realIndex+')" ondragend="dragEnd()" onclick="sortData('+realIndex+')" class="'+(sortCol===realIndex?'sorted':'')+'">'+hdr+' <span class="sort-icon">'+(sortCol===realIndex?(sortDir==='asc'?'▲':'▼'):'⇅')+'</span><div class="resize-handle" onmousedown="event.stopPropagation();startResize(event,'+realIndex+')"></div></th>';
    });
    h+='</tr></thead><tbody>';
    d.forEach(function(r,i){
        h+='<tr><td>'+(i+1)+'</td>';
        visibleHeaders.forEach(function(hdr){ h+='<td>'+(r[hdr]||'')+'</td>'; });
        h+='</tr>';
    });
    h+='</tbody></table>';
    document.getElementById('tableContainer').innerHTML=h;
    document.getElementById('resultCount').textContent='Показано: '+d.length+' / '+currentData.length;
}

function dragStart(e, index){dragSrcIndex=index;e.target.classList.add('dragging');e.dataTransfer.effectAllowed='move';e.dataTransfer.setData('text/plain',index)}
function dragOver(e){e.preventDefault();e.dataTransfer.dropEffect='move';var th=e.target.closest('th');if(th)th.classList.add('drag-over')}
function drop(e, targetIndex){
    e.preventDefault();document.querySelectorAll('th').forEach(function(th){th.classList.remove('drag-over')});
    if(dragSrcIndex===null||dragSrcIndex===targetIndex)return;
    var moved=currentHeaders.splice(dragSrcIndex,1)[0];currentHeaders.splice(targetIndex,0,moved);
    currentData.forEach(function(row){var keys=Object.keys(row);var vals=keys.map(function(k){return row[k]});var movedVal=vals.splice(dragSrcIndex,1)[0];vals.splice(targetIndex,0,movedVal);keys.forEach(function(k,i){row[k]=vals[i]})});
    sortCol=-1;renderTable(currentData);renderToggleButtons();
}
function dragEnd(){document.querySelectorAll('th').forEach(function(th){th.classList.remove('dragging','drag-over')});dragSrcIndex=null}

var resizing=null,startX,resizeIndex;
function startResize(e,index){e.preventDefault();e.stopPropagation();resizing=e.target;resizeIndex=index;startX=e.clientX;var th=resizing.parentElement;th.style.width=th.offsetWidth+'px';resizing.classList.add('active');document.addEventListener('mousemove',onResize);document.addEventListener('mouseup',stopResize)}
function onResize(e){if(!resizing)return;var th=resizing.parentElement;var cw=parseInt(th.style.width)||th.offsetWidth;var diff=e.clientX-startX;var nw=cw+diff;if(nw<5)nw=5;th.style.width=nw+'px';columnWidths[currentHeaders[resizeIndex]]=nw;startX=e.clientX}
function stopResize(){if(resizing)resizing.classList.remove('active');resizing=null;document.removeEventListener('mousemove',onResize);document.removeEventListener('mouseup',stopResize)}

renderSubTabs();
if(allData.length>0) selectSheet(0);
</script>
'@

$script = $script -replace '__ALL_DATA__', $allDataJson

$html = @"
<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8"><title>Excel Таблица v6.1</title>$style</head>
<body>
<div class="header"><div class="header-left"><span class="header-logo">📊 Excel Таблица</span><span class="header-version">v6.1</span></div><div class="header-right"><span class="header-time" id="currentTime">—</span></div></div>
<div class="main-container">
<div class="sub-tabs" id="subTabs"></div>
<div class="toggle-panel" id="togglePanel"></div>
<div class="toolbar" id="toolbar">
    <input type="text" class="search-input" id="searchInput" placeholder="🔍 Поиск..." oninput="applyFilters()">
    <select class="filter-select" id="filterColumn" onchange="applyFilters()"><option value="">Все колонки</option></select>
    <span class="result-count" id="resultCount"></span>
    <button class="btn-reset-width" onclick="resetWidths()">↔ Сброс ширины</button>
    <button class="btn-refresh" onclick="location.reload()">🔄 Обновить</button>
</div>
<div class="table-outer"><div class="table-scroll" id="tableContainer"><div class="placeholder">👆 Выберите лист</div></div></div>
<div class="footer">C:\Прокст сайта\Excel\monitoring.html | Обновлено: <span id="updateTime">—</span></div>
</div>
$script
</body>
</html>
"@

try { 
    $html | Set-Content $htmlFile -Encoding UTF8
    Write-Output "HTML: $htmlFile"
    Start-Process $htmlFile
    Write-Output "ГОТОВО"
} catch { Write-Output "ОШИБКА: $_" }