window.App=(()=>{
  const routes={
    dashboard:{title:'Dashboard',render:()=>Dashboard.render()},
    styles:{title:'Style Management',render:()=>SheetTable.render('styles')},
    requests:{title:'Sample Request Management',render:()=>SheetTable.render('requests')},
    monitoring:{title:'Sample Room Monitoring',render:()=>SheetTable.render('monitoring')},
    quality:{title:'Data Quality',render:renderQuality},
    export:{title:'Export',render:()=>ExportView.render()},
    import:{title:'Import workbook',render:()=>ImportWizard.render()}
  };
  let route='dashboard';
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  function navigate(next){
    if(!routes[next])return;
    route=next;
    document.querySelectorAll('.nav-item').forEach(button=>button.classList.toggle('active',button.dataset.route===next));
    document.getElementById('breadcrumbs').textContent=`AOI Materials  /  ${routes[next].title}`;
    document.getElementById('sidebar').classList.remove('open');
    const badge=document.querySelector('.demo-badge');
    badge.lastChild.textContent=AppState.data.sheets.length?'Workspace active':'Workspace empty';
    routes[next].render();
    updateCreateButton();
    window.scrollTo(0,0);
  }
  function updateCreateButton(){
    const create=document.getElementById('createRecord');
    create.disabled=!AppState.data.sheets.some(sheet=>sheet.columnCount>0);
  }
  function activeSheetIndex(){
    if(['styles','requests','monitoring'].includes(route))return SheetTable.currentSheetIndex(route);
    if(route==='export')return Number(document.getElementById('csvSheet')?.value||0);
    return 0;
  }
  function refresh(){
    routes[route].render();
    updateCreateButton();
  }
  function toast(message,type='success'){
    const element=document.createElement('div');
    element.className=`toast${type==='error'?' toast-error':''}`;
    element.textContent=message;
    document.getElementById('toastRegion').appendChild(element);
    setTimeout(()=>element.remove(),5000);
  }
  function renderQuality(){
    const app=document.getElementById('app'),summary=Validate.summary();
    if(!AppState.data.sheets.length){
      app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Data Quality</h1><p class="page-subtitle">Import a workbook to inspect source-cell parse issues.</p></div></div><div class="card empty-export"><strong>No workbook loaded</strong><span>There are no imported cells to validate yet.</span><button id="emptyQualityImport" class="primary-btn">Import workbook</button></div>`;
      document.getElementById('emptyQualityImport').onclick=()=>navigate('import');
      return;
    }
    const query=(document.getElementById('globalSearch').value||'').trim().toLocaleLowerCase();
    const issues=summary.all.filter(issue=>!query||`${issue.title} ${issue.detail}`.toLocaleLowerCase().includes(query));
    app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Data Quality</h1><p class="page-subtitle">Reports Excel error cells found during import; no additional assumptions are applied.</p></div></div>
    <div class="kpi-grid import-summary-grid mb-4"><div class="kpi card"><div class="kpi-value">${summary.sheets}</div><div class="kpi-label">Worksheets</div></div><div class="kpi card"><div class="kpi-value">${summary.rows}</div><div class="kpi-label">Data rows</div></div><div class="kpi card"><div class="kpi-value">${summary.columns}</div><div class="kpi-label">Columns</div></div><div class="kpi card"><div class="kpi-value">${summary.errors}</div><div class="kpi-label">Parse issues</div></div></div>
    <section class="card quality-panel"><div class="sheet-list-heading"><div><h2>Imported cell issues</h2><p>${issues.length} of ${summary.errors} reported</p></div></div>${issues.length?`<div class="quality-issue-list">${issues.map(issue=>`<div class="quality-issue"><strong>${esc(issue.title)}</strong><span>${esc(issue.detail)}</span></div>`).join('')}</div>`:'<p class="import-no-errors">No imported Excel error cells match this search.</p>'}</section>`;
  }
  document.getElementById('navList').querySelectorAll('.nav-item').forEach(button=>button.onclick=()=>navigate(button.dataset.route));
  document.getElementById('topImport').onclick=()=>navigate('import');
  document.getElementById('createRecord').onclick=()=>Drawer.openCreate(activeSheetIndex());
  document.getElementById('mobileMenu').onclick=()=>document.getElementById('sidebar').classList.toggle('open');
  const settingsModal=document.getElementById('settingsModal'),themeToggle=document.getElementById('themeToggle');
  function closeSettings(){settingsModal.hidden=true;document.getElementById('settingsButton').focus()}
  document.getElementById('settingsButton').onclick=()=>{
    themeToggle.checked=document.body.classList.contains('dark');
    document.getElementById('themeState').textContent=themeToggle.checked?'Dark theme is active':'Light theme is active';
    settingsModal.hidden=false;
    document.getElementById('closeSettings').focus();
  };
  document.getElementById('closeSettings').onclick=closeSettings;
  settingsModal.onclick=event=>{if(event.target===settingsModal)closeSettings()};
  themeToggle.onchange=()=>{
    document.body.classList.toggle('dark',themeToggle.checked);
    document.getElementById('themeState').textContent=themeToggle.checked?'Dark theme is active':'Light theme is active';
  };
  document.getElementById('globalSearch').addEventListener('input',()=>{
    if(route==='dashboard')Dashboard.search();
    if(route==='styles')SheetTable.search('styles');
    if(route==='requests')SheetTable.search('requests');
    if(route==='monitoring')SheetTable.search('monitoring');
    if(route==='quality')renderQuality();
  });
  document.addEventListener('keydown',event=>{
    if(event.key==='Escape'&&!settingsModal.hidden)closeSettings();
    if(event.key==='/'&&document.activeElement?.tagName!=='INPUT'){event.preventDefault();document.getElementById('globalSearch').focus()}
  });
  AppState.ready.then(()=>navigate('dashboard')).catch(error=>{
    document.getElementById('app').innerHTML=`<section class="card storage-error"><strong>Workbook storage is unavailable</strong><p>${esc(error.message)}</p><p>Check browser storage permissions, then refresh the page.</p></section>`;
  });
  return {navigate,toast,refresh};
})();
