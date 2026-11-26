window.Dashboard=(()=>{
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  let files=[];
  let loaded=false;
  let renderId=0;

  function dateText(value){
    if(!value)return '—';
    const date=new Date(value);
    return Number.isNaN(date.getTime())?'—':date.toLocaleDateString(undefined,{day:'2-digit',month:'short',year:'numeric'});
  }

  function fileType(file){
    const extension=String(file.originalFileName||file.name||'').split('.').pop().toUpperCase();
    return ['XLS','XLSX','CSV'].includes(extension)?extension:'FILE';
  }

  function searchFiles(query){
    const folded=query.trim().toLocaleLowerCase();
    if(!folded)return files;
    return files.filter(file=>
      `${file.name} ${file.originalFileName} ${(file.rowCounts||[]).map(sheet=>sheet.name).join(' ')}`
        .toLocaleLowerCase().includes(folded)
    );
  }

  async function render(){
    const app=document.getElementById('app');
    const requestId=++renderId;
    app.innerHTML='<div class="files-loading" role="status">Loading workbook overview…</div>';
    try{
      const allFiles=await AppState.listFiles();
      if(requestId!==renderId)return;
      files=allFiles.filter(file=>file.status!=='Archived');
      loaded=true;
      renderOverview();
    }catch(error){
      if(requestId!==renderId)return;
      loaded=false;
      app.innerHTML=`<section class="card storage-error"><strong>Could not load workbook overview</strong><p>${esc(error.message)}</p></section>`;
    }
  }

  function renderOverview(){
    if(!loaded)return;
    const app=document.getElementById('app');
    const query=document.getElementById('globalSearch')?.value||'';
    const visibleFiles=searchFiles(query);
    const worksheetCount=visibleFiles.reduce((total,file)=>total+(file.rowCounts||[]).length,0);
    const rowCount=visibleFiles.reduce((total,file)=>total+(file.rowCounts||[]).reduce((sum,sheet)=>sum+(Number(sheet.rows)||0),0),0);
    const columnCount=visibleFiles.reduce((total,file)=>total+(file.rowCounts||[]).reduce((sum,sheet)=>sum+(Number(sheet.columns)||0),0),0);
    const hasImportedFiles=files.length>0;

    app.innerHTML=`<header class="page-heading dashboard-overview-heading">
      <div><p class="eyebrow">WORKBOOK OVERVIEW</p><h1>Dashboard</h1></div>
    </header>
    <div class="kpi-grid import-summary-grid mb-4">
      <div class="kpi card"><div class="kpi-value">${visibleFiles.length}</div><div class="kpi-label">Workbooks</div></div>
      <div class="kpi card"><div class="kpi-value">${worksheetCount}</div><div class="kpi-label">Worksheets</div></div>
      <div class="kpi card"><div class="kpi-value">${rowCount}</div><div class="kpi-label">Data rows</div></div>
      <div class="kpi card"><div class="kpi-value">${columnCount}</div><div class="kpi-label">Columns</div></div>
    </div>
    <section class="card dashboard-library-overview" aria-labelledby="dashboardWorkbooksTitle">
      <div class="dashboard-overview-section-heading"><div><h2 id="dashboardWorkbooksTitle">Imported workbooks</h2><p>${query.trim()?`${visibleFiles.length} matching workbook${visibleFiles.length===1?'':'s'}`:`${visibleFiles.length} workbook${visibleFiles.length===1?'':'s'} in your library`}</p></div></div>
      ${visibleFiles.length?`<div class="dashboard-workbook-list">${visibleFiles.map(renderWorkbook).join('')}</div>`:
        `<div class="dashboard-overview-empty"><strong>${hasImportedFiles?'No workbooks or worksheets match your search.':'No active workbooks yet.'}</strong><span>${hasImportedFiles?'Try a different search.':'Use the Import control in the top bar to add workbooks. Archived workbooks are available from Files.'}</span></div>`}
    </section>`;

    app.querySelectorAll('[data-open-workbook]').forEach(button=>{
      button.onclick=()=>App.openFile(button.dataset.openWorkbook);
    });
  }

  function renderWorkbook(file){
    const sheets=file.rowCounts||[];
    const rows=sheets.reduce((total,sheet)=>total+(Number(sheet.rows)||0),0);
    return `<article class="dashboard-workbook-card">
      <header class="dashboard-workbook-header">
        <div class="dashboard-workbook-title">
          <span class="file-type-mark" aria-hidden="true">${fileType(file)}</span>
          <div><h3 title="${esc(file.name)}">${esc(file.name)}</h3><span>${esc(file.originalFileName||file.name)}</span></div>
        </div>
        <button type="button" class="select" data-open-workbook="${esc(file.id)}">Open workbook</button>
      </header>
      <div class="dashboard-workbook-meta"><span>Version ${Number(file.currentVersion)||1}</span><span>${sheets.length} worksheets</span><span>${rows} data rows</span><span>Modified ${dateText(file.lastModifiedAt)}</span></div>
      <div class="dashboard-sheet-overview" aria-label="Worksheet overview">
        ${sheets.length?sheets.map(sheet=>`<div class="dashboard-sheet-summary"><strong>${esc(sheet.name)}</strong><span>${Number(sheet.rows)||0} rows</span><span>${Number(sheet.columns)||0} columns</span></div>`).join(''):'<span class="dashboard-sheet-empty">No worksheet summary available.</span>'}
      </div>
    </article>`;
  }

  function search(){
    renderOverview();
  }

  return {render,search};
})();
