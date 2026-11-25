window.Dashboard=(()=>{
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  function render(){
    const app=document.getElementById('app'),data=AppState.data;
    if(!data.sheets.length){
      app.innerHTML=`<section class="empty-workspace"><div class="empty-topline"><span class="eyebrow">WORKBOOK OVERVIEW</span><span class="workspace-status"><i></i> Awaiting import</span></div><div class="empty-hero"><div class="empty-copy"><p class="empty-kicker">NO WORKBOOK LOADED</p><h1>Start with your source file.</h1><p>Import an Excel workbook to make all its worksheets, rows, columns, and values available throughout this workspace.</p><div class="empty-actions"><button id="emptyImport" class="primary-btn large">Import workbook <span>→</span></button></div></div></div><div class="empty-kpis"><article><span>Worksheets</span><strong>—</strong><small>Waiting for import</small></article><article><span>Rows</span><strong>—</strong><small>Waiting for import</small></article><article><span>Columns</span><strong>—</strong><small>Waiting for import</small></article><article><span>Parse issues</span><strong>—</strong><small>Waiting for import</small></article></div></section>`;
      document.getElementById('emptyImport').onclick=()=>App.navigate('import');
      return;
    }
    app.innerHTML=`<div class="sheet-view-heading"><div><p class="eyebrow">WORKBOOK OVERVIEW</p><h1>Dashboard</h1><p class="page-subtitle">A live inventory of worksheets from the imported source file.</p></div><div class="sheet-source"><span>Imported workbook</span><strong title="${esc(data.fileName)}">${esc(data.fileName)}</strong></div></div>
    <div class="kpi-grid import-summary-grid mb-4"><div class="kpi card"><div class="kpi-value">${data.summary.sheets}</div><div class="kpi-label">Worksheets</div></div><div class="kpi card"><div class="kpi-value">${data.summary.rows}</div><div class="kpi-label">Data rows</div></div><div class="kpi card"><div class="kpi-value">${data.summary.columns}</div><div class="kpi-label">Columns</div></div><div class="kpi card"><div class="kpi-value">${data.parseErrors.length}</div><div class="kpi-label">Parse issues</div></div></div>
    <section class="card sheet-list-panel"><div class="sheet-list-heading"><div><h2>Worksheets</h2><p id="dashboardSearchSummary"></p></div><button id="dashboardImport" class="select">Replace workbook</button></div><div class="grid-scroll sheet-list-scroll"><table class="data-table"><thead><tr><th>Worksheet</th><th>Rows</th><th>Columns</th><th>Matches</th><th></th></tr></thead><tbody id="dashboardSheets"></tbody></table></div></section>`;
    document.getElementById('dashboardImport').onclick=()=>App.navigate('import');
    renderSheets();
  }
  function renderSheets(){
    const body=document.getElementById('dashboardSheets');
    if(!body)return;
    const query=(document.getElementById('globalSearch')?.value||'').trim().toLocaleLowerCase();
    const result=AppState.data.sheets.map((sheet,index)=>{
      const rows=AppState.rowsFor(sheet);
      const matches=rows.filter(row=>!query||row.values.some(value=>String(value??'').toLocaleLowerCase().includes(query))).length;
      return {sheet,index,rows,matches};
    }).filter(item=>!query||item.matches>0||item.sheet.name.toLocaleLowerCase().includes(query));
    document.getElementById('dashboardSearchSummary').textContent=query?`${result.reduce((sum,item)=>sum+item.matches,0)} matching rows across ${result.length} worksheets`:`${AppState.data.summary.rows} rows across ${AppState.data.sheets.length} worksheets`;
    body.innerHTML=result.map(({sheet,index,rows,matches})=>`<tr><td><strong>${esc(sheet.name)}</strong></td><td>${rows.length}</td><td>${sheet.columnCount}</td><td>${matches}</td><td><button class="select" data-open-sheet="${index}">Open worksheet</button></td></tr>`).join('')||'<tr><td colspan="5" class="sheet-empty-cell">No imported rows match this search.</td></tr>';
    body.querySelectorAll('[data-open-sheet]').forEach(button=>button.onclick=()=>{
      const index=Number(button.dataset.openSheet);
      App.navigate('requests');
      SheetTable.selectSheet('requests',index);
    });
  }
  function search(){renderSheets()}
  return {render,search};
})();
