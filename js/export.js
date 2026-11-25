window.ExportView=(()=>{
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  function render(){
    const data=AppState.data,app=document.getElementById('app');
    if(!data.sheets.length){
      app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Export</h1><p class="page-subtitle">No workbook is currently available to export.</p></div></div><div class="card empty-export"><strong>Nothing to export yet</strong><span>Import a workbook to export its complete sheets and source values.</span><button id="emptyExportImport" class="primary-btn">Import workbook</button></div>`;
      document.getElementById('emptyExportImport').onclick=()=>App.navigate('import');
      return;
    }
    app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Export</h1><p class="page-subtitle">Export the complete imported workbook or an individual worksheet.</p></div></div>
    <div class="kpi-grid import-summary-grid mb-4"><div class="kpi card"><div class="kpi-value">${data.summary.sheets}</div><div class="kpi-label">Worksheets</div></div><div class="kpi card"><div class="kpi-value">${data.summary.rows}</div><div class="kpi-label">Data rows</div></div><div class="kpi card"><div class="kpi-value">${data.summary.columns}</div><div class="kpi-label">Columns</div></div></div>
    <section class="card export-panel"><div><h2>Complete workbook</h2><p>Includes every imported worksheet, row, column, cell value, and merged range.</p><button id="xlsxExport" class="primary-btn">Download full workbook (.xlsx)</button></div><div class="export-divider"></div><div><h2>Single worksheet</h2><p>Download one worksheet as CSV without changing the saved workbook.</p><div class="flex gap-2"><select id="csvSheet" class="select" aria-label="Worksheet to export">${data.sheets.map((sheet,index)=>`<option value="${index}">${esc(sheet.name)}</option>`).join('')}</select><button id="csvExport" class="select">Download CSV</button></div></div></section>`;
    document.getElementById('xlsxExport').onclick=downloadWorkbook;
    document.getElementById('csvExport').onclick=downloadCsv;
  }
  function worksheetForExport(sheet){
    const origin={r:sheet.startRow||0,c:sheet.startColumn||0};
    const worksheet=XLSX.utils.aoa_to_sheet(sheet.matrix,{origin});
    for(let r=0;r<sheet.cellMeta.length;r++){
      for(let c=0;c<sheet.cellMeta[r].length;c++){
        const meta=sheet.cellMeta[r][c];
        if(!meta)continue;
        const address=XLSX.utils.encode_cell({r:origin.r+r,c:origin.c+c});
        if(meta.error)worksheet[address]={t:'e',v:meta.errorValue,w:meta.error};
        else if(meta.formula&&worksheet[address])worksheet[address].f=meta.formula;
        if(meta.format&&worksheet[address])worksheet[address].z=meta.format;
      }
    }
    if(sheet.endRow>=sheet.startRow&&sheet.endColumn>=sheet.startColumn){
      worksheet['!ref']=XLSX.utils.encode_range({
        s:{r:sheet.startRow,c:sheet.startColumn},
        e:{r:sheet.endRow,c:sheet.endColumn}
      });
    }
    worksheet['!merges']=(sheet.merges||[]).map(range=>XLSX.utils.decode_range(range));
    return worksheet;
  }
  function downloadWorkbook(){
    try{
      const workbook=XLSX.utils.book_new();
      AppState.data.sheets.forEach(sheet=>XLSX.utils.book_append_sheet(workbook,worksheetForExport(sheet),sheet.name));
      XLSX.writeFile(workbook,'aoi-workbook-export.xlsx');
    }catch(error){
      App.toast(`Workbook export failed: ${error.message}`,'error');
    }
  }
  function downloadCsv(){
    const index=Number(document.getElementById('csvSheet').value);
    const sheet=AppState.data.sheets[index];
    if(!sheet){App.toast('Choose a valid worksheet to export.','error');return}
    try{
      const workbook=XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook,worksheetForExport(sheet),sheet.name);
      XLSX.writeFile(workbook,`${sheet.name}.csv`,{bookType:'csv'});
    }catch(error){
      App.toast(`CSV export failed: ${error.message}`,'error');
    }
  }
  return {render};
})();
