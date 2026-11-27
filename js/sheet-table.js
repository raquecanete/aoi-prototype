window.SheetTable=(()=>{
  const styleHeaders=[
    'Season','Brand','RBU','Style','Style Name','Colorway Count','Dev Type',
    'Gender','Age Group','Unit Type','Size Table','Base Size','Product Type',
    'Product Category','SMV','Merchandiser','Pattern Maker','CFT','IE/TSD','OPD',
    'Buy-Ready Date','EMC Approval Date','Last Updated BOM'
  ];
  const monitoringHeaders=[
    'Document No.','Ex-Factory Date','Shipment Status','Season','Brand','Style Name',
    'Style Number','Article #','Colorway Name','Sample Size','Order Qty.',
    'Sample Status','Merchandiser'
  ];
  const styleSourceHeaders={
    'Brand':['Brand','Brand Name'],
    'Style Name':['Style Name','styleName'],
    'Colorway Count':['Colorway Count','ColorwayCount'],
    'Dev Type':['Dev Type','Development Type','devType'],
    'Merchandiser':['Merchandiser','User'],
    'IE/TSD':['IE/TSD','IE TSD'],
    'Buy-Ready Date':['Buy-Ready Date','Buy Ready Date'],
    'EMC Approval Date':['EMC Approval Date','EMC Approval'],
    'Last Updated BOM':['Last Updated BOM'],
    'Article Number/Colourways':['Article Number/Colourways','Article Number / Colourways','Article Number/Colorways','Article No. / Colourways','Article/Colourways'],
    'Article Number':['Article Number','Article No.','Article No','Article #'],
    'Colourways':['Colourways','Colorways'],
    'Fabrics':['Fabrics','Fabric','Fabric Details','Fabric Composition'],
    'Trims':['Trims','Trim','Trim Details'],
    'Size Breakdown':['Size Breakdown','Size Break Down','Size Distribution'],
    'Product Analysis':['Product Analysis','Product Details Analysis'],
    'Attach picture':['Attach picture','Attach Picture','Picture','Image','Image URL','Photo'],
    'Notes/Remarks':['Notes/Remarks','Notes / Remarks','Notes','Remarks','Comments']
  };
  const styleDetailHeaders=[
    'Fabrics','Trims','Size Breakdown','Product Analysis',
    'Attach picture','Notes/Remarks'
  ];
  const monitoringSourceHeaders={
    'Document No.':['Document No.','Document No','Document Number','Document #','Doc No.','Doc No'],
    'Ex-Factory Date':['Ex-Factory Date','Ex Factory Date','Exfactory Date'],
    'Shipment Status':['Shipment Status','Status Shipment','Shipping Status'],
    'Brand':['Brand','Brand Name'],
    'Style Number':['Style Number','Style No.','Style No','Style #','Style Code','Style'],
    'Article #':['Article #','Article Number','Article No.','Article No'],
    'Colorway Name':['Colorway Name','Colorway','Color Way Name'],
    'Sample Size':['Sample Size','Size'],
    'Order Qty.':['Order Qty.','Order Qty','Ordered Qty','Order Quantity','Quantity Ordered','Qty Total'],
    'Sample Status':['Sample Status','Status Sample'],
    'Merchandiser':['Merchandiser','User']
  };
  const views={
    styles:{title:'Style Management',description:'Browse every imported style record and all of its source columns.',suggest:/style/i},
    requests:{title:'Imported File',description:'Browse, review, and search rows from every worksheet in the active workbook.',suggest:null},
    monitoring:{title:'Sample Room Monitoring',description:'Inspect all imported sample-room records and source fields.',suggest:/sample|monitor/i}
  };
  const states={};
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  const normalizeHeader=value=>String(value??'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  function sourceIndexFor(mode,sheet,header){
    const aliasMap=mode==='styles'?styleSourceHeaders:mode==='monitoring'?monitoringSourceHeaders:null;
    const aliases=aliasMap?(aliasMap[header]||[header]):[header];
    for(const alias of aliases){
      const normalizedAlias=normalizeHeader(alias);
      const index=sheet.headers.findIndex(sourceHeader=>normalizeHeader(sourceHeader)===normalizedAlias);
      if(index>=0)return index;
    }
    return -1;
  }
  function columnsFor(mode,sheet){
    if(mode==='styles')return styleHeaders.map(header=>({header,sourceIndex:sourceIndexFor(mode,sheet,header)}));
    if(mode==='monitoring')return monitoringHeaders.map(header=>({header,sourceIndex:sourceIndexFor(mode,sheet,header)}));
    return sheet.headers.map((header,sourceIndex)=>({header,sourceIndex}));
  }
  function detailColumnsFor(mode,sheet){
    const columns=columnsFor(mode,sheet);
    if(mode!=='styles')return columns;
    const combinedArticleIndex=sourceIndexFor(mode,sheet,'Article Number/Colourways');
    let articleColumns;
    if(combinedArticleIndex>=0){
      articleColumns=[{header:'Article Number/Colourways',sourceIndex:combinedArticleIndex}];
    }else{
      articleColumns=['Article Number','Colourways']
        .map(header=>({header,sourceIndex:sourceIndexFor(mode,sheet,header)}))
        .filter(({sourceIndex})=>sourceIndex>=0);
      if(!articleColumns.length)articleColumns=[{header:'Article Number/Colourways',sourceIndex:-1}];
    }
    return columns.concat(articleColumns,styleDetailHeaders.map(header=>({header,sourceIndex:sourceIndexFor(mode,sheet,header)})));
  }
  function stateFor(mode){
    if(!states[mode])states[mode]={sheetName:null,page:1,size:25,sortIndex:-1,dir:1,brand:'',season:'',version:''};
    return states[mode];
  }
  function initialSheetIndex(mode,sheets){
    const suggestion=views[mode].suggest;
    if(mode==='styles'||mode==='monitoring'){
      const categories=mode==='styles'?styleHeaders:monitoringHeaders;
      const suggestedIndex=suggestion?sheets.findIndex(sheet=>suggestion.test(sheet.name)):-1;
      let bestIndex=-1,bestScore=0;
      sheets.forEach((sheet,index)=>{
        const score=categories.filter(header=>sourceIndexFor(mode,sheet,header)>=0).length;
        if(score>bestScore){bestScore=score;bestIndex=index}
      });
      if(suggestedIndex>=0&&bestIndex>=0){
        const suggestedScore=categories.filter(header=>sourceIndexFor(mode,sheets[suggestedIndex],header)>=0).length;
        if(suggestedScore===bestScore)return suggestedIndex;
      }
      if(bestIndex>=0)return bestIndex;
    }
    if(suggestion){const match=sheets.findIndex(sheet=>suggestion.test(sheet.name));if(match>=0)return match}
    return 0;
  }
  function getSheet(mode){
    const sheets=AppState.data.sheets;
    const state=stateFor(mode);
    if(state.version!==AppState.data.importedAt){
      state.version=AppState.data.importedAt;
      state.sheetName=sheets[initialSheetIndex(mode,sheets)]?.name||null;
      state.page=1;state.sortIndex=-1;state.brand='';state.season='';
    }
    let index=sheets.findIndex(sheet=>sheet.name===state.sheetName);
    if(index<0)index=initialSheetIndex(mode,sheets);
    state.sheetName=sheets[index]?.name||null;
    return {sheet:sheets[index],index,state};
  }
  function render(mode){
    const app=document.getElementById('app');
    const config=views[mode],sheets=AppState.data.sheets;
    if(!sheets.length){
      app.innerHTML=`<section class="empty-workspace"><div class="empty-topline"><span class="eyebrow">WORKBOOK DATA</span><span class="workspace-status"><i></i> Awaiting import</span></div><div class="empty-hero"><div class="empty-copy"><p class="empty-kicker">NO WORKBOOK LOADED</p><h1>${esc(config.title)}</h1><p>Import an Excel workbook to view its complete sheets, rows, and columns here. This page will show only imported data.</p><div class="empty-actions"><button id="emptySheetImport" class="primary-btn large">Import workbook <span>→</span></button></div></div></div></section>`;
      document.getElementById('emptySheetImport').onclick=()=>App.navigate('import');
      return;
    }
    const {sheet,index,state}=getSheet(mode);
    const rows=AppState.rowsFor(sheet);
    const preamble=AppState.preambleFor(sheet);
    const preambleHtml=mode==='requests'&&preamble.length?`<div class="sheet-preamble" aria-label="Worksheet report parameters">${preamble.map(row=>`<div class="sheet-preamble-row">${row.values.map((value,column)=>value===''||value===null||value===undefined?'':`<span class="sheet-preamble-cell">${esc(AppState.displayCell(sheet,row.index,column))}</span>`).join('')}</div>`).join('')}</div>`:'';
    const brandIndex=sourceIndexFor(mode,sheet,'Brand');
    const seasonIndex=sourceIndexFor(mode,sheet,'Season');
    const brands=brandIndex<0?[]:uniqueValues(rows,brandIndex);
    const seasons=seasonIndex<0?[]:uniqueValues(rows,seasonIndex);
    app.innerHTML=`<div class="sheet-view-heading"><div><p class="eyebrow">IMPORTED WORKBOOK</p><h1>${esc(config.title)}</h1><p class="page-subtitle">${esc(config.description)}</p></div><div class="sheet-source"><span>Source file</span><strong title="${esc(AppState.data.fileName)}">${esc(AppState.data.fileName)}</strong></div></div>
    ${preambleHtml}
    <section class="card sheet-browser">
      <div class="sheet-toolbar">
        <div><label for="sheetSelect" class="sheet-label">Worksheet</label><select id="sheetSelect" class="select">${sheets.map((item,i)=>`<option value="${i}" ${i===index?'selected':''}>${esc(item.name)} · ${AppState.rowsFor(item).length} rows · ${item.columnCount} columns</option>`).join('')}</select></div>
        <div class="sheet-row-count"><strong id="sheetCount"></strong><span>rows in this worksheet</span></div>
      </div>
      <div class="sheet-filterbar">
        <div class="filters">${brandIndex<0?'':`<select id="sheetBrand" class="select" aria-label="Filter by brand"><option value="">All brands</option>${brands.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join('')}</select>`}${seasonIndex<0?'':`<select id="sheetSeason" class="select" aria-label="Filter by season"><option value="">All seasons</option>${seasons.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join('')}</select>`}</div>
        <div class="flex gap-2"><select id="sheetPageSize" class="select" aria-label="Rows per page"><option value="25">25 rows</option><option value="50">50 rows</option><option value="100">100 rows</option></select><button id="sheetClearFilters" class="select">Clear filters</button></div>
      </div>
      <div class="grid-scroll sheet-table-scroll"><table class="data-table workbook-table"><thead id="sheetHead"></thead><tbody id="sheetBody"></tbody></table></div>
      <div class="sheet-pager"><span id="sheetPagerLabel"></span><div><button id="sheetPrevious" class="select" aria-label="Previous page">‹</button><button id="sheetNext" class="select" aria-label="Next page">›</button></div></div>
    </section>`;
    document.getElementById('sheetSelect').onchange=event=>{
      state.sheetName=sheets[Number(event.target.value)].name;
      state.page=1;state.sortIndex=-1;state.brand='';state.season='';
      render(mode);
    };
    document.getElementById('sheetPageSize').value=String(state.size);
    document.getElementById('sheetPageSize').onchange=event=>{state.size=Number(event.target.value);state.page=1;renderTable(mode)};
    const brandSelect=document.getElementById('sheetBrand');
    if(brandSelect){brandSelect.value=state.brand;brandSelect.onchange=event=>{state.brand=event.target.value;state.page=1;renderTable(mode)}}
    const seasonSelect=document.getElementById('sheetSeason');
    if(seasonSelect){seasonSelect.value=state.season;seasonSelect.onchange=event=>{state.season=event.target.value;state.page=1;renderTable(mode)}}
    document.getElementById('sheetClearFilters').onclick=()=>{
      state.brand='';state.season='';state.page=1;
      if(brandSelect)brandSelect.value='';
      if(seasonSelect)seasonSelect.value='';
      renderTable(mode);
    };
    document.getElementById('sheetPrevious').onclick=()=>{if(state.page>1){state.page--;renderTable(mode)}};
    document.getElementById('sheetNext').onclick=()=>{state.page++;renderTable(mode)};
    renderTable(mode);
  }
  function uniqueValues(rows,columnIndex){
    return [...new Set(rows.map(row=>String(row.values[columnIndex]??'').trim()).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:'base'}));
  }
  function selectSheet(mode,index){
    const sheet=AppState.data.sheets[index];
    if(!sheet)return;
    const state=stateFor(mode);
    state.version=AppState.data.importedAt;
    state.sheetName=sheet.name;
    state.page=1;state.sortIndex=-1;state.brand='';state.season='';
    render(mode);
  }
  function renderTable(mode){
    const head=document.getElementById('sheetHead');
    if(!head)return;
    const {sheet,index,state}=getSheet(mode);
    const localBrand=document.getElementById('sheetBrand'),localSeason=document.getElementById('sheetSeason');
    if(localBrand)localBrand.value=state.brand;
    if(localSeason)localSeason.value=state.season;
    const rows=AppState.rowsFor(sheet);
    const columns=columnsFor(mode,sheet);
    const query=(document.getElementById('globalSearch')?.value||'').trim().toLocaleLowerCase();
    const brandIndex=sourceIndexFor(mode,sheet,'Brand'),seasonIndex=sourceIndexFor(mode,sheet,'Season');
    let filtered=AppState.searchRows(sheet,query).filter(row=>
      (brandIndex<0||!state.brand||String(row.values[brandIndex]??'').trim()===state.brand)&&
      (seasonIndex<0||!state.season||String(row.values[seasonIndex]??'').trim()===state.season)
    );
    if(state.sortIndex>=0){
      const sourceIndex=columns[state.sortIndex]?.sourceIndex??-1;
      filtered.sort((a,b)=>String(sourceIndex<0?'':a.values[sourceIndex]??'').localeCompare(String(sourceIndex<0?'':b.values[sourceIndex]??''),undefined,{numeric:true,sensitivity:'base'})*state.dir||a.index-b.index);
    }
    const count=filtered.length,pages=Math.max(1,Math.ceil(count/state.size));
    if(state.page>pages)state.page=pages;
    const visible=filtered.slice((state.page-1)*state.size,state.page*state.size);
    document.getElementById('sheetCount').textContent=String(rows.length);
    document.getElementById('sheetPagerLabel').textContent=count?`${(state.page-1)*state.size+1}–${Math.min(state.page*state.size,count)} of ${count} rows`:'0 of 0 rows';
    document.getElementById('sheetPrevious').disabled=state.page<=1;
    document.getElementById('sheetNext').disabled=state.page>=pages;
    head.innerHTML=`<tr>${columns.map(({header},column)=>`<th data-sort-column="${column}" title="${esc(header)}">${esc(header)} <span>↕</span></th>`).join('')}</tr>`;
    document.getElementById('sheetBody').innerHTML=visible.map(row=>`<tr data-row-index="${row.index}">${columns.map(({sourceIndex})=>{
      const raw=sourceIndex<0?null:row.values[sourceIndex];
      const value=sourceIndex<0?'—':AppState.displayCell(sheet,row.index,sourceIndex);
      const numeric=typeof raw==='number'||typeof raw==='string'&&raw.trim()!==''&&Number.isFinite(Number(raw));
      return `<td class="${numeric?'numeric-cell':''}" title="${esc(value)}">${esc(value)}</td>`;
    }).join('')}</tr>`).join('');
    head.querySelectorAll('[data-sort-column]').forEach(th=>th.onclick=()=>{
      const column=Number(th.dataset.sortColumn);
      state.dir=state.sortIndex===column?state.dir*-1:1;
      state.sortIndex=column;renderTable(mode);
    });
    document.getElementById('sheetBody').querySelectorAll('tr').forEach(row=>row.onclick=()=>Drawer.open(index,Number(row.dataset.rowIndex),detailColumnsFor(mode,sheet)));
  }
  function search(mode){renderTable(mode)}
  function currentSheetIndex(mode){return getSheet(mode).index}
  return {render,search,selectSheet,currentSheetIndex};
})();
