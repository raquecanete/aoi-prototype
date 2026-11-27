window.SampleRequests=(()=>{
  const fields=[
    {label:'No.'},
    {label:'MO Created'},
    {label:'MO Updated'},
    {label:'Tarikan'},
    {label:'Tgl MO Completed'},
    {label:'Document No.'},
    {label:'Season'},
    {label:'Style'},
    {label:'Article'},
    {label:'Size'},
    {label:'Status Sample'},
    {label:'Qty MO'},
    {label:'PIC'},
    {label:'Marker Remark'},
    {label:'Description'},
    {label:'Product Category'},
    {label:'Color'},
    {label:'Allocation'},
    {label:'Doc Status'}
  ];
  const filters=[
    {key:'season',header:'Season',label:'Season',allLabel:'All seasons'},
    {key:'status-sample',header:'Status Sample',label:'Status Sample',allLabel:'All statuses'},
    {key:'product-category',header:'Product Category',label:'Product Ctg',allLabel:'All categories'}
  ];
  const documentNoFieldIndex=fields.findIndex(field=>field.label==='Document No.');
  const dateFieldIndices=new Set(['MO Created','MO Updated','Tgl MO Completed']
    .map(label=>fields.findIndex(field=>field.label===label)));
  const headerAliases={
    'MO Created':['MO Created Date','Date MO Created','Created Date','Date Created','MO Created At'],
    'MO Updated':['MO Updated Date','Date MO Updated','Updated Date','Date Updated','MO Updated At'],
    'Tgl MO Completed':['Tgl. MO Completed','Tanggal MO Completed','Tanggal MO Selesai','Date MO Completed','MO Completed Date','MO Completion Date','Completed Date'],
    'Product Category':['Product Ctg','Product Cat','Category'],
    'Doc Status':['Document Status','Doc. Status','Document Status Code','Doc Status Code']
  };
  const pageSourceKey='sampleRequestsPageWorkbook';
  let pageSource=null;
  let pageSourceLoaded=false;
  let pageSourcePromise=null;
  let renderSequence=0;
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  const normalizeHeader=value=>String(value??'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();

  function loadPageSource(){
    if(pageSourceLoaded)return Promise.resolve(pageSource);
    if(!pageSourcePromise){
      pageSourcePromise=AOIDB.get('settings',pageSourceKey).then(record=>{
        pageSource=record?.value||null;
        pageSourceLoaded=true;
        return pageSource;
      });
    }
    return pageSourcePromise;
  }

  function pageHeading(){
    return `<div class="page-heading sample-request-heading"><div><p class="eyebrow">SAMPLE WORKFLOW</p><h1>Sample Request Management</h1><p class="page-subtitle">Import a workbook for this page only, without changing your active workbook or library.</p></div><div class="sample-request-import-actions"><input id="sample-request-file-input" type="file" accept=".xlsx,.xls,.csv" hidden><button id="sample-request-import" class="primary-btn" type="button">⇧ Import file</button>${pageSource?'<button id="sample-request-use-active" class="select" type="button">Use active workbook</button>':''}</div></div><p id="sample-request-import-status" class="sample-request-import-status" role="status" hidden></p>`;
  }

  function bindPageImport(app){
    const input=app.querySelector('#sample-request-file-input');
    const button=app.querySelector('#sample-request-import');
    const status=app.querySelector('#sample-request-import-status');
    button.onclick=()=>input.click();
    input.onchange=async()=>{
      const file=input.files?.[0];
      if(!file)return;
      button.disabled=true;
      button.textContent='Importing…';
      status.hidden=false;
      status.classList.remove('is-error');
      status.textContent=`Reading ${file.name}…`;
      try{
        const workbook=await ImportWizard.parseFile(await file.arrayBuffer(),file.name);
        if(!workbook.sheets?.length)throw new Error('The workbook contains no readable worksheets.');
        pageSource={fileName:file.name,importedAt:new Date().toISOString(),workbook};
        pageSourceLoaded=true;
        pageSourcePromise=Promise.resolve(pageSource);
        try{
          await AOIDB.put('settings',{key:pageSourceKey,value:pageSource});
        }catch(error){
          App.toast(`The workbook is available for this session, but could not be saved: ${error.message}`,'error');
        }
        App.toast(`${file.name} imported for Sample Request Management only.`);
        render();
      }catch(error){
        status.classList.add('is-error');
        status.textContent=`Could not import ${file.name}: ${error.message}`;
        App.toast(`Could not import ${file.name}: ${error.message}`,'error');
        button.disabled=false;
        button.textContent='⇧ Import file';
      }finally{
        input.value='';
      }
    };
    const useActive=app.querySelector('#sample-request-use-active');
    if(useActive)useActive.onclick=async()=>{
      useActive.disabled=true;
      try{
        await AOIDB.delete('settings',pageSourceKey);
        pageSource=null;
        pageSourceLoaded=true;
        pageSourcePromise=Promise.resolve(null);
        render();
      }catch(error){
        useActive.disabled=false;
        App.toast(`Could not switch back to the active workbook: ${error.message}`,'error');
      }
    };
  }

  function openPageSourceRow(entry,displayColumns){
    const drawer=document.getElementById('drawer');
    const editButton=AppState.currentUser?.role==='admin'?'<button id="sample-request-row-edit" class="select" type="button">Edit row</button>':'';
    drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">${esc(entry.sheet.name)}</div><div class="text-xs text-slate-500">Page-only import · source row ${entry.row.sourceRow}</div></div><div class="drawer-head-actions">${editButton}<button id="drawerClose" class="icon-btn" aria-label="Close row details">×</button></div></div><div class="drawer-body">${displayColumns.map(({header,displayValue})=>`<div class="field-grid"><div class="field-label">${esc(header||'—')}</div><div class="field-value">${esc(displayValue)}</div></div>`).join('')}</div>`;
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden','false');
    document.getElementById('drawerBackdrop').classList.add('open');
    document.getElementById('drawerClose').onclick=Drawer.close;
    drawer.querySelector('#sample-request-row-edit')?.addEventListener('click',()=>openPageSourceRowEditor(entry,displayColumns));
  }

  function openPageSourceRowEditor(entry,displayColumns){
    const drawer=document.getElementById('drawer');
    const fields=displayColumns.map(({header,sourceIndex,displayValue},index)=>{
      if(sourceIndex<0)return `<div class="field-grid"><div class="field-label">${esc(header||'—')}</div><div class="field-value">${esc(displayValue)}</div></div>`;
      const rawValue=entry.row.values[sourceIndex];
      const value=dateFieldIndices.has(index)?AppState.formatDateValue(rawValue)||'':String(rawValue??'');
      const type=dateFieldIndices.has(index)?'date':'text';
      return `<label class="edit-row-field"><span>${esc(header||'—')}</span><input class="filter-input" type="${type}" data-page-source-column="${sourceIndex}" value="${esc(value)}"></label>`;
    }).join('');
    drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">Edit row · ${esc(entry.sheet.name)}</div><div class="text-xs text-slate-500">Changes are saved to this page-only import.</div></div><button id="drawerClose" class="icon-btn" aria-label="Close row editor">×</button></div><div class="drawer-body"><form id="sample-request-row-form"><div class="edit-row-form">${fields}</div><div class="create-row-actions"><button id="sample-request-row-cancel" class="select" type="button">Cancel</button><button id="sample-request-row-save" class="primary-btn" type="submit">Save changes</button></div></form></div>`;
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden','false');
    document.getElementById('drawerBackdrop').classList.add('open');
    drawer.querySelector('#drawerClose').onclick=Drawer.close;
    drawer.querySelector('#sample-request-row-cancel').onclick=()=>openPageSourceRow(entry,displayColumns);
    const form=drawer.querySelector('#sample-request-row-form');
    const saveButton=drawer.querySelector('#sample-request-row-save');
    const inputs=[...form.querySelectorAll('[data-page-source-column]')];
    const initialValues=new Map(inputs.map(input=>[input.dataset.pageSourceColumn,input.value]));
    const updateSaveState=()=>{saveButton.disabled=!inputs.some(input=>input.value!==initialValues.get(input.dataset.pageSourceColumn))};
    inputs.forEach(input=>{
      input.oninput=updateSaveState;
      input.onchange=updateSaveState;
    });
    updateSaveState();
    form.onsubmit=event=>{
      event.preventDefault();
      savePageSourceRow(entry,form,initialValues);
    };
  }

  async function savePageSourceRow(entry,form,initialValues){
    const row=entry.sheet.matrix?.[entry.row.index];
    if(!Array.isArray(row)){
      App.toast('Could not save this page-only row because its source row is unavailable.','error');
      return;
    }
    const changes=[...form.querySelectorAll('[data-page-source-column]')]
      .map(input=>({column:Number(input.dataset.pageSourceColumn),value:input.value}))
      .filter(change=>change.value!==initialValues.get(String(change.column)));
    if(!changes.length)return;
    const previous=changes.map(change=>row[change.column]);
    const button=form.querySelector('#sample-request-row-save');
    button.disabled=true;
    button.textContent='Saving…';
    try{
      changes.forEach(change=>{row[change.column]=change.value});
      await AOIDB.put('settings',{key:pageSourceKey,value:pageSource});
      App.toast(`Saved ${changes.length} field${changes.length===1?'':'s'} to the page-only workbook.`);
      Drawer.close();
      render();
    }catch(error){
      changes.forEach((change,index)=>{row[change.column]=previous[index]});
      button.disabled=false;
      button.textContent='Save changes';
      App.toast(`Could not save the page-only row: ${error.message}`,'error');
    }
  }

  function sourceColumn(sheet,header){
    const aliases=new Set([header,...(headerAliases[header]||[])].map(normalizeHeader));
    if(header==='Document No.')aliases.add(normalizeHeader('Document Number'));
    return sheet.headers.findIndex(sourceHeader=>aliases.has(normalizeHeader(sourceHeader)));
  }

  function fieldValue(entry,index,rowNumber){
    const column=entry.columns[index];
    const sourceValue=column<0?'':AppState.displayCell(entry.sheet,entry.row.index,column);
    if(index===documentNoFieldIndex&&(!sourceValue||sourceValue==='—')&&entry.linkedDocumentNo)return entry.linkedDocumentNo;
    if(dateFieldIndices.has(index)&&column>=0){
      const date=AppState.formatDateValue(entry.row.values[column]);
      if(date){
        const [year,month,day]=date.split('-').map(Number);
        return new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'})
          .format(new Date(Date.UTC(year,month-1,day)));
      }
    }
    return index===0&&(!sourceValue||sourceValue==='—')?String(rowNumber+1):sourceValue;
  }

  function filterValue(entry,header){
    const index=fields.findIndex(field=>field.label===header);
    if(index<0||entry.columns[index]<0)return '';
    const value=AppState.displayCell(entry.sheet,entry.row.index,entry.columns[index]);
    return String(value??'').trim()==='—'?'':String(value??'').trim();
  }

  function documentsByPartNo(sheets){
    const documents=new Map();
    for(const sheet of sheets){
      const partColumn=sourceColumn(sheet,'Part No.');
      const documentColumn=sourceColumn(sheet,'Document No.');
      if(partColumn<0||documentColumn<0)continue;
      for(const row of AppState.rowsFor(sheet)){
        const partNo=String(row.values[partColumn]??'').trim();
        const documentNo=String(AppState.displayCell(sheet,row.index,documentColumn)??'').trim();
        if(!partNo||!documentNo||documentNo==='—')continue;
        if(!documents.has(partNo))documents.set(partNo,new Set());
        documents.get(partNo).add(documentNo);
      }
    }
    return documents;
  }

  function render(){
    const app=document.getElementById('app');
    const sequence=++renderSequence;
    const workbook=pageSource?.workbook||AppState.data;
    const sheets=workbook.sheets||[];
    if(!pageSourceLoaded){
      loadPageSource().then(()=>{
        if(sequence===renderSequence&&document.querySelector('.nav-item.active[data-route="sample-requests"]'))render();
      }).catch(error=>{
        pageSourceLoaded=true;
        pageSourcePromise=null;
        App.toast(`Could not load the page-specific workbook: ${error.message}`,'error');
      });
    }
    if(!sheets.length){
      app.innerHTML=`${pageHeading()}<section class="card empty-export"><strong>${pageSource?'The page-only workbook has no worksheets':'No workbook loaded'}</strong><span>Import an Excel or CSV file here to use it only in Sample Request Management.</span></section>`;
      bindPageImport(app);
      return;
    }

    let rows=[];
    let matchedFieldCount=0;
    const documentMap=documentsByPartNo(sheets);
    for(let sheetIndex=0;sheetIndex<sheets.length;sheetIndex++){
      const sheet=sheets[sheetIndex];
      const columns=fields.map(field=>sourceColumn(sheet,field.label));
      matchedFieldCount=Math.max(matchedFieldCount,columns.filter((column,index)=>index>0&&column>=0).length);
      if(!columns.some((column,index)=>index>0&&column>=0))continue;
      for(const row of AppState.rowsFor(sheet)){
        const hasValues=columns.some((column,index)=>index>0&&column>=0&&String(row.values[column]??'').trim()!=='');
        if(hasValues){
          const entry={sheet,sheetIndex,row,columns};
          const documentColumn=columns[documentNoFieldIndex];
          const documentValue=documentColumn<0?'':String(AppState.displayCell(sheet,row.index,documentColumn)??'').trim();
          const styleColumn=sourceColumn(sheet,'Style Number')>=0
            ?sourceColumn(sheet,'Style Number')
            :sourceColumn(sheet,'Style');
          const styleNumber=styleColumn<0?'':String(row.values[styleColumn]??'').trim();
          if((!documentValue||documentValue==='—')&&styleNumber){
            entry.linkedDocumentNo=[...(documentMap.get(styleNumber)||[])].join(', ');
          }
          rows.push(entry);
        }
      }
    }

    const filterOptions=filters.map(filter=>({
      ...filter,
      values:[...new Set(rows.map(entry=>filterValue(entry,filter.header)).filter(Boolean))]
        .sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:'base'}))
    }));

    app.innerHTML=`${pageHeading()}
      ${matchedFieldCount?`<section class="card sample-request-table-panel" aria-label="Sample request records"><div class="sample-request-table-summary"><strong id="sample-request-count" aria-live="polite">${rows.length} matching ${rows.length===1?'record':'records'}</strong><span>Source: ${esc(AppState.data.fileName||AppState.activeFile?.name||'Active workbook')}</span></div>
        <div class="sample-request-filters" aria-label="Filter and search sample requests">${filterOptions.map(filter=>`<label class="sample-request-filter-field" for="sample-request-filter-${filter.key}"><span>${esc(filter.label)}</span><select id="sample-request-filter-${filter.key}" class="select sample-request-filter" data-filter-key="${filter.key}" ${filter.values.length?'':'disabled'}><option value="">${esc(filter.allLabel)}</option>${filter.values.map(value=>`<option value="${esc(value)}">${esc(value)}</option>`).join('')}</select></label>`).join('')}<label class="sample-request-filter-field sample-request-search-field" for="sample-request-search"><span>Search all details</span><input id="sample-request-search" class="filter-input" type="search" placeholder="Search table details" aria-label="Search all details in the sample request table" autocomplete="off"></label></div>
        <div class="grid-scroll sample-request-table-scroll"><table class="data-table sample-requests-table"><thead><tr>${fields.map(field=>`<th scope="col">${esc(field.label)}</th>`).join('')}</tr></thead><tbody>${rows.length?`${rows.map((entry,rowNumber)=>`<tr class="sample-request-record" data-request-sheet="${entry.sheetIndex}" data-request-row="${entry.row.index}" ${filterOptions.map(filter=>`data-filter-${filter.key}="${esc(filterValue(entry,filter.header))}"`).join(' ')} tabindex="0" aria-label="View all data from ${esc(entry.sheet.name)}, source row ${entry.row.sourceRow}">${fields.map((field,index)=>{
          const value=fieldValue(entry,index,rowNumber);
          const source=index===0?` title="${esc(`${entry.sheet.name} · source row ${entry.row.sourceRow}`)}"`:'';
          return `<td${source}>${esc(value)}</td>`;
        }).join('')}</tr>`).join('')}<tr class="sample-request-filter-empty" hidden><td class="sample-request-no-matches" colspan="${fields.length}">No records match the selected filters.</td></tr>`:`<tr><td class="sample-request-no-matches" colspan="${fields.length}">No populated rows match the available sample request columns.</td></tr>`}</tbody></table></div></section>`:
        `<section class="card empty-export sample-request-empty" aria-label="No matching imported columns"><strong>No matching sample request columns</strong><span>This workbook does not contain columns that match the sample request table. Import a workbook with sample request fields to populate it.</span></section>`}`;

    const sourceLabel=pageSource?`Page-only import: ${pageSource.fileName}`:AppState.data.fileName||AppState.activeFile?.name||'Active workbook';
    const sourceSummary=app.querySelector('.sample-request-table-summary span');
    if(sourceSummary)sourceSummary.textContent=`Source: ${sourceLabel}`;
    bindPageImport(app);

    const filterSelects=[...app.querySelectorAll('.sample-request-filter')];
    const searchInput=app.querySelector('#sample-request-search');
    const records=[...app.querySelectorAll('.sample-request-record')];
    const filterEmpty=app.querySelector('.sample-request-filter-empty');
    const countLabel=app.querySelector('#sample-request-count');
    const applyFilters=()=>{
      const activeFilters=filterSelects.filter(select=>select.value);
      const query=String(searchInput?.value||'').trim().toLocaleLowerCase();
      let visibleCount=0;
      records.forEach(row=>{
        const matchesFilters=activeFilters.every(select=>row.getAttribute(`data-filter-${select.dataset.filterKey}`)===select.value);
        const matchesSearch=!query||row.textContent.toLocaleLowerCase().includes(query);
        const matches=matchesFilters&&matchesSearch;
        row.hidden=!matches;
        if(matches)visibleCount++;
      });
      if(filterEmpty)filterEmpty.hidden=visibleCount!==0;
      if(countLabel)countLabel.textContent=`${visibleCount} matching ${visibleCount===1?'record':'records'}`;
    };
    filterSelects.forEach(select=>select.addEventListener('change',applyFilters));
    if(searchInput)searchInput.addEventListener('input',applyFilters);

    app.querySelectorAll('.sample-request-record').forEach((row,rowNumber)=>{
      const entry=rows[rowNumber];
      const displayColumns=fields.map((field,index)=>({
        header:field.label,
        sourceIndex:entry.columns[index],
        displayValue:fieldValue(entry,index,rowNumber)
      }));
      const openDetails=()=>pageSource
        ?openPageSourceRow(entry,displayColumns)
        :Drawer.open(entry.sheetIndex,entry.row.index,displayColumns);
      row.onclick=openDetails;
      row.onkeydown=event=>{
        if(event.key!=='Enter'&&event.key!==' ')return;
        event.preventDefault();
        openDetails();
      };
    });
  }

  return {render};
})();
