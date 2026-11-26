window.FilesView=(()=>{
  let searchText='';
  let filter='all';
  let sort='recent';
  let view=null;
  let auditEntries=[];
  let auditPage=1;
  let activityRenderId=0;
  const AUDIT_PAGE_SIZE=10;
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  const dateText=value=>value?new Date(value).toLocaleDateString(undefined,{day:'2-digit',month:'short',year:'numeric'}):'—';
  const dateTime=value=>value?new Date(value).toLocaleString():'—';
  const activeView=()=>view||(window.matchMedia('(max-width: 760px)').matches?'cards':'table');

  function rowsForFile(file){
    return (file.rowCounts||[]).reduce((total,sheet)=>total+Number(sheet.rows||0),0);
  }

  function recentTime(file){
    return Math.max(Date.parse(file.lastOpenedAt||'')||0,Date.parse(file.lastModifiedAt||'')||0);
  }

  function matchesFilter(file){
    if(filter==='archived')return file.status==='Archived';
    if(file.status==='Archived')return false;
    if(filter==='pinned')return file.pinned;
    if(filter==='active')return file.status==='Active';
    if(filter==='edited'){
      const date=new Date(file.lastModifiedAt||0);
      const today=new Date();
      return date.toDateString()===today.toDateString();
    }
    return true;
  }

  function sortedFiles(files){
    const filtered=files.filter(file=>
      matchesFilter(file)&&file.name.toLocaleLowerCase().includes(searchText.toLocaleLowerCase())
    );
    return filtered.sort((a,b)=>{
      if(sort==='name')return a.name.localeCompare(b.name,undefined,{sensitivity:'base'});
      if(sort==='size')return b.sizeBytes-a.sizeBytes;
      if(sort==='rows')return rowsForFile(b)-rowsForFile(a);
      return recentTime(b)-recentTime(a);
    });
  }

  async function render(){
    const app=document.getElementById('app');
    app.innerHTML='<div class="files-loading" role="status">Loading file library…</div>';
    try{
      const [files,estimate]=await Promise.all([
        AppState.listFiles(),AOIDB.estimate()
      ]);
      const recent=files.filter(file=>file.status!=='Archived')
        .sort((a,b)=>recentTime(b)-recentTime(a)).slice(0,8);
      const visible=sortedFiles(files);
      const storage=estimate?.quota?`${formatBytes(estimate.usage)} used of ${formatBytes(estimate.quota)}`:`${formatBytes(estimate?.usage||0)} used`;
      app.innerHTML=`<header class="page-heading files-page-heading"><div><p class="eyebrow">WORKBOOK LIBRARY</p><h1>Files</h1><p class="page-subtitle">Find, open, and manage your workbooks.</p></div></header>
      ${recent.length?`<section class="recent-files-section" aria-labelledby="recentFilesTitle"><div class="files-section-heading"><h2 id="recentFilesTitle">Recent</h2><span>Quick access to recently used workbooks</span></div><div class="recent-files-strip">${recent.map(file=>`<button class="recent-file-card ${file.id===AppState.activeFileId?'is-active':''}" type="button" data-open-file="${esc(file.id)}"><span class="file-type-mark">${fileType(file.originalFileName)}</span><span class="recent-file-copy"><strong title="${esc(file.name)}">${esc(file.name)}</strong><small>v${file.currentVersion} · ${rowsForFile(file)} rows</small></span></button>`).join('')}</div></section>`:''}
      <section class="card files-library-panel" aria-labelledby="allFilesTitle">
        <div class="files-library-heading"><div><h2 id="allFilesTitle">All files</h2><span id="filesCount">${visible.length} ${visible.length===1?'file':'files'}</span></div>
          <div class="files-toolbar"><label class="files-search"><span class="sr-only">Search files</span><input id="filesSearch" class="filter-input" type="search" placeholder="Search by file name…" value="${esc(searchText)}"></label><label class="sr-only" for="filesSort">Sort files</label><select id="filesSort" class="select" aria-label="Sort files"><option value="recent">Recently used</option><option value="name">Name</option><option value="size">Size</option><option value="rows">Rows</option></select><div class="files-view-toggle" role="group" aria-label="File display"><button type="button" data-view="table" class="${activeView()==='table'?'selected':''}" aria-pressed="${activeView()==='table'}">List</button><button type="button" data-view="cards" class="${activeView()==='cards'?'selected':''}" aria-pressed="${activeView()==='cards'}">Cards</button></div></div>
        </div>
        <div class="files-filterbar" role="group" aria-label="Filter files">${renderFilters()}</div>
        <div id="filesResults">${visible.length?renderFileRows(visible):renderEmptyState(files.length===0)}</div>
        <div class="storage-meter"><div class="storage-meter-copy"><strong>Browser storage</strong><span>${esc(storage)}</span></div><progress max="100" value="${estimate?.quota?Math.min(100,(estimate.usage/estimate.quota)*100):0}" aria-label="Browser storage usage"></progress></div>
      </section>
      `;
      document.getElementById('filesSort').value=sort;
      bindEvents();
    }catch(error){
      app.innerHTML=`<section class="card storage-error"><strong>Could not load Files</strong><p>${esc(error.message)}</p></section>`;
    }
  }

  async function renderRecentActivity(){
    const app=document.getElementById('app');
    const requestId=++activityRenderId;
    app.innerHTML='<div class="files-loading" role="status">Loading recent activity…</div>';
    try{
      const audit=await AppState.getAudit();
      if(requestId!==activityRenderId||document.querySelector('.nav-item.active')?.dataset.route!=='activity')return;
      auditEntries=audit;
      auditPage=1;
      app.innerHTML=`<header class="page-heading recent-activity-heading">
        <div><p class="eyebrow">WORKSPACE HISTORY</p><h1>Recent Activity</h1><p class="page-subtitle">A chronological log of recent actions in this workspace.</p></div>
        <span class="audit-count">${audit.length} ${audit.length===1?'entry':'entries'}</span>
      </header>
      <section class="card audit-panel recent-activity-panel" aria-label="Recent activity"><div id="auditActivity"></div></section>`;
      renderAuditPage();
    }catch(error){
      if(requestId!==activityRenderId||document.querySelector('.nav-item.active')?.dataset.route!=='activity')return;
      app.innerHTML=`<section class="card storage-error"><strong>Could not load Recent Activity</strong><p>${esc(error.message)}</p></section>`;
    }
  }

  function renderAuditPage(){
    const activity=document.getElementById('auditActivity');
    if(!activity)return;
    if(!auditEntries.length){
      activity.innerHTML='<p class="files-empty-recent">No file activity yet.</p>';
      return;
    }

    const pageCount=Math.ceil(auditEntries.length/AUDIT_PAGE_SIZE);
    auditPage=Math.min(Math.max(auditPage,1),pageCount);
    const start=(auditPage-1)*AUDIT_PAGE_SIZE;
    const pageEntries=auditEntries.slice(start,start+AUDIT_PAGE_SIZE);
    activity.innerHTML=`<div class="audit-list">${pageEntries.map(item=>`<div class="audit-item"><strong>${esc(item.user)}</strong><span>${esc(item.action)}${item.details?` · ${esc(item.details)}`:''}</span><time datetime="${esc(item.at)}">${esc(dateTime(item.at))}</time></div>`).join('')}</div>
      ${pageCount>1?`<nav class="audit-pagination" aria-label="Recent activity pages"><span class="audit-page-status" aria-live="polite">Page ${auditPage} of ${pageCount} · Showing ${start+1}-${start+pageEntries.length} of ${auditEntries.length}</span><div class="audit-pagination-controls"><button type="button" class="select audit-page-button" data-audit-step="-1" aria-label="Previous activity page" ${auditPage===1?'disabled':''}>Previous</button><button type="button" class="select audit-page-button" data-audit-step="1" aria-label="Next activity page" ${auditPage===pageCount?'disabled':''}>Next</button></div></nav>`:''}`;
    activity.querySelectorAll('[data-audit-step]').forEach(button=>{
      button.onclick=()=>{
        auditPage+=Number(button.dataset.auditStep);
        renderAuditPage();
      };
    });
  }

  function renderFilters(){
    const filters=[
      ['all','All files'],
      ['pinned','Pinned'],
      ['active','Active'],
      ['archived','Archived'],
      ['edited','Edited today']
    ];
    return filters.map(([value,label])=>`<button type="button" class="file-filter-chip ${filter===value?'selected':''}" data-filter="${value}" aria-pressed="${filter===value}">${label}</button>`).join('');
  }

  function renderEmptyState(noFiles){
    if(!noFiles)return '<div class="files-empty-state"><strong>No files match these filters</strong><span>Try changing the search or filter.</span></div>';
    return `<div class="files-empty-state"><span class="empty-file-icon" aria-hidden="true">▤</span><strong>No workbooks yet</strong><span>Use the Import control in the top bar to add an Excel or CSV file.</span></div>`;
  }

  function renderFileRows(files){
    if(activeView()==='cards'){
      return `<div class="files-card-grid">${files.map(file=>`<article class="file-library-card"><div class="file-card-top"><span class="file-type-mark">${fileType(file.originalFileName)}</span>${file.pinned?'<span class="pinned-label">Pinned</span>':''}</div><strong class="file-card-name" title="${esc(file.name)}">${esc(file.name)}</strong><span class="file-card-rows">${rowsForFile(file)} rows · ${file.rowCounts?.length||0} sheets</span><div class="file-card-meta"><span>Imported ${dateText(file.importedAt)}</span><span>Modified ${dateText(file.lastModifiedAt)}</span><span>Opened ${dateText(file.lastOpenedAt)}</span><span>${formatBytes(file.sizeBytes)}</span></div><div class="file-card-status"><span>v${file.currentVersion}</span><span class="file-status ${file.status==='Archived'?'archived':''}">${esc(file.status)}</span></div><div class="file-card-actions"><button type="button" class="file-open-button" data-file-action="open" data-id="${esc(file.id)}">Open workbook</button>${renderFileActions(file,false)}</div></article>`).join('')}</div>`;
    }
    return `<div class="grid-scroll files-table-scroll"><table class="data-table files-table"><thead><tr><th>Workbook</th><th>Sheets &amp; rows</th><th>Last modified</th><th>Last opened</th><th>Version</th><th>Status</th><th><span class="sr-only">Actions</span></th></tr></thead><tbody>${files.map(file=>`<tr><td><div class="file-name-cell"><span class="file-type-mark">${fileType(file.originalFileName)}</span><span class="file-name-copy"><strong title="${esc(file.name)}">${esc(file.name)}</strong><small>${esc(file.originalFileName)} · ${dateText(file.importedAt)} · ${formatBytes(file.sizeBytes)}</small></span>${file.pinned?'<span class="pinned-indicator" title="Pinned">★</span>':''}</div></td><td><strong>${rowsForFile(file)}</strong><span class="file-sheet-count">${(file.rowCounts||[]).map(sheet=>`${esc(sheet.name)} ${sheet.rows}`).join(' · ')}</span></td><td>${dateText(file.lastModifiedAt)}</td><td>${dateText(file.lastOpenedAt)}</td><td><span class="file-version">v${file.currentVersion}</span></td><td><span class="file-status ${file.status==='Archived'?'archived':''}">${esc(file.status)}</span></td><td>${renderFileActions(file)}</td></tr>`).join('')}</tbody></table></div>`;
  }

  function renderFileActions(file,includeOpen=true){
    return `<div class="file-row-actions">${includeOpen?`<button type="button" class="file-open-button" data-file-action="open" data-id="${esc(file.id)}">Open</button>`:''}<label class="sr-only" for="fileActions-${esc(file.id)}">More actions for ${esc(file.name)}</label><select id="fileActions-${esc(file.id)}" class="file-more-select" data-file-menu data-id="${esc(file.id)}" aria-label="More actions for ${esc(file.name)}"><option value="">More</option><option value="history">Version history</option><option value="export">Export workbook</option><option value="update">Update from Excel</option><option value="rename">Rename</option><option value="duplicate">Duplicate</option><option value="pin">${file.pinned?'Unpin':'Pin'}</option><option value="archive">${file.status==='Archived'?'Restore':'Archive'}</option><option value="delete">Delete</option></select></div>`;
  }

  function fileType(name){
    const ext=String(name||'').split('.').pop().toUpperCase();
    return ['XLS','XLSX','CSV'].includes(ext)?ext:'FILE';
  }

  function formatBytes(value){
    const bytes=Number(value)||0;
    if(bytes<1024)return `${bytes} B`;
    if(bytes<1024*1024)return `${(bytes/1024).toFixed(1)} KB`;
    return `${(bytes/(1024*1024)).toFixed(1)} MB`;
  }

  function bindEvents(){
    const app=document.getElementById('app');
    app.querySelectorAll('[data-open-file]').forEach(button=>button.onclick=()=>App.openFile(button.dataset.openFile));
    app.querySelectorAll('[data-view]').forEach(button=>button.onclick=()=>{
      view=button.dataset.view;render();
    });
    document.getElementById('filesSearch').oninput=event=>{
      searchText=event.target.value;renderResults();
    };
    app.querySelectorAll('[data-filter]').forEach(button=>button.onclick=()=>{
      filter=button.dataset.filter;
      renderResults();
    });
    document.getElementById('filesSort').onchange=event=>{
      sort=event.target.value;renderResults();
    };
    bindFileActions(app);
  }

  function bindFileActions(root){
    root.querySelectorAll('[data-file-action]').forEach(button=>button.onclick=()=>{
      runAction(button.dataset.fileAction,button.dataset.id);
    });
    root.querySelectorAll('[data-file-menu]').forEach(select=>select.onchange=()=>{
      const action=select.value;
      select.value='';
      if(action)runAction(action,select.dataset.id);
    });
  }

  async function renderResults(){
    const files=await AppState.listFiles();
    const visible=sortedFiles(files);
    document.getElementById('filesResults').innerHTML=visible.length?renderFileRows(visible):renderEmptyState(files.length===0);
    document.getElementById('filesCount').textContent=`${visible.length} ${visible.length===1?'file':'files'}`;
    document.querySelectorAll('.files-filterbar [data-filter]').forEach(button=>{
      const selected=button.dataset.filter===filter;
      button.classList.toggle('selected',selected);
      button.setAttribute('aria-pressed',String(selected));
    });
    bindFileActions(document.getElementById('filesResults'));
  }

  async function runAction(action,fileId){
    const file=await AppState.getFile(fileId);
    if(!file)return;
    try{
      if(action==='open'){
        if(AppState.activeFileId===fileId)App.navigate('dashboard');
        else await App.openFile(fileId);
        return;
      }
      if(action==='history'){await showHistory(file);return}
      if(action==='export'){await exportFile(file);return}
      AppState.requireAdmin();
      if(action==='rename'){
        const name=window.prompt('Rename file:',file.name);
        if(name===null)return;
        await AppState.renameFile(fileId,name);
      }else if(action==='duplicate'){
        const name=window.prompt('Name for the duplicate:',`${file.name} (2)`);
        if(name===null)return;
        await AppState.duplicateFile(fileId,name);
      }else if(action==='pin'){
        await AppState.setPinned(fileId,!file.pinned);
      }else if(action==='archive'){
        await AppState.setArchived(fileId,file.status!=='Archived');
      }else if(action==='delete'){
        const confirmation=window.prompt(`Type "${file.name}" to permanently delete this file and its history:`);
        if(confirmation!==file.name)return;
        await AppState.deleteFile(fileId);
      }else if(action==='update'){
        ImportWizard.setUpdateTarget(fileId);
        App.navigate('import');
        return;
      }
      App.toast('File updated.');
      await render();
      App.updateAppChrome();
    }catch(error){
      App.toast(error.message,'error');
    }
  }

  async function exportFile(file,versionNo=file.currentVersion){
    const version=await AppState.getVersion(file.id,versionNo);
    if(!version)throw new Error('The selected version could not be found.');
    ExportView.downloadWorkbookData(version.data,`${file.name}.xlsx`);
  }

  async function showHistory(file){
    const versions=(await AppState.versionsFor(file.id)).sort((a,b)=>b.versionNo-a.versionNo);
    const drawer=document.getElementById('drawer');
    drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">Version history · ${esc(file.name)}</div><div class="text-xs text-slate-500">${versions.length} saved version${versions.length===1?'':'s'}</div></div><button id="historyClose" class="icon-btn" aria-label="Close version history">×</button></div><div class="drawer-body"><div class="version-history-list">${versions.map(version=>`<article class="version-history-item"><div class="version-history-heading"><strong>Version ${version.versionNo}</strong><time datetime="${esc(version.createdAt)}">${esc(dateTime(version.createdAt))}</time></div><p>${esc(version.createdBy)}${version.note?` · ${esc(version.note)}`:''}</p><small>${esc(version.changeSummary)}</small><div class="version-actions"><button class="select" data-version-action="view" data-version="${version.versionNo}">View</button><button class="select" data-version-action="compare" data-version="${version.versionNo}">Compare with current</button><button class="select" data-version-action="export" data-version="${version.versionNo}">Export</button><button class="select" data-version-action="restore" data-version="${version.versionNo}">Restore</button></div></article>`).join('')}</div><div id="versionPreview"></div></div>`;
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden','false');
    document.getElementById('drawerBackdrop').classList.add('open');
    document.getElementById('historyClose').onclick=Drawer.close;
    drawer.querySelectorAll('[data-version-action]').forEach(button=>{
      button.onclick=()=>runVersionAction(file,Number(button.dataset.version),button.dataset.versionAction);
    });
  }

  async function runVersionAction(file,versionNo,action){
    const version=await AppState.getVersion(file.id,versionNo);
    if(!version)return;
    if(action==='export'){
      await exportFile(file,versionNo);
      return;
    }
    if(action==='restore'){
      AppState.requireAdmin();
      if(!window.confirm(`Restore version ${versionNo}? This creates a new version and keeps the existing history.`))return;
      await AppState.restoreVersion(file.id,versionNo);
      await showHistory(await AppState.getFile(file.id));
      App.toast(`Restored version ${versionNo} as a new version.`);
      return;
    }
    if(action==='view'){
      renderVersionPreview(version.data);
      return;
    }
    if(action==='compare'){
      const current=await AppState.getVersion(file.id,(await AppState.getFile(file.id)).currentVersion);
      renderVersionDiff(version.data,current.data);
    }
  }

  function renderVersionPreview(workbook){
    const sheet=workbook.sheets[0];
    if(!sheet){document.getElementById('versionPreview').innerHTML='<p class="import-no-errors">This version contains no worksheets.</p>';return}
    const rows=AppState.rowsFor(sheet).slice(0,10);
    document.getElementById('versionPreview').innerHTML=`<section class="version-preview"><h3>${esc(sheet.name)} · ${sheet.rowCount} rows</h3><div class="grid-scroll"><table class="data-table"><thead><tr>${sheet.headers.map(header=>`<th>${esc(header)}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${sheet.headers.map((_,column)=>`<td title="${esc(AppState.displayCell(sheet,row.index,column))}">${esc(AppState.displayCell(sheet,row.index,column))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>${sheet.rowCount>rows.length?`<small>Showing the first ${rows.length} rows.</small>`:''}</section>`;
  }

  function keyedRows(sheet){
    const keyHeaders=['Document No.','Part No.','Code Raw Material','Material','Color Code Raw Material'];
    const indices=keyHeaders.map(header=>AppState.columnIndex(sheet,header));
    const hasKey=indices.every(index=>index>=0);
    const groups=new Map();
    AppState.rowsFor(sheet).forEach((row,index)=>{
      const key=hasKey?indices.map(column=>String(row.values[column]??'').trim()).join('\u001f'):`row:${index}`;
      if(!groups.has(key))groups.set(key,[]);
      groups.get(key).push(row);
    });
    return groups;
  }

  function renderVersionDiff(oldData,newData){
    const changes=[];
    oldData.sheets.forEach(oldSheet=>{
      const newSheet=newData.sheets.find(sheet=>sheet.name===oldSheet.name);
      if(!newSheet){changes.push({kind:'Removed sheet',text:oldSheet.name});return}
      const oldGroups=keyedRows(oldSheet),newGroups=keyedRows(newSheet);
      for(const [key,oldRows] of oldGroups){
        const newRows=newGroups.get(key)||[];
        oldRows.forEach((oldRow,index)=>{
          const newRow=newRows[index];
          if(!newRow){changes.push({kind:'Removed',text:key.replace(/\u001f/g,' · ')});return}
          const changed=oldSheet.headers.filter((header,column)=>!Object.is(oldRow.values[column],newRow.values[column]));
          if(changed.length)changes.push({kind:'Changed',text:`${key.replace(/\u001f/g,' · ')} — ${changed.join(', ')}`});
        });
      }
      for(const [key,newRows] of newGroups){
        const oldRows=oldGroups.get(key)||[];
        newRows.slice(oldRows.length).forEach(()=>changes.push({kind:'Added',text:key.replace(/\u001f/g,' · ')}));
      }
    });
    newData.sheets.filter(sheet=>!oldData.sheets.some(item=>item.name===sheet.name))
      .forEach(sheet=>changes.push({kind:'Added sheet',text:sheet.name}));
    document.getElementById('versionPreview').innerHTML=`<section class="version-preview"><h3>Changes compared with current</h3>${changes.length?`<ul class="version-diff-list">${changes.map(change=>`<li class="diff-${change.kind.toLocaleLowerCase().replace(/\s/g,'-')}"><strong>${esc(change.kind)}</strong><span>${esc(change.text)}</span></li>`).join('')}</ul>`:'<p class="import-no-errors">No row changes.</p>'}</section>`;
  }

  return {render,renderRecentActivity,exportFile,keyedRows,renderVersionDiff};
})();
