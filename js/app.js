window.App=(()=>{
  function renderMonitoringPlaceholder(){
    document.getElementById('app').innerHTML=`<div class="page-heading"><div><p class="eyebrow">IN PROGRESS</p><h1>Sample Room Monitoring</h1></div></div><section class="card empty-export" role="status"><strong>This page is still being developed.</strong></section>`;
  }
  const routes={
    dashboard:{title:'Dashboard',render:()=>Dashboard.render()},
    styles:{title:'Style Management',render:()=>SheetTable.render('styles')},
    requests:{title:'Imported File',render:()=>SheetTable.render('requests')},
    'sample-requests':{title:'Sample Request Management',render:()=>SampleRequests.render()},
    monitoring:{title:'Sample Room Monitoring',render:renderMonitoringPlaceholder},
    quality:{title:'Data Quality',render:renderQuality},
    files:{title:'Files',render:()=>FilesView.render()},
    export:{title:'Export',render:()=>ExportView.render()},
    activity:{title:'Recent Activity',render:()=>FilesView.renderRecentActivity()},
    import:{title:'Import workbook',render:()=>ImportWizard.render()}
  };
  let route='dashboard';
  let searchTimer;
  let switcherLabelLimit=null;
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  function getSwitcherLabelLimit(){
    if(window.matchMedia('(max-width: 420px)').matches)return 20;
    if(window.matchMedia('(max-width: 760px)').matches)return 28;
    if(window.matchMedia('(max-width: 1023px)').matches)return 36;
    return Infinity;
  }
  function renderSearchSuggestions(matches,query){
    const input=document.getElementById('globalSearch');
    const suggestions=document.getElementById('searchSuggestions');
    if(!query||document.activeElement!==input){
      suggestions.hidden=true;
      suggestions.innerHTML='';
      input.setAttribute('aria-expanded','false');
      return;
    }
    if(!AppState.data.sheets.length){
      suggestions.innerHTML='<div class="search-suggestions-empty">Import a workbook to search its data.</div>';
    }else if(!matches.length){
      suggestions.innerHTML='<div class="search-suggestions-empty">No matching records.</div>';
    }else{
      suggestions.innerHTML=matches.slice(0,8).map(({sheet,row,cells},index)=>
        `<button type="button" class="search-suggestion" role="option" aria-selected="false" data-search-suggestion="${index}"><span class="search-suggestion-title">${esc(sheet.name)} · Row ${row.sourceRow}</span><span class="search-suggestion-detail">${cells.slice(0,2).map(({header,value})=>`${esc(header)}: ${esc(value)}`).join(' · ')}</span></button>`
      ).join('');
      suggestions.querySelectorAll('[data-search-suggestion]').forEach(button=>button.onclick=()=>{
        const match=matches[Number(button.dataset.searchSuggestion)];
        input.blur();
        suggestions.hidden=true;
        input.setAttribute('aria-expanded','false');
        navigate('requests');
        SheetTable.selectSheet('requests',match.sheetIndex);
      });
    }
    suggestions.hidden=false;
    input.setAttribute('aria-expanded','true');
  }
  function navigate(next){
    if(!routes[next])return;
    route=next;
    document.querySelectorAll('.nav-item').forEach(button=>button.classList.toggle('active',button.dataset.route===next));
    document.getElementById('breadcrumbs').textContent=`AOI Materials  /  ${routes[next].title}`;
    routes[next].render();
    updateCreateButton();
    renderGlobalSearch();
    window.scrollTo(0,0);
  }
  function updateCreateButton(){
    const create=document.getElementById('createRecord');
    create.hidden=route==='sample-requests'||route==='monitoring';
    create.disabled=!AppState.data.sheets.some(sheet=>sheet.columnCount>0);
    document.getElementById('topImport').hidden=false;
    document.getElementById('saveWorkbook').hidden=!AppState.dirty;
    document.getElementById('discardWorkbook').hidden=!AppState.dirty;
  }
  async function updateAppChrome(){
    const files=await AppState.listFiles();
    const recent=files.filter(file=>file.status!=='Archived').sort((a,b)=>{
      const recentTime=file=>Math.max(Date.parse(file.lastOpenedAt||'')||0,Date.parse(file.lastModifiedAt||'')||0);
      return recentTime(b)-recentTime(a);
    }).slice(0,8);
    const active=AppState.activeFile;
    const control=document.getElementById('activeFileControl');
    const switcher=document.getElementById('activeFileSwitcher');
    const maxLength=getSwitcherLabelLimit();
    switcherLabelLimit=maxLength;
    const optionLabel=file=>{
      const name=String(file.name||'');
      const compactName=name.length>maxLength?`${name.slice(0,maxLength-1).trimEnd()}…`:name;
      return `<option value="${esc(file.id)}" title="${esc(name)}">${esc(compactName)} · v${file.currentVersion}</option>`;
    };
    control.hidden=!active;
    if(active){
      switcher.innerHTML=recent.map(optionLabel).join('');
      if(!recent.some(file=>file.id===active.id)){
        switcher.insertAdjacentHTML('beforeend',optionLabel(active));
      }
      switcher.value=active.id;
      switcher.title=active.name;
      switcher.setAttribute('aria-label',`Switch active workbook. Current workbook: ${active.name}`);
      document.getElementById('activeFileMeta').textContent=AppState.dirty?'• Unsaved changes':`v${active.currentVersion}`;
    }
    updateCreateButton();
  }
  function chooseUnsavedAction(){
    const dialog=document.getElementById('unsavedDialog');
    return new Promise(resolve=>{
      let settled=false;
      const finish=value=>{
        if(settled)return;
        settled=true;
        dialog.hidden=true;
        resolve(value);
      };
      document.getElementById('switchSave').onclick=()=>finish('save');
      document.getElementById('switchDiscard').onclick=()=>finish('discard');
      document.getElementById('switchCancel').onclick=()=>finish('cancel');
      dialog.onclick=event=>{if(event.target===dialog)finish('cancel')};
      dialog.hidden=false;
      dialog.querySelector('.settings-dialog').focus();
    });
  }
  async function switchFile(fileId,openDashboard=false){
    if(fileId===AppState.activeFileId)return true;
    try{
      if(AppState.dirty){
        const action=await chooseUnsavedAction();
        if(action==='cancel')return false;
        if(action==='save'){
          const note=window.prompt('Optional note for this version:','');
          if(note===null)return false;
          await AppState.saveActive(note);
        }else{
          await AppState.discardActive();
        }
      }
      await AppState.openFile(fileId);
      await updateAppChrome();
      if(openDashboard)navigate('dashboard');
      else refresh();
      return true;
    }catch(error){
      toast(`Could not open file: ${error.message}`,'error');
      await updateAppChrome();
      return false;
    }
  }
  function activeSheetIndex(){
    if(['styles','requests','monitoring'].includes(route))return SheetTable.currentSheetIndex(route);
    if(route==='export')return Number(document.getElementById('csvSheet')?.value||0);
    return 0;
  }
  function refresh(){
    routes[route].render();
    updateCreateButton();
    renderGlobalSearch();
    updateAppChrome();
  }
  function toast(message,type='success'){
    const element=document.createElement('div');
    element.className=`toast${type==='error'?' toast-error':''}`;
    element.textContent=message;
    document.getElementById('toastRegion').appendChild(element);
    setTimeout(()=>element.remove(),5000);
  }
  function renderGlobalSearch(){
    const query=(document.getElementById('globalSearch').value||'').trim();
    if(!query){
      renderSearchSuggestions([],query);
      return;
    }
    const foldedQuery=query.toLocaleLowerCase();
    const matches=[];
    AppState.data.sheets.forEach((sheet,sheetIndex)=>{
      AppState.searchRows(sheet,query).forEach(row=>{
        const cells=row.values.reduce((found,value,column)=>{
          const raw=String(value??'').toLocaleLowerCase();
          const display=AppState.displayCell(sheet,row.index,column);
          if(raw.includes(foldedQuery)||display.toLocaleLowerCase().includes(foldedQuery)){
            found.push({header:sheet.headers[column]||`Column ${column+1}`,value:display});
          }
          return found;
        },[]);
        matches.push({sheet,sheetIndex,row,cells});
      });
    });
    renderSearchSuggestions(matches,query);
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
  document.getElementById('activeFileSwitcher').onchange=event=>{
    const target=event.target.value;
    switchFile(target).then(opened=>{
      if(!opened)event.target.value=AppState.activeFileId;
    });
  };
  document.getElementById('saveWorkbook').onclick=async()=>{
    const note=window.prompt('Optional note for this version:','');
    if(note===null)return;
    try{
      await AppState.saveActive(note);
      toast(`Saved ${AppState.activeFile.name} as version ${AppState.activeVersion.versionNo}`);
      refresh();
    }catch(error){toast(`Could not save changes: ${error.message}`,'error')}
  };
  document.getElementById('discardWorkbook').onclick=async()=>{
    if(!window.confirm('Discard all unsaved changes to this file?'))return;
    try{
      await AppState.discardActive();
      toast('Unsaved changes discarded.');
      refresh();
    }catch(error){toast(`Could not discard changes: ${error.message}`,'error')}
  };
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
    AOIDB.get('settings','uiPreferences').then(current=>{
      const value={...(current?.value||{}),theme:themeToggle.checked?'dark':'light'};
      return AOIDB.put('settings',{key:'uiPreferences',value});
    }).catch(error=>toast(`Could not save appearance preference: ${error.message}`,'error'));
  };
  document.getElementById('globalSearch').addEventListener('input',()=>{
    clearTimeout(searchTimer);
    if(route==='dashboard')Dashboard.search();
    if(route==='styles')SheetTable.search('styles');
    if(route==='requests')SheetTable.search('requests');
    if(route==='quality')renderQuality();
    searchTimer=setTimeout(renderGlobalSearch,180);
  });
  document.getElementById('globalSearch').addEventListener('focus',()=>renderGlobalSearch());
  document.getElementById('globalSearch').addEventListener('blur',()=>{
    setTimeout(()=>{
      const suggestions=document.getElementById('searchSuggestions');
      if(!suggestions.contains(document.activeElement)){
        suggestions.hidden=true;
        document.getElementById('globalSearch').setAttribute('aria-expanded','false');
      }
    },100);
  });
  document.getElementById('globalSearch').addEventListener('keydown',event=>{
    if(event.key==='ArrowDown'&&!document.getElementById('searchSuggestions').hidden){
      event.preventDefault();
      document.querySelector('#searchSuggestions [data-search-suggestion]')?.focus();
    }else if(event.key==='Escape'){
      document.getElementById('searchSuggestions').hidden=true;
      document.getElementById('globalSearch').setAttribute('aria-expanded','false');
    }
  });
  document.getElementById('searchSuggestions').addEventListener('keydown',event=>{
    const options=[...document.querySelectorAll('#searchSuggestions [data-search-suggestion]')];
    const current=options.indexOf(document.activeElement);
    if(event.key==='ArrowDown'&&current<options.length-1){
      event.preventDefault();
      options[current+1].focus();
    }else if(event.key==='ArrowUp'){
      event.preventDefault();
      if(current>0)options[current-1].focus();
      else document.getElementById('globalSearch').focus();
    }
  });
  window.addEventListener('resize',()=>{
    const nextLimit=getSwitcherLabelLimit();
    if(nextLimit===switcherLabelLimit)return;
    switcherLabelLimit=nextLimit;
    updateAppChrome().catch(error=>toast(`Could not refresh the workbook switcher: ${error.message}`,'error'));
  });
  document.addEventListener('keydown',event=>{
    if(event.key==='Escape'&&!settingsModal.hidden)closeSettings();
    if(event.key==='Escape'&&!document.getElementById('unsavedDialog').hidden){
      document.getElementById('unsavedDialog').hidden=true;
    }
    if(event.key==='/'&&document.activeElement?.tagName!=='INPUT'){event.preventDefault();document.getElementById('globalSearch').focus()}
  });
  window.addEventListener('appstate:change',()=>{
    updateAppChrome().catch(error=>toast(`Could not refresh file status: ${error.message}`,'error'));
    if(route==='files')FilesView.render();
    if(route==='activity')FilesView.renderRecentActivity();
  });
  window.addEventListener('storage-warning',event=>{
    const warning=document.getElementById('storageWarning');
    warning.hidden=false;
    if(event.detail?.message)warning.textContent=`Browser storage failed. Changes are kept in memory only: ${event.detail.message}`;
  });
  AppState.ready.then(async()=>{
    const preferences=await AOIDB.get('settings','uiPreferences');
    const dark=preferences?.value?.theme==='dark';
    document.body.classList.toggle('dark',dark);
    themeToggle.checked=dark;
    document.getElementById('themeState').textContent=dark?'Dark theme is active':'Light theme is active';
    navigate('dashboard');
  }).catch(error=>{
    document.getElementById('app').innerHTML=`<section class="card storage-error"><strong>Workbook storage is unavailable</strong><p>${esc(error.message)}</p><p>Check browser storage permissions, then refresh the page.</p></section>`;
  });
  return {navigate,toast,refresh,switchFile,openFile:fileId=>switchFile(fileId,true),updateAppChrome};
})();
