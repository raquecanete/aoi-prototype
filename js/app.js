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
  let searchTimer;
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
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
    document.getElementById('sidebar').classList.remove('open');
    const badge=document.querySelector('.demo-badge');
    badge.lastChild.textContent=AppState.data.sheets.length?'Workspace active':'Workspace empty';
    routes[next].render();
    updateCreateButton();
    renderGlobalSearch();
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
    renderGlobalSearch();
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
    clearTimeout(searchTimer);
    if(route==='dashboard')Dashboard.search();
    if(route==='styles')SheetTable.search('styles');
    if(route==='requests')SheetTable.search('requests');
    if(route==='monitoring')SheetTable.search('monitoring');
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
  document.addEventListener('keydown',event=>{
    if(event.key==='Escape'&&!settingsModal.hidden)closeSettings();
    if(event.key==='/'&&document.activeElement?.tagName!=='INPUT'){event.preventDefault();document.getElementById('globalSearch').focus()}
  });
  AppState.ready.then(()=>navigate('dashboard')).catch(error=>{
    document.getElementById('app').innerHTML=`<section class="card storage-error"><strong>Workbook storage is unavailable</strong><p>${esc(error.message)}</p><p>Check browser storage permissions, then refresh the page.</p></section>`;
  });
  return {navigate,toast,refresh};
})();
