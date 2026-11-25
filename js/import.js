window.ImportWizard=(()=>{
  let step=1,workbook=null,parseError='';
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  function render(){
    const app=document.getElementById('app');
    app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Import workbook</h1><p class="page-subtitle">Read every worksheet and preserve its rows, columns, blank cells, and merged ranges.</p></div></div>
    <div class="card wizard"><aside class="step-list"><div class="step ${step===1?'active':''}">1. Upload</div><div class="step ${step===2?'active':''}">2. Review</div><div class="step ${step===3?'active':''}">3. Confirm</div></aside><section id="wizardBody" class="p-5"></section></div>`;
    if(step===1)renderUpload();
    if(step===2)renderReview();
    if(step===3)renderConfirm();
  }
  function renderUpload(){
    const body=document.getElementById('wizardBody');
    body.innerHTML=`<div id="drop" class="upload-zone"><div class="text-lg font-semibold mb-2">Drop an Excel or CSV file here</div><div class="text-sm text-slate-500 mb-4">.xlsx, .xls, or .csv · every sheet and row will be imported</div><input id="fileInput" type="file" accept=".xlsx,.xls,.csv" class="hidden"><button id="browse" class="primary-btn">Browse files</button>${parseError?`<p class="login-error mt-3" role="alert">${esc(parseError)}</p>`:''}</div>`;
    document.getElementById('browse').onclick=()=>document.getElementById('fileInput').click();
    document.getElementById('fileInput').onchange=event=>loadFile(event.target.files[0]);
    const drop=document.getElementById('drop');
    drop.ondragover=event=>{event.preventDefault();drop.style.borderColor='#e34438'};
    drop.ondragleave=()=>{drop.style.borderColor=''};
    drop.ondrop=event=>{event.preventDefault();drop.style.borderColor='';loadFile(event.dataTransfer.files[0])};
  }
  async function loadFile(file){
    if(!file)return;
    parseError='';
    const body=document.getElementById('wizardBody');
    body.innerHTML='<div class="upload-zone" role="status">Reading all workbook sheets…</div>';
    try{
      const buffer=await file.arrayBuffer();
      let result;
      try{
        result=await parseInWorker(buffer,file.name);
      }catch(workerError){
        body.innerHTML='<div class="upload-zone" role="status">Background parsing is unavailable here. Parsing on this page…</div>';
        App.toast(`Background parsing unavailable; using in-page parsing. ${workerError.message}`,'info');
        await new Promise(resolve=>requestAnimationFrame(resolve));
        result=WorkbookParser.parse(buffer,file.name);
      }
      if(!result.sheets?.length)throw new Error('The workbook contains no readable worksheets.');
      workbook=result;
      step=2;
      render();
    }catch(error){
      parseError=`Import failed: ${error.message}`;
      render();
    }
  }
  function parseInWorker(buffer,fileName){
    const parserSource=`importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');const parseWorkbook=(${WorkbookParser.parse.toString()});self.onmessage=event=>{try{self.postMessage(parseWorkbook(event.data.buffer,event.data.fileName))}catch(error){self.postMessage({error:error?.message||String(error)})}};`;
    const url=URL.createObjectURL(new Blob([parserSource],{type:'text/javascript'}));
    let worker;
    try{
      worker=new Worker(url);
    }catch(error){
      URL.revokeObjectURL(url);
      return Promise.reject(error);
    }
    return new Promise((resolve,reject)=>{
      const finish=()=>{worker.terminate();URL.revokeObjectURL(url)};
      worker.onmessage=event=>{
        finish();
        if(event.data.error)reject(new Error(event.data.error));
        else resolve(event.data);
      };
      worker.onerror=event=>{
        event.preventDefault();
        finish();
        reject(new Error(event.message||'The workbook parsing worker failed.'));
      };
      try{worker.postMessage({buffer,fileName})}
      catch(error){finish();reject(error)}
    });
  }
  function renderReview(){
    const summary=workbook.summary;
    document.getElementById('wizardBody').innerHTML=`<div class="import-file-heading"><div><strong>${esc(workbook.fileName)}</strong><span>Review detected content before replacing the current workbook. Adjust a header row when the worksheet has title rows above its column names.</span></div></div>
    <div class="kpi-grid import-summary-grid mb-4"><div class="kpi card"><div class="kpi-value">${summary.sheets}</div><div class="kpi-label">Sheets</div></div><div class="kpi card"><div class="kpi-value">${summary.rows}</div><div class="kpi-label">Data rows</div></div><div class="kpi card"><div class="kpi-value">${summary.columns}</div><div class="kpi-label">Columns</div></div><div class="kpi card"><div class="kpi-value">${workbook.parseErrors.length}</div><div class="kpi-label">Parse issues</div></div></div>
    <div class="grid-scroll import-sheet-list"><table class="data-table"><thead><tr><th>Worksheet</th><th>Header row (Excel)</th><th>Detected headers</th><th>Rows</th><th>Columns</th><th>Merged ranges</th></tr></thead><tbody>${workbook.sheets.map((sheet,index)=>`<tr><td>${esc(sheet.name)}</td><td><input class="select" type="number" min="${sheet.startRow+1}" max="${sheet.endRow+1}" value="${sheet.headerRow<0?'':sheet.startRow+sheet.headerRow+1}" data-sheet-header="${index}" aria-label="Header row for ${esc(sheet.name)}" ${sheet.endRow<sheet.startRow?'disabled':''}></td><td title="${esc(sheet.headers.join(' · '))}">${esc(sheet.headers.join(' · '))}</td><td>${sheet.rowCount}</td><td>${sheet.columnCount}</td><td>${sheet.merges.length}</td></tr>`).join('')}</tbody></table></div>
    ${renderErrors(workbook.parseErrors)}
    <div class="flex justify-between mt-4"><button id="importBack" class="select">Choose another file</button><button id="importContinue" class="primary-btn">Review import</button></div>`;
    document.querySelectorAll('[data-sheet-header]').forEach(input=>input.onchange=()=>{
      const sheet=workbook.sheets[Number(input.dataset.sheetHeader)];
      const absoluteRow=input.value===''?-1:Number(input.value);
      if(absoluteRow!==-1&&(!Number.isInteger(absoluteRow)||absoluteRow<sheet.startRow+1||absoluteRow>sheet.endRow+1)){
        input.value=sheet.headerRow<0?'':sheet.startRow+sheet.headerRow+1;
        App.toast(`Header row must be between ${sheet.startRow+1} and ${sheet.endRow+1}.`,'error');
        return;
      }
      sheet.headerRow=absoluteRow<0?-1:absoluteRow-sheet.startRow-1;
      sheet.headers=sheet.headerRow<0?Array.from({length:sheet.columnCount},()=> ''):sheet.matrix[sheet.headerRow].map(value=>String(value??'').trim());
      sheet.rowCount=sheet.matrix.length-(sheet.headerRow<0?0:sheet.headerRow+1);
      workbook.summary.rows=workbook.sheets.reduce((total,item)=>total+item.rowCount,0);
      renderReview();
    });
    document.getElementById('importBack').onclick=()=>{step=1;render()};
    document.getElementById('importContinue').onclick=()=>{step=3;render()};
  }
  function renderConfirm(){
    const summary=workbook.summary;
    document.getElementById('wizardBody').innerHTML=`<div class="page-heading"><div><p class="eyebrow">REPLACE WORKBOOK DATA</p><h2>Confirm import</h2><p class="page-subtitle">Importing this file replaces the previously saved workbook on every page.</p></div></div>
    <div class="card import-confirm-card"><strong>${esc(workbook.fileName)}</strong><div>${summary.sheets} sheets · ${summary.rows} data rows · ${summary.columns} columns · ${workbook.parseErrors.length} parse issues</div></div>
    ${renderErrors(workbook.parseErrors)}
    <div class="flex justify-between mt-4"><button id="confirmBack" class="select">Back</button><button id="confirmImport" class="primary-btn">Import workbook</button></div>`;
    document.getElementById('confirmBack').onclick=()=>{step=2;render()};
    document.getElementById('confirmImport').onclick=async event=>{
      const button=event.currentTarget;
      button.disabled=true;button.textContent='Saving workbook…';
      try{
        await AppState.replaceWorkbook(workbook);
        App.toast(`Imported ${summary.sheets} sheets and ${summary.rows} rows`);
        workbook=null;step=1;App.navigate('dashboard');
      }catch(error){
        button.disabled=false;button.textContent='Import workbook';
        App.toast(`Could not save workbook: ${error.message}`,'error');
      }
    };
  }
  function renderErrors(errors){
    if(!errors.length)return '<p class="import-no-errors">No Excel error cells were found.</p>';
    return `<details class="import-errors" open><summary>${errors.length} parse issue${errors.length===1?'':'s'} detected</summary><ul>${errors.map(item=>`<li><strong>${esc(item.sheet)}!${esc(item.cell)}</strong> — ${esc(item.message)}</li>`).join('')}</ul></details>`;
  }
  return {render};
})();
