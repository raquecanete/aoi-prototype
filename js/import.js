window.ImportWizard=(()=>{
  let step=1;
  let queue=[];
  let updateTargetId=null;
  let files=[];
  let importing=false;
  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  const extension=name=>String(name||'').split('.').pop().toLowerCase();
  const baseName=name=>String(name||'Workbook').replace(/\.[^.]+$/,'').trim()||'Workbook';

  function render(){
    const app=document.getElementById('app');
    app.innerHTML=`<div class="page-heading"><div><p class="eyebrow">WORKBOOK DATA</p><h1>Import files</h1><p class="page-subtitle">Import multiple Excel or CSV workbooks independently. Existing files are never silently overwritten.</p></div></div>
    <div class="card wizard"><aside class="step-list"><div class="step ${step===1?'active':''}">1. Upload</div><div class="step ${step===2?'active':''}">2. Review</div><div class="step ${step===3?'active':''}">3. Confirm</div></aside><section id="wizardBody" class="p-5"></section></div>`;
    if(step===1)renderUpload();
    if(step===2)renderReview();
    if(step===3)renderConfirm();
  }

  function renderUpload(){
    const body=document.getElementById('wizardBody');
    body.innerHTML=`<div id="drop" class="upload-zone"><div class="text-lg font-semibold mb-2">Drop Excel or CSV files here</div><div class="text-sm text-slate-500 mb-4">Select multiple .xlsx, .xls, or .csv files. Each file is parsed independently.</div><input id="fileInput" type="file" accept=".xlsx,.xls,.csv" class="hidden" multiple><button id="browse" class="primary-btn">Browse files</button></div>${queue.length?`<div class="import-queue">${renderQueue()}</div><div class="flex justify-between mt-4"><button id="clearQueue" class="select" ${importing?'disabled':''}>Clear queue</button><button id="reviewQueue" class="primary-btn" ${queue.some(job=>job.status==='Ready')?'':'disabled'}>Review ready files</button></div>`:''}`;
    const input=document.getElementById('fileInput');
    document.getElementById('browse').onclick=()=>input.click();
    input.onchange=event=>addFiles(event.target.files);
    const drop=document.getElementById('drop');
    drop.ondragover=event=>{event.preventDefault();drop.classList.add('drag-over')};
    drop.ondragleave=()=>drop.classList.remove('drag-over');
    drop.ondrop=event=>{event.preventDefault();drop.classList.remove('drag-over');addFiles(event.dataTransfer.files)};
    const clear=document.getElementById('clearQueue');
    if(clear)clear.onclick=()=>{queue=[];render()};
    const review=document.getElementById('reviewQueue');
    if(review)review.onclick=()=>{step=2;render()};
  }

  function renderQueue(){
    return `<div class="import-queue-list">${queue.map(job=>`<article class="import-queue-item"><div class="import-queue-head"><strong title="${esc(job.file.name)}">${esc(job.file.name)}</strong><span class="import-status-chip status-${job.status.toLowerCase()}">${esc(job.status)}</span></div><progress max="100" value="${job.progress}" aria-label="Import progress for ${esc(job.file.name)}"></progress>${job.error?`<p class="login-error" role="alert">${esc(job.error)}</p>`:''}${job.workbook?`<p>${job.workbook.sheets.map(sheet=>`${esc(sheet.name)}: ${sheet.rowCount} rows`).join(' · ')}</p><small class="import-queue-validation">${job.validation.errors} errors · ${job.validation.warnings} warnings · ${job.validation.duplicates} possible duplicates</small>`:''}</article>`).join('')}</div>`;
  }

  async function addFiles(fileList){
    const incoming=[...fileList].filter(file=>['xlsx','xls','csv'].includes(extension(file.name)));
    if(!incoming.length){App.toast('Choose .xlsx, .xls, or .csv files.','error');return}
    for(const file of incoming){
      if(queue.some(job=>job.file.name===file.name&&job.file.size===file.size))continue;
      queue.push({file,status:'Queued',progress:0,workbook:null,error:'',mode:'',targetFileId:'',name:baseName(file.name),diff:[],validation:null});
    }
    step=1;
    render();
    await processQueue();
  }

  async function processQueue(){
    if(importing)return;
    importing=true;
    files=await AppState.listFiles();
    for(const job of queue.filter(item=>item.status==='Queued')){
      job.status='Reading';
      job.progress=12;
      render();
      try{
        const buffer=await job.file.arrayBuffer();
        job.progress=35;
        const result=await parseFile(buffer,job.file.name);
        if(!result.sheets?.length)throw new Error('The workbook contains no readable worksheets.');
        job.workbook=result;
        job.validation=validateWorkbook(result);
        job.status='Ready';
        job.progress=65;
        const exact=files.find(file=>file.originalFileName.toLocaleLowerCase()===job.file.name.toLocaleLowerCase());
        const named=files.find(file=>file.name.toLocaleLowerCase()===job.name.toLocaleLowerCase());
        if(updateTargetId){
          job.mode='update';job.targetFileId=updateTargetId;
        }else if(exact){job.mode='update';job.targetFileId=exact.id}
        else if(named){job.mode='';job.targetFileId=''}
        else job.mode='new';
      }catch(error){
        job.status='Failed';
        job.error=`Import failed: ${error.message}`;
        job.progress=100;
      }
      render();
    }
    importing=false;
    render();
  }

  function validateWorkbook(workbook){
    let warnings=0,duplicates=0;
    workbook.sheets.forEach(sheet=>{
      if(sheet.headerRow<0||!sheet.headers.some(Boolean))warnings++;
      const indices=['Document No.','Part No.','Code Raw Material','Material','Color Code Raw Material']
        .map(header=>AppState.columnIndex(sheet,header));
      if(indices.some(index=>index<0))return;
      const seen=new Set();
      AppState.rowsFor(sheet).forEach(row=>{
        const key=indices.map(index=>String(row.values[index]??'').trim()).join('\u001f');
        if(!key.replace(/\u001f/g,'').trim())return;
        if(seen.has(key))duplicates++;
        seen.add(key);
      });
    });
    return {errors:workbook.parseErrors.length,warnings,duplicates};
  }

  function parseFile(buffer,fileName){
    return parseInWorker(buffer,fileName).catch(async workerError=>{
      App.toast(`Worker unavailable for ${fileName}; parsing on this page. ${workerError.message}`,'info');
      await new Promise(resolve=>requestAnimationFrame(resolve));
      return WorkbookParser.parse(buffer,fileName);
    });
  }

  function parseInWorker(buffer,fileName){
    const parserSource=`importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');const parseWorkbook=(${WorkbookParser.parse.toString()});self.onmessage=event=>{try{self.postMessage(parseWorkbook(event.data.buffer,event.data.fileName))}catch(error){self.postMessage({error:error?.message||String(error)})}};`;
    const url=URL.createObjectURL(new Blob([parserSource],{type:'text/javascript'}));
    let worker;
    try{worker=new Worker(url)}
    catch(error){URL.revokeObjectURL(url);return Promise.reject(error)}
    return new Promise((resolve,reject)=>{
      const finish=()=>{worker.terminate();URL.revokeObjectURL(url)};
      worker.onmessage=event=>{
        finish();
        if(event.data.error)reject(new Error(event.data.error));
        else resolve(event.data);
      };
      worker.onerror=event=>{
        event.preventDefault();finish();
        reject(new Error(event.message||'The workbook parsing worker failed.'));
      };
      try{worker.postMessage({buffer,fileName})}
      catch(error){finish();reject(error)}
    });
  }

  function updateHeader(job,sheetIndex,absoluteRow){
    const sheet=job.workbook.sheets[sheetIndex];
    if(absoluteRow!==-1&&(!Number.isInteger(absoluteRow)||absoluteRow<sheet.startRow+1||absoluteRow>sheet.endRow+1)){
      App.toast(`Header row must be between ${sheet.startRow+1} and ${sheet.endRow+1}.`,'error');
      return;
    }
    sheet.headerRow=absoluteRow<0?-1:absoluteRow-sheet.startRow-1;
    sheet.headers=sheet.headerRow<0?Array.from({length:sheet.columnCount},()=> ''):sheet.matrix[sheet.headerRow].map(value=>String(value??'').trim());
    sheet.rowCount=sheet.matrix.length-(sheet.headerRow<0?0:sheet.headerRow+1);
    job.workbook.summary.rows=job.workbook.sheets.reduce((total,item)=>total+item.rowCount,0);
    job.validation=validateWorkbook(job.workbook);
    render();
  }

  function collision(job){
    return files.find(file=>file.name.toLocaleLowerCase()===job.name.toLocaleLowerCase());
  }

  function renderReview(){
    document.getElementById('wizardBody').innerHTML=`<div class="import-review-heading"><div><h2>Review import queue</h2><p>${queue.filter(job=>job.status==='Ready').length} ready · ${queue.filter(job=>job.status==='Failed').length} failed</p></div><label class="select-files-more">Add more files<input id="moreFiles" type="file" accept=".xlsx,.xls,.csv" multiple></label></div>
    ${queue.map((job,jobIndex)=>renderJobReview(job,jobIndex)).join('')}
    <div class="flex justify-between mt-4"><button id="reviewBack" class="select">Back to upload</button><button id="toConfirm" class="primary-btn" ${queue.some(job=>job.status==='Ready')?'':'disabled'}>Continue to confirm</button></div>`;
    const more=document.getElementById('moreFiles');
    more.onchange=event=>addFiles(event.target.files);
    document.querySelectorAll('[data-header-job]').forEach(input=>input.onchange=()=>{
      const job=queue[Number(input.dataset.headerJob)];
      updateHeader(job,Number(input.dataset.headerSheet),input.value===''?-1:Number(input.value));
    });
    document.querySelectorAll('[data-mode-job]').forEach(select=>select.onchange=()=>{
      const job=queue[Number(select.dataset.modeJob)];
      job.mode=select.value;
      render();
    });
    document.querySelectorAll('[data-target-job]').forEach(select=>select.onchange=()=>{
      queue[Number(select.dataset.targetJob)].targetFileId=select.value;
    });
    document.querySelectorAll('[data-name-job]').forEach(input=>input.oninput=()=>{
      queue[Number(input.dataset.nameJob)].name=input.value;
    });
    document.getElementById('reviewBack').onclick=()=>{step=1;render()};
    document.getElementById('toConfirm').onclick=async()=>{await prepareDiffs();step=3;render()};
  }

  function renderJobReview(job,index){
    if(job.status==='Failed')return `<article class="import-review-file"><div class="import-queue-head"><strong>${esc(job.file.name)}</strong><span class="import-status-chip status-failed">Failed</span></div><p class="login-error">${esc(job.error)}</p></article>`;
    if(job.status!=='Ready')return `<article class="import-review-file"><div class="import-queue-head"><strong>${esc(job.file.name)}</strong><span class="import-status-chip status-${job.status.toLowerCase()}">${esc(job.status)}</span></div></article>`;
    const nameCollision=collision(job);
    const sameOriginal=files.find(file=>file.originalFileName.toLocaleLowerCase()===job.file.name.toLocaleLowerCase());
    const choices=nameCollision?`<option value="" disabled ${job.mode===''?'selected':''}>Choose duplicate action…</option><option value="update" ${job.mode==='update'?'selected':''}>Update existing file (new version)</option><option value="new" ${job.mode==='new'?'selected':''}>Keep both (auto-rename)</option><option value="skip" ${job.mode==='skip'?'selected':''}>Cancel / skip</option>`:`<option value="new" ${job.mode==='new'?'selected':''}>Import as new file</option><option value="update" ${job.mode==='update'?'selected':''}>Update existing file</option><option value="skip" ${job.mode==='skip'?'selected':''}>Skip</option>`;
    return `<article class="import-review-file"><div class="import-queue-head"><div><strong>${esc(job.file.name)}</strong><p>${job.workbook.summary.sheets} sheets · ${job.workbook.summary.rows} rows · ${job.workbook.summary.columns} columns</p></div><span class="import-status-chip status-ready">Ready</span></div>
    <div class="import-validation-summary"><span>${job.validation.errors} errors</span><span>${job.validation.warnings} warnings</span><span>${job.validation.duplicates} possible duplicates</span></div>
    <div class="import-options"><label>File name<input class="filter-input" data-name-job="${index}" value="${esc(job.name)}"></label><label>Action<select class="select" data-mode-job="${index}">${choices}</select></label>${job.mode==='update'?`<label>Update target<select class="select" data-target-job="${index}">${files.map(file=>`<option value="${esc(file.id)}" ${file.id===job.targetFileId?'selected':''}>${esc(file.name)} · v${file.currentVersion}</option>`).join('')}</select></label>`:''}</div>
    ${nameCollision?`<p class="duplicate-name-notice">A file named “${esc(job.name)}” already exists. Choose Update, Keep both, or Cancel; nothing will be overwritten without your choice.</p>`:''}
    <details class="import-sheet-details" open><summary>Detected sheets and header rows</summary><div class="grid-scroll import-sheet-list"><table class="data-table"><thead><tr><th>Sheet</th><th>Header row</th><th>Detected headers</th><th>Rows</th><th>Columns</th></tr></thead><tbody>${job.workbook.sheets.map((sheet,sheetIndex)=>`<tr><td>${esc(sheet.name)}</td><td><input class="select" type="number" min="${sheet.startRow+1}" max="${sheet.endRow+1}" value="${sheet.headerRow<0?'':sheet.startRow+sheet.headerRow+1}" data-header-job="${index}" data-header-sheet="${sheetIndex}" aria-label="Header row for ${esc(sheet.name)}" ${sheet.endRow<sheet.startRow?'disabled':''}></td><td title="${esc(sheet.headers.join(' · '))}">${esc(sheet.headers.join(' · '))}</td><td>${sheet.rowCount}</td><td>${sheet.columnCount}</td></tr>`).join('')}</tbody></table></div></details>
    ${job.workbook.parseErrors.length?`<details class="import-errors"><summary>${job.workbook.parseErrors.length} Excel parse issues</summary><ul>${job.workbook.parseErrors.slice(0,100).map(item=>`<li><strong>${esc(item.sheet)}!${esc(item.cell)}</strong> — ${esc(item.message)}</li>`).join('')}</ul></details>`:''}</article>`;
  }

  async function prepareDiffs(){
    files=await AppState.listFiles();
    for(const job of queue){
      if(job.status!=='Ready'||job.mode!=='update')continue;
      const target=files.find(file=>file.id===job.targetFileId);
      if(!target){job.diff=[];job.diffError='Choose an existing update target.';continue}
      const version=await AppState.getVersion(target.id,target.currentVersion);
      if(!version){job.diff=[];job.diffError='The saved version could not be loaded.';continue}
      job.diff=buildDiff(version.data,job.workbook);
      job.diffError='';
      job.targetVersion=version;
    }
  }

  function rowsGrouped(sheet){
    const keys=['Document No.','Part No.','Code Raw Material','Material','Color Code Raw Material'];
    const indices=keys.map(header=>AppState.columnIndex(sheet,header));
    const hasKeys=indices.every(index=>index>=0);
    const groups=new Map();
    AppState.rowsFor(sheet).forEach((row,index)=>{
      const parts=hasKeys?indices.map(column=>String(row.values[column]??'').trim()):[];
      const key=hasKeys&&parts.some(Boolean)?parts.join('\u001f'):`row:${index}`;
      if(!groups.has(key))groups.set(key,[]);
      groups.get(key).push(row);
    });
    return groups;
  }

  function buildDiff(current,incoming){
    const changes=[];
    current.sheets.forEach(oldSheet=>{
      const newSheet=incoming.sheets.find(sheet=>sheet.name===oldSheet.name);
      if(!newSheet){changes.push({kind:'sheet-removed',sheet:oldSheet.name,oldRowIndex:-1});return}
      const oldGroups=rowsGrouped(oldSheet),newGroups=rowsGrouped(newSheet);
      for(const [key,newRows] of newGroups){
        const oldRows=oldGroups.get(key)||[];
        newRows.forEach((newRow,occurrence)=>{
          const oldRow=oldRows[occurrence];
          if(!oldRow){
            changes.push({kind:'added',sheet:oldSheet.name,key,occurrence,values:newRow.values.slice(),headers:newSheet.headers.slice()});
            return;
          }
          const edits=[];
          newSheet.headers.forEach((header,newColumn)=>{
            const oldColumn=AppState.columnIndex(oldSheet,header);
            const oldValue=oldColumn<0?'':oldRow.values[oldColumn]??'';
            const newValue=newRow.values[newColumn]??'';
            if(!Object.is(oldValue,newValue))edits.push({header,value:newValue});
          });
          if(edits.length)changes.push({kind:'changed',sheet:oldSheet.name,key,occurrence,oldRowIndex:oldRow.index,edits});
        });
      }
      for(const [key,oldRows] of oldGroups){
        const newRows=newGroups.get(key)||[];
        oldRows.slice(newRows.length).forEach((row,offset)=>changes.push({
          kind:'removed',sheet:oldSheet.name,key,occurrence:newRows.length+offset,oldRowIndex:row.index
        }));
      }
    });
    incoming.sheets.filter(sheet=>!current.sheets.some(item=>item.name===sheet.name))
      .forEach(sheet=>changes.push({kind:'sheet-added',sheet:sheet.name,data:sheet}));
    return changes;
  }

  function renderDiff(job,index){
    const counts=job.diff.reduce((result,item)=>(result[item.kind]=(result[item.kind]||0)+1,result),{});
    return `<section class="update-diff-preview"><h3>Diff preview · ${esc(job.targetFile?.name||'existing file')}</h3><div class="diff-counts"><span>${counts.added||0} added</span><span>${counts.removed||0} removed</span><span>${counts.changed||0} changed</span></div>${job.diff.length?`<div class="diff-change-list">${job.diff.map((item,changeIndex)=>`<label class="diff-change diff-${item.kind}"><input type="checkbox" data-diff-job="${index}" data-diff-index="${changeIndex}" checked><span><strong>${esc(item.kind.replace('-',' '))}</strong> · ${esc(item.sheet)} · ${esc(item.key||'Worksheet')} ${item.edits?.length?`<small>Changed: ${item.edits.map(edit=>esc(edit.header)).join(', ')}</small>`:''}</span></label>`).join('')}</div>`:'<p class="import-no-errors">No row changes were detected.</p>'}</section>`;
  }

  function renderConfirm(){
    const jobs=queue.filter(job=>job.status==='Ready'&&job.mode!=='skip');
    document.getElementById('wizardBody').innerHTML=`<div class="page-heading"><div><p class="eyebrow">IMPORT SUMMARY</p><h2>Confirm file actions</h2><p class="page-subtitle">Review each file action. Updates create a new version; duplicate names are never overwritten silently.</p></div></div>
    ${jobs.map((job,index)=>{
      const sourceIndex=queue.indexOf(job);
      if(job.mode==='update'){
        const target=files.find(file=>file.id===job.targetFileId);
        job.targetFile=target;
      }
      return `<article class="import-confirm-file"><div class="import-queue-head"><strong>${esc(job.file.name)}</strong><span class="import-status-chip status-ready">${job.mode==='update'?'Update existing':job.mode==='new'?'Import as new':'Choose an action'}</span></div><p>${job.mode==='update'?`Create a new version of ${esc(job.targetFile?.name||'selected file')}.`:`New library file: ${esc(job.name||baseName(job.file.name))}`}</p>${job.mode==='update'?(job.diffError?`<p class="login-error">${esc(job.diffError)}</p>`:renderDiff(job,sourceIndex)):''}</article>`;
    }).join('')||'<p class="import-no-errors">No files selected for import.</p>'}`;
    const readyToImport=jobs.length>0&&jobs.every(job=>
      job.mode==='new'||job.mode==='update'&&job.targetFileId&&!job.diffError
    );
    document.getElementById('wizardBody').innerHTML+=`<div class="import-confirm-actions"><button id="confirmBack" class="select">Back to review</button><button id="runImport" class="primary-btn" ${readyToImport?'':'disabled'}>Import selected files</button></div>`;
    document.getElementById('confirmBack').onclick=()=>{step=2;render()};
    document.querySelectorAll('[data-diff-job]').forEach(input=>input.onchange=()=>{
      queue[Number(input.dataset.diffJob)].diff[Number(input.dataset.diffIndex)].selected=input.checked;
    });
    document.getElementById('runImport').onclick=runImports;
  }

  function ensureUniqueName(name,allFiles){
    let candidate=name,index=2;
    while(allFiles.some(file=>file.name.toLocaleLowerCase()===candidate.toLocaleLowerCase())){
      candidate=`${name} (${index++})`;
    }
    return candidate;
  }

  function ensureColumn(sheet,header){
    let index=sheet.headers.findIndex(item=>String(item).trim().toLocaleLowerCase()===header.trim().toLocaleLowerCase());
    if(index>=0)return index;
    index=sheet.headers.length;
    sheet.headers.push(header);
    sheet.matrix.forEach((row,rowIndex)=>{
      row.push(rowIndex===sheet.headerRow?header:'');
      if(!Array.isArray(sheet.cellMeta[rowIndex]))sheet.cellMeta[rowIndex]=[];
      sheet.cellMeta[rowIndex].push(null);
    });
    sheet.columnCount=sheet.headers.length;
    sheet.endColumn=sheet.startColumn+sheet.columnCount-1;
    return index;
  }

  function applyDiff(base,incoming,diff){
    const result=JSON.parse(JSON.stringify(base));
    const selected=diff.filter(item=>item.selected!==false);
    selected.filter(item=>item.kind==='sheet-added').forEach(item=>{
      result.sheets.push(JSON.parse(JSON.stringify(item.data)));
    });
    selected.filter(item=>item.kind==='sheet-removed').forEach(item=>{
      result.sheets=result.sheets.filter(sheet=>sheet.name!==item.sheet);
    });
    selected.filter(item=>item.kind==='changed').forEach(item=>{
      const sheet=result.sheets.find(current=>current.name===item.sheet);
      if(!sheet||!sheet.matrix[item.oldRowIndex])return;
      item.edits.forEach(edit=>{
        const column=ensureColumn(sheet,edit.header);
        sheet.matrix[item.oldRowIndex][column]=edit.value;
        if(!Array.isArray(sheet.cellMeta[item.oldRowIndex]))sheet.cellMeta[item.oldRowIndex]=[];
        sheet.cellMeta[item.oldRowIndex][column]=null;
      });
    });
    selected.filter(item=>item.kind==='added').forEach(item=>{
      const sheet=result.sheets.find(current=>current.name===item.sheet);
      if(!sheet)return;
      const row=Array.from({length:sheet.headers.length},()=> '');
      item.headers.forEach((header,column)=>{
        row[ensureColumn(sheet,header)]=item.values[column]??'';
      });
      sheet.matrix.push(row);
      sheet.cellMeta.push(row.map(()=>null));
      sheet.endRow=sheet.startRow+sheet.matrix.length-1;
    });
    selected.filter(item=>item.kind==='removed').sort((a,b)=>b.oldRowIndex-a.oldRowIndex).forEach(item=>{
      const sheet=result.sheets.find(current=>current.name===item.sheet);
      if(!sheet||!sheet.matrix[item.oldRowIndex])return;
      sheet.matrix.splice(item.oldRowIndex,1);
      sheet.cellMeta.splice(item.oldRowIndex,1);
      sheet.endRow=sheet.startRow+sheet.matrix.length-1;
    });
    result.importedAt=new Date().toISOString();
    result.fileName=incoming.fileName;
    result.summary.rows=result.sheets.reduce((total,sheet)=>{
      sheet.rowCount=sheet.matrix.length-(sheet.headerRow<0?0:sheet.headerRow+1);
      return total+sheet.rowCount;
    },0);
    result.summary.sheets=result.sheets.length;
    result.summary.columns=result.sheets.reduce((total,sheet)=>total+sheet.columnCount,0);
    return result;
  }

  async function runImports(event){
    const button=event.currentTarget;
    button.disabled=true;
    button.textContent='Importing files…';
    const actionable=queue.filter(job=>job.status==='Ready'&&job.mode!=='skip');
    const imported=[];
    for(const job of actionable){
      job.status='Reading';
      job.progress=75;
      render();
      try{
        if(job.mode==='update'){
          const target=files.find(file=>file.id===job.targetFileId);
          if(!target)throw new Error('Choose a valid existing file to update.');
          const selectedDiff=job.diff.map((item,index)=>({...item,selected:document.querySelector(`[data-diff-job="${queue.indexOf(job)}"][data-diff-index="${index}"]`)?.checked!==false}));
          const base=job.targetVersion?.data;
          if(!base)throw new Error('The file version used for this preview is no longer available.');
          const merged=applyDiff(base,job.workbook,selectedDiff);
          const updated=await AppState.updateFromWorkbook(target.id,merged,{sizeBytes:job.file.size,note:'Updated from Excel'});
          imported.push(updated.file);
        }else if(job.mode==='new'){
          const currentFiles=await AppState.listFiles();
          const name=ensureUniqueName(job.name||baseName(job.file.name),currentFiles);
          const created=await AppState.createFile(job.workbook,{
            name,originalFileName:job.file.name,sizeBytes:job.file.size,open:false,allowDuplicate:false
          });
          imported.push(created);
        }else{
          job.status='Skipped';
          job.progress=100;
          continue;
        }
        job.status='Imported';
        job.progress=100;
      }catch(error){
        job.status='Failed';
        job.error=`Could not import: ${error.message}`;
        job.progress=100;
      }
      render();
    }
    const skipped=queue.filter(job=>job.status==='Ready'&&job.mode==='skip');
    skipped.forEach(job=>{job.status='Skipped';job.progress=100});
    button.disabled=false;
    button.textContent='Import selected files';
    step=2;
    render();
    if(imported.length)App.toast(`Imported ${imported.length} file${imported.length===1?'':'s'}`);
    if(queue.some(job=>job.status==='Failed'))App.toast('One or more files failed. Review their individual error messages.','error');
    await App.updateAppChrome();
    updateTargetId=null;
  }

  function setUpdateTarget(fileId){
    updateTargetId=fileId;
    queue=[];
    step=1;
    render();
    AppState.getFile(fileId).then(file=>App.toast(`Choose the newer Excel file to update ${file?.name||'the selected file'}.`,'info'));
  }

  async function initFiles(){
    files=await AppState.listFiles();
    return files;
  }

  return {render,setUpdateTarget,initFiles,buildDiff,applyDiff};
})();
