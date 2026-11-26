window.AppState=(()=>{
  const USERS={
    admin:{id:'admin',name:'Admin',role:'admin'}
  };
  const emptyData=()=>({fileName:'',importedAt:'',sheets:[],parseErrors:[],summary:null});
  let data=emptyData();
  let activeFile=null;
  let activeVersion=null;
  let currentUser=USERS.admin;
  let dirty=false;
  let draftTimer=null;

  function clone(value){
    return JSON.parse(JSON.stringify(value));
  }

  function makeId(){
    return window.crypto?.randomUUID?crypto.randomUUID():`file-${Date.now()}-${Math.random().toString(36).slice(2,10)}`;
  }

  function versionId(fileId,versionNo){
    return `${fileId}:v${versionNo}`;
  }

  function updateSummary(workbook){
    workbook.sheets.forEach(sheet=>{
      const headerRow=Number.isInteger(sheet.headerRow)?sheet.headerRow:-1;
      sheet.rowCount=sheet.matrix.length-(headerRow<0?0:headerRow+1);
    });
    workbook.summary={
      sheets:workbook.sheets.length,
      rows:workbook.sheets.reduce((total,sheet)=>total+sheet.rowCount,0),
      columns:workbook.sheets.reduce((total,sheet)=>total+sheet.columnCount,0)
    };
    return workbook;
  }

  function seedSheet(name,headers,rows){
    const matrix=[headers,...rows];
    return {
      name,headers:[...headers],matrix,cellMeta:matrix.map(row=>row.map(()=>null)),merges:[],
      startRow:0,endRow:matrix.length-1,startColumn:0,endColumn:headers.length-1,
      headerRow:0,rowCount:rows.length,columnCount:headers.length
    };
  }

  function seedWorkbook(){
    const csiHeaders=['Season','Brand','Style Number','Style Name','Sample Type','Color Code','Status','Date Promised','Notes'];
    const moHeaders=['Document No.','Part No.','Code Raw Material','Material','Color Code Raw Material','Cons','FBC','PO','Status'];
    const data=updateSummary({
      fileName:'Sample - CSI MO AH27.xlsx',
      importedAt:new Date().toISOString(),
      parseErrors:[],
      sheets:[
        seedSheet('CSI',csiHeaders,[
          ['AH27','AOI','0004821','Essential Jersey Tee','Fit Sample','0012','In Progress','2026-12-15','Demo sample record'],
          ['AH27','AOI','0005170','Lightweight Training Short','Size Set','0004','Pending','2026-12-22','Demo sample record']
        ]),
        seedSheet('MO',moHeaders,[
          ['MO-270001','0004821','00001001','Cotton Jersey 180 GSM','0012',1.125,0.875,'#N/A','Open'],
          ['MO-270001','0004821','00002008','Polyester Sewing Thread','0001',0.035,0.025,'001234567','Open'],
          ['MO-270002','0005170','00001044','Recycled Polyester Interlock','0004',0.940,0.720,'001234568','Pending']
        ])
      ]
    });
    return data;
  }

  function requireAdmin(){
    if(currentUser.role!=='admin')throw new Error('Admin permissions are required for this action.');
  }

  function emitChange(){
    window.dispatchEvent(new CustomEvent('appstate:change',{detail:{file:activeFile,dirty,user:currentUser}}));
  }

  function rowCountsFor(workbook){
    return workbook.sheets.map(sheet=>({name:sheet.name,rows:sheet.rowCount,columns:sheet.columnCount}));
  }

  function calculateChangeSummary(previous,next){
    let rowsAdded=0,rowsDeleted=0,rowsEdited=0;
    const fields=new Set();
    const previousSheets=new Map(previous.sheets.map(sheet=>[sheet.name,sheet]));
    next.sheets.forEach(sheet=>{
      const oldSheet=previousSheets.get(sheet.name);
      if(!oldSheet){
        rowsAdded+=sheet.rowCount;
        return;
      }
      const keyHeaders=['Document No.','Part No.','Code Raw Material','Material','Color Code Raw Material'];
      const oldKeys=keyHeaders.map(header=>columnIndex(oldSheet,header));
      const newKeys=keyHeaders.map(header=>columnIndex(sheet,header));
      const useKey=oldKeys.every(index=>index>=0)&&newKeys.every(index=>index>=0);
      const groupRows=(target,keys)=>{
        const groups=new Map();
        rowsFor(target).forEach((row,index)=>{
          const parts=useKey?keys.map(column=>String(row.values[column]??'').trim()):[];
          const key=useKey&&parts.some(Boolean)?parts.join('\u001f'):`row:${index}`;
          if(!groups.has(key))groups.set(key,[]);
          groups.get(key).push(row);
        });
        return groups;
      };
      const oldGroups=groupRows(oldSheet,oldKeys),newGroups=groupRows(sheet,newKeys);
      for(const [key,oldRows] of oldGroups){
        const newRows=newGroups.get(key)||[];
        const shared=Math.min(oldRows.length,newRows.length);
        rowsDeleted+=oldRows.length-shared;
        for(let occurrence=0;occurrence<shared;occurrence++){
          const oldRow=oldRows[occurrence],newRow=newRows[occurrence];
          let changed=false;
          sheet.headers.forEach((header,newColumn)=>{
            const oldColumn=columnIndex(oldSheet,header);
            const oldValue=oldColumn<0?'':oldRow.values[oldColumn]??'';
            const newValue=newRow.values[newColumn]??'';
            if(!Object.is(oldValue,newValue)){
              changed=true;
              fields.add(header||`Column ${newColumn+1}`);
            }
          });
          if(changed)rowsEdited++;
        }
      }
      for(const [key,newRows] of newGroups){
        const oldRows=oldGroups.get(key)||[];
        rowsAdded+=Math.max(0,newRows.length-oldRows.length);
      }
    });
    previous.sheets.forEach(sheet=>{
      if(!next.sheets.some(item=>item.name===sheet.name))rowsDeleted+=sheet.rowCount;
    });
    const details=[];
    if(rowsAdded)details.push(`${rowsAdded} row${rowsAdded===1?'':'s'} added`);
    if(rowsEdited)details.push(`${rowsEdited} row${rowsEdited===1?'':'s'} edited`);
    if(rowsDeleted)details.push(`${rowsDeleted} row${rowsDeleted===1?'':'s'} deleted`);
    if(fields.size)details.push(`fields: ${[...fields].slice(0,8).join(', ')}`);
    return details.length?details.join(' · '):'No row changes';
  }

  function scheduleDraftSave(){
    window.clearTimeout(draftTimer);
    draftTimer=window.setTimeout(()=>persistDraft().catch(error=>{
      window.dispatchEvent(new CustomEvent('storage-warning',{detail:{message:error.message}}));
    }),2000);
  }

  async function persistDraft(){
    if(!activeFile||!dirty)return;
    requireAdmin();
    await AOIDB.put('drafts',{
      fileId:activeFile.id,
      savedAt:new Date().toISOString(),
      savedBy:currentUser.id,
      data:clone(data)
    });
  }

  function markDirty(){
    dirty=true;
    scheduleDraftSave();
    emitChange();
  }

  async function addAudit(action,details=''){
    const current=await AOIDB.get('settings','audit');
    const entries=Array.isArray(current?.value)?current.value:[];
    entries.unshift({id:makeId(),user:currentUser.name,action,details,at:new Date().toISOString()});
    await AOIDB.put('settings',{key:'audit',value:entries.slice(0,50)});
  }

  async function listFiles(){
    const files=await AOIDB.getAll('files');
    return files.sort((a,b)=>String(b.lastOpenedAt||'').localeCompare(String(a.lastOpenedAt||'')));
  }

  async function getFile(fileId){
    return AOIDB.get('files',fileId);
  }

  async function getVersion(fileId,versionNo){
    return AOIDB.get('versions',versionId(fileId,versionNo));
  }

  async function writeNewFile(workbook,options={}){
    const now=new Date().toISOString();
    const fileId=options.id||makeId();
    const name=String(options.name||workbook.fileName||'Imported workbook').trim();
    const file={
      id:fileId,name,originalFileName:options.originalFileName||workbook.fileName||name,
      importedAt:now,lastOpenedAt:now,lastModifiedAt:now,importedBy:options.importedBy||currentUser.id,
      sizeBytes:Number(options.sizeBytes)||new Blob([JSON.stringify(workbook)]).size,
      rowCounts:rowCountsFor(workbook),pinned:Boolean(options.pinned),status:'Active',currentVersion:1
    };
    const version={
      id:versionId(fileId,1),fileId,versionNo:1,createdAt:now,createdBy:currentUser.id,
      note:options.note||'Initial import',changeSummary:'Initial import',data:clone(workbook)
    };
    await AOIDB.put('versions',version);
    await AOIDB.put('files',file);
    await addAudit('Imported',file.name);
    return file;
  }

  async function loadFile(fileId,options={}){
    const file=await getFile(fileId);
    if(!file)throw new Error('This file is no longer available.');
    const version=await getVersion(fileId,file.currentVersion);
    if(!version)throw new Error('The saved version for this file could not be found.');
    let nextData=clone(version.data);
    let nextDirty=false;
    const draft=await AOIDB.get('drafts',fileId);
    if(draft){
      let restore=options.restoreDraft;
      if(typeof restore!=='boolean'){
        restore=window.confirm(`Restore the unsaved draft for "${file.name}" from ${new Date(draft.savedAt).toLocaleString()}?`);
      }
      if(restore){
        nextData=clone(draft.data);
        nextDirty=true;
      }else{
        await AOIDB.delete('drafts',fileId);
      }
    }
    activeFile={...file,lastOpenedAt:new Date().toISOString()};
    activeVersion=version;
    data=updateSummary(nextData);
    dirty=nextDirty;
    await AOIDB.put('files',activeFile);
    await AOIDB.put('settings',{key:'activeFileId',value:fileId});
    await addAudit('Opened',file.name);
    if(dirty)scheduleDraftSave();
    emitChange();
    return data;
  }

  async function bootstrap(){
    await AOIDB.ready;
    const userSetting=await AOIDB.get('settings','currentUser');
    if(userSetting?.value&&USERS[userSetting.value.id])currentUser=USERS[userSetting.value.id];
    let files=await listFiles();
    if(!files.length){
      let legacy=null;
      try{legacy=await readLegacyWorkbook()}catch(error){legacy=null}
      if(legacy?.sheets?.length){
        await writeNewFile(legacy,{name:legacy.fileName,originalFileName:legacy.fileName,note:'Migrated from saved workbook'});
      }else{
        await writeNewFile(seedWorkbook(),{
          name:'Sample - CSI MO AH27',
          originalFileName:'Sample - CSI MO AH27.xlsx',
          importedBy:'system',
          note:'Bundled CSI and MO sample data'
        });
      }
      files=await listFiles();
    }
    const selected=await AOIDB.get('settings','activeFileId');
    const activeId=files.some(file=>file.id===selected?.value)?selected.value:files[0].id;
    await loadFile(activeId);
    await AOIDB.put('settings',{key:'currentUser',value:currentUser});
    return data;
  }

  function readLegacyWorkbook(){
    if(!window.indexedDB)return Promise.resolve(null);
    return new Promise(resolve=>{
      let request;
      try{request=indexedDB.open('aoi-workbook-store',1)}
      catch(error){resolve(null);return}
      request.onupgradeneeded=event=>{
        if(event.oldVersion===0)request.transaction.abort();
      };
      request.onerror=()=>resolve(null);
      request.onsuccess=()=>{
        const db=request.result;
        if(!db.objectStoreNames.contains('workbooks')){
          db.close();
          indexedDB.deleteDatabase('aoi-workbook-store');
          resolve(null);
          return;
        }
        const transaction=db.transaction('workbooks','readonly');
        const get=transaction.objectStore('workbooks').get('active');
        get.onsuccess=()=>{db.close();resolve(get.result||null)};
        get.onerror=()=>{db.close();resolve(null)};
      };
    });
  }

  async function createFile(workbook,options={}){
    requireAdmin();
    const name=String(options.name||workbook.fileName||'Imported workbook').trim();
    if(!name)throw new Error('Enter a name for this file.');
    const duplicate=(await listFiles()).find(file=>file.name.toLocaleLowerCase()===name.toLocaleLowerCase());
    if(duplicate&&!options.allowDuplicate)throw new Error('A file with this name already exists. Choose update or keep both.');
    const file=await writeNewFile(workbook,{...options,name});
    if(options.open!==false)await loadFile(file.id,{restoreDraft:false});
    return file;
  }

  async function updateFile(fileId,workbook,options={}){
    requireAdmin();
    if(dirty&&activeFile?.id===fileId)throw new Error('Save or discard unsaved changes before updating this file.');
    const file=await getFile(fileId);
    if(!file)throw new Error('Choose an existing file to update.');
    const versions=await AOIDB.getVersionsForFile(fileId);
    const latest=versions.find(version=>version.versionNo===file.currentVersion);
    if(!latest)throw new Error('The current file version could not be loaded.');
    const versionNo=file.currentVersion+1;
    const now=new Date().toISOString();
    const updatedFile={
      ...file,lastModifiedAt:now,currentVersion:versionNo,sizeBytes:Number(options.sizeBytes)||new Blob([JSON.stringify(workbook)]).size,
      rowCounts:rowCountsFor(workbook),status:'Active'
    };
    const version={
      id:versionId(fileId,versionNo),fileId,versionNo,createdAt:now,createdBy:currentUser.id,
      note:options.note||'Updated from Excel',
      changeSummary:calculateChangeSummary(latest.data,workbook),data:clone(workbook)
    };
    await AOIDB.put('versions',version);
    await AOIDB.put('files',updatedFile);
    await AOIDB.delete('drafts',fileId);
    await addAudit('Updated from Excel',updatedFile.name);
    if(activeFile?.id===fileId){
      activeFile=updatedFile;
      activeVersion=version;
      data=updateSummary(clone(workbook));
      dirty=false;
      emitChange();
    }
    return {file:updatedFile,version};
  }

  async function openFile(fileId,options={}){
    if(dirty&&activeFile?.id!==fileId&&!options.allowDirty)throw new Error('Save or discard the current unsaved changes before switching files.');
    return loadFile(fileId,options);
  }

  async function saveActive(note=''){
    requireAdmin();
    if(!activeFile||!dirty)return activeVersion;
    const now=new Date().toISOString();
    const versionNo=activeFile.currentVersion+1;
    const version={
      id:versionId(activeFile.id,versionNo),fileId:activeFile.id,versionNo,
      createdAt:now,createdBy:currentUser.id,note:String(note||'').trim(),
      changeSummary:calculateChangeSummary(activeVersion.data,data),data:clone(data)
    };
    const updatedFile={
      ...activeFile,currentVersion:versionNo,lastModifiedAt:now,
      sizeBytes:new Blob([JSON.stringify(data)]).size,rowCounts:rowCountsFor(data)
    };
    await AOIDB.put('versions',version);
    await AOIDB.put('files',updatedFile);
    await AOIDB.delete('drafts',activeFile.id);
    await addAudit('Saved version',`${updatedFile.name} · v${versionNo}`);
    activeFile=updatedFile;
    activeVersion=version;
    dirty=false;
    window.clearTimeout(draftTimer);
    emitChange();
    return version;
  }

  async function discardActive(){
    requireAdmin();
    if(!activeFile)return;
    window.clearTimeout(draftTimer);
    await AOIDB.delete('drafts',activeFile.id);
    const version=await getVersion(activeFile.id,activeFile.currentVersion);
    if(!version)throw new Error('The saved version could not be restored.');
    activeVersion=version;
    data=updateSummary(clone(version.data));
    dirty=false;
    emitChange();
    await addAudit('Discarded changes',activeFile.name);
  }

  async function versionsFor(fileId){
    return AOIDB.getVersionsForFile(fileId);
  }

  async function restoreVersion(fileId,versionNo){
    requireAdmin();
    if(dirty&&activeFile?.id===fileId)throw new Error('Save or discard unsaved changes before restoring a version.');
    const file=await getFile(fileId);
    const source=await getVersion(fileId,versionNo);
    if(!file||!source)throw new Error('The requested version is no longer available.');
    const now=new Date().toISOString();
    const nextNo=file.currentVersion+1;
    const version={
      id:versionId(fileId,nextNo),fileId,versionNo:nextNo,createdAt:now,createdBy:currentUser.id,
      note:`Restored version ${versionNo}`,changeSummary:`Restored version ${versionNo}`,data:clone(source.data)
    };
    const updated={...file,currentVersion:nextNo,lastModifiedAt:now,rowCounts:rowCountsFor(version.data)};
    await AOIDB.put('versions',version);
    await AOIDB.put('files',updated);
    await AOIDB.delete('drafts',fileId);
    await addAudit('Restored version',`${file.name} · v${versionNo} → v${nextNo}`);
    if(activeFile?.id===fileId){
      activeFile=updated;activeVersion=version;data=updateSummary(clone(version.data));dirty=false;emitChange();
    }
    return version;
  }

  async function renameFile(fileId,name){
    requireAdmin();
    const clean=String(name||'').trim();
    if(!clean)throw new Error('File names cannot be blank.');
    const files=await listFiles();
    if(files.some(file=>file.id!==fileId&&file.name.toLocaleLowerCase()===clean.toLocaleLowerCase())){
      throw new Error('A file with this name already exists.');
    }
    const file=await getFile(fileId);
    if(!file)throw new Error('This file is no longer available.');
    const updated={...file,name:clean};
    await AOIDB.put('files',updated);
    await addAudit('Renamed',`${file.name} → ${clean}`);
    if(activeFile?.id===fileId){activeFile=updated;emitChange()}
    return updated;
  }

  async function duplicateFile(fileId,name){
    requireAdmin();
    const source=await getFile(fileId);
    if(!source)throw new Error('This file is no longer available.');
    const version=await getVersion(fileId,source.currentVersion);
    const files=await listFiles();
    const base=String(name||`${source.name} (2)`).trim();
    let newName=base,index=2;
    while(files.some(file=>file.name.toLocaleLowerCase()===newName.toLocaleLowerCase())){
      newName=`${source.name} (${index++})`;
    }
    return createFile(version.data,{
      name:newName,originalFileName:source.originalFileName,sizeBytes:source.sizeBytes,
      importedBy:currentUser.id,note:`Duplicated from ${source.name}`,open:false
    });
  }

  async function setPinned(fileId,pinned){
    requireAdmin();
    const file=await getFile(fileId);
    if(!file)throw new Error('This file is no longer available.');
    const updated={...file,pinned:Boolean(pinned)};
    await AOIDB.put('files',updated);
    await addAudit(pinned?'Pinned':'Unpinned',file.name);
    if(activeFile?.id===fileId)activeFile=updated;
    return updated;
  }

  async function setArchived(fileId,archived){
    requireAdmin();
    const file=await getFile(fileId);
    if(!file)throw new Error('This file is no longer available.');
    const updated={...file,status:archived?'Archived':'Active'};
    await AOIDB.put('files',updated);
    await addAudit(archived?'Archived':'Restored',file.name);
    if(activeFile?.id===fileId)activeFile=updated;
    return updated;
  }

  async function deleteFile(fileId){
    requireAdmin();
    if(dirty&&activeFile?.id===fileId)throw new Error('Save or discard unsaved changes before deleting the active file.');
    const files=await listFiles();
    if(files.length<=1)throw new Error('Keep at least one file in the library.');
    const file=files.find(item=>item.id===fileId);
    if(!file)throw new Error('This file is no longer available.');
    for(const version of await versionsFor(fileId))await AOIDB.delete('versions',version.id);
    await AOIDB.delete('drafts',fileId);
    await AOIDB.delete('files',fileId);
    await addAudit('Deleted',file.name);
    if(activeFile?.id===fileId){
      const next=files.find(item=>item.id!==fileId);
      await loadFile(next.id,{restoreDraft:false});
    }
  }

  async function setCurrentUser(userId){
    if(userId!=='admin')throw new Error('Only the Admin user is available.');
    currentUser=USERS[userId];
    await AOIDB.put('settings',{key:'currentUser',value:currentUser});
    await addAudit('Signed in',currentUser.name);
    emitChange();
    return currentUser;
  }

  async function getAudit(){
    const record=await AOIDB.get('settings','audit');
    return Array.isArray(record?.value)?record.value:[];
  }

  function formatDateValue(value){
    const formatParts=(year,month,day)=>{
      const date=new Date(Date.UTC(year,month-1,day));
      if(date.getUTCFullYear()!==year||date.getUTCMonth()!==month-1||date.getUTCDate()!==day)return null;
      return `${String(year).padStart(4,'0')}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    };
    const fromSerial=serial=>{
      if(!Number.isFinite(serial)||serial<1||serial>2958465)return null;
      const parts=XLSX.SSF.parse_date_code(serial);
      return parts?formatParts(parts.y,parts.m,parts.d):null;
    };
    if(typeof value==='number')return fromSerial(value);
    if(typeof value!=='string')return null;
    const text=value.trim();
    let match=text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:$|[T\s])/);
    if(match)return formatParts(Number(match[1]),Number(match[2]),Number(match[3]));
    match=text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:$|[T\s])/);
    if(match){
      const first=Number(match[1]),second=Number(match[2]),year=Number(match[3]);
      if(first>12)return formatParts(year,second,first);
      if(second>12)return formatParts(year,first,second);
      return null;
    }
    if(/^\d{1,7}(?:\.\d+)?$/.test(text))return fromSerial(Number(text));
    return null;
  }

  function displayCell(sheet,rowIndex,columnIndex){
    const value=sheet.matrix[rowIndex]?.[columnIndex];
    if(value===undefined||value===null||value==='')return '—';
    const meta=sheet.cellMeta?.[rowIndex]?.[columnIndex];
    if(meta?.error)return meta.error;
    const header=String(sheet.headers?.[columnIndex]||'').trim().toLocaleLowerCase();
    if(['date promised','date start schedule','date finish schedule','date finished schedule','last updated bom'].includes(header)){
      const date=formatDateValue(value);
      if(date){
        const parts=date.split('-').map(Number);
        return new Intl.DateTimeFormat('en-GB',{
          day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'
        }).format(new Date(Date.UTC(parts[0],parts[1]-1,parts[2])));
      }
    }
    if(meta?.format&&typeof value==='number')return XLSX.SSF.format(meta.format,value);
    return String(value);
  }

  function rowsFor(sheet){
    if(!sheet)return [];
    const headerRow=Number.isInteger(sheet.headerRow)?sheet.headerRow:-1;
    return sheet.matrix.reduce((rows,values,index)=>{
      if(headerRow<0||index>headerRow)rows.push({index,sourceRow:sheet.startRow+index+1,values});
      return rows;
    },[]);
  }

  function searchRows(sheet,query){
    const text=String(query??'').trim().toLocaleLowerCase();
    const rows=rowsFor(sheet);
    if(!text)return rows;
    return rows.filter(row=>row.values.some((value,column)=>
      String(value??'').toLocaleLowerCase().includes(text)||
      displayCell(sheet,row.index,column).toLocaleLowerCase().includes(text)
    ));
  }

  function preambleFor(sheet){
    if(!sheet||sheet.headerRow<=0)return [];
    return sheet.matrix.slice(0,sheet.headerRow).reduce((rows,values,index)=>{
      if(values.some(value=>value!==''&&value!==null&&value!==undefined)){
        rows.push({index,sourceRow:sheet.startRow+index+1,values});
      }
      return rows;
    },[]);
  }

  function columnIndex(sheet,name){
    const target=String(name).trim().toLocaleLowerCase();
    return sheet.headers.findIndex(header=>String(header).trim().toLocaleLowerCase()===target);
  }

  async function appendRow(sheetIndex,values){
    requireAdmin();
    const sheet=data.sheets[sheetIndex];
    if(!sheet)throw new Error('The selected worksheet is no longer available.');
    if(values.length!==sheet.columnCount)throw new Error('The row does not match the worksheet column count.');
    sheet.matrix.push(values);
    sheet.cellMeta.push(values.map(()=>null));
    sheet.endRow=sheet.startRow+sheet.matrix.length-1;
    markDirty();
    return updateSummary(data);
  }

  async function updateRow(sheetIndex,rowIndex,changes,newColumns={}){
    requireAdmin();
    const sheet=data.sheets[sheetIndex];
    if(!sheet||!Number.isInteger(rowIndex)||rowIndex<0||!Array.isArray(sheet.matrix[rowIndex])){
      throw new Error('The selected worksheet row is no longer available.');
    }
    const entries=Object.entries(changes);
    if(entries.some(([column])=>!Number.isInteger(Number(column))||Number(column)<0||Number(column)>=sheet.columnCount)){
      throw new Error('One or more edited columns are not available in this worksheet.');
    }
    const newColumnEntries=Object.entries(newColumns);
    if(newColumnEntries.some(([header,value])=>!header.trim()||typeof value!=='string')){
      throw new Error('One or more new worksheet columns are invalid.');
    }
    if(!Array.isArray(sheet.cellMeta))sheet.cellMeta=[];
    if(!Array.isArray(sheet.cellMeta[rowIndex]))sheet.cellMeta[rowIndex]=[];
    const changedCells=new Set();
    newColumnEntries.forEach(([header,value])=>{
      let columnIndex=sheet.headers.findIndex(item=>String(item??'').trim().toLocaleLowerCase()===header.trim().toLocaleLowerCase());
      if(columnIndex<0){
        columnIndex=sheet.headers.length;
        sheet.headers.push(header.trim());
        sheet.matrix.forEach((values,index)=>{
          values.push('');
          if(!Array.isArray(sheet.cellMeta[index]))sheet.cellMeta[index]=[];
          sheet.cellMeta[index].push(null);
        });
        if(Number.isInteger(sheet.headerRow)&&sheet.headerRow>=0&&sheet.matrix[sheet.headerRow]){
          sheet.matrix[sheet.headerRow][columnIndex]=header.trim();
        }
        sheet.columnCount=sheet.headers.length;
        sheet.endColumn=sheet.startColumn+sheet.columnCount-1;
      }
      sheet.matrix[rowIndex][columnIndex]=value;
      sheet.cellMeta[rowIndex][columnIndex]=null;
      changedCells.add(XLSX.utils.encode_cell({r:sheet.startRow+rowIndex,c:sheet.startColumn+columnIndex}));
    });
    entries.forEach(([column,value])=>{
      const index=Number(column);
      sheet.matrix[rowIndex][index]=value;
      sheet.cellMeta[rowIndex][index]=null;
      changedCells.add(XLSX.utils.encode_cell({r:sheet.startRow+rowIndex,c:sheet.startColumn+index}));
    });
    data.parseErrors=(data.parseErrors||[]).filter(error=>error.sheet!==sheet.name||!changedCells.has(error.cell));
    updateSummary(data);
    markDirty();
    return data;
  }

  async function replaceWorkbook(workbook){
    return createFile(workbook,{name:workbook.fileName,originalFileName:workbook.fileName});
  }

  async function updateFromWorkbook(fileId,workbook,options={}){
    return updateFile(fileId,workbook,options);
  }

  async function clearWorkbook(){
    requireAdmin();
    throw new Error('Files are managed in the Files library.');
  }

  const ready=bootstrap();
  return {
    get data(){return data},
    get activeFile(){return activeFile},
    get activeVersion(){return activeVersion},
    get activeFileId(){return activeFile?.id||null},
    get dirty(){return dirty},
    get currentUser(){return currentUser},
    ready,requireAdmin,listFiles,getFile,getVersion,versionsFor,openFile,createFile,updateFile,updateFromWorkbook,
    saveActive,discardActive,restoreVersion,renameFile,duplicateFile,setPinned,setArchived,deleteFile,
    setCurrentUser,getAudit,appendRow,updateRow,replaceWorkbook,clearWorkbook,persistDraft,
    rowsFor,searchRows,preambleFor,columnIndex,displayCell,formatDateValue,calculateChangeSummary
  };
})();
