window.AppState=(()=>{
  const DB_NAME='aoi-workbook-store';
  const STORE_NAME='workbooks';
  const ACTIVE_KEY='active';
  const emptyData=()=>({fileName:'',importedAt:'',sheets:[],parseErrors:[],summary:null});
  let data=emptyData();
  let databasePromise;

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

  function openDatabase(){
    if(!window.indexedDB)return Promise.reject(new Error('IndexedDB is unavailable in this browser.'));
    if(!databasePromise){
      databasePromise=new Promise((resolve,reject)=>{
        const request=indexedDB.open(DB_NAME,1);
        request.onupgradeneeded=()=>{if(!request.result.objectStoreNames.contains(STORE_NAME))request.result.createObjectStore(STORE_NAME)};
        request.onsuccess=()=>{request.result.onversionchange=()=>request.result.close();resolve(request.result)};
        request.onerror=()=>reject(request.error||new Error('Could not open workbook storage.'));
        request.onblocked=()=>reject(new Error('Workbook storage is blocked by another browser tab.'));
      });
    }
    return databasePromise;
  }

  function readActiveWorkbook(){
    return openDatabase().then(db=>new Promise((resolve,reject)=>{
      const request=db.transaction(STORE_NAME,'readonly').objectStore(STORE_NAME).get(ACTIVE_KEY);
      request.onsuccess=()=>{data=request.result?updateSummary(request.result):emptyData();resolve(data)};
      request.onerror=()=>reject(request.error||new Error('Could not read the saved workbook.'));
    }));
  }

  function replaceWorkbook(workbook){
    return openDatabase().then(db=>new Promise((resolve,reject)=>{
      const transaction=db.transaction(STORE_NAME,'readwrite');
      transaction.objectStore(STORE_NAME).put(workbook,ACTIVE_KEY);
      transaction.oncomplete=()=>{data=updateSummary(workbook);resolve(data)};
      transaction.onerror=()=>reject(transaction.error||new Error('Could not save the imported workbook.'));
      transaction.onabort=()=>reject(transaction.error||new Error('Workbook save was cancelled.'));
    }));
  }

  function appendRow(sheetIndex,values){
    return openDatabase().then(db=>new Promise((resolve,reject)=>{
      const transaction=db.transaction(STORE_NAME,'readwrite');
      const store=transaction.objectStore(STORE_NAME);
      const request=store.get(ACTIVE_KEY);
      let validationError=null;
      request.onsuccess=()=>{
        const workbook=request.result;
        const sheet=workbook?.sheets?.[sheetIndex];
        if(!sheet){
          validationError=new Error('The selected worksheet is no longer available.');
          transaction.abort();
          return;
        }
        if(values.length!==sheet.columnCount){
          validationError=new Error('The row does not match the worksheet column count.');
          transaction.abort();
          return;
        }
        sheet.matrix.push(values);
        sheet.cellMeta.push(values.map(()=>null));
        sheet.endRow=sheet.startRow+sheet.matrix.length-1;
        sheet.endColumn=sheet.startColumn+sheet.columnCount-1;
        store.put(workbook,ACTIVE_KEY);
      };
      transaction.oncomplete=()=>{data=updateSummary(request.result);resolve(data)};
      transaction.onerror=()=>reject(transaction.error||new Error('Could not save the new row.'));
      transaction.onabort=()=>reject(validationError||transaction.error||new Error('Row creation was cancelled.'));
    }));
  }

  function updateRow(sheetIndex,rowIndex,changes,newColumns={}){
    return openDatabase().then(db=>new Promise((resolve,reject)=>{
      const transaction=db.transaction(STORE_NAME,'readwrite');
      const store=transaction.objectStore(STORE_NAME);
      const request=store.get(ACTIVE_KEY);
      let validationError=null;
      request.onsuccess=()=>{
        const workbook=request.result;
        const sheet=workbook?.sheets?.[sheetIndex];
        if(!sheet||!Number.isInteger(rowIndex)||rowIndex<0||!Array.isArray(sheet.matrix[rowIndex])){
          validationError=new Error('The selected worksheet row is no longer available.');
          transaction.abort();
          return;
        }
        const entries=Object.entries(changes);
        if(entries.some(([column])=>!Number.isInteger(Number(column))||Number(column)<0||Number(column)>=sheet.columnCount)){
          validationError=new Error('One or more edited columns are not available in this worksheet.');
          transaction.abort();
          return;
        }
        const newColumnEntries=Object.entries(newColumns);
        if(newColumnEntries.some(([header,value])=>!header.trim()||typeof value!=='string')){
          validationError=new Error('One or more new worksheet columns are invalid.');
          transaction.abort();
          return;
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
          changedCells.add(XLSX.utils.encode_cell({
            r:sheet.startRow+rowIndex,
            c:sheet.startColumn+columnIndex
          }));
        });
        entries.forEach(([column,value])=>{
          const columnIndex=Number(column);
          sheet.matrix[rowIndex][columnIndex]=value;
          sheet.cellMeta[rowIndex][columnIndex]=null;
          changedCells.add(XLSX.utils.encode_cell({
            r:sheet.startRow+rowIndex,
            c:sheet.startColumn+columnIndex
          }));
        });
        workbook.parseErrors=(workbook.parseErrors||[]).filter(error=>
          error.sheet!==sheet.name||!changedCells.has(error.cell)
        );
        store.put(workbook,ACTIVE_KEY);
      };
      transaction.oncomplete=()=>{data=updateSummary(request.result);resolve(data)};
      transaction.onerror=()=>reject(transaction.error||new Error('Could not save the edited row.'));
      transaction.onabort=()=>reject(validationError||transaction.error||new Error('Row update was cancelled.'));
    }));
  }

  function clearWorkbook(){
    return openDatabase().then(db=>new Promise((resolve,reject)=>{
      const transaction=db.transaction(STORE_NAME,'readwrite');
      transaction.objectStore(STORE_NAME).delete(ACTIVE_KEY);
      transaction.oncomplete=()=>{data=emptyData();resolve(data)};
      transaction.onerror=()=>reject(transaction.error||new Error('Could not clear the saved workbook.'));
      transaction.onabort=()=>reject(transaction.error||new Error('Workbook removal was cancelled.'));
    }));
  }

  function rowsFor(sheet){
    if(!sheet)return [];
    const headerRow=Number.isInteger(sheet.headerRow)?sheet.headerRow:-1;
    return sheet.matrix.reduce((rows,values,index)=>{
      if(headerRow<0||index>headerRow)rows.push({index,sourceRow:sheet.startRow+index+1,values});
      return rows;
    },[]);
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
    let     match=text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:$|[T\s])/);
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
    if(meta?.format&&typeof value==='number'&&XLSX.SSF.is_date(meta.format))return XLSX.SSF.format(meta.format,value);
    const header=String(sheet.headers?.[columnIndex]||'').trim().toLocaleLowerCase();
    if(['date promised','date start schedule','date finish schedule','date finished schedule','last updated bom'].includes(header)){
      const date=formatDateValue(value);
      if(date)return date;
    }
    return String(value);
  }

  const ready=readActiveWorkbook();

  return {
    get data(){return data},
    ready,
    replaceWorkbook,
    appendRow,
    updateRow,
    clearWorkbook,
    rowsFor,
    preambleFor,
    columnIndex,
    displayCell
  };
})();
