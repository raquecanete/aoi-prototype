window.Drawer=(()=>{
  function esc(value){return String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]))}
  const editOptions={
    Brand:['Adidas','PUMA','JAKO'],
    RBU:['Football','Originals','Olympics','Teamsports','Denim','Running','Motorsports'],
    'Dev Type':['New','Carry-over']
  };
  function editControl(header,sourceIndex,value){
    const binding=sourceIndex<0?`data-new-column="${esc(header)}"`:`data-edit-column="${sourceIndex}"`;
    const options=editOptions[header];
    if(!options)return `<input class="filter-input" ${binding} value="${esc(value)}">`;
    const choices=value&&!options.includes(value)?[value,...options]:options;
    return `<select class="select" ${binding}><option value="" ${value===''?'selected':''}>Select…</option>${choices.map(option=>`<option value="${esc(option)}" ${value===option?'selected':''}>${esc(option)}</option>`).join('')}</select>`;
  }
  function open(sheetIndex,rowIndex,displayColumns=null){
    const sheet=AppState.data.sheets[sheetIndex],row=AppState.rowsFor(sheet).find(item=>item.index===rowIndex);
    if(!sheet||!row)return;
    const fields=displayColumns
      ?displayColumns.map(({header,sourceIndex})=>({header,value:sourceIndex<0?'—':AppState.displayCell(sheet,row.index,sourceIndex)}))
      :sheet.headers.map((header,column)=>({header:header||'—',value:AppState.displayCell(sheet,row.index,column)}));
    const drawer=document.getElementById('drawer');
    drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">${esc(sheet.name)}</div><div class="text-xs text-slate-500">Source row ${row.sourceRow}</div></div><div class="drawer-head-actions"><button id="drawerEdit" class="select">Edit row</button><button id="drawerClose" class="icon-btn" aria-label="Close row details">×</button></div></div><div class="drawer-body">${fields.map(({header,value})=>`<div class="field-grid"><div class="field-label">${esc(header||'—')}</div><div class="field-value">${esc(value)}</div></div>`).join('')}</div>`;
    drawer.classList.add('open');drawer.setAttribute('aria-hidden','false');document.getElementById('drawerBackdrop').classList.add('open');
    document.getElementById('drawerClose').onclick=close;
    document.getElementById('drawerEdit').onclick=()=>openEdit(sheetIndex,rowIndex,displayColumns);
  }
  function openEdit(sheetIndex,rowIndex,displayColumns=null){
    const sheet=AppState.data.sheets[sheetIndex];
    if(!sheet||!Array.isArray(sheet.matrix[rowIndex]))return;
    const columns=displayColumns||sheet.headers.map((header,sourceIndex)=>({header,sourceIndex}));
    const description=displayColumns
      ?'Blank fields are editable. A source column is added when you save a value for a field not in the workbook.'
      :'Update any source column. Changes are saved to this workbook.';
    const initialValues=sheet.headers.map((_,column)=>{
      const value=sheet.matrix[rowIndex][column];
      return value===undefined||value===null||value===''?'':AppState.displayCell(sheet,rowIndex,column);
    });
    const drawer=document.getElementById('drawer');
    drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">Edit row · ${esc(sheet.name)}</div><div class="text-xs text-slate-500">${esc(description)}</div></div><button id="drawerClose" class="icon-btn" aria-label="Close row editor">×</button></div><div class="drawer-body"><div class="edit-row-form">${columns.map(({header,sourceIndex})=>{const value=sourceIndex<0?'':initialValues[sourceIndex];return `<label class="edit-row-field"><span>${esc(header||`Column ${sourceIndex+1}`)}</span>${editControl(header||`Column ${sourceIndex+1}`,sourceIndex,value)}</label>`}).join('')}</div><div class="create-row-actions"><button id="editCancel" class="select">Cancel</button><button id="editSave" class="primary-btn">Save changes</button></div></div>`;
    drawer.classList.add('open');drawer.setAttribute('aria-hidden','false');document.getElementById('drawerBackdrop').classList.add('open');
    document.getElementById('drawerClose').onclick=close;
    document.getElementById('editCancel').onclick=()=>open(sheetIndex,rowIndex,displayColumns);
    const saveButton=document.getElementById('editSave');
    const updateSaveState=()=>{
      const hasChanges=[...drawer.querySelectorAll('[data-edit-column]')].some(field=>field.value!==initialValues[Number(field.dataset.editColumn)]);
      const hasNewValues=[...drawer.querySelectorAll('[data-new-column]')].some(field=>field.value!=='');
      saveButton.disabled=!hasChanges&&!hasNewValues;
    };
    drawer.querySelectorAll('[data-edit-column], [data-new-column]').forEach(input=>{
      input.oninput=updateSaveState;
      input.onchange=updateSaveState;
    });
    saveButton.disabled=true;
    saveButton.onclick=()=>saveEdit(sheetIndex,rowIndex,initialValues);
  }
  async function saveEdit(sheetIndex,rowIndex,initialValues){
    const drawer=document.getElementById('drawer');
    const changes={};
    const newColumns={};
    drawer.querySelectorAll('[data-edit-column]').forEach(input=>{
      const column=Number(input.dataset.editColumn);
      if(input.value!==initialValues[column])changes[column]=input.value;
    });
    drawer.querySelectorAll('[data-new-column]').forEach(input=>{
      if(input.value!=='')newColumns[input.dataset.newColumn]=input.value;
    });
    if(!Object.keys(changes).length&&!Object.keys(newColumns).length)return;
    const button=document.getElementById('editSave');
    button.disabled=true;button.textContent='Saving…';
    try{
      await AppState.updateRow(sheetIndex,rowIndex,changes,newColumns);
      const count=Object.keys(changes).length+Object.keys(newColumns).length;
      App.toast(`Updated ${count} column${count===1?'':'s'}`);
      close();
      App.refresh();
    }catch(error){
      button.disabled=false;button.textContent='Save changes';
      App.toast(`Could not update row: ${error.message}`,'error');
    }
  }
  function openCreate(initialSheetIndex=0){
    if(!AppState.data.sheets.length){
      App.toast('Import a workbook before creating a row.','error');
      return;
    }
    let selected=AppState.data.sheets[initialSheetIndex]?initialSheetIndex:0;
    const drafts=new Map();
    const drawer=document.getElementById('drawer');
    function saveDraft(){
      drafts.set(selected,[...drawer.querySelectorAll('[data-create-column]')].map(input=>input.value));
    }
    function renderForm(){
      const sheet=AppState.data.sheets[selected];
      const draft=drafts.get(selected)||[];
      drawer.innerHTML=`<div class="drawer-head"><div><div class="font-semibold">Create workbook row</div><div class="text-xs text-slate-500">Values are saved to the selected worksheet.</div></div><button id="drawerClose" class="icon-btn" aria-label="Close create form">×</button></div><div class="drawer-body"><label class="create-sheet-label" for="createSheet">Worksheet</label><select id="createSheet" class="select">${AppState.data.sheets.map((item,index)=>`<option value="${index}" ${index===selected?'selected':''}>${esc(item.name)}</option>`).join('')}</select>${sheet.headers.length?`<div class="create-row-form">${sheet.headers.map((header,column)=>`<label class="create-row-field"><span>${esc(header||'—')}</span><input class="filter-input" data-create-column="${column}" value="${esc(draft[column]||'')}"></label>`).join('')}</div>`:'<p class="import-no-errors">This worksheet has no detected columns. Choose a worksheet with columns to create a row.</p>'}<div class="create-row-actions"><button id="createCancel" class="select">Cancel</button><button id="createSave" class="primary-btn" ${sheet.headers.length?'':'disabled'}>Save row</button></div></div>`;
      drawer.classList.add('open');drawer.setAttribute('aria-hidden','false');document.getElementById('drawerBackdrop').classList.add('open');
      document.getElementById('drawerClose').onclick=close;
      document.getElementById('createCancel').onclick=close;
      document.getElementById('createSheet').onchange=event=>{saveDraft();selected=Number(event.target.value);renderForm()};
      drawer.querySelectorAll('[data-create-column]').forEach(input=>input.oninput=saveDraft);
      document.getElementById('createSave').onclick=save;
    }
    async function save(){
      saveDraft();
      const sheet=AppState.data.sheets[selected],values=drafts.get(selected)||[];
      if(!values.some(value=>value.trim())){
        App.toast('Enter at least one value before saving.','error');
        return;
      }
      const button=document.getElementById('createSave');
      button.disabled=true;button.textContent='Saving…';
      try{
        await AppState.appendRow(selected,values.map(value=>value.trim()));
        App.toast(`Added row to ${sheet.name}`);
        close();
        App.refresh();
      }catch(error){
        button.disabled=false;button.textContent='Save row';
        App.toast(`Could not save row: ${error.message}`,'error');
      }
    }
    renderForm();
  }
  function close(){
    document.getElementById('drawer').classList.remove('open');
    document.getElementById('drawer').setAttribute('aria-hidden','true');
    document.getElementById('drawerBackdrop').classList.remove('open');
  }
  document.addEventListener('keydown',event=>{if(event.key==='Escape')close()});
  document.getElementById('drawerBackdrop').addEventListener('click',close);
  return {open,openCreate,close};
})();
