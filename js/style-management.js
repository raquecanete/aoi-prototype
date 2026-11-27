window.StyleManagement=(()=>{
  const fieldAliases={
    style:['Style','Style No.','Style No','Style Number','Style #'],
    article:['Article Number','Article No.','Article No','Article #'],
    season:['Season'],
    pic:['User','PIC','Merchandiser'],
    rbu:['RBU','Region Business Unit','Business Unit'],
    styleName:['Style Name','Style Description'],
    colorwayCount:['Colorway Count','Colorways','No. of Colorways'],
    devType:['Dev Type','Development Type'],
    gender:['Gender'],
    ageGroup:['Age Group'],
    unitType:['Unit Type'],
    sizeTable:['Size Table'],
    baseSize:['Base Size'],
    smv:['SMV'],
    merchandiser:['Merchandiser','User','PIC'],
    patternMaker:['Pattern Maker'],
    cft:['CFT'],
    eiTsd:['EI/TSD','IE/TSD','EI TSD','IE TSD'],
    opd:['OPD'],
    buyReadyDate:['Buy-Ready-Date','Buy-Ready Date','Buy Ready Date'],
    emcApprovalDate:['EMC Approval Date','EMC Approval'],
    category:['Product Category'],
    document:['Document No.','Document No','Document Number'],
    part:['Part No.','Part No','Part Number'],
    rawCode:['Code Raw Material','Raw Material Code'],
    material:['Material','Material Description'],
    supplier:['Supplier Name','Supplier'],
    status:['Status Sample','Sample Status'],
    ordered:['Ordered Qty','Order Qty','Order Quantity'],
    cons:['Cons','Consumption'],
    uom:['UOM','Unit of Measure'],
    fbc:['FBC'],
    delivered:['Qty Delivered','Qty Shipment','Delivered Qty'],
    balance:['Balance ORI','Balance'],
    po:['PO','PO No.','PO Number'],
    productFg:['Product FG','Product FG Code','Finish Goods Code'],
    mo:['MO','Manufacturing Order'],
    finishGoods:['Finish Goods Code','Product FG','Product FG Code'],
    qtyKeeping:['Qty Keeping','Quantity Keeping'],
    qtyShipment:['Qty Shipment','Quantity Shipment'],
    qtyTotal:['Qty Total','Total Qty'],
    dateOrdered:['Date Ordered','MO Created','Created Date'],
    datePromised:['Date Promised','Promised Date'],
    dateStart:['Date Start Schedule','Start','Start Date','Date Start'],
    dateFinish:['Date Finish Schedule','Finish','Finish Date','Date Finish'],
    lastBom:['Last Updated BOM']
  };
  const styleDetailFields=[
    {key:'rbu',label:'RBU'},
    {key:'styleName',label:'Style Name'},
    {key:'colorwayCount',label:'Colorway Count',type:'article-count'},
    {key:'devType',label:'Dev Type',type:'select',options:['New','Carry-over']},
    {key:'gender',label:'Gender',type:'select',options:['Men','Women']},
    {key:'ageGroup',label:'Age Group'},
    {key:'unitType',label:'Unit Type'},
    {key:'sizeTable',label:'Size Table'},
    {key:'baseSize',label:'Base Size',type:'select',options:['S','L','M','XL']},
    {key:'smv',label:'SMV',type:'duration'},
    {key:'merchandiser',label:'Merchandiser'},
    {key:'patternMaker',label:'Pattern Maker',type:'staff'},
    {key:'cft',label:'CFT',type:'staff'},
    {key:'eiTsd',label:'EI/TSD',type:'staff'},
    {key:'opd',label:'OPD',type:'date'},
    {key:'buyReadyDate',label:'Buy-Ready Date',type:'date'},
    {key:'emcApprovalDate',label:'EMC Approval Date',type:'date'},
    {key:'lastBom',label:'Last Updated BOM',type:'date'}
  ];
  const columns=[
    {key:'style',label:'Style No.'},
    {key:'season',label:'Season'},
    {key:'category',label:'Product Category'},
    {key:'pic',label:'PIC'},
    {key:'articles',label:'Articles'},
    {key:'mos',label:'MOs'},
    {key:'lines',label:'Material Lines'},
    {key:'status',label:'Status Sample'},
    {key:'datePromised',label:'Earliest Date Promised'}
  ];
  const filters=[
    {key:'season',label:'Season'},
    {key:'category',label:'Product Category'},
    {key:'pic',label:'PIC'},
    {key:'status',label:'Status Sample'}
  ];
  const state={query:'',lineQuery:'',sortKey:'style',sortDirection:1,page:1,pageSize:25,filter:{season:'',category:'',pic:'',status:''}};
  let model={groups:[],csiIndex:-1,moIndex:-1};
  let activeTab='overview';
  let selectedStyle='';
  let selectedArticle='';
  let selectedMo='';
  let lastTrigger=null;
  let activeFileId=null;
  let sourceEditor=false;
  let styleDetailsEditing=false;
  let renderToken=0;
  let initialized=false;

  const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[char]));
  const normalizeHeader=value=>String(value??'').normalize('NFKD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
  const trimText=value=>String(value??'').trim();
  const sheetColumn=(sheet,field)=>{
    const aliases=(fieldAliases[field]||[field]).map(normalizeHeader);
    return sheet.headers.findIndex(header=>aliases.includes(normalizeHeader(header)));
  };
  const readValue=(sheet,row,column)=>{
    if(column<0)return '';
    const value=AppState.displayCell(sheet,row.index,column);
    return value==='—'?'':String(value??'').trim();
  };
  const readField=(sheet,row,field)=>readValue(sheet,row,sheetColumn(sheet,field));
  const readRawField=(sheet,row,field)=>{
    const column=sheetColumn(sheet,field);
    if(column<0)return '';
    const value=row.values[column];
    return value===undefined||value===null||value===''?'':String(value).trim();
  };
  const rowEntries=sheetIndex=>{
    const sheet=AppState.data.sheets[sheetIndex];
    return sheet?AppState.rowsFor(sheet).map(row=>({sheet,sheetIndex,row})): [];
  };
  const distinct=values=>{
    const seen=new Set(),result=[];
    values.forEach(value=>{
      const clean=trimText(value);
      if(clean&&!seen.has(clean)){seen.add(clean);result.push(clean)}
    });
    return result;
  };
  const normalizedText=value=>trimText(value).toLocaleLowerCase();
  const statusClass=value=>{
    const status=normalizedText(value);
    if(status==='garment test')return 'style-status-chip garment-test';
    if(status==='sealing')return 'style-status-chip sealing';
    return 'style-status-chip';
  };
  const statusChips=values=>{
    const items=distinct(values);
    return items.length?items.map(value=>`<span class="${statusClass(value)}">${esc(value)}</span>`).join(' '):'<span class="style-muted">-</span>';
  };
  const textCell=value=>value?esc(value):'<span class="style-muted">-</span>';
  const cellDate=(sheet,row,column)=>{
    if(column<0)return '';
    const raw=row.values[column];
    const iso=AppState.formatDateValue(raw);
    if(!iso)return readValue(sheet,row,column);
    const [year,month,day]=iso.split('-').map(Number);
    return new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'})
      .format(new Date(Date.UTC(year,month-1,day)));
  };
  const dateIso=(sheet,row,field)=>{
    const column=sheetColumn(sheet,field);
    return column<0?'':AppState.formatDateValue(row.values[column])||'';
  };
  const formatDecimalSum=values=>{
    const parsed=values.map(value=>{
      const text=trimText(value);
      const match=text.match(/^([+-]?)(\d+)(?:\.(\d+))?$/);
      if(!match)return null;
      return {negative:match[1]==='-',whole:match[2],fraction:match[3]||''};
    }).filter(Boolean);
    if(!parsed.length)return '';
    const scale=Math.max(...parsed.map(item=>item.fraction.length));
    const factor=10n**BigInt(scale);
    const total=parsed.reduce((sum,item)=>{
      const integer=BigInt(item.whole)*factor+BigInt((item.fraction+'0'.repeat(scale)).slice(0,scale)||'0');
      return sum+(item.negative?-integer:integer);
    },0n);
    const negative=total<0n;
    const absolute=negative?-total:total;
    const whole=absolute/factor;
    const fraction=scale?String(absolute%factor).padStart(scale,'0').replace(/0+$/,''):'';
    return `${negative?'-':''}${whole}${fraction?`.${fraction}`:''}`;
  };

  function findSheetIndex(kind){
    const sheets=AppState.data.sheets||[];
    const score=sheet=>{
      if(kind==='csi')return ['style','article'].filter(field=>sheetColumn(sheet,field)>=0).length;
      return ['finishGoods','mo'].filter(field=>sheetColumn(sheet,field)>=0).length;
    };
    const nameMatch=sheets.findIndex(sheet=>kind==='csi'?/csi/i.test(sheet.name):/\bmo\b|manufactur/i.test(sheet.name));
    if(nameMatch>=0)return nameMatch;
    let bestIndex=-1,bestScore=0;
    sheets.forEach((sheet,index)=>{
      const current=score(sheet);
      if(current>bestScore){bestScore=current;bestIndex=index}
    });
    return bestIndex;
  }

  function isBundledWorkbook(){
    return AppState.activeFile?.demoData===true||AppState.activeFile?.importedBy==='system';
  }

  function buildModel(){
    if(isBundledWorkbook())return {groups:[],csiIndex:-1,moIndex:-1};
    const csiIndex=findSheetIndex('csi');
    const moIndex=findSheetIndex('mo');
    const csi=csiIndex<0?null:AppState.data.sheets[csiIndex];
    const moSheet=moIndex<0?null:AppState.data.sheets[moIndex];
    const groupMap=new Map();
    const identifiers=new Map();
    if(!csi)return {groups:[],csiIndex,moIndex};
    const addIdentifier=(key,group,articleKey='')=>{
      const clean=trimText(key);
      if(!clean)return;
      if(!identifiers.has(clean))identifiers.set(clean,[]);
      const found=identifiers.get(clean);
      let item=found.find(entry=>entry.group===group);
      if(!item){item={group,articles:new Set()};found.push(item)}
      if(articleKey)item.articles.add(articleKey);
    };
    const csiRows=rowEntries(csiIndex);
    csiRows.forEach(entry=>{
      const styleKey=readField(csi,entry.row,'style');
      if(!styleKey)return;
      let group=groupMap.get(styleKey);
      if(!group){
        group={
          key:styleKey,lines:[],articles:new Map(),mos:new Map(),
          seasons:[],categories:[],pics:[],statuses:[],datePromised:[],
          startDates:[],finishDates:[],bomDates:[]
        };
        groupMap.set(styleKey,group);
      }
      const articleNumber=readField(csi,entry.row,'article');
      const productFg=readField(csi,entry.row,'productFg');
      const articleKey=articleNumber||productFg;
      const line={
        ...entry,styleKey,articleNumber,articleKey,productFg,
        document:readField(csi,entry.row,'document'),
        part:readField(csi,entry.row,'part'),
        rawCode:readField(csi,entry.row,'rawCode'),
        material:readField(csi,entry.row,'material'),
        supplier:readField(csi,entry.row,'supplier'),
        status:readField(csi,entry.row,'status'),
        ordered:readRawField(csi,entry.row,'ordered'),
        cons:readRawField(csi,entry.row,'cons'),
        uom:readField(csi,entry.row,'uom'),
        fbc:readRawField(csi,entry.row,'fbc'),
        delivered:readRawField(csi,entry.row,'delivered'),
        balance:readRawField(csi,entry.row,'balance'),
        po:readField(csi,entry.row,'po')
      };
      group.lines.push(line);
      const season=readField(csi,entry.row,'season');
      const category=readField(csi,entry.row,'category');
      const pic=readField(csi,entry.row,'pic');
      const status=line.status;
      if(season)group.seasons.push(season);
      if(category)group.categories.push(category);
      if(pic)group.pics.push(pic);
      if(status)group.statuses.push(status);
      const promised=dateIso(csi,entry.row,'datePromised');
      const start=dateIso(csi,entry.row,'dateStart');
      const finish=dateIso(csi,entry.row,'dateFinish');
      const bom=dateIso(csi,entry.row,'lastBom');
      if(promised)group.datePromised.push(promised);
      if(start)group.startDates.push(start);
      if(finish)group.finishDates.push(finish);
      if(bom)group.bomDates.push(bom);
      if(articleNumber){
        if(!group.articles.has(articleNumber)){
          group.articles.set(articleNumber,{key:articleNumber,number:articleNumber,lines:[],productFgs:[],statuses:[],mOs:new Set()});
        }
        const article=group.articles.get(articleNumber);
        article.lines.push(line);
        if(productFg)article.productFgs.push(productFg);
        if(status)article.statuses.push(status);
      }
      addIdentifier(styleKey,group);
      addIdentifier(articleNumber,group,articleNumber);
      addIdentifier(productFg,group,articleNumber);
    });

    if(moSheet){
      const moRows=rowEntries(moIndex);
      const moGroups=new Map();
      moRows.forEach(entry=>{
        const moName=readField(moSheet,entry.row,'mo');
        if(!moName)return;
        const finishGoods=readField(moSheet,entry.row,'finishGoods');
        const explicitStyle=readField(moSheet,entry.row,'style');
        const explicitArticle=readField(moSheet,entry.row,'article');
        const possibleKeys=distinct([explicitStyle,finishGoods,explicitArticle,readField(moSheet,entry.row,'productFg')]);
        const matches=[];
        possibleKeys.forEach(key=>(identifiers.get(key)||[]).forEach(match=>{
          if(!matches.some(item=>item.group===match.group))matches.push(match);
        }));
        matches.forEach(match=>{
          const group=match.group;
          if(!group.mos.has(moName)){
            group.mos.set(moName,{
              key:moName,name:moName,entries:[],articleKeys:new Set(),finishGoods
            });
          }
          const mo=group.mos.get(moName);
          mo.entries.push(entry);
          const articleKeys=new Set(match.articles);
          if(!articleKeys.size&&possibleKeys.includes(group.key)){
            group.articles.forEach(article=>articleKeys.add(article.key));
          }
          articleKeys.forEach(key=>mo.articleKeys.add(key));
        });
      });
    }

    const csiMoColumn=sheetColumn(csi,'mo');
    if(csiMoColumn>=0){
      groupMap.forEach(group=>group.lines.forEach(line=>{
        const name=readValue(csi,line.row,csiMoColumn);
        if(!name)return;
        if(!group.mos.has(name))group.mos.set(name,{key:name,name,entries:[],articleKeys:new Set(),finishGoods:line.productFg});
        const mo=group.mos.get(name);
        mo.entries.push(line);
        if(line.articleKey)mo.articleKeys.add(line.articleKey);
      }));
    }

    const groups=[...groupMap.values()].map(group=>{
      group.articles.forEach(article=>{
        group.mos.forEach(mo=>{
          if(!mo.articleKeys.size||mo.articleKeys.has(article.key))article.mOs.add(mo.key);
        });
      });
      group.seasons=distinct(group.seasons);
      group.categories=distinct(group.categories);
      group.pics=distinct(group.pics);
      group.statuses=distinct(group.statuses);
      group.datePromised=distinct(group.datePromised).sort();
      group.startDates=distinct(group.startDates).sort();
      group.finishDates=distinct(group.finishDates).sort();
      group.bomDates=distinct(group.bomDates).sort();
      group.articles=[...group.articles.values()];
      group.mos=[...group.mos.values()];
      group.orderedTotal=formatDecimalSum(group.lines.map(line=>line.ordered));
      group.suppliers=distinct(group.lines.map(line=>line.supplier));
      group.latestPromised=group.datePromised[0]||'';
      group.duplicateLines=findDuplicateLines(group.lines);
      group.noPoLines=group.lines.filter(line=>!line.po||normalizedText(line.po)==='no po').length;
      return group;
    });
    return {groups,csiIndex,moIndex};
  }

  function findDuplicateLines(lines){
    const counts=new Map();
    lines.forEach(line=>{
      const values=line.sheet.headers.map((_,column)=>trimText(line.row.values[column]));
      const key=JSON.stringify(values);
      counts.set(key,(counts.get(key)||0)+1);
    });
    return [...counts.values()].reduce((total,count)=>total+Math.max(0,count-1),0);
  }

  function qualityIssues(){
    const current=buildModel();
    const issues=[];
    current.groups.forEach(group=>{
      if(group.duplicateLines){
        issues.push({
          severity:'warning',
          title:`Style ${group.key} · duplicate material lines`,
          detail:`${group.duplicateLines} exact duplicate material line${group.duplicateLines===1?'':'s'} found.`
        });
      }
      if(group.noPoLines){
        issues.push({
          severity:'warning',
          title:`Style ${group.key} · No PO lines`,
          detail:`${group.noPoLines} material line${group.noPoLines===1?'':'s'} have no PO value.`
        });
      }
      if(group.seasons.length>1){
        issues.push({
          severity:'warning',
          title:`Style ${group.key} · inconsistent Season`,
          detail:`Rows for this style contain ${group.seasons.join(', ')}.`
        });
      }
      if(group.categories.length>1){
        issues.push({
          severity:'warning',
          title:`Style ${group.key} · inconsistent Product Category`,
          detail:`Rows for this style contain ${group.categories.join(', ')}.`
        });
      }
    });
    return issues;
  }

  function groupValues(group,key){
    if(key==='articles')return group.articles.length;
    if(key==='mos')return group.mos.length;
    if(key==='lines')return group.lines.length;
    if(key==='status')return group.statuses.join(', ');
    if(key==='datePromised')return group.latestPromised;
    if(key==='season')return group.seasons.join(', ');
    if(key==='category')return group.categories.join(', ');
    if(key==='pic')return group.pics.join(', ');
    return group.key;
  }

  function filterOptions(key){
    if(key==='status')return distinct(model.groups.flatMap(group=>group.statuses)).sort((a,b)=>a.localeCompare(b,undefined,{sensitivity:'base'}));
    const values=model.groups.flatMap(group=>key==='season'?group.seasons:key==='category'?group.categories:group.pics);
    return distinct(values).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:'base'}));
  }

  function matchingGroups(){
    const query=normalizedText(state.query);
    let filtered=model.groups.map(group=>{
      const styleMatch=!query||normalizedText(group.key).includes(query);
      const matchedArticles=query&&!styleMatch
        ?group.articles.filter(article=>normalizedText(article.number).includes(query))
        :[];
      return {...group,matchedArticles};
    }).filter(group=>{
      const queryMatch=!query||normalizedText(group.key).includes(query)||group.matchedArticles.length>0;
      const seasonMatch=!state.filter.season||group.seasons.includes(state.filter.season);
      const categoryMatch=!state.filter.category||group.categories.includes(state.filter.category);
      const picMatch=!state.filter.pic||group.pics.includes(state.filter.pic);
      const statusMatch=!state.filter.status||group.statuses.includes(state.filter.status);
      return queryMatch&&seasonMatch&&categoryMatch&&picMatch&&statusMatch;
    });
    filtered.sort((a,b)=>{
      const aValue=groupValues(a,state.sortKey),bValue=groupValues(b,state.sortKey);
      if(state.sortKey==='articles'||state.sortKey==='mos'||state.sortKey==='lines'){
        return (aValue-bValue)*state.sortDirection||a.key.localeCompare(b.key,undefined,{sensitivity:'base'});
      }
      const options={numeric:state.sortKey!=='style',sensitivity:'base'};
      return String(aValue).localeCompare(String(bValue),undefined,options)*state.sortDirection||
        a.key.localeCompare(b.key,undefined,{sensitivity:'base'});
    });
    return filtered;
  }

  function render(){
    const fileId=AppState.activeFileId;
    if(activeFileId!==null&&fileId!==activeFileId)closeDetail({clearHash:true,restoreFocus:false});
    activeFileId=fileId;
    model=buildModel();
    const app=document.getElementById('app');
    const all=matchingGroups();
    const count=all.length;
    const pageCount=Math.max(1,Math.ceil(count/state.pageSize));
    if(state.page>pageCount)state.page=pageCount;
    const visible=all.slice((state.page-1)*state.pageSize,state.page*state.pageSize);
    const filtersHtml=filters.map(filter=>{
      const options=filterOptions(filter.key);
      return `<label class="style-filter" for="style-filter-${filter.key}"><span>${esc(filter.label)}</span><select id="style-filter-${filter.key}" class="select"><option value="">All ${esc(filter.label.toLocaleLowerCase())}</option>${options.map(value=>`<option value="${esc(value)}" ${state.filter[filter.key]===value?'selected':''}>${esc(value)}</option>`).join('')}</select></label>`;
    }).join('');
    const fileName=AppState.activeFile?.name||AppState.data.fileName||'No workbook';
    const tableRows=visible.length?visible.map(group=>renderStyleRow(group)).join(''):`<tr><td class="style-no-results" colspan="${columns.length+1}">${count===0&&model.groups.length?'No styles match these search and filter settings.':'No styles are available in the active CSI sheet.'}${count===0&&model.groups.length?'<br><button id="style-empty-clear" class="style-text-link" type="button">Clear filters</button>':''}</td></tr>`;
    const sortHeaders=columns.map(column=>{
      const direction=state.sortKey===column.key?state.sortDirection:0;
      const indicator=direction?(direction===1?'↑':'↓'):'↕';
      const ariaSort=direction?(direction===1?'ascending':'descending'):'none';
      return `<th scope="col" aria-sort="${ariaSort}"><button type="button" data-style-sort="${column.key}" aria-label="Sort by ${esc(column.label)}">${esc(column.label)} <span aria-hidden="true">${indicator}</span></button></th>`;
    }).join('');
    const range=count?`${(state.page-1)*state.pageSize+1}-${Math.min(state.page*state.pageSize,count)} of ${count} styles`:'0 styles';
    const pagePanel=isBundledWorkbook()
      ?'<section class="card style-empty-state"><strong>Import or select your CSI workbook</strong><span>Style Management only displays rows from an imported CSI workbook. The active workbook is bundled sample/demo data.</span></section>'
      :model.csiIndex<0
      ?'<section class="card style-empty-state"><strong>No CSI style data found</strong><span>Open a workbook with a CSI sheet containing Style and Article Number columns.</span></section>'
      :`<section class="card style-table-panel" aria-label="Styles">
        <div class="style-toolbar"><label class="style-search"><span>Search styles or articles</span><input id="style-table-search" class="filter-input" type="search" value="${esc(state.query)}" placeholder="Style No. or Article Number" aria-label="Search styles and article numbers"></label><div class="style-pagination-size"><label for="style-page-size">Rows</label><select id="style-page-size" class="select"><option value="25" ${state.pageSize===25?'selected':''}>25</option><option value="50" ${state.pageSize===50?'selected':''}>50</option><option value="100" ${state.pageSize===100?'selected':''}>100</option></select></div></div>
        <div class="style-filterbar">${filtersHtml}<button id="style-clear-filters" type="button" class="select">Clear filters</button></div>
        <div class="style-table-scroll"><table class="data-table style-list-table"><thead><tr>${sortHeaders}<th scope="col"><span class="sr-only">Open style</span></th></tr></thead><tbody>${tableRows}</tbody></table></div>
        <div class="style-pager"><span>${range}</span><div><button id="style-prev" class="select" type="button" ${state.page<=1?'disabled':''}>Previous</button><button id="style-next" class="select" type="button" ${state.page>=pageCount?'disabled':''}>Next</button></div></div>
      </section>`;
    app.innerHTML=`<header class="page-heading style-page-heading"><div><p class="eyebrow">IMPORTED WORKBOOK</p><h1>Style Management</h1><p class="page-subtitle">Style and material line data come from the active workbook's CSI sheet.</p></div><div class="style-source-name"><span>Active file</span><strong title="${esc(fileName)}">${esc(fileName)}</strong></div></header>${pagePanel}`;
    bindTableEvents();
    const hashStyle=getHashStyle();
    if(hashStyle){
      if(model.groups.some(group=>group.key===hashStyle)){
        const drawer=document.getElementById('drawer');
        if(!isOpen||selectedStyle!==hashStyle||sourceEditor||!drawer.classList.contains('open'))openDetail(hashStyle,{fromHash:true});
        else renderDetail();
      }
      else clearStyleHash();
    }else if(isOpen&&!sourceEditor)closeDetail({clearHash:false,restoreFocus:false});
  }

  function renderStyleRow(group){
    const articleHint=group.matchedArticles?.length
      ?`<small class="style-match-hint" title="${esc(group.matchedArticles.map(article=>article.number).join(', '))}">Matched article: ${esc(group.matchedArticles.map(article=>article.number).join(', '))}</small>`
      :'';
    const promised=group.latestPromised?formatIsoDate(group.latestPromised):'';
    return `<tr class="style-summary-row" data-style-key="${esc(group.key)}" tabindex="0" aria-label="Open details for style ${esc(group.key)}">
      <td><a class="style-number-link" href="#style=${encodeURIComponent(group.key)}" data-open-style="${esc(group.key)}">${esc(group.key)}</a>${articleHint}</td>
      <td>${textCell(group.seasons.join(', '))}</td><td>${textCell(group.categories.join(', '))}</td><td>${textCell(group.pics.join(', '))}</td>
      <td>${group.articles.length}</td><td>${group.mos.length}</td><td>${group.lines.length}</td>
      <td>${statusChips(group.statuses)}</td><td>${promised?esc(promised):'<span class="style-muted">-</span>'}</td>
      <td class="style-open-arrow" aria-hidden="true">→</td></tr>`;
  }

  function bindTableEvents(){
    const app=document.getElementById('app');
    const search=app.querySelector('#style-table-search');
    if(search)search.oninput=()=>{
      state.query=search.value;
      state.page=1;
      const global=document.getElementById('globalSearch');
      if(global)global.value=state.query;
      render();
      document.getElementById('style-table-search')?.focus();
    };
    app.querySelectorAll('[data-style-sort]').forEach(button=>button.onclick=()=>{
      const key=button.dataset.styleSort;
      state.sortDirection=state.sortKey===key?state.sortDirection*-1:1;
      state.sortKey=key;state.page=1;render();
    });
    filters.forEach(filter=>{
      const select=app.querySelector(`#style-filter-${filter.key}`);
      if(select)select.onchange=()=>{state.filter[filter.key]=select.value;state.page=1;render()};
    });
    const pageSize=app.querySelector('#style-page-size');
    if(pageSize)pageSize.onchange=()=>{state.pageSize=Number(pageSize.value);state.page=1;render()};
    const clear=app.querySelector('#style-clear-filters');
    const clearFilters=()=>{
      state.query='';state.page=1;
      Object.keys(state.filter).forEach(key=>state.filter[key]='');
      const global=document.getElementById('globalSearch');
      if(global)global.value='';
      render();
    };
    if(clear)clear.onclick=clearFilters;
    const emptyClear=app.querySelector('#style-empty-clear');
    if(emptyClear)emptyClear.onclick=clearFilters;
    const prev=app.querySelector('#style-prev'),next=app.querySelector('#style-next');
    if(prev)prev.onclick=()=>{if(state.page>1){state.page--;render()}};
    if(next)next.onclick=()=>{state.page++;render()};
    app.querySelectorAll('.style-summary-row').forEach(row=>{
      const open=()=>openDetail(row.dataset.styleKey,{trigger:row});
      row.onclick=event=>{
        if(event.target.closest('button,a'))return;
        open();
      };
      row.onkeydown=event=>{
        if(event.key!=='Enter'&&event.key!==' ')return;
        if(event.target.closest('a,button'))return;
        event.preventDefault();
        open();
      };
    });
    app.querySelectorAll('[data-open-style]').forEach(link=>link.onclick=event=>{
      event.preventDefault();event.stopPropagation();
      openDetail(link.dataset.openStyle,{trigger:link.closest('tr')});
    });
  }

  function formatIsoDate(iso){
    if(!iso)return '';
    const [year,month,day]=iso.split('-').map(Number);
    return new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'})
      .format(new Date(Date.UTC(year,month-1,day)));
  }
  function formatDateField(sheet,row,field){return cellDate(sheet,row,sheetColumn(sheet,field))}
  function articleFor(group,key){return group.articles.find(article=>article.key===key)||null}
  function selectedGroup(){return model.groups.find(group=>group.key===selectedStyle)||null}
  function currentFilteredLines(group){
    let lines=group.lines;
    if(selectedArticle)lines=lines.filter(line=>line.articleKey===selectedArticle);
    if(selectedMo){
      const mo=group.mos.find(item=>item.key===selectedMo);
      if(mo){
        const directColumn=sheetColumn(lineSheet(group), 'mo');
        if(directColumn>=0){
          const direct=lines.filter(line=>readValue(line.sheet,line.row,directColumn)===selectedMo);
          if(direct.length)lines=direct;
          else if(mo.articleKeys.size)lines=lines.filter(line=>mo.articleKeys.has(line.articleKey));
        }else if(mo.articleKeys.size)lines=lines.filter(line=>mo.articleKeys.has(line.articleKey));
      }
    }
    return lines;
  }
  function lineSheet(group){return group.lines[0]?.sheet||null}
  function currentFilteredMos(group){
    let result=group.mos;
    if(selectedArticle)result=result.filter(mo=>!mo.articleKeys.size||mo.articleKeys.has(selectedArticle));
    return result;
  }

  function openDetail(styleKey,options={}){
    const group=model.groups.find(item=>item.key===styleKey);
    if(!group)return;
    if(selectedStyle!==styleKey){
      selectedArticle='';
      selectedMo='';
      state.lineQuery='';
    }
    selectedStyle=styleKey;
    sourceEditor=false;
    styleDetailsEditing=false;
    if(options.trigger)lastTrigger=options.trigger;
    setStyleHash(styleKey);
    isOpen=true;
    const drawer=document.getElementById('drawer');
    drawer.classList.add('style-detail-drawer');
    drawer.classList.remove('open');
    void drawer.offsetWidth;
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden','false');
    drawer.setAttribute('role','dialog');
    drawer.setAttribute('aria-modal','true');
    drawer.setAttribute('aria-label',`Style ${styleKey} details`);
    const backdrop=document.getElementById('drawerBackdrop');
    backdrop.classList.add('open');
    backdrop.setAttribute('aria-hidden','true');
    const token=++renderToken;
    drawer.innerHTML=`<div class="style-drawer-skeleton" role="status" aria-label="Loading style details"><div class="style-skeleton-line wide"></div><div class="style-skeleton-kpis">${Array.from({length:5},()=>'<span></span>').join('')}</div><div class="style-skeleton-line"></div><div class="style-skeleton-line medium"></div><div class="style-skeleton-grid">${Array.from({length:8},()=>'<span></span>').join('')}</div></div>`;
    requestAnimationFrame(()=>{if(token===renderToken&&isOpen)renderDetail(true)});
  }

  function renderDetail(focusClose=false){
    const group=selectedGroup();
    if(!group)return;
    const drawer=document.getElementById('drawer');
    const previousFocus=drawer.contains(document.activeElement)?document.activeElement:null;
    const previousTab=previousFocus?.dataset.styleTab||'';
    const previousArticleControl=previousFocus?.dataset.filterArticle||'';
    const previousControlId=previousFocus?.id||'';
    const allGroups=matchingGroups();
    const navGroups=allGroups.some(item=>item.key===group.key)?allGroups:model.groups;
    const currentIndex=navGroups.findIndex(item=>item.key===group.key);
    const previous=navGroups[currentIndex-1],next=navGroups[currentIndex+1];
    const season=group.seasons.join(', '),category=group.categories.join(', '),pic=group.pics.join(', ');
    const articleFilter=selectedArticle?articleFor(group,selectedArticle):null;
    const moFilter=selectedMo?group.mos.find(mo=>mo.key===selectedMo):null;
    const tabs=[
      ['overview','Overview'],['articles','Articles'],['mos','Manufacturing Orders'],
      ['lines','Material Lines'],['suppliers','Suppliers']
    ];
    drawer.innerHTML=`<header class="style-detail-header">
      <div class="style-detail-topline"><div><p class="eyebrow">STYLE DETAILS</p><h2 id="style-detail-title">${esc(group.key)}</h2></div><div class="style-detail-actions"><button type="button" class="select" id="style-previous" ${previous?'':'disabled'} aria-label="Previous style">Previous</button><button type="button" class="select" id="style-next" ${next?'':'disabled'} aria-label="Next style">Next</button><button type="button" class="icon-btn" id="style-detail-close" aria-label="Close style details">×</button></div></div>
      <div class="style-detail-meta"><span><strong>Season</strong>${textCell(season)}</span><span><strong>Product Category</strong>${textCell(category)}</span><span><strong>PIC</strong>${textCell(pic)}</span><span class="style-detail-status">${statusChips(group.statuses)}</span></div>
      <div class="style-kpis"><div><strong>${group.articles.length}</strong><span>Articles</span></div><div><strong>${group.mos.length}</strong><span>MOs</span></div><div><strong>${group.lines.length}</strong><span>Material Lines</span></div><div><strong>${group.suppliers.length}</strong><span>Suppliers</span></div><div><strong>${esc(group.orderedTotal||'-')}</strong><span>Total Ordered Qty</span></div></div>
    </header>
    <div class="style-detail-body">
      <div class="style-detail-filters">${articleFilter?`<button type="button" class="style-filter-chip" data-clear-article aria-label="Remove article filter">Article: ${esc(articleFilter.number)} <span aria-hidden="true">×</span></button>`:''}${moFilter?`<button type="button" class="style-filter-chip" data-clear-mo aria-label="Remove MO filter">MO: ${esc(moFilter.name)} <span aria-hidden="true">×</span></button>`:''}</div>
      <div class="style-detail-tabs" role="tablist" aria-label="Style detail sections">${tabs.map(([key,label])=>`<button type="button" role="tab" id="style-tab-${key}" aria-controls="style-panel-${key}" aria-selected="${activeTab===key}" tabindex="${activeTab===key?'0':'-1'}" data-style-tab="${key}">${label}</button>`).join('')}</div>
      <section id="style-panel-${activeTab}" class="style-tab-panel" role="tabpanel" aria-labelledby="style-tab-${activeTab}" tabindex="0">${renderTab(group)}</section>
    </div>`;
    bindDetailEvents(group);
      const focusTarget=(previousArticleControl&&drawer.querySelector(`[data-filter-article="${CSS.escape(previousArticleControl)}"]`))||
        (previousControlId&&drawer.querySelector(`#${CSS.escape(previousControlId)}`))||
        (previousTab&&drawer.querySelector(`[data-style-tab="${previousTab}"]`))||
        drawer.querySelector(`[data-style-tab="${activeTab}"]`)||
        drawer.querySelector('#style-detail-close');
      if(focusClose)drawer.querySelector('#style-detail-close')?.focus();
      else focusTarget?.focus();
  }

  function renderTab(group){
    if(activeTab==='articles')return renderArticles(group);
    if(activeTab==='mos')return renderMos(group);
    if(activeTab==='lines')return renderLines(group);
    if(activeTab==='suppliers')return renderSuppliers(group);
    if(styleDetailsEditing)return renderStyleDetailsEditor(group);
    return renderOverview(group);
  }

  function styleFieldRawValues(group,field){
    return distinct(group.lines.map(line=>{
      const column=sheetColumn(line.sheet,field.key);
      if(column<0)return '';
      const value=line.row.values[column];
      return value===undefined||value===null||value===''?'':String(value).trim();
    }));
  }

  function styleFieldEditValue(group,field){
    if(field.type==='article-count')return String(group.articles.length);
    if(field.type==='duration'){
      const line=group.lines.find(entry=>sheetColumn(entry.sheet,field.key)>=0);
      if(!line)return '';
      const column=sheetColumn(line.sheet,field.key);
      const value=AppState.displayCell(line.sheet,line.row.index,column);
      return value==='—'?'':String(value??'').trim();
    }
    const value=styleFieldRawValues(group,field)[0]||'';
    return field.type==='date'?AppState.formatDateValue(value)||'':value;
  }

  function styleFieldDisplayValue(group,field){
    if(field.type==='article-count')return String(group.articles.length);
    const values=distinct(group.lines.map(line=>{
      const column=sheetColumn(line.sheet,field.key);
      if(column<0)return '';
      const iso=field.type==='date'?AppState.formatDateValue(line.row.values[column]):null;
      return iso?formatIsoDate(iso):readValue(line.sheet,line.row,column);
    }).filter(Boolean));
    return values.join(', ');
  }

  function renderStyleDetailsCard(group){
    const editButton=AppState.currentUser?.role==='admin'
      ?'<button type="button" class="select" id="style-edit-details">Edit details</button>'
      :'';
    return `<section class="style-detail-card style-fields-card"><div class="style-detail-card-heading"><h3>Style Details</h3>${editButton}</div><dl class="style-style-fields">${styleDetailFields.map(field=>summaryItem(field.label,styleFieldDisplayValue(group,field))).join('')}</dl></section>`;
  }

  function renderStyleDetailsEditor(group){
    const staffHeaders=new Set(['Pattern Maker','CFT','IE/TSD','EI/TSD','Merchandiser','User','Staff','Staff Name','Staff List','Employee','Employee Name','Employee List','Full Name','Name'].map(normalizeHeader));
    const staffOptions=new Set(['John','Ken','Renzie','James']);
    AppState.data.sheets.forEach(sheet=>{
      sheet.headers.forEach((header,column)=>{
        if(!staffHeaders.has(normalizeHeader(header)))return;
        AppState.rowsFor(sheet).forEach(row=>{
          const value=trimText(AppState.displayCell(sheet,row.index,column));
          if(value&&value!=='—')staffOptions.add(value);
        });
      });
    });
    const controls=styleDetailFields.map(field=>{
      const value=styleFieldEditValue(group,field);
      const common=`class="${field.type==='select'||field.type==='staff'?'select':'filter-input'}" data-style-detail-field="${field.key}"`;
      if(field.type==='select'||field.type==='staff'){
        const options=field.type==='staff'
          ?[...staffOptions].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true,sensitivity:'base'}))
          :field.options;
        const choices=value&&!options.includes(value)?[value,...options]:options;
        return `<label class="style-detail-edit-field"><span>${esc(field.label)}</span><select ${common}><option value="" ${value===''?'selected':''}>Select…</option>${choices.map(option=>`<option value="${esc(option)}" ${value===option?'selected':''}>${esc(option)}</option>`).join('')}</select></label>`;
      }
      const durationAttributes=field.type==='duration'?'inputmode="numeric" placeholder="HH:mm"':'';
      const articleCountAttributes=field.type==='article-count'?'min="0" step="1" readonly':'';
      const help=field.type==='duration'
        ?'<small class="style-detail-field-help">Duration, e.g. 00:12</small>'
        :field.type==='article-count'
          ?'<small class="style-detail-field-help">Calculated from unique Article Number values in the CSI sheet.</small>'
          :'';
      const inputType=field.type==='date'?'date':field.type==='article-count'?'number':'text';
      return `<label class="style-detail-edit-field"><span>${esc(field.label)}</span><input ${common} type="${inputType}" value="${esc(value)}" title="${esc(styleFieldDisplayValue(group,field))}" autocomplete="off" ${durationAttributes} ${articleCountAttributes}>${help}</label>`;
    }).join('');
    return `<section class="style-detail-card style-fields-editor"><div class="style-detail-card-heading"><div><h3>Edit Style Details</h3><p>Changes apply to all rows for this style and follow the workbook’s normal save/version flow.</p></div></div><form id="style-details-form"><div class="style-detail-edit-grid">${controls}</div><div class="style-detail-edit-actions"><button type="button" class="select" data-cancel-style-details>Cancel</button><button type="submit" class="primary-btn">Save details</button></div></form></section>`;
  }

  function renderOverview(group){
    const issues=[];
    if(group.duplicateLines)issues.push(`${group.duplicateLines} duplicate material line${group.duplicateLines===1?'':'s'}`);
    if(group.noPoLines)issues.push(`${group.noPoLines} line${group.noPoLines===1?'':'s'} without a PO`);
    if(group.seasons.length>1)issues.push(`Inconsistent Season: ${group.seasons.join(', ')}`);
    if(group.categories.length>1)issues.push(`Inconsistent Product Category: ${group.categories.join(', ')}`);
    const schedule=[
      ['Date Start Schedule',group.startDates[0]||''],
      ['Date Finish Schedule',group.finishDates[group.finishDates.length-1]||''],
      ['Date Promised',group.datePromised[0]||'']
    ];
    return `<div class="style-overview-grid"><section class="style-detail-card"><h3>Style summary</h3><dl>${summaryItem('Style No.',group.key)}${summaryItem('Season',group.seasons.join(', '))}${summaryItem('Product Category',group.categories.join(', '))}${summaryItem('PIC',group.pics.join(', '))}${summaryItem('Articles',group.articles.length)}${summaryItem('Manufacturing Orders',group.mos.length)}${summaryItem('Material Lines',group.lines.length)}</dl></section>
      <section class="style-detail-card"><h3>Schedule</h3><dl>${schedule.map(([label,value])=>summaryItem(label,value?formatIsoDate(value):'')).join('')}</dl></section>
      ${renderStyleDetailsCard(group)}
      <section class="style-detail-card style-quality-card"><div class="style-detail-card-heading"><h3>Data quality</h3><button type="button" class="style-text-link" data-open-quality>View Data Quality</button></div>${issues.length?`<ul>${issues.map(issue=>`<li>${esc(issue)}</li>`).join('')}</ul>`:'<p class="style-muted">No duplicate material lines, No PO lines, or style inconsistencies found.</p>'}</section></div>`;
  }
  function summaryItem(label,value){
    const content=value===0?'0':value?esc(value):'<span class="style-muted">-</span>';
    return `<div class="style-summary-item"><dt>${esc(label)}</dt><dd>${content}</dd></div>`;
  }

  function renderArticles(group){
    if(!group.articles.length)return '<p class="style-empty-note">No Article Numbers were found for this style.</p>';
    const articles=group.articles.map(article=>{
      const productFgs=distinct(article.productFgs);
      const moNames=distinct(group.mos.filter(mo=>!mo.articleKeys.size||mo.articleKeys.has(article.key)).map(mo=>mo.name));
      return `<tr><td><button type="button" class="style-article-link" data-filter-article="${esc(article.key)}" title="${esc(article.number)}">${esc(article.number)}</button></td><td>${textCell(productFgs.join(', '))}</td><td>${textCell(moNames.join(', '))}</td><td>${article.lines.length}</td><td>${statusChips(article.statuses)}</td></tr>`;
    }).join('');
    return `<div class="style-tab-heading"><div><h3>Articles</h3><p>${group.articles.length} unique Article Number${group.articles.length===1?'':'s'}</p></div></div><div class="style-detail-table-wrap"><table class="style-detail-table"><thead><tr><th>Article Number</th><th>Product FG</th><th>MOs using it</th><th>Lines</th><th>Status Sample</th></tr></thead><tbody>${articles}</tbody></table></div>`;
  }

  function renderMos(group){
    const items=currentFilteredMos(group);
    if(!items.length)return '<p class="style-empty-note">No manufacturing orders match this style and article filter.</p>';
    const headers=['Finish Goods Code','MO','Qty Keeping','Qty Shipment','Qty Total','Date Ordered','Date Promised','Start','Finish','Status Sample'];
    const body=items.map(mo=>{
      const entry=mo.entries[0];
      if(!entry)return '';
      const sheet=entry.sheet,row=entry.row;
      const values=[
        readField(sheet,row,'finishGoods'),mo.name,readRawField(sheet,row,'qtyKeeping'),
        readRawField(sheet,row,'qtyShipment'),readRawField(sheet,row,'qtyTotal'),
        formatDateField(sheet,row,'dateOrdered'),formatDateField(sheet,row,'datePromised'),
        formatDateField(sheet,row,'dateStart'),formatDateField(sheet,row,'dateFinish'),
        readField(sheet,row,'status')
      ];
      return `<tr data-open-mo="${esc(mo.key)}">${values.map((value,index)=>`<td>${index===9?statusChips([value]):textCell(value)}</td>`).join('')}<td>${adminEditButton(entry,'Edit MO')}</td></tr>`;
    }).join('');
    return `<div class="style-tab-heading"><div><h3>Manufacturing Orders</h3><p>${items.length} unique MO${items.length===1?'':'s'}${selectedArticle?' for the selected article':''}</p></div></div><div class="style-detail-table-wrap"><table class="style-detail-table"><thead><tr>${headers.map(header=>`<th>${esc(header)}</th>`).join('')}<th><span class="sr-only">Actions</span></th></tr></thead><tbody>${body}</tbody></table></div><p class="style-help-text">Select an MO row to view its material lines.</p>`;
  }

  function renderLines(group){
    const lines=currentFilteredLines(group);
    if(!lines.length)return '<p class="style-empty-note">No material lines match the selected filters.</p>';
    const rows=lines.map(line=>{
      const cons=[line.cons,line.uom].filter(Boolean).join(' ');
      const cells=[line.document,line.part,line.rawCode,line.material,line.supplier,cons,line.fbc,line.ordered,line.delivered,line.balance,line.po];
      return renderMaterialLineRow(line,cells);
    }).join('');
    const editHint=AppState.currentUser?.role==='admin'?'<span class="style-line-edit-hint">Click a row to edit its details.</span>':'';
    return `<div class="style-tab-heading style-lines-heading"><div><h3>Material Lines</h3><p>${lines.length} line${lines.length===1?'':'s'}${selectedArticle?' for the selected article':''}${selectedMo?' for the selected MO':''}</p>${editHint}</div><div class="style-line-actions"><label class="sr-only" for="style-line-search">Search material lines</label><input id="style-line-search" class="filter-input" type="search" value="${esc(state.lineQuery||'')}" placeholder="Search these lines"><button type="button" class="select" data-open-main-lines>Open in Material Lines</button></div></div><div class="style-detail-table-wrap style-lines-table-wrap"><table class="style-detail-table style-material-lines"><thead><tr>${['Document No','Part No','Code Raw Material','Material','Supplier Name','Cons + UOM','FBC','Ordered Qty','Qty Delivered','Balance ORI','PO'].map(header=>`<th>${esc(header)}</th>`).join('')}<th><span class="sr-only">Actions</span></th></tr></thead><tbody>${rows}</tbody></table></div>`;
  }

  function renderMaterialLineRow(line,cells){
    const editable=AppState.currentUser?.role==='admin';
    const lineLabel=line.document||line.part||line.rawCode||`source row ${line.row.index+1}`;
    const rowAttributes=editable
      ?`data-edit-line-sheet="${line.sheetIndex}" data-edit-line-row="${line.row.index}" tabindex="0" aria-label="Edit material line, ${esc(lineLabel)}"`
      :'';
    return `<tr ${rowAttributes}>${cells.map(value=>`<td><span title="${esc(value)}">${textCell(value)}</span></td>`).join('')}<td>${materialLineEditButton(line)}</td></tr>`;
  }

  function materialLineEditButton(line){
    if(AppState.currentUser?.role!=='admin')return '';
    return `<button type="button" class="style-edit-source" data-edit-material-line-sheet="${line.sheetIndex}" data-edit-material-line-row="${line.row.index}">Edit line</button>`;
  }

  function bindMaterialLineRows(container){
    container?.querySelectorAll('[data-edit-line-sheet]').forEach(row=>{
      const openEditor=()=>openMaterialLineEditor(Number(row.dataset.editLineSheet),Number(row.dataset.editLineRow));
      row.onclick=event=>{
        if(event.target.closest('button,a'))return;
        openEditor();
      };
      row.onkeydown=event=>{
        if((event.key!=='Enter'&&event.key!==' ')||event.target!==row)return;
        event.preventDefault();
        openEditor();
      };
    });
    container?.querySelectorAll('[data-edit-material-line-sheet]').forEach(button=>{
      button.onclick=event=>{
        event.stopPropagation();
        openMaterialLineEditor(Number(button.dataset.editMaterialLineSheet),Number(button.dataset.editMaterialLineRow));
      };
    });
  }

  function renderSuppliers(group){
    const lines=currentFilteredLines(group);
    const suppliers=new Map();
    lines.forEach(line=>{
      const name=line.supplier||'';
      if(!suppliers.has(name))suppliers.set(name,{name,codes:new Set(),count:0});
      const item=suppliers.get(name);
      item.count++;
      if(line.rawCode)item.codes.add(line.rawCode);
    });
    const rows=[...suppliers.values()].sort((a,b)=>a.name.localeCompare(b.name,undefined,{sensitivity:'base'}));
    return rows.length?`<div class="style-tab-heading"><div><h3>Suppliers</h3><p>${rows.length} unique supplier${rows.length===1?'':'s'}${selectedArticle||selectedMo?' in the selected lines':''}</p></div></div><div class="style-detail-table-wrap"><table class="style-detail-table"><thead><tr><th>Supplier Name</th><th>Line Count</th><th>Raw Material Codes</th></tr></thead><tbody>${rows.map(item=>`<tr><td>${textCell(item.name)}</td><td>${item.count}</td><td class="style-code-list">${textCell([...item.codes].join(', '))}</td></tr>`).join('')}</tbody></table></div>`:'<p class="style-empty-note">No supplier data is available for this style.</p>';
  }

  function adminEditButton(entry,label){
    if(AppState.currentUser?.role!=='admin')return '';
    return `<button type="button" class="style-edit-source" data-edit-sheet="${entry.sheetIndex}" data-edit-row="${entry.row.index}">${esc(label)}</button>`;
  }

  function bindDetailEvents(group){
    const drawer=document.getElementById('drawer');
    drawer.querySelector('#style-detail-close').onclick=()=>closeDetail();
    drawer.querySelector('#style-previous').onclick=()=>{
      const filtered=matchingGroups(),groups=filtered.some(item=>item.key===group.key)?filtered:model.groups,index=groups.findIndex(item=>item.key===group.key);
      if(index>0)openDetail(groups[index-1].key);
    };
    drawer.querySelector('#style-next').onclick=()=>{
      const filtered=matchingGroups(),groups=filtered.some(item=>item.key===group.key)?filtered:model.groups,index=groups.findIndex(item=>item.key===group.key);
      if(index>=0&&index<groups.length-1)openDetail(groups[index+1].key);
    };
    drawer.querySelectorAll('[data-style-tab]').forEach(button=>{
      button.onclick=()=>{activeTab=button.dataset.styleTab;renderDetail()};
      button.onkeydown=event=>{
        if(event.key!=='ArrowLeft'&&event.key!=='ArrowRight')return;
        const tabs=[...drawer.querySelectorAll('[data-style-tab]')];
        const index=tabs.indexOf(button),step=event.key==='ArrowRight'?1:-1;
        const next=tabs[(index+step+tabs.length)%tabs.length];
        event.preventDefault();next.focus();next.click();
      };
    });
    drawer.querySelector('[data-clear-article]')?.addEventListener('click',()=>{selectedArticle='';selectedMo='';renderDetail()});
    drawer.querySelector('[data-clear-mo]')?.addEventListener('click',()=>{selectedMo='';renderDetail()});
    drawer.querySelectorAll('[data-filter-article]').forEach(button=>button.onclick=()=>{
      selectedArticle=button.dataset.filterArticle;
      selectedMo='';
      activeTab='lines';
      renderDetail();
    });
    drawer.querySelectorAll('[data-open-mo]').forEach(row=>row.onclick=event=>{
      if(event.target.closest('button,a'))return;
      selectedMo=row.dataset.openMo;
      activeTab='lines';
      renderDetail();
    });
    const lineSearch=drawer.querySelector('#style-line-search');
    if(lineSearch)lineSearch.oninput=()=>{state.lineQuery=lineSearch.value;renderLineRows(group)};
    drawer.querySelectorAll('[data-open-main-lines]').forEach(button=>button.onclick=()=>openMainLines(group));
    drawer.querySelectorAll('[data-open-quality]').forEach(button=>button.onclick=()=>App.navigate('quality'));
    drawer.querySelectorAll('[data-edit-sheet]').forEach(button=>button.onclick=()=>{
      openSourceEditor(Number(button.dataset.editSheet),Number(button.dataset.editRow));
    });
    bindMaterialLineRows(drawer.querySelector('.style-material-lines tbody'));
    const editDetails=drawer.querySelector('#style-edit-details');
    if(editDetails)editDetails.onclick=()=>{
      styleDetailsEditing=true;
      renderDetail();
      drawer.querySelector('[data-style-detail-field]')?.focus();
    };
    const cancelDetails=drawer.querySelector('[data-cancel-style-details]');
    if(cancelDetails)cancelDetails.onclick=()=>{
      styleDetailsEditing=false;
      renderDetail();
    };
    const detailsForm=drawer.querySelector('#style-details-form');
    if(detailsForm)detailsForm.onsubmit=event=>{
      event.preventDefault();
      saveStyleDetails(group,detailsForm);
    };
  }

  async function saveStyleDetails(group,form){
    const button=form.querySelector('button[type="submit"]');
    const submitted=new Map([...form.querySelectorAll('[data-style-detail-field]')]
      .map(input=>[input.dataset.styleDetailField,input.value.trim()]));
    const smvField=styleDetailFields.find(field=>field.key==='smv');
    const smvValue=submitted.get('smv')||'';
    const originalSmv=styleFieldEditValue(group,smvField);
    if(smvValue!==originalSmv&&smvValue&&!/^\d{2,}:[0-5]\d$/.test(smvValue)){
      App.toast('Enter SMV as a duration in HH:mm format, such as 00:12.','error');
      form.querySelector('[data-style-detail-field="smv"]')?.focus();
      return;
    }
    const changedFields=styleDetailFields.filter(field=>
      submitted.get(field.key)!==styleFieldEditValue(group,field));
    if(!changedFields.length){
      styleDetailsEditing=false;
      renderDetail();
      return;
    }
    const first=group.lines[0];
    if(!first){
      App.toast(`Could not save style details for ${group.key}: no source rows were found.`,'error');
      return;
    }
    button.disabled=true;
    button.textContent='Saving…';
    try{
      AppState.requireAdmin();
      const changes={},newColumns={};
      changedFields.forEach(field=>{
        const column=sheetColumn(first.sheet,field.key),value=submitted.get(field.key);
        if(column<0)newColumns[field.label]=value;
        else changes[column]=value;
      });
      await AppState.updateRow(first.sheetIndex,first.row.index,changes,newColumns);
      for(const entry of group.lines.slice(1)){
        const rowChanges={};
        changedFields.forEach(field=>{
          const column=sheetColumn(entry.sheet,field.key);
          if(column>=0)rowChanges[column]=submitted.get(field.key);
        });
        if(Object.keys(rowChanges).length){
          await AppState.updateRow(entry.sheetIndex,entry.row.index,rowChanges);
        }
      }
      styleDetailsEditing=false;
      model=buildModel();
      render();
      App.toast(`Saved style details for ${group.key}.`);
    }catch(error){
      button.disabled=false;
      button.textContent='Save details';
      App.toast(`Could not save style details for ${group.key}: ${error.message}`,'error');
    }
  }

  function renderLineRows(group){
    const drawer=document.getElementById('drawer');
    const body=drawer.querySelector('.style-material-lines tbody');
    if(!body)return;
    const query=normalizedText(state.lineQuery||'');
    const lines=currentFilteredLines(group).filter(line=>!query||[
      line.document,line.part,line.rawCode,line.material,line.supplier,line.cons,line.uom,line.fbc,
      line.ordered,line.delivered,line.balance,line.po
    ].some(value=>normalizedText(value).includes(query)));
    body.innerHTML=lines.map(line=>{
      const cons=[line.cons,line.uom].filter(Boolean).join(' ');
      const cells=[line.document,line.part,line.rawCode,line.material,line.supplier,cons,line.fbc,line.ordered,line.delivered,line.balance,line.po];
      return renderMaterialLineRow(line,cells);
    }).join('');
    bindMaterialLineRows(body);
  }

  function openMainLines(group){
    const csiIndex=model.csiIndex;
    if(csiIndex<0)return;
    closeDetail({clearHash:true,restoreFocus:false});
    const global=document.getElementById('globalSearch');
    if(global)global.value=group.key;
    App.navigate('requests');
    SheetTable.selectSheet('requests',csiIndex);
  }

  function openMaterialLineEditor(sheetIndex,rowIndex){
    const sheet=AppState.data.sheets[sheetIndex];
    if(!sheet||AppState.currentUser?.role!=='admin')return;
    const fields=['document','part','rawCode','material','supplier','cons','uom','fbc','ordered','delivered','balance','po'];
    const usedColumns=new Set();
    const columns=fields.map(field=>{
      const sourceIndex=sheetColumn(sheet,field);
      if(sourceIndex<0||usedColumns.has(sourceIndex))return null;
      usedColumns.add(sourceIndex);
      return {header:sheet.headers[sourceIndex],sourceIndex};
    }).filter(Boolean);
    if(!columns.length){
      App.toast('No Material Lines fields were found in this CSI row.','error');
      return;
    }
    openSourceEditor(sheetIndex,rowIndex,columns);
  }

  function openSourceEditor(sheetIndex,rowIndex,displayColumns=null){
    const sheet=AppState.data.sheets[sheetIndex];
    if(!sheet||AppState.currentUser?.role!=='admin')return;
    sourceEditor=true;
    const styleKey=selectedStyle;
    Drawer.open(sheetIndex,rowIndex,displayColumns);
    const drawer=document.getElementById('drawer');
    drawer.classList.add('style-detail-drawer','open');
    const returnToStyle=()=>{
      Drawer.close();
      sourceEditor=false;
      selectedStyle=styleKey;
      isOpen=true;
      drawer.classList.add('style-detail-drawer','open');
      drawer.setAttribute('aria-hidden','false');
      document.getElementById('drawerBackdrop').classList.add('open');
      renderDetail();
    };
    const bindEditorReturn=()=>{
      const close=drawer.querySelector('#drawerClose');
      const cancel=drawer.querySelector('#editCancel');
      if(close)close.onclick=returnToStyle;
      if(cancel)cancel.onclick=returnToStyle;
    };
    const closeButton=drawer.querySelector('#drawerClose');
    if(closeButton)closeButton.onclick=returnToStyle;
    const editButton=drawer.querySelector('#drawerEdit');
    if(editButton){
      const openEdit=editButton.onclick;
      editButton.onclick=()=>{openEdit?.();bindEditorReturn()};
    }
  }

  function getHashStyle(){
    const match=window.location.hash.match(/^#style=(.*)$/);
    if(!match)return '';
    try{return decodeURIComponent(match[1])}catch(error){return match[1]}
  }
  function setStyleHash(styleKey){
    const hash=`#style=${encodeURIComponent(styleKey)}`;
    if(window.location.hash!==hash)history.replaceState(history.state,'',`${window.location.href.split('#')[0]}${hash}`);
  }
  function clearStyleHash(){
    if(window.location.hash.startsWith('#style=')){
      history.replaceState(history.state,'',window.location.href.split('#')[0]);
    }
  }
  let isOpen=false;
  function closeDetail(options={}){
    if(!isOpen&&!sourceEditor)return;
    const restoreKey=selectedStyle;
    const focusTarget=lastTrigger?.isConnected
      ?lastTrigger
      :[...document.querySelectorAll('.style-summary-row')].find(row=>row.dataset.styleKey===restoreKey);
    renderToken++;
    isOpen=false;
    sourceEditor=false;
    styleDetailsEditing=false;
    selectedStyle='';
    selectedArticle='';
    selectedMo='';
    activeTab='overview';
    const drawer=document.getElementById('drawer');
    drawer.classList.remove('open');
    drawer.setAttribute('aria-hidden','true');
    drawer.removeAttribute('role');
    drawer.removeAttribute('aria-modal');
    drawer.removeAttribute('aria-label');
    document.getElementById('drawerBackdrop').classList.remove('open');
    if(options.clearHash!==false)clearStyleHash();
    if(options.restoreFocus!==false){
      (focusTarget||document.getElementById('style-table-search'))?.focus();
    }
    lastTrigger=null;
    window.setTimeout(()=>{
      if(!isOpen)drawer.classList.remove('style-detail-drawer');
    },270);
  }

  function onKeydown(event){
    if(!isOpen)return;
    const drawer=document.getElementById('drawer');
    if(event.key==='Escape'){
      event.preventDefault();event.stopImmediatePropagation();
      if(sourceEditor){
        const close=drawer.querySelector('#drawerClose');
        close?.click();
      }else closeDetail();
      return;
    }
    if(event.key!=='Tab')return;
    const focusable=[...drawer.querySelectorAll('button:not(:disabled),a[href],input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex]:not([tabindex="-1"])')]
      .filter(element=>element.offsetParent!==null);
    if(!focusable.length){event.preventDefault();drawer.focus();return}
    const first=focusable[0],last=focusable[focusable.length-1];
    if(event.shiftKey&&document.activeElement===first){event.preventDefault();last.focus()}
    else if(!event.shiftKey&&document.activeElement===last){event.preventDefault();first.focus()}
  }

  function onBackdropClick(event){
    if(event.target.id!=='drawerBackdrop'||!isOpen)return;
    event.preventDefault();event.stopImmediatePropagation();
    if(sourceEditor){
      document.getElementById('drawerClose')?.click();
    }else closeDetail();
  }

  function search(){
    const global=document.getElementById('globalSearch');
    state.query=global?.value||'';
    state.page=1;
    if(document.querySelector('.style-page-heading'))render();
  }
  function currentSheetIndex(){return model.csiIndex>=0?model.csiIndex:findSheetIndex('csi')}

  function openFromHash(){
    const styleKey=getHashStyle();
    if(!styleKey)return;
    if(document.getElementById('appShell').hidden)return;
    App.navigate('styles');
  }

  function demoWorkbook(){
    const csiHeaders=['Style','Article Number','Season','User','Product Category','Document No','Part No','Code Raw Material','Material','Supplier Name','Cons','UOM','FBC','Ordered Qty','Qty Delivered','Balance ORI','PO','Status Sample','Product FG','Date Promised','Date Start Schedule','Date Finish Schedule','Last Updated BOM'];
    const moHeaders=['Finish Goods Code','MO','Qty Keeping','Qty Shipment','Qty Total','Date Ordered','Date Promised','Start','Finish','Status Sample'];
    const articles=[
      {article:'01 Demo Navy-Electro Royal Blue / Inline Product',fg:'FG-DEMO-794213-01',part:'000000101',mo:['MO/DEMO/1099580','MO/DEMO/1099581']},
      {article:'02 Demo Bright Red / Inline Product',fg:'FG-DEMO-794213-02',part:'000000102',mo:['MO/DEMO/1099582']},
      {article:'03 Demo White-Black / Inline Product',fg:'FG-DEMO-794213-03',part:'000000103',mo:['MO/DEMO/1099583']},
      {article:'04 Demo Green-Volt / Inline Product',fg:'FG-DEMO-794213-04',part:'000000104',mo:['MO/DEMO/1099584']}
    ];
    const materials=[
      ['00001001','Cotton Jersey 180 GSM','AOI Textiles',1.125,'KG'],
      ['00002008','Polyester Sewing Thread','Indo Thread',0.035,'CONE'],
      ['00003012','Printed Size Label','Label Works',1,'PCS']
    ];
    const csiRows=[];
    const moRows=[];
    articles.forEach((article,articleIndex)=>{
      materials.forEach((material,lineIndex)=>{
        csiRows.push([
          'DEMO-STYLE-001',article.article,'AH27','John','Demo Performance Top',
          `DOC-DEMO-${String(articleIndex+1).padStart(2,'0')}`,article.part,material[0],material[1],material[2],
          material[3],material[4],String(0.875+lineIndex/10),String(24+lineIndex*4),'0',String(24+lineIndex*4),
          `PO-DEMO-${articleIndex+1}`,articleIndex%2?'SEALING':'GARMENT TEST',article.fg,
          '2026-09-10','2026-08-01','2026-08-20','2026-07-28'
        ]);
      });
      article.mo.forEach((mo,index)=>{
        moRows.push([article.fg,mo,'24','0','24','2026-07-25','2026-09-10','2026-08-01','2026-08-20',articleIndex%2?'SEALING':'GARMENT TEST']);
      });
    });
    const makeSheet=(name,headers,rows)=>{
      const matrix=[headers,...rows];
      return {
        name,headers:[...headers],matrix,cellMeta:matrix.map(row=>row.map(()=>null)),merges:[],
        startRow:0,endRow:matrix.length-1,startColumn:0,endColumn:headers.length-1,
        headerRow:0,rowCount:rows.length,columnCount:headers.length
      };
    };
    return {
      fileName:'Demo - Multi-Article Styles.xlsx',importedAt:new Date().toISOString(),parseErrors:[],
      sheets:[makeSheet('CSI',csiHeaders,csiRows),makeSheet('MO',moHeaders,moRows)],
      summary:{sheets:2,rows:csiRows.length+moRows.length,columns:csiHeaders.length+moHeaders.length}
    };
  }

  async function seedDemoFile(){
    const setting=await AOIDB.get('settings','multiArticleStylesDemoSeeded');
    if(setting?.value)return;
    const files=await AppState.listFiles();
    const existing=files.find(file=>file.name==='Demo - Multi-Article Styles');
    if(!existing){
      const created=await AppState.createFile(demoWorkbook(),{
        name:'Demo - Multi-Article Styles',
        originalFileName:'Demo - Multi-Article Styles.xlsx',
        note:'Multi-article style detail demonstration',
        importedBy:'system',
        open:false
      });
      await AOIDB.put('files',{...created,demoData:true});
    }else if(!existing.demoData){
      await AOIDB.put('files',{...existing,demoData:true});
    }
    await AOIDB.put('settings',{key:'multiArticleStylesDemoSeeded',value:true});
  }

  async function addDemoBadges(){
    const demoFiles=(await AppState.listFiles()).filter(file=>file.demoData);
    const actions=[...document.querySelectorAll('[data-id]')];
    demoFiles.forEach(file=>{
      const action=actions.find(element=>element.dataset.id===file.id);
      const root=action?.closest('.file-library-card, tr');
      const name=root?.querySelector('.file-card-name, .file-name-copy strong');
      if(!name||root.querySelector('.demo-file-badge'))return;
      const badge=document.createElement('span');
      badge.className='demo-file-badge';
      badge.textContent='Demo data';
      name.insertAdjacentElement('afterend',badge);
    });
  }

  function updateQualityLabels(){
    const app=document.getElementById('app');
    if(app.querySelector('h1')?.textContent.trim()!=='Data Quality')return;
    const subtitle=app.querySelector('.page-heading .page-subtitle');
    if(subtitle)subtitle.textContent='Reports source-cell errors and style consistency issues in the active workbook.';
    const label=app.querySelector('.import-summary-grid .kpi:last-child .kpi-label');
    if(label)label.textContent='Data quality issues';
    const heading=app.querySelector('.quality-panel h2');
    if(heading)heading.textContent='Data quality issues';
    const empty=app.querySelector('.import-no-errors');
    if(empty)empty.textContent='No source-cell or style consistency issues match this search.';
  }

  function install(){
    if(initialized)return;
    initialized=true;
    const oldRender=SheetTable.render,oldSearch=SheetTable.search,oldCurrent=SheetTable.currentSheetIndex;
    SheetTable.render=mode=>mode==='styles'?render():oldRender(mode);
    SheetTable.search=mode=>mode==='styles'?search():oldSearch(mode);
    SheetTable.currentSheetIndex=mode=>mode==='styles'?currentSheetIndex():oldCurrent(mode);
    const originalNavigate=App.navigate;
    App.navigate=route=>{
      if(route!=='styles')closeDetail({clearHash:true});
      return originalNavigate(route);
    };
    const navigateWithQuality=App.navigate;
    App.navigate=route=>{
      const result=navigateWithQuality(route);
      if(route==='quality')requestAnimationFrame(updateQualityLabels);
      return result;
    };
    const originalFilesRender=FilesView.render;
    FilesView.render=async(...args)=>{
      const result=await originalFilesRender(...args);
      await addDemoBadges();
      return result;
    };
    document.getElementById('globalSearch').addEventListener('input',()=>{
      if(document.querySelector('#app .quality-panel'))requestAnimationFrame(updateQualityLabels);
    });
    document.addEventListener('keydown',onKeydown,true);
    document.getElementById('drawerBackdrop').addEventListener('click',onBackdropClick,true);
    document.addEventListener('click',event=>{
      const routeButton=event.target.closest('[data-route]');
      if(routeButton&&routeButton.dataset.route!=='styles')closeDetail({clearHash:true});
      if(routeButton?.dataset.route==='quality')requestAnimationFrame(updateQualityLabels);
      if(event.target.closest('#topImport'))closeDetail({clearHash:true});
      if(event.target.closest('[data-search-suggestion]'))closeDetail({clearHash:true});
    },true);
    window.addEventListener('hashchange',()=>{
      const hashStyle=getHashStyle();
      if(hashStyle){
        openFromHash();
      }else if(isOpen)closeDetail({clearHash:false});
    });
    window.addEventListener('appstate:change',()=>{
      if(activeFileId!==null&&activeFileId!==AppState.activeFileId){
        closeDetail({clearHash:true,restoreFocus:false});
        activeFileId=AppState.activeFileId;
      }
      requestAnimationFrame(()=>{if(!isOpen)openFromHash()});
    });
    const drawer= document.getElementById('drawer');
    const oldDrawerOpen=Drawer.open,oldDrawerCreate=Drawer.openCreate;
    Drawer.open=(...args)=>{
      if(!sourceEditor&&!isOpen)drawer.classList.remove('style-detail-drawer');
      return oldDrawerOpen(...args);
    };
    Drawer.openCreate=(...args)=>{
      if(!sourceEditor&&!isOpen)drawer.classList.remove('style-detail-drawer');
      return oldDrawerCreate(...args);
    };
    drawer.addEventListener('transitionend',event=>{
      if(event.propertyName==='transform'&&!isOpen)drawer.classList.remove('style-detail-drawer');
    });
    AppState.ready.then(async()=>{
      activeFileId=AppState.activeFileId;
      try{
        await seedDemoFile();
        await App.updateAppChrome();
      }catch(error){App.toast(`Could not create the multi-article demo file: ${error.message}`,'error')}
      openFromHash();
    });
    const oldIssues=Validate.issues,oldSummary=Validate.summary;
    Validate.issues=()=>[...oldIssues(),...qualityIssues()];
    Validate.summary=()=>{
      const summary=oldSummary();
      const all=Validate.issues();
      return {...summary,errors:all.length,all};
    };
  }

  return {
    render,search,currentSheetIndex,qualityIssues,buildModel,
    openStyle:openDetail,close:closeDetail,init:install
  };
})();
window.StyleManagement.init();
