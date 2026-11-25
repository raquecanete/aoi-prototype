window.WorkbookParser={
  parse:function parseWorkbook(buffer,fileName){
    function errorText(code){
      const errors={0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A',43:'#GETTING_DATA'};
      return errors[code]||`#ERROR (${String(code??'unknown')})`;
    }
    const workbook=XLSX.read(buffer,{type:'array',cellDates:false,cellText:false,cellFormula:true});
    const parseErrors=[];
    const sheets=workbook.SheetNames.map(name=>{
      const worksheet=workbook.Sheets[name];
      if(!worksheet['!ref'])return {
        name,headers:[],matrix:[],cellMeta:[],merges:[],startRow:0,endRow:-1,
        startColumn:0,endColumn:-1,headerRow:-1,rowCount:0,columnCount:0
      };
      const range=XLSX.utils.decode_range(worksheet['!ref']);
      const matrix=[],cellMeta=[];
      for(let r=range.s.r;r<=range.e.r;r++){
        const values=[],metadata=[];
        for(let c=range.s.c;c<=range.e.c;c++){
          const address=XLSX.utils.encode_cell({r,c});
          const cell=worksheet[address];
          const error=cell?.t==='e'?(cell.w||errorText(cell.v)):'';
          const value=cell?.v;
          values.push(cell?((error||(typeof value==='string'?value.trim():value))??''):'');
          metadata.push(cell&&(cell.z||cell.f||error)?{
            format:cell.z||'',
            formula:cell.f||'',
            error,
            errorValue:cell.t==='e'?cell.v:null
          }:null);
          if(error)parseErrors.push({sheet:name,cell:address,message:`Excel error cell: ${error}`});
        }
        matrix.push(values);
        cellMeta.push(metadata);
      }
      let headerRow=-1,bestScore=0;
      matrix.forEach((row,index)=>{
        const values=row.map(value=>String(value??'').trim()).filter(Boolean);
        const score=values.length*2+new Set(values).size;
        if(score>bestScore){bestScore=score;headerRow=index}
      });
      const headers=headerRow<0?Array.from({length:range.e.c-range.s.c+1},()=> ''):matrix[headerRow].map(value=>String(value??'').trim());
      const merges=(worksheet['!merges']||[]).map(merge=>XLSX.utils.encode_range(merge));
      const rowCount=matrix.length-(headerRow<0?0:headerRow+1);
      return {
        name,headers,matrix,cellMeta,merges,
        startRow:range.s.r,endRow:range.e.r,
        startColumn:range.s.c,endColumn:range.e.c,
        headerRow,rowCount,
        columnCount:range.e.c-range.s.c+1
      };
    });
    const summary={
      sheets:sheets.length,
      rows:sheets.reduce((total,sheet)=>total+sheet.rowCount,0),
      columns:sheets.reduce((total,sheet)=>total+sheet.columnCount,0)
    };
    return {fileName,importedAt:new Date().toISOString(),sheets,parseErrors,summary};
  }
};
