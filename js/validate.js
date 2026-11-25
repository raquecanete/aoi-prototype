window.Validate=(()=>{
  function issues(){
    return AppState.data.parseErrors.map(error=>({
      severity:'error',
      title:'Excel error cell',
      detail:`${error.sheet}!${error.cell}: ${error.message}`
    }));
  }
  function summary(){
    return {
      sheets:AppState.data.summary?.sheets||0,
      rows:AppState.data.summary?.rows||0,
      columns:AppState.data.summary?.columns||0,
      errors:AppState.data.parseErrors.length,
      all:issues()
    };
  }
  return {issues,summary};
})();
