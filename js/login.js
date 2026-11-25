window.DemoLogin=(()=>{
  const SESSION_KEY='aoi-admin-demo-session';
  const USERNAME='admin';
  const PASSWORD='AOI2026!';
  const loginView=document.getElementById('loginView');
  const appShell=document.getElementById('appShell');
  const form=document.getElementById('loginForm');
  const error=document.getElementById('loginError');

  function showLogin(){
    loginView.hidden=false;
    appShell.hidden=true;
    error.hidden=true;
    document.getElementById('loginPassword').value='';
    document.getElementById('loginUsername').focus();
  }
  function showApp(){
    loginView.hidden=true;
    appShell.hidden=false;
  }

  if(sessionStorage.getItem(SESSION_KEY)==='active')showApp();
  else showLogin();

  form.addEventListener('submit',event=>{
    event.preventDefault();
    const username=document.getElementById('loginUsername').value.trim().toLowerCase();
    const password=document.getElementById('loginPassword').value;
    if(username!==USERNAME||password!==PASSWORD){
      error.hidden=false;
      document.getElementById('loginPassword').focus();
      return;
    }
    sessionStorage.setItem(SESSION_KEY,'active');
    error.hidden=true;
    showApp();
  });

  form.addEventListener('input',()=>{error.hidden=true});
  document.getElementById('signOut').addEventListener('click',()=>{
    sessionStorage.removeItem(SESSION_KEY);
    showLogin();
  });

  return {};
})();
