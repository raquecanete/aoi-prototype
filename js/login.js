window.DemoLogin=(()=>{
  const SESSION_KEY='aoi-admin-demo-session';
  const loginView=document.getElementById('loginView');
  const appShell=document.getElementById('appShell');
  const form=document.getElementById('loginForm');
  const error=document.getElementById('loginError');
  const passwordInput=document.getElementById('loginPassword');
  const passwordToggle=document.getElementById('togglePassword');

  passwordToggle.addEventListener('click',()=>{
    const showPassword=passwordInput.type==='password';
    passwordInput.type=showPassword?'text':'password';
    passwordToggle.setAttribute('aria-pressed',String(showPassword));
    passwordToggle.setAttribute('aria-label',showPassword?'Hide password':'Show password');
  });

  function resetPassword(){
    passwordInput.value='';
    passwordInput.type='password';
    passwordToggle.setAttribute('aria-pressed','false');
    passwordToggle.setAttribute('aria-label','Show password');
  }

  function showLogin(){
    loginView.hidden=false;
    appShell.hidden=true;
    error.hidden=true;
    resetPassword();
    document.getElementById('loginUsername').focus();
  }
  function showApp(){
    resetPassword();
    loginView.hidden=true;
    appShell.hidden=false;
  }

  function readSession(){
    try{return sessionStorage.getItem(SESSION_KEY)}catch(error){
      App.toast('This browser blocked session storage. Sign in again after closing this tab.','error');
      return '';
    }
  }
  function writeSession(userId){
    try{sessionStorage.setItem(SESSION_KEY,userId)}catch(error){
      App.toast('This browser blocked session storage. You can continue, but sign-in will not persist for this tab.','error');
    }
  }
  function clearSession(){
    try{sessionStorage.removeItem(SESSION_KEY)}catch(error){
      App.toast('This browser blocked session storage. Close this tab to end the mock session.','error');
    }
  }

  const savedUser=readSession();
  if(savedUser){
    AppState.ready.then(()=>AppState.setCurrentUser('admin')).then(()=>{
      writeSession('admin');
      showApp();
    }).catch(error=>{
      App.toast(`Could not restore the mock session: ${error.message}`,'error');
      showLogin();
    });
  }else showLogin();

  form.addEventListener('submit',async event=>{
    event.preventDefault();
    const username=document.getElementById('loginUsername').value.trim().toLowerCase();
    const password=document.getElementById('loginPassword').value;
    if(username!=='admin'||password!=='AOI2026!'){
      error.hidden=false;
      document.getElementById('loginUsername').setAttribute('aria-invalid','true');
      passwordInput.setAttribute('aria-invalid','true');
      document.getElementById('loginPassword').focus();
      return;
    }
    try{
      await AppState.ready;
      await AppState.setCurrentUser('admin');
      writeSession('admin');
      error.hidden=true;
      showApp();
    }catch(loginError){
      error.textContent=`Sign-in failed: ${loginError.message}`;
      error.hidden=false;
    }
  });

  form.addEventListener('input',()=>{
    error.hidden=true;
    passwordInput.removeAttribute('aria-invalid');
    document.getElementById('loginUsername').removeAttribute('aria-invalid');
  });
  document.getElementById('signOut').addEventListener('click',()=>{
    clearSession();
    showLogin();
  });

  return {};
})();
