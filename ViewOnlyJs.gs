/**
 * ViewOnlyJs.gs   (NEW FILE)
 *
 * Browser-side part of "Open To All". doGet() in WebApp.gs inserts this code
 * into Index.html just before the page starts, so Index.html itself is NOT edited.
 *
 * It does 3 things:
 *   1. adds the "Open to all" selector to Settings > Contacts
 *   2. for a VIEWER: hides every changing button, locks every field, opens records read-only
 *   3. shows a "VIEW ONLY" badge under the viewer's e-mail
 * The real protection is on the server (withLock_ / denyViewer_ in Utils.gs).
 *
 * NOTE: keep this code free of backticks, dollar-brace and backslashes.
 */
var VIEW_ONLY_JS_ = `
function RO(){return !!(S.boot&&S.boot.role==='VIEWER');}
SET_IDS.push('openAll');

/* ---- hide buttons: Settings (view only) and Calendar Sync (viewers cannot use it) ---- */
(function(){
  var setN=['saveSettings(','sendTest(','addItem(','saveList(','rmItem(','resetLogo(','logoFile'];
  var syncN=['syncOne(','runSync(','openReset(','toggleAuto(','resolveC('];
  var a=setN.map(function(n){return 'body.ro [onclick*="'+n+'"]';});
  a.push('body.ro .ladd','body.ro .savebar','body.ro .chip b');
  var b=syncN.map(function(n){return 'body.vw [onclick*="'+n+'"]';});
  var st=document.createElement('style');
  st.textContent=a.concat(b).join(',')+'{display:none!important}'+
    'body.vw .who::after{content:"VIEWER";display:inline-block;margin-top:6px;padding:2px 8px;border-radius:99px;background:#D4AF37;color:#2a2100;font-weight:700;font-size:10px}';
  document.head.appendChild(st);
})();

/* ---- ro = viewer on the Settings page, vw = viewer anywhere ---- */
function setRoClass(){
  document.body.classList.toggle('vw',RO());
  document.body.classList.toggle('ro',RO()&&S.view==='settings');
}
var _applyBrand=applyBrand;
applyBrand=function(b){_applyBrand(b);setRoClass();};
var _buildNav=buildNav;
buildNav=function(){_buildNav();setRoClass();};
var _go=go;
go=function(){var r=_go.apply(this,arguments);setRoClass();return r;};

var _call=call;
call=function(fn){
  var p=_call.apply(this,arguments);
  if(fn==='appGetSettings'){p.then(function(s){S.openAllVal=s&&s.openAll;},function(){});}
  return p;
};

function saveSettings(){
  var d={};SET_IDS.forEach(function(k){var e=$('s_'+k);if(e)d[k]=e.value;});
  var b=$('ssave'),er=$('serr');er.textContent='';b.disabled=true;b.textContent='Saving…';
  call('appSaveSettings',d).then(function(bt){
    S.boot=bt;applyBrand(bt.brand);buildNav();persist();
    b.disabled=false;b.textContent='Save settings';
    toast(bt.shareNote||'Settings saved',!!bt.shareNote);
  }).catch(function(e){b.disabled=false;b.textContent='Save settings';er.textContent=e.message||String(e);});
}

/* ---- second line of defence: only Settings and Calendar Sync actions are blocked ---- */
['saveSettings','sendTest','saveList','addItem','rmItem','pickLogo','resetLogo',
 'syncOne','runSync','openReset','startReset','toggleAuto','resolveC'].forEach(function(n){
  var f=window[n];if(typeof f!=='function')return;
  window[n]=function(){
    if(RO()){toast('Viewers cannot change Settings or Calendar Sync.',1);return;}
    return f.apply(this,arguments);
  };
});

function injectOpenAll(){
  var f=document.querySelector('#p_contact .form');
  if(!f||$('s_openAll'))return;
  var d=document.createElement('div');
  d.className='fld full';
  d.innerHTML='<label>Open to all (everyone can use the system, only Owner / EA can change Settings)</label>'+
    '<select id="s_openAll"><option value="No">No: only the e-mails listed above can enter</option>'+
    '<option value="Yes">Yes: anyone with a Google account can enter, add and edit records, but only view Settings</option></select>'+
    '<small>Visitors cannot change Settings or Calendar Sync. Saving this switch also changes the spreadsheet sharing (see your Share settings).</small>';
  f.appendChild(d);
  d.querySelector('select').value=S.openAllVal||'No';
}
new MutationObserver(function(){
  injectOpenAll();
  if(!RO())return;
  [].forEach.call(document.querySelectorAll('#main .panel input,#main .panel select,#main .panel textarea'),function(e){if(!e.disabled)e.disabled=true;});
}).observe($('main'),{childList:true,subtree:true});
`;
