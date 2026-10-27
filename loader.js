/* Personal web app loader — OneDrive Excel backend.
   Sign-in with a personal Microsoft account (OAuth 2.0 code flow + PKCE, no library),
   then reads/writes OneDrive › <folder> › Finance.xlsx through Microsoft Graph and loads the private app.html from there.
   Exposes window.APP and a window.claude.use() shim so the app code runs unchanged. */
(function(){
'use strict';
const CFGX=Object.assign({clientId:'',folder:'/Finance',workbook:'Finance.xlsx',config:'finance.json',app:'app.html',authority:'consumers'},window.APP_CONFIG||{});
const GRAPH='https://graph.microsoft.com/v1.0';
const LOGIN='https://login.microsoftonline.com/'+CFGX.authority+'/oauth2/v2.0';
const SCOPES='openid profile offline_access User.Read Files.ReadWrite';
const REDIRECT=location.origin+location.pathname;
const TXCOLS=['id','type','date','account','category','sub','people','vendor','note','amount','currency','sar','bulk','toAccount','received','receivedCurrency','fee','rate','fxRate','sarCharged','recurringId','dueKey','matched','receiptUrl','createdAt','extra'];
const NUMCOLS=new Set(['amount','sar','received','fee','rate','fxRate','sarCharged']);
const BOOLCOLS=new Set(['bulk','matched']);
const TEXTDATE=new Set(['date']);
const enc=encodeURIComponent;
const wbPath=()=>GRAPH+'/me/drive/root:'+CFGX.folder.replace(/\/$/,'')+'/'+enc(CFGX.workbook)+':/workbook';

/* ---------------- auth ---------------- */
const LS={get(k){try{return localStorage.getItem(k)}catch(e){return null}},set(k,v){try{localStorage.setItem(k,v)}catch(e){}},del(k){try{localStorage.removeItem(k)}catch(e){}}};
const SS={get(k){try{return sessionStorage.getItem(k)}catch(e){return null}},set(k,v){try{sessionStorage.setItem(k,v)}catch(e){}},del(k){try{sessionStorage.removeItem(k)}catch(e){}}};
const b64url=buf=>btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
const rand=n=>{const a=new Uint8Array(n);crypto.getRandomValues(a);return b64url(a).slice(0,n)};
let TOK=null;try{TOK=JSON.parse(LS.get('fin.tok')||'null')}catch(e){}
function saveTok(j){TOK={access:j.access_token,exp:Date.now()+((j.expires_in||3600)-90)*1000,refresh:j.refresh_token||(TOK&&TOK.refresh)};LS.set('fin.tok',JSON.stringify(TOK))}
async function login(prompt){
  if(!CFGX.clientId){throw new Error('no-client-id')}
  const verifier=rand(64),state=rand(16);
  const challenge=b64url(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(verifier)));
  SS.set('fin.pkce',JSON.stringify({verifier,state,hash:location.hash}));
  const q=new URLSearchParams({client_id:CFGX.clientId,response_type:'code',redirect_uri:REDIRECT,scope:SCOPES,code_challenge:challenge,code_challenge_method:'S256',state,response_mode:'query'});
  if(prompt)q.set('prompt',prompt);
  location.assign(LOGIN+'/authorize?'+q);
  return new Promise(()=>{});
}
async function tokenPost(params){
  const r=await fetch(LOGIN+'/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams(Object.assign({client_id:CFGX.clientId,scope:SCOPES},params))});
  const j=await r.json().catch(()=>({}));if(!r.ok)throw Object.assign(new Error(j.error_description||j.error||('token '+r.status)),{code:j.error||'token'});
  return j;
}
async function handleRedirect(){
  const u=new URL(location.href);const code=u.searchParams.get('code'),err=u.searchParams.get('error');
  if(!code&&!err)return false;
  let p={};try{p=JSON.parse(SS.get('fin.pkce')||'{}')}catch(e){}
  SS.del('fin.pkce');history.replaceState(null,'',REDIRECT+(p.hash||''));
  if(err)throw Object.assign(new Error(u.searchParams.get('error_description')||err),{code:err});
  if(!p.verifier||u.searchParams.get('state')!==p.state)throw Object.assign(new Error('Sign-in state mismatch, please try again'),{code:'state'});
  saveTok(await tokenPost({grant_type:'authorization_code',code,redirect_uri:REDIRECT,code_verifier:p.verifier}));
  return true;
}
async function token(){
  if(TOK&&TOK.access&&Date.now()<TOK.exp)return TOK.access;
  if(TOK&&TOK.refresh){try{saveTok(await tokenPost({grant_type:'refresh_token',refresh_token:TOK.refresh}));return TOK.access}catch(e){}}
  LS.del('fin.tok');TOK=null;return login();
}
function signOut(){LS.del('fin.tok');TOK=null;location.assign(LOGIN+'/logout?post_logout_redirect_uri='+enc(REDIRECT))}

/* ---------------- graph ---------------- */
async function g(method,url,body,opt={}){
  for(let attempt=0;attempt<4;attempt++){
    const t=await token();
    const h={Authorization:'Bearer '+t};if(body!==undefined&&!opt.raw)h['Content-Type']='application/json';if(opt.type)h['Content-Type']=opt.type;
    const r=await fetch(url.startsWith('http')?url:GRAPH+url,{method,headers:h,body:body===undefined?undefined:(opt.raw?body:JSON.stringify(body))});
    if(r.status===401&&attempt===0){TOK.exp=0;continue}
    if((r.status===429||r.status===503||r.status===504)&&attempt<3){const ra=+(r.headers.get('Retry-After')||2);await new Promise(ok=>setTimeout(ok,Math.min(ra,10)*1000));continue}
    if(!r.ok){let j={};try{j=await r.json()}catch(e){}const e=new Error((j.error&&j.error.message)||('Graph '+r.status));e.code=(j.error&&j.error.code)||String(r.status);e.status=r.status;throw e}
    if(opt.blob)return r.blob();
    if(r.status===204)return null;
    const ct=r.headers.get('content-type')||'';return ct.includes('json')?r.json():r.text();
  }
}
const tableValues=async name=>{const j=await g('GET',wbPath()+'/tables/'+name+'/range?$select=values');return j.values||[]};
const colValues=async(name,idx)=>{const j=await g('GET',wbPath()+'/tables/'+name+'/columns/itemAt(index='+idx+')/range?$select=values');return (j.values||[]).map(r=>r[0])};
const addRow=(name,row)=>g('POST',wbPath()+'/tables/'+name+'/rows',{values:[row]});
const patchRow=(name,i,row)=>g('PATCH',wbPath()+'/tables/'+name+'/rows/itemAt(index='+i+')',{values:[row]});
const delRow=(name,i)=>g('DELETE',wbPath()+'/tables/'+name+'/rows/itemAt(index='+i+')');

/* ---------------- value helpers ---------------- */
const serialToIso=n=>{const d=new Date(Math.round((n-25569)*86400000));return d.toISOString().slice(0,10)};
const isBlankRow=r=>r.every(v=>v===''||v===null);
function rowToTx(h,r){
  const o={};let id='';
  h.forEach((c,i)=>{let v=r[i];if(v===''||v===null||v===undefined)return;
    if(c==='id'){id=String(v);return}
    if(TEXTDATE.has(c)&&typeof v==='number')v=serialToIso(v);
    if(c==='createdAt'&&typeof v==='number')v=new Date(Math.round((v-25569)*86400000)).toISOString();
    if(NUMCOLS.has(c)){const n=typeof v==='number'?v:parseFloat(String(v).replace(/,/g,''));if(!isNaN(n))v=n}
    if(BOOLCOLS.has(c))v=(v===true||String(v).toUpperCase()==='TRUE');
    if(c==='people'){v=String(v).split(',').map(s=>s.trim()).filter(Boolean)}
    if(c==='receiptUrl'){o.receipt={id:'onedrive',url:String(v)};return}
    if(c==='extra'){try{Object.assign(o,JSON.parse(v))}catch(e){}return}
    if(!NUMCOLS.has(c)&&!BOOLCOLS.has(c)&&c!=='people')v=String(v);
    o[c]=v});
  if(!o.people)o.people=[];
  return {id,data:o};
}
function txToRow(id,d){
  const extra={};Object.keys(d).forEach(k=>{if(!TXCOLS.includes(k)&&k!=='receipt')extra[k]=d[k]});
  return TXCOLS.map(c=>{
    if(c==='id')return id;
    if(c==='people')return (d.people||[]).join(', ');
    if(c==='receiptUrl')return d.receipt&&d.receipt.url||'';
    if(c==='extra')return Object.keys(extra).length?JSON.stringify(extra):'';
    const v=d[c];if(v===undefined||v===null)return '';
    if(typeof v==='object')return JSON.stringify(v);
    return v});
}
const newId=()=>{const a='abcdefghijklmnopqrstuvwxyz0123456789';let s='';const r=new Uint8Array(20);crypto.getRandomValues(r);r.forEach(x=>s+=a[x%36]);return s};

/* ---------------- local store + listeners ---------------- */
const TX=new Map();          // id -> data
const ST=new Map();          // 'col/id' -> data
let HEAD=TXCOLS.slice();
const colL=[],docL=[];
let pending=0,lastErr=null;
const snapDocs=docs=>({docs:docs.map(([id,d])=>({id,data:()=>JSON.parse(JSON.stringify(d)),exists:true})),size:docs.length,empty:!docs.length});
function colDocs(col){if(col==='tx')return [...TX.entries()];const p=col+'/';return [...ST.entries()].filter(([k])=>k.startsWith(p)&&k.indexOf('/',p.length)<0).map(([k,d])=>[k.slice(p.length),d])}
function getDoc(path){const [col,id]=path.split('/');return col==='tx'?TX.get(id):ST.get(path)}
function emit(col){colL.filter(l=>!col||l.col===col).forEach(l=>{try{l.cb(snapDocs(colDocs(l.col)))}catch(e){console.error(e)}});
  docL.filter(l=>!col||l.path.startsWith(col+'/')).forEach(l=>{const d=getDoc(l.path);try{l.cb({exists:!!d,id:l.path.split('/')[1],data:()=>d?JSON.parse(JSON.stringify(d)):undefined})}catch(e){console.error(e)}})}
function status(msg){const el=document.getElementById('app-sync');if(el)el.textContent=msg}

async function readAll(){
  const [t,s]=await Promise.all([tableValues('Transactions'),tableValues('Store')]);
  HEAD=(t[0]||TXCOLS).map(String);TX.clear();
  t.slice(1).forEach(r=>{if(isBlankRow(r))return;const x=rowToTx(HEAD,r);if(x.id)TX.set(x.id,x.data)});
  ST.clear();
  s.slice(1).forEach(r=>{if(isBlankRow(r))return;const col=String(r[0]),id=String(r[1]);try{ST.set(col+'/'+id,JSON.parse(r[2]))}catch(e){console.warn('bad json',col,id)}});
}
async function refresh(){if(pending)return;try{await readAll();emit();status('Synced with OneDrive · '+new Date().toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}))}catch(e){status('Could not reach OneDrive ('+(e.code||e.message)+')')}}

/* write queue keeps writes in order */
let chain=Promise.resolve();
function queue(fn){pending++;status('Saving to OneDrive…');const p=chain.then(fn);chain=p.catch(()=>{}).then(()=>{pending--;if(!pending)status('Saved to OneDrive · '+new Date().toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'}))});
  return p.catch(e=>{lastErr=e;status('Not saved: '+(e.code||e.message)+' — reloading from OneDrive');setTimeout(refresh,1500);throw e})}
async function findTxIndex(id){const ids=await colValues('Transactions',HEAD.indexOf('id'));return ids.slice(1).findIndex(v=>String(v)===id)}
async function findStoreIndex(col,id){const [a,b]=await Promise.all([colValues('Store',0),colValues('Store',1)]);for(let i=1;i<a.length;i++)if(String(a[i])===col&&String(b[i])===id)return i-1;return -1}
function rowForHead(id,d){const full=txToRow(id,d);return HEAD.map(c=>{const i=TXCOLS.indexOf(c);return i<0?'':full[i]})}
async function writeTx(id,d){const i=await findTxIndex(id);const row=rowForHead(id,d);if(i<0)await addRow('Transactions',row);else await patchRow('Transactions',i,row)}
async function removeTx(id){const i=await findTxIndex(id);if(i>=0)await delRow('Transactions',i)}
async function writeStore(col,id,d){const i=await findStoreIndex(col,id);const row=[col,id,JSON.stringify(d),new Date().toISOString()];if(i<0)await addRow('Store',row);else await patchRow('Store',i,row)}
async function removeStore(col,id){const i=await findStoreIndex(col,id);if(i>=0)await delRow('Store',i)}

function docRef(path){
  const [col,id]=path.split('/');
  const put=d=>{if(col==='tx')TX.set(id,d);else ST.set(path,d);emit(col);return queue(()=>col==='tx'?writeTx(id,d):writeStore(col,id,d))};
  return {id,
    set:data=>put(JSON.parse(JSON.stringify(data))),
    update:data=>{const cur=getDoc(path)||{};const n=Object.assign({},cur);Object.entries(data).forEach(([k,v])=>{if(v&&v.__delete__)delete n[k];else n[k]=v});return put(n)},
    delete:()=>{if(col==='tx')TX.delete(id);else ST.delete(path);emit(col);return queue(()=>col==='tx'?removeTx(id):removeStore(col,id))},
    get:async()=>{const d=getDoc(path);return {exists:!!d,id,data:()=>d}},
    onSnapshot:(cb,err)=>{const l={path,cb,err};docL.push(l);const d=getDoc(path);setTimeout(()=>cb({exists:!!d,id,data:()=>d?JSON.parse(JSON.stringify(d)):undefined}),0);return ()=>{const i=docL.indexOf(l);if(i>=0)docL.splice(i,1)}}};
}
function colRef(col){
  const q={orderBy:()=>q,limit:()=>q,where:()=>q,
    add:async data=>{const id=newId();await docRef(col+'/'+id).set(data);return {id}},
    doc:id=>docRef(col+'/'+(id||newId())),
    get:async()=>snapDocs(colDocs(col)),
    onSnapshot:(cb,err)=>{const l={col,cb,err};colL.push(l);setTimeout(()=>cb(snapDocs(colDocs(col))),0);return ()=>{const i=colL.indexOf(l);if(i>=0)colL.splice(i,1)}}};
  return q;
}
const DBSHIM={collection:colRef,doc:docRef};

/* ---------------- files ---------------- */
const ASSETS={upload:async f=>{
  const name=new Date().toISOString().slice(0,10)+'-'+Date.now().toString(36)+'-'+(f.name||'file').replace(/[^\w.\-]+/g,'_');
  const item=await g('PUT',GRAPH+'/me/drive/root:'+CFGX.folder.replace(/\/$/,'')+'/Files/'+enc(name)+':/content',f,{raw:true,type:f.type||'application/octet-stream'});
  return {id:item.id,url:item.webUrl,name:item.name}}};
const dataUrl=f=>new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=no;r.readAsDataURL(f)});
const DL={save:async({filename,data,mimeType})=>{const blob=data instanceof Blob?data:new Blob([typeof data==='string'?data:JSON.stringify(data,null,1)],{type:mimeType||(String(filename).endsWith('.csv')?'text/csv':'application/json')});
  const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=filename||'download';document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000);return {ok:true}}};

/* ---------------- rates (FX + silver), refreshed from the browser every 12 h ---------------- */
const FXCODES=['EUR','GBP','TRY','KWD','JOD','MAD','CHF','JPY','CNY','INR','MYR','THB','IDR','SGD','AUD','CAD','ZAR','GEL','AZN','KRW','PKR'];
function riyadhNow(){const d=new Date(Date.now()+3*3600000);return d.toISOString().slice(0,16).replace('T',' ')}
async function refreshRates(force){
  try{
    const cur=ST.get('settings/rates')||{};const age=cur.asOf?(Date.now()-new Date(cur.asOf.replace(' ','T')+':00+03:00').getTime()):Infinity;
    if(!force&&age<12*3600000)return;
    const [fx,ag]=await Promise.all([fetch('https://open.er-api.com/v6/latest/USD').then(r=>r.json()).catch(()=>null),fetch('https://api.gold-api.com/price/XAG').then(r=>r.json()).catch(()=>null)]);
    const n=Object.assign({},cur);const today=riyadhNow().slice(0,10);
    if(cur.asOf&&cur.asOf.slice(0,10)<today)n.prev={silverSarPerGram:cur.silverSarPerGram,egpPerSar:cur.egpPerSar,asOf:cur.asOf};
    let ok=false;
    if(fx&&fx.rates&&fx.rates.EGP){ok=true;n.egpPerSar=Math.round(fx.rates.EGP/3.75*1000)/1000;n.usdSar=3.75;n.aedSar=1.0211;
      const m=Object.assign({},cur.fx||{});FXCODES.forEach(c=>{if(fx.rates[c])m[c]=Math.round(3.75/fx.rates[c]*1e6)/1e6});m.QAR=Math.round(3.75/3.64*1e6)/1e6;m.BHD=Math.round(3.75/0.376*1e6)/1e6;m.OMR=Math.round(3.75/0.3845*1e6)/1e6;n.fx=m}
    if(ag&&ag.price){ok=true;n.silverSarPerGram=Math.round(ag.price*3.75/31.1034768*100)/100}
    if(!ok)return;
    n.asOf=riyadhNow();n.source='open.er-api.com FX; gold-api.com silver (refreshed by the web app)';
    await docRef('settings/rates').set(n);
    await docRef('rateHistory/'+today).set({silver:n.silverSarPerGram,egp:n.egpPerSar,asOf:n.asOf});
  }catch(e){console.warn('rates',e)}
}

/* ---------------- boot ---------------- */
function overlay(html){let o=document.getElementById('app-ov');if(!o){o=document.createElement('div');o.id='app-ov';o.style.cssText='position:fixed;inset:0;z-index:999;background:var(--bg,#FBF7F1);display:grid;place-items:center;padding:24px;font-family:system-ui,sans-serif;color:#2E2420;text-align:center';document.body.appendChild(o)}o.innerHTML='<div style="max-width:420px">'+html+'</div>';return o}
function hideOverlay(){const o=document.getElementById('app-ov');if(o)o.remove()}
function boot(){return BOOT||(BOOT=bootOnce())}
async function bootOnce(){
  try{
    if(!CFGX.clientId||/PASTE/i.test(CFGX.clientId)){overlay('<h2>One step left</h2><p>Open <b>config.js</b> in your GitHub repository and paste your Application (client) ID from Microsoft Entra. Then reload this page.</p>');return new Promise(()=>{})}
    await handleRedirect();
    if(!TOK){overlay('<div style="font-size:48px">💰</div><h2 style="font-family:Georgia,serif">Family Finance</h2><p>Sign in with the Microsoft account that holds your OneDrive.</p><button id="app-login" style="font:600 16px system-ui;padding:12px 22px;border-radius:12px;border:0;background:#4F6B40;color:#fff;cursor:pointer">Sign in with Microsoft</button>');
      await new Promise(ok=>document.getElementById('app-login').onclick=ok);await login('select_account')}
    overlay('<p>Loading your data from OneDrive…</p>');
    const base=GRAPH+'/me/drive/root:'+CFGX.folder.replace(/\/$/,'');
    const [conf,hist]=await Promise.all([
      g('GET',base+'/'+enc(CFGX.config)+':/content').then(x=>typeof x==='string'?JSON.parse(x):x),
      tableValues('History'),readAll()]);
    const hh=(hist[0]||[]).map(String);const ix=c=>hh.indexOf(c);
    const rows=hist.slice(1).filter(r=>!isBlankRow(r)).map(r=>{let d=r[ix('date')];if(typeof d==='number')d=serialToIso(d);
      return [String(d),String(r[ix('account')]),String(r[ix('category')]),String(r[ix('sub')]),String(r[ix('people')]||'').split(',').map(s=>s.trim()).filter(Boolean),String(r[ix('vendor')]||''),String(r[ix('note')]||''),+r[ix('amount')]||0,String(r[ix('currency')]||'SAR'),+r[ix('sar')]||0]});
    hideOverlay();
    setInterval(()=>{if(document.visibilityState==='visible')refresh()},90000);
    document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')refresh()});
    setTimeout(()=>refreshRates(false),4000);
    g('GET','/me?$select=displayName').then(me=>{APP.user=me&&me.displayName;status('Synced with OneDrive'+(APP.user?' · '+APP.user:''))}).catch(()=>{});
    return {c:conf,h:{cols:['date','account','category','sub','people','vendor','note','amount','currency','sar'],rows}};
  }catch(e){
    console.error(e);
    const msg=e.status===404?'Could not find <b>'+CFGX.folder+'/'+CFGX.workbook+'</b> and <b>'+CFGX.config+'</b> in your OneDrive. Upload both files to that folder and reload.':esc(e.message||String(e));
    overlay('<h2>Could not open your data</h2><p>'+msg+'</p><p><button onclick="location.reload()" style="padding:10px 18px;border-radius:10px;border:1px solid #ccc;background:#fff">Try again</button> <button onclick="APP.signOut()" style="padding:10px 18px;border-radius:10px;border:1px solid #ccc;background:#fff">Sign out</button></p>');
    return new Promise(()=>{});
  }
}
let BOOT=null;
async function start(){
  await boot();
  overlay('<p>Opening the app…</p>');
  const base=GRAPH+'/me/drive/root:'+CFGX.folder.replace(/\/$/,'');
  let html;try{html=await g('GET',base+'/'+enc(CFGX.app)+':/content');if(typeof html!=='string')html=await html.text()}catch(e){overlay('<h2>App file missing</h2><p>Upload <b>'+CFGX.app+'</b> to OneDrive › '+CFGX.folder+' and reload.</p>');return}
  const doc=new DOMParser().parseFromString(html,'text/html');
  doc.head.querySelectorAll('style,link[rel=stylesheet],link[rel=preconnect]').forEach(n=>document.head.appendChild(document.importNode(n,true)));
  const scripts=[...doc.body.querySelectorAll('script')];scripts.forEach(n=>n.remove());
  doc.body.querySelectorAll('style,link[rel=stylesheet]').forEach(n=>document.head.appendChild(document.importNode(n,true)));
  const t=doc.querySelector('title');if(t)document.title=t.textContent;
  document.body.querySelectorAll('[data-shell]').forEach(n=>n.remove());
  [...doc.body.childNodes].forEach(n=>document.body.appendChild(document.importNode(n,true)));
  hideOverlay();
  for(const sc of scripts){if(sc.src)continue;const n=document.createElement('script');n.textContent=sc.textContent;document.body.appendChild(n)}
  setTimeout(()=>status('Synced with OneDrive'+(APP.user?' · '+APP.user:'')),500);
}
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

window.APP={boot,start,signOut,refresh,refreshRates,dataUrl,config:CFGX,user:null};
window.claude={use:async name=>{
  if(name==='db')return DBSHIM;
  if(name==='assets')return ASSETS;
  if(name==='downloads')return DL;
  throw Object.assign(new Error(name+' is not available in the web version'),{code:'unavailable'})}};
})();
