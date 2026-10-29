const vm=require('vm'),fs=require('fs');
function makeGas(dir){
  const props={},cache={},sheets={},drive={files:{},seq:0};
  class Range{constructor(s,r,c,nr,nc){Object.assign(this,{s,r,c,nr:nr||1,nc:nc||1})}
    getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++){const v=(this.s.d[this.r-1+i]||[])[this.c-1+j];row.push(v===undefined?'':v)}o.push(row)}return o}
    getValue(){return this.getValues()[0][0]}
    setValues(v){v.forEach((row,i)=>row.forEach((x,j)=>this.s.set(this.r+i,this.c+j,x)));return this}
    setValue(x){this.s.set(this.r,this.c,x);return this}
    setNumberFormat(){return this}}
  class Sheet{constructor(n){this.n=n;this.d=[]}
    set(r,c,x){while(this.d.length<r)this.d.push([]);const row=this.d[r-1];while(row.length<c)row.push('');row[c-1]=x}
    appendRow(a){this.d.push(a.slice())}
    getLastRow(){return this.d.length}
    getLastColumn(){return this.d.reduce((m,r)=>Math.max(m,r.length),0)}
    getRange(r,c,nr,nc){return new Range(this,r,c,nr,nc)}
    setFrozenRows(){}}
  const ss={getSheetByName:n=>sheets[n]||null,insertSheet:n=>sheets[n]=new Sheet(n),getName:()=>'mock',copy:()=>ss};
  const ctx={console,Date,JSON,Math,String,Number,Object,Array,RegExp,Error,parseInt,parseFloat,isNaN,
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>k in props?props[k]:null,setProperty:(k,v)=>{props[k]=String(v)},deleteProperty:k=>{delete props[k]}})},
    CacheService:{getScriptCache:()=>({get:k=>k in cache?cache[k]:null,put:(k,v)=>{cache[k]=v}})},
    LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})},
    ContentService:{MimeType:{JSON:'json'},createTextOutput:t=>({t,setMimeType(){return this},getContent(){return t}})},
    SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss},
    Utilities:{formatDate:d=>new Date(d).toISOString(),getUuid:()=>'u'+Math.random().toString(36).slice(2)+Math.random().toString(36).slice(2),base64Decode:s=>Buffer.from(s,'base64'),newBlob:(b,m,n)=>({b,m,n,getContentType:()=>m,getBytes:()=>b})},
    DriveApp:{createFolder:n=>mkFolder(n),getFolderById:id=>drive.files[id],getFoldersByName:()=>({hasNext:()=>false}),getRootFolder:()=>root,Access:{ANYONE:1,DOMAIN:2,PRIVATE:3},Permission:{VIEW:1,EDIT:2}},
    ScriptApp:{newTrigger:()=>({timeBased(){return this},everyWeeks(){return this},onWeekDay(){return this},atHour(){return this},everyDays(){return this},create(){}}),WeekDay:{SUNDAY:0}},
    MailApp:{sendEmail(){}},UrlFetchApp:{fetch(){throw new Error('no net')}},DocumentApp:{},MimeType:{PLAIN_TEXT:'text/plain'}};
  function mkFolder(n){const id='F'+(++drive.seq);const kids=[];const it=a=>{let i=0;return{hasNext:()=>i<a.length,next:()=>a[i++]}};
 const f={id,n,getId:()=>id,getName:()=>n,getUrl:()=>'https://drive.google.com/drive/folders/'+id,setSharing(){return this},
  createFolder:x=>{const c=mkFolder(x);kids.push(c);c.kind='folder';return c},
  getFoldersByName:x=>it(kids.filter(k=>k.kind==='folder'&&k.n===x)),
  getFilesByName:x=>it(kids.filter(k=>k.kind==='file'&&k.nm===x)),
  createFile:(a,b,c)=>{const fid='D'+(++drive.seq);const isBlob=typeof a==='object';const nm=isBlob?a.n:a,data=isBlob?a.b:b;
   const fl={kind:'file',nm,n:nm,getId:()=>fid,getName:()=>nm,getUrl:()=>'https://drive.google.com/file/d/'+fid+'/view',getSize:()=>(data&&data.length)||0,
    getBlob:()=>({getDataAsString:()=>String(data),getContentType:()=>isBlob?a.m:c}),getParents:()=>it([f]),setName(x){this.nm=x;this.n=x;return this},moveTo(){return this},setTrashed(t){fl.trashed=t;if(t){const i=kids.indexOf(fl);if(i>=0)kids.splice(i,1)}return this},setSharing(){return this},getMimeType:()=>isBlob?a.m:c};
   drive.files[fid]=fl;kids.push(fl);return fl}};
 drive.files[id]=f;return f}
 const root=mkFolder('root');root.kind='folder';vm.createContext(ctx);
  ['Code.gs','Setup.gs'].forEach(f=>vm.runInContext(fs.readFileSync(dir+'/gas/'+f,'utf8'),ctx,{filename:f}));
  return {ctx,props,sheets,post:async b=>{const r=ctx.doPost({postData:{contents:JSON.stringify(b)}});return JSON.parse(r.getContent())},get:()=>JSON.parse(ctx.doGet().getContent())};
}
module.exports=makeGas;
