const mk=require('./gasmock'),ROOT=require('path').resolve(__dirname,'..');
let pass=0,fail=0;const ok=(c,m)=>{c?pass++:fail++;console.log((c?'ok   ':'FAIL ')+m)};
(async()=>{
 const G=mk(ROOT);
 const data=Buffer.from('hello hospital file').toString('base64'),half=Math.ceil(data.length/2/3)*3;
 const parts=[data.slice(0,half),data.slice(half)].filter(Boolean);
 let r;for(let i=0;i<parts.length;i++){r=await G.post({action:'fileChunk',uid:'abcdef12',idx:i,total:parts.length,name:'ทดสอบ.txt',size:19,mime:'text/plain',data:parts[i],cat:'HAIT',by:'t'});}
 ok(r.ok&&r.files&&r.files.length===1,'อัปโหลดไฟล์ 2 ชิ้น -> Files มี 1 รายการ '+(r.error||''));
 const fid=r.files&&r.files[0]&&r.files[0].id;
 r=await G.post({action:'fileUpdate',id:fid,patch:{name:'ชื่อใหม่.txt',note:'x',cat:'HA'},by:'t'});ok(r.ok&&r.files[0].name==='ชื่อใหม่.txt','แก้ชื่อ/หมวดไฟล์ '+(r.error||''));
 r=await G.post({action:'fileDelete',id:fid,by:'t'});ok(r.ok&&r.files.length===0,'ลบไฟล์ '+(r.error||''));
 r=await G.post({action:'fileChunk',uid:'abcdef13',idx:0,total:1,name:'a.exe',size:5,data:'AAAA'});ok(r.error==='badtype','ปฏิเสธไฟล์ .exe');
 r=await G.post({action:'fileChunk',uid:'abcdef14',idx:0,total:1,name:'big.pdf',size:99999999,data:'AAAA'});ok(r.error==='toolarge','ปฏิเสธไฟล์ใหญ่เกิน 20MB');
 r=await G.post({action:'memoCreate',memo:{type:'general',subject:'',to:'',body:''}});ok(!r.ok,'memoCreate ข้อมูลไม่ครบ -> ปฏิเสธ ('+r.error+')');
 // proxy
 process.env.GAS_URL='https://x/exec';const proxy=require(ROOT+'/api/gas.js');
 const run=async(req,fetchImpl)=>{let o;const res={setHeader(){},status(c){this.c=c;return this},json(j){o={c:this.c||200,j}}};global.fetch=fetchImpl||(async()=>({text:async()=>'{"ok":true}'}));await proxy(req,res);return o};
 let o=await run({method:'GET'});ok(o.c===405,'proxy GET -> 405');
 o=await run({method:'POST',body:{action:'evil'}});ok(o.c===400,'proxy action ไม่อนุญาต -> 400');
 o=await run({method:'POST',body:'{bad'});ok(o.c===400,'proxy JSON เสีย -> 400');
 o=await run({method:'POST',body:{action:'put',x:'a'.repeat(250000)}});ok(o.c===413,'proxy body ใหญ่ -> 413');
 o=await run({method:'POST',body:{action:'ver'}},async()=>({text:async()=>'<html>login</html>'}));ok(o.c===502&&o.j.error==='notjson','proxy GAS ตอบ HTML (ยังไม่เปิดสิทธิ์ Anyone) -> 502 notjson');
 o=await run({method:'POST',body:{action:'ver'}},async()=>{throw new Error('x')});ok(o.c===502&&o.j.error==='upstream','proxy เชื่อม GAS ไม่ได้ -> 502 upstream');
 delete process.env.GAS_URL;o=await run({method:'POST',body:{action:'ver'}});ok(o.c===500&&o.j.error==='server_config','ไม่มี GAS_URL -> 500 server_config');
 console.log(pass+' pass, '+fail+' fail');process.exit(fail?1:0);
})().catch(e=>{console.log('CRASH',e);process.exit(2)});
