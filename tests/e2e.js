const {JSDOM}=require('jsdom'),fs=require('fs'),mk=require('./gasmock');
const ROOT=require('path').resolve(__dirname,'..'),G=mk(ROOT);
process.env.GAS_URL='https://script.google.com/macros/s/TEST/exec';delete process.env.GAS_SECRET;
const proxy=require(ROOT+'/api/gas.js');
const calls=[];
async function apiGas(init){ // เลียนแบบ Vercel: req/res
  const body=JSON.parse(init.body);calls.push(body.action);
  let out;const res={_h:{},setHeader(){},status(c){this.c=c;return this},json(j){out={code:this.c||200,j}}};
  global.fetch=async(u,o)=>{ if(u!==process.env.GAS_URL)throw new Error('unexpected '+u); const sent=JSON.parse(o.body); if('secret' in sent) global.__secretSent=true; const t=await G.post(sent); return {text:async()=>JSON.stringify(t)}};
  await proxy({method:'POST',body},res);
  return {ok:out.code<400,status:out.code,json:async()=>out.j};
}
let errs=[];
(async()=>{
 const html=fs.readFileSync(ROOT+'/public/index.html','utf8').replace(/<script>[\s\S]*<\/script>/,'');
 const dom=new JSDOM(html,{url:'http://localhost/',runScripts:'outside-only',pretendToBeVisual:true});
 const w=dom.window;w.DATA=JSON.parse(fs.readFileSync(ROOT+'/public/data/plan.json','utf8'));
 w.fetch=(u,i)=>u==='/api/gas'?apiGas(i):Promise.reject(new Error('net '+u));
 w.matchMedia=w.matchMedia||(()=>({matches:false,addEventListener(){},addListener(){}}));
 w.scrollTo=()=>{};w.HTMLElement.prototype.scrollIntoView=function(){};
 w.addEventListener('error',e=>errs.push('onerror '+e.message));
 const ce=w.console.error;w.console.error=(...a)=>{errs.push('console.error '+a.map(String).join(' ').slice(0,200))};
 w.eval(fs.readFileSync(ROOT+'/public/js/app.js','utf8'));
 const $=(s,r)=>(r||w.document).querySelector(s),$$=(s,r)=>[...(r||w.document).querySelectorAll(s)];
 const sleep=ms=>new Promise(r=>setTimeout(r,ms));
 const click=e=>e.dispatchEvent(new w.MouseEvent('click',{bubbles:true}));
 const set=(e,v)=>{e.value=v;e.dispatchEvent(new w.Event('input',{bubbles:true}));e.dispatchEvent(new w.Event('change',{bubbles:true}))};
 let pass=0,fail=0;const ok=(c,m)=>{c?pass++:fail++;console.log((c?'ok   ':'FAIL ')+m)};
 await sleep(600);
 ok($('#mode').textContent.includes('ซิงก์แล้ว'),'เชื่อมหลังบ้านสำเร็จ: '+$('#mode').textContent);
 ok(calls.includes('snapshot'),'โหลด snapshot ผ่าน proxy โดยไม่มี GAS_SECRET; secret ไม่ถูกส่ง='+!global.__secretSent);
 // ปฏิทิน: กดวันที่
 const days=$$('#v-cal .day:not(.empty)');ok(days.length>=28,'ปฏิทินแสดง '+days.length+' วัน');
 click(days[14]);await sleep(100);
 let sheet=$('#layer .sheet');ok(!!sheet,'กดวันที่ -> เปิดหน้าต่างกำหนดการ');
 const addTask=async(kindIdx,title,tm)=>{
   const chips=$$('.fchip',sheet);click(chips[kindIdx]);
   const inputs=$$('input[type=text]',sheet);set(inputs[0],title);
   if(tm){const ts=$$('input[type=time]',sheet);set(ts[0],tm)}
   const nm=$$('input[type=text]',sheet).find(i=>/ชื่อของคุณ/.test(i.getAttribute('aria-label')||''));if(nm)set(nm,'ผู้ทดสอบ');
   const sb=$$('button',sheet).find(b=>/เพิ่มกำหนดการ$/.test(b.textContent)&&b.className==='btn');click(sb);await sleep(250)};
 await addTask(0,'อบรมมาตรฐาน','09:00');await addTask(1,'ประชุม PCT','13:00');await addTask(2,'ส่งเอกสาร');
 let plus=Object.keys(G.ctx&&JSON.parse(G.ctx.doPost({postData:{contents:JSON.stringify({action:'snapshot'})}}).getContent()).hait.plus);
 ok(plus.filter(k=>k.startsWith('task:')).length===3,'บันทึก 3 ประเภท (กิจกรรม/ประชุม/งาน) ลงชีต: '+plus.length);
 ok($$('.taskrow',sheet).length===3,'แสดงรายการในหน้าต่างครบ 3');
 // ติ๊กเสร็จ
 const cb=$$('.taskrow input[type=checkbox]',sheet)[0];cb.checked=true;cb.dispatchEvent(new w.Event('change',{bubbles:true}));await sleep(250);
 const snap=()=>JSON.parse(G.ctx.doPost({postData:{contents:JSON.stringify({action:'snapshot'})}}).getContent());
 ok(Object.values(snap().hait.plus).some(v=>v.done===1),'ติ๊กเสร็จ -> บันทึกลงชีต');
 // แก้ไข
 const eb=$$('.taskrow button',sheet).find(b=>b.textContent==='แก้ไข');click(eb);await sleep(50);
 const t1=$$('input[type=text]',sheet)[0];set(t1,'ชื่อที่แก้แล้ว');click($$('button',sheet).find(b=>/บันทึกการแก้ไข/.test(b.textContent)));await sleep(250);
 ok(Object.values(snap().hait.plus).some(v=>v.t==='ชื่อที่แก้แล้ว'),'แก้ไขรายการ -> บันทึก');
 // ลบ (กด 2 ครั้ง)
 const db=$$('.taskrow button',sheet).find(b=>b.textContent==='ลบ');click(db);click($$('.taskrow button',sheet).find(b=>/กดอีกครั้ง/.test(b.textContent)));await sleep(250);
 ok(Object.values(snap().hait.plus).filter(v=>v.del).length===1,'ลบรายการ (soft delete) -> บันทึก');
 click($('.hd .iconbtn,.hd button',sheet)||$('[aria-label="ปิด"]'));await sleep(50);
 // เลื่อนเดือน
 const nxt=$('[aria-label="เดือนถัดไป"]'),prv=$('[aria-label="เดือนก่อน"]');click(nxt);await sleep(60);ok(!/ตุลาคม/.test($('#v-cal h2').textContent),'เลื่อนไปเดือนถัดไป: '+$('#v-cal h2').textContent);
 // ทุกแท็บ
 for(const t of $$('#tabs .tab:not(.tabadd)')){click(t);await sleep(120);const vis=$$('main > section').filter(s=>!s.hidden);ok(vis.length===1&&vis[0].children.length>0,'แท็บ "'+t.textContent.trim().slice(0,12)+'" แสดงผลได้');}
 // บันทึกประเด็น
 click($('#tabs .tabadd'));await sleep(100);sheet=$('#layer .sheet');ok(!!sheet,'เปิดฟอร์มบันทึกประเด็น');
 if(sheet){const ta=$('textarea',sheet);set(ta,'ประเด็นทดสอบระบบ');const nm=$$('input[type=text]',sheet).find(i=>/ชื่อของคุณ/.test(i.placeholder||''));if(nm)set(nm,'ผู้ทดสอบ');
  const sv=$$('button',sheet).find(b=>/^บันทึก/.test(b.textContent.trim())&&!/ยกเลิก/.test(b.textContent));click(sv);await sleep(400);
  ok(snap().issues.some(i=>i.text==='ประเด็นทดสอบระบบ'),'บันทึกประเด็น -> ลงชีต (issues='+snap().issues.length+')');}
 console.log('\nerrors captured:',errs.length);errs.slice(0,8).forEach(e=>console.log('  ',e));
 console.log('API actions:',[...new Set(calls)].join(','));
 console.log(pass+' pass, '+fail+' fail');process.exit(fail?1:0);
})().catch(e=>{console.log('CRASH',e);process.exit(2)});
