const L=require('./lib');
let pass=0,fail=0;const ok=(c,m)=>{c?pass++:fail++;console.log((c?'ok   ':'FAIL ')+m)};
(async()=>{
 const G=L.newGas(),A=await L.boot(G),{$,$$,sleep,click,set,btn,w}=A;
 const snap=()=>JSON.parse(G.ctx.doPost({postData:{contents:'{"action":"snapshot"}'}}).getContent());
 const tab=async n=>{click($$('#tabs .tab').find(t=>t.textContent.includes(n)));await sleep(150)};
 // --- QMR: เปลี่ยนวัน ---
 await tab('ประชุม');click(btn(document=A.w.document,/กำหนดวันประชุม/));await sleep(100);
 let sh=$('#layer .sheet');ok(!!sh,'เปิดหน้า "กำหนดวันประชุม QMR"');
 click($('.ev.qmr',sh));await sleep(100);sh=$('#layer .sheet');
 const di=$$('input[type=date]',sh)[0];ok(!!di,'ฟอร์มแก้ QMR มีช่องวันที่');
 if(di){set(di,'2026-11-20');const nm=$$('input[type=text]',sh).find(i=>/ชื่อของคุณ/.test(i.placeholder||''));if(nm)set(nm,'ผู้ทดสอบ');
  const sv=btn(sh,/^บันทึก/);click(sv);await sleep(400);
  const q=Object.entries(snap().hait.plus).filter(([k])=>k.startsWith('qmr:'));ok(q.length>=1,'เลื่อนวัน QMR -> บันทึกคีย์ qmr: ลงชีต ('+q.length+') '+(q[0]?JSON.stringify(q[0][1]).slice(0,80):''));}
 // --- บันทึกการประชุม ---
 await tab('ประชุม');click(btn(w.document,/\+ บันทึกการประชุม/));await sleep(150);sh=$('#layer .sheet');ok(!!sh,'เปิดฟอร์ม "บันทึกการประชุม"');
 if(sh){const tas=$$('textarea',sh);tas.forEach((t,i)=>set(t,i===0?'สรุปทดสอบ':'ทดสอบ'));const nm=$$('input[type=text]',sh).find(i=>/ชื่อของคุณ/.test(i.placeholder||''));if(nm)set(nm,'ผู้ทดสอบ');
  const d=$$('input[type=date]',sh)[0];if(d&&!d.value)set(d,'2026-10-29');
  const sv=btn(sh,/^บันทึก/);click(sv);await sleep(500);
  const m=Object.keys(snap().hait.plus).filter(k=>k.startsWith('mtg:'));ok(m.length===1,'บันทึกประชุม -> คีย์ mtg: ลงชีต ('+m.length+') errs='+A.errs.length);}
 // --- มาตรฐาน: ให้คะแนน ---
 await tab('มาตรฐาน');const scoreBtn=$$('#v-std button').find(b=>/ประเมิน|คะแนน|ให้คะแนน/.test(b.textContent)&&!/พิมพ์|ส่งออก/.test(b.textContent));
 console.log('   std buttons:',$$('#v-std button').slice(0,12).map(b=>b.textContent.trim().slice(0,14)).join(' | '));
 // --- ซิงก์ข้ามเครื่อง ---
 const B=await L.boot(G,{poll:300});
 ok(B.$$('#v-mtg, #v-cal').length>0&&B.$('#mode').textContent.includes('ซิงก์แล้ว'),'เครื่อง B เปิดแล้วเชื่อมได้');
 const ver0=B.w.document.querySelector('#mode').textContent;
 // A เพิ่มงาน -> B ต้องเห็นหลัง poll
 await tab('ปฏิทิน');const days=$$('#v-cal .day:not(.empty)');click(days[20]);await sleep(100);sh=$('#layer .sheet');
 set($$('input[type=text]',sh)[0],'งานจากเครื่อง A');const nm2=$$('input[type=text]',sh).find(i=>/ชื่อของคุณ/.test(i.placeholder||''));if(nm2)set(nm2,'A');click(btn(sh,/เพิ่มกำหนดการ$/));await sleep(300);
 await B.sleep(900);
 const bDays=B.$$('#v-cal .day .tkn');ok(bDays.length>=1,'เครื่อง B เห็นงานของเครื่อง A อัตโนมัติ (ป้ายรายการ '+bDays.length+')');
 // รีโหลดหน้า: ข้อมูลยังอยู่
 const C=await L.boot(G);click(C.$$('#v-cal .day:not(.empty)')[20]);await C.sleep(100);
 ok(/งานจากเครื่อง A/.test(C.$('#layer').textContent),'เปิดใหม่ (รีโหลด) ข้อมูลยังอยู่จากชีต');
 console.log('errors A/B/C:',A.errs.length,B.errs.length,C.errs.length,[...A.errs,...B.errs,...C.errs].slice(0,5));
 console.log(pass+' pass, '+fail+' fail');process.exit(fail?1:0);
})().catch(e=>{console.log('CRASH',e);process.exit(2)});
