const L=require('./lib');let pass=0,fail=0;const ok=(c,m)=>{c?pass++:fail++;console.log((c?'ok   ':'FAIL ')+m)};
(async()=>{const G=L.newGas(),A=await L.boot(G),{$,$$,sleep,click,set,btn,w}=A;
 const snap=()=>JSON.parse(G.ctx.doPost({postData:{contents:'{"action":"snapshot"}'}}).getContent());
 click($$('#tabs .tab').find(t=>t.textContent.includes('มาตรฐาน')));await sleep(150);
 click($$('#v-std button').find(b=>/ยังไม่ประเมิน/.test(b.textContent)&&b.closest('.card,.std,li,div')&&!b.classList.contains('fchip')&&b.textContent.length>14)||$$('#v-std .fchip')[1]);await sleep(100);
 console.log('   std sheet?',!!$('#layer .sheet'),'| buttons:',$$('#layer button, #v-std .card button').slice(0,10).map(b=>b.textContent.trim().slice(0,12)).join('|'));
 // ประเด็น: แก้สถานะ/ปิด + ค้นหา/กรอง
 click($$('#tabs .tab').find(t=>t.textContent.includes('ประเด็น')));click($('#tabs .tabadd'));await sleep(100);let sh=$('#layer .sheet');
 set($('textarea',sh),'ประเด็น X');const nm=$$('input[type=text]',sh).find(i=>/ชื่อของคุณ/.test(i.placeholder||''));if(nm)set(nm,'T');click(btn(sh,/^บันทึก/));await sleep(400);
 ok(snap().issues.length===1,'สร้างประเด็น 1');
 click($$('#tabs .tab').find(t=>t.textContent.includes('ประเด็น')));await sleep(100);
 const card=$$('#v-iss button').find(b=>/แก้ไข/.test(b.textContent));ok(!!card,'ประเด็นมีปุ่มแก้ไข');
 if(card){click(card);await sleep(100);sh=$('#layer .sheet');const st=$$('select',sh).find(s=>[...s.options].some(o=>/^ปิด/.test(o.text)));if(st){set(st,[...st.options].find(o=>/^ปิด/.test(o.text)).value||[...st.options].find(o=>/^ปิด/.test(o.text)).text);click(btn(sh,/^บันทึก/));await sleep(400);ok(snap().issues[0].status.startsWith('ปิด'),'แก้สถานะประเด็นเป็น "ปิด" -> ลงชีต');}else ok(false,'ไม่พบช่องสถานะ')}
 console.log('errors',A.errs.length,A.errs.slice(0,3));console.log(pass+' pass, '+fail+' fail');process.exit(fail?1:0)})().catch(e=>{console.log('CRASH',e);process.exit(2)});
