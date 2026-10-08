// Vercel Serverless Function: proxy บาง ๆ ไปยัง Google Apps Script
// ซ่อน GAS_URL และ GAS_SECRET ไว้ฝั่งเซิร์ฟเวอร์ ไม่ให้อยู่ใน JS ของหน้าเว็บ
const ALLOWED = new Set(['ver', 'snapshot', 'put', 'patch', 'delete', 'dates', 'hait', 'fileChunk', 'fileUpdate', 'fileDelete', 'memoCreate', 'memoGet', 'memoDelete']);
// ชิ้นไฟล์ (base64) ใหญ่กว่าคำสั่งทั่วไป แต่ต้องไม่เกินเพดาน 4.5 MB ของ Vercel
const MAX_BODY = 200000, MAX_CHUNK = 4000000;

module.exports = async (req, res) => {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') return res.status(405).json({ ok: false, error: 'method' });

  const { GAS_URL, GAS_SECRET } = process.env;
  if (!GAS_URL || !GAS_SECRET) return res.status(500).json({ ok: false, error: 'server_config' });

  let body = req.body;
  if (typeof body === 'string') { try { body = JSON.parse(body); } catch (e) { body = null; } }
  if (!body || typeof body !== 'object' || !ALLOWED.has(body.action)) {
    return res.status(400).json({ ok: false, error: 'bad_request' });
  }
  if (JSON.stringify(body).length > (body.action === 'fileChunk' ? MAX_CHUNK : MAX_BODY)) return res.status(413).json({ ok: false, error: 'too_large' });

  try {
    const r = await fetch(GAS_URL, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({}, body, { secret: GAS_SECRET })),
    });
    const text = await r.text();
    let j;
    try { j = JSON.parse(text); } catch (e) { return res.status(502).json({ ok: false, error: 'notjson' }); }
    return res.status(200).json(j);
  } catch (e) {
    return res.status(502).json({ ok: false, error: 'upstream' });
  }
};
