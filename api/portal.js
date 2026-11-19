const GAS_URL = process.env.GAS_WEB_APP_URL;
module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (req.method !== 'POST') { res.setHeader('Allow','POST'); return res.status(405).json({ok:false,error:'Method not allowed'}); }
  if (!GAS_URL || !/^https:\/\/script\.google\.com\/macros\/s\//.test(GAS_URL)) return res.status(500).json({ok:false,error:'Environment GAS_WEB_APP_URL belum diatur dengan URL Web App Apps Script yang benar.'});
  try {
    const payload = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
    if (!payload.action || !Array.isArray(payload.args)) return res.status(400).json({ok:false,error:'Format permintaan tidak valid.'});
    const upstream = await fetch(GAS_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload),redirect:'follow'});
    const raw = await upstream.text(); let data;
    try { data=JSON.parse(raw); } catch (_) { throw new Error('Respons Apps Script bukan JSON. Pastikan Web App dapat diakses oleh pengguna portal.'); }
    return res.status(upstream.ok?200:upstream.status).json(data);
  } catch(err) { return res.status(502).json({ok:false,error:err&&err.message?err.message:String(err)}); }
};
module.exports.config={api:{bodyParser:{sizeLimit:'3mb'}}};
