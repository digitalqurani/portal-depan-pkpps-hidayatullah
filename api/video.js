const WORKER_URL = process.env.AI_VIDEO_WORKER_URL;
module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');
  if (!WORKER_URL || !/^https:\/\/[a-z0-9.-]+\/?$/i.test(WORKER_URL))
    return res.status(500).json({ok:false,error:'Environment AI_VIDEO_WORKER_URL belum diatur.'});
  const authorization=req.headers.authorization||'';
  if(!/^Bearer\s+\S+$/i.test(authorization))
    return res.status(401).json({ok:false,error:'Sesi Portal diperlukan.'});
  try {
    let target, method=req.method, headers={authorization};
    if(method==='POST'){
      const body=typeof req.body==='string'?JSON.parse(req.body):(req.body||{});
      if(body.action!=='create')return res.status(400).json({ok:false,error:'Aksi tidak dikenal.'});
      target=new URL('/api/video/jobs',WORKER_URL);
      headers['content-type']='application/json';
      const upstream=await fetch(target,{method:'POST',headers,body:JSON.stringify(body)});
      const raw=await upstream.text(); let data;
      try{data=JSON.parse(raw)}catch(_){throw new Error('Respons Worker bukan JSON.')}
      return res.status(upstream.status).json(data);
    }
    if(method==='GET'){
      const jobId=String(req.query.jobId||'');
      if(!/^[a-f0-9-]{20,}$/i.test(jobId))return res.status(400).json({ok:false,error:'ID pekerjaan tidak valid.'});
      target=new URL('/api/video/jobs/'+encodeURIComponent(jobId),WORKER_URL);
      const upstream=await fetch(target,{method:'GET',headers});
      const raw=await upstream.text(); let data;
      try{data=JSON.parse(raw)}catch(_){throw new Error('Respons Worker bukan JSON.')}
      return res.status(upstream.status).json(data);
    }
    res.setHeader('Allow','GET, POST');
    return res.status(405).json({ok:false,error:'Method not allowed'});
  }catch(err){return res.status(502).json({ok:false,error:err.message||String(err)});}
};
module.exports.config={api:{bodyParser:{sizeLimit:'64kb'}}};
