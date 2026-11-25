const json=(data,status=200)=>new Response(JSON.stringify(data),{status,headers:{"content-type":"application/json; charset=utf-8","cache-control":"no-store"}});
async function getIdentity(request,env){
 const h=request.headers.get("authorization")||"", token=h.startsWith("Bearer ")?h.slice(7).trim():"";
 if(!token)throw new Error("Token Portal tidak ditemukan.");
 const r=await fetch(env.GAS_WEB_APP_URL,{method:"POST",headers:{"content-type":"text/plain;charset=utf-8"},body:JSON.stringify({action:"getVideoStudioIdentity",args:[token]})});
 const data=await r.json();
 if(!r.ok||!data.ok||!data.data)throw new Error(data.error||"Sesi Portal tidak valid.");
 return data.data;
}
async function createJob(request,env,u){
 if(!env.FAL_KEY)throw new Error("Secret FAL_KEY belum diatur.");
 if(!env.FAL_MODEL_ID)throw new Error("FAL_MODEL_ID belum dipilih.");
 const b=await request.json(), prompt=String(b.prompt||"").trim(), duration=Number(b.duration||5);
 if(prompt.length<8||prompt.length>4000)throw new Error("Prompt harus 8–4000 karakter.");
 if(![5,10].includes(duration))throw new Error("Durasi uji hanya 5 atau 10 detik.");
 const ratio=["16:9","9:16","1:1"].includes(b.aspect_ratio)?b.aspect_ratio:"16:9";
 const r=await fetch("https://queue.fal.run/"+env.FAL_MODEL_ID,{method:"POST",headers:{"authorization":"Key "+env.FAL_KEY,"content-type":"application/json"},body:JSON.stringify({prompt,duration:String(duration),aspect_ratio:ratio,resolution:"720p",enable_safety_checker:true})});
 const fal=await r.json().catch(()=>({}));
 if(!r.ok)throw new Error(fal.detail||fal.error||"Permintaan fal.ai gagal.");
 const id=crypto.randomUUID(), org=u.organizationId||"pkpps-hidayatullah-karimun";
 await env.DB.prepare(`INSERT INTO video_jobs
 (id,organization_id,owner_portal_user_id,provider,model_id,provider_request_id,status,prompt,duration_seconds,aspect_ratio,provider_status_url,provider_response_url)
 VALUES(?,?,?,'fal',?,?,'queued',?,?,?,?,?)`)
 .bind(id,org,u.id,env.FAL_MODEL_ID,fal.request_id||null,prompt,duration,ratio,fal.status_url||null,fal.response_url||null).run();
 return {id,status:"queued",provider:"fal",providerRequestId:fal.request_id||null};
}
async function readJob(id,u,env){
 const org=u.organizationId||"pkpps-hidayatullah-karimun";
 const row=await env.DB.prepare("SELECT * FROM video_jobs WHERE id=? AND organization_id=? AND owner_portal_user_id=?").bind(id,org,u.id).first();
 if(!row)return null;
 if(!row.provider_status_url||!row.provider_request_id)return row;
 const r=await fetch(row.provider_status_url,{headers:{authorization:"Key "+env.FAL_KEY}});
 const st=await r.json().catch(()=>({}));
 if(!r.ok)throw new Error("Status fal.ai gagal dibaca.");
 let status=row.status;
 if(st.status==="IN_QUEUE")status="queued";
 if(st.status==="IN_PROGRESS")status="processing";
 if(st.status==="COMPLETED")status="completed";
 if(st.status==="FAILED")status="failed";
 let output=row.output_url||null, err=st.error||null;
 if(status==="completed"&&!output&&row.provider_response_url){
  const rr=await fetch(row.provider_response_url,{headers:{authorization:"Key "+env.FAL_KEY}});
  const result=await rr.json().catch(()=>({}));
  output=result.video?.url||result.video_url||null;
 }
 await env.DB.prepare("UPDATE video_jobs SET status=?,output_url=?,error_message=?,updated_at=CURRENT_TIMESTAMP WHERE id=?").bind(status,output,err,id).run();
 return {id,status,outputUrl:output,errorMessage:err};
}
export default {async fetch(request,env){
 const url=new URL(request.url);
 try{
  if(url.pathname==="/health"&&request.method==="GET")return json({ok:true,service:"pkpps-ai-video-gateway"});
  if(url.pathname==="/api/video/jobs"&&request.method==="POST"){
   const u=await getIdentity(request,env), job=await createJob(request,env,u);
   return json({ok:true,data:job},202);
  }
  const m=url.pathname.match(/^\\/api\\/video\\/jobs\\/([a-f0-9-]+)$/i);
  if(m&&request.method==="GET"){
   const u=await getIdentity(request,env), job=await readJob(m[1],u,env);
   return job?json({ok:true,data:job}):json({ok:false,error:"Pekerjaan tidak ditemukan."},404);
  }
  return json({ok:false,error:"Endpoint tidak ditemukan."},404);
 }catch(e){return json({ok:false,error:e.message||String(e)},400);}
}};
