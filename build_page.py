import base64, os

logo = open('logo.jpg', 'rb').read()
b64 = 'data:image/jpeg;base64,' + base64.b64encode(logo).decode()

tsx = r'''
"use client";
import React, { useState, useRef, useEffect, useCallback } from "react";

const API_BASE = "http://localhost:8001/api";
const LOGO = "''' + b64 + r'''";

type Stage = "idle"|"uploading"|"processing"|"downloading"|"done"|"error";
interface Pdf { id:string; filename:string; stockNo:string; productName:string; url:string; status:string; }

export default function App() {
  const [stage, setStage] = useState<Stage>("idle");
  const [drag, setDrag]   = useState(false);
  const [file, setFile]   = useState<File|null>(null);
  const [jobId, setJobId] = useState<string|null>(null);
  const [pct, setPct]     = useState(0);
  const [total, setTotal] = useState(0);
  const [done, setDone]   = useState(0);
  const [pdfs, setPdfs]   = useState<Pdf[]>([]);
  const [err, setErr]     = useState("");
  const [lbl, setLbl]     = useState("");
  const ref  = useRef<HTMLInputElement>(null);
  const poll = useRef<any>(null);

  const dl = useCallback((href:string, name:string, delay=0) =>
    setTimeout(()=>{ const a=document.createElement("a"); a.href=href; a.download=name; document.body.appendChild(a); a.click(); document.body.removeChild(a); }, delay)
  , []);

  const startPoll = useCallback((id:string)=>{
    poll.current = setInterval(async ()=>{
      try {
        const r = await fetch(`${API_BASE}/job/${id}`);
        if(!r.ok) return;
        const j = await r.json();
        setPct(j.progress??0); setDone(j.successCount??0); setTotal(j.totalRows??0);
        const p = j.progress??0;
        if(p<30) setLbl("Parsing Excel columns and data rows...");
        else if(p<60) setLbl("Compiling HTML templates and resolving images...");
        else if(p<90) setLbl("Rendering PDFs via headless browser engine...");
        else setLbl("Finalising and packaging output files...");
        if(j.status==="COMPLETED"||j.status==="FAILED"){
          clearInterval(poll.current); poll.current=null;
          if(j.status==="FAILED"){ setErr("PDF generation failed. Check the Excel format."); setStage("error"); return; }
          const ok:Pdf[] = (j.details??[]).filter((d:Pdf)=>d.status==="COMPLETED");
          setPdfs(ok);
          setStage("downloading");
          setLbl(`Downloading ${ok.length} PDF${ok.length!==1?"s":""}…`);
          if(ok.length===1) dl(`${API_BASE}/pdf/download/${encodeURIComponent(ok[0].filename)}`, ok[0].filename, 400);
          else if(ok.length>1) dl(`${API_BASE}/job/${id}/zip`, `datasheets_${id}.zip`, 400);
          setTimeout(()=>{ setStage("done"); setLbl(""); }, 1800);
        }
      } catch{}
    }, 1500);
  },[dl]);

  const run = useCallback(async (f:File) => {
    if(poll.current) clearInterval(poll.current);
    setStage("uploading"); setPct(0); setPdfs([]); setErr(""); setLbl("Uploading spreadsheet to server…");
    const fd = new FormData(); fd.append("file", f);
    try {
      const r = await fetch(`${API_BASE}/job/upload`,{method:"POST",body:fd});
      if(!r.ok){ const e=await r.json().catch(()=>({})); throw new Error((e as any).message||"Upload failed"); }
      const j = await r.json();
      setJobId(j.id); setTotal(j.totalRows??0);
      setStage("processing"); setLbl("Parsing Excel columns and data rows…");
      startPoll(j.id);
    } catch(e:any){ setErr(e.message||"Connection failed."); setStage("error"); }
  }, [startPoll]);

  useEffect(()=>{ return ()=>{ if(poll.current) clearInterval(poll.current); }; }, []);
  useEffect(()=>{ if(file) run(file); }, [file, run]);

  const validate = (f:File) => {
    if(!/\.(xlsx|xls)$/i.test(f.name)){ setErr("Only .xlsx / .xls files accepted."); setStage("error"); return; }
    setFile(f);
  };

  const reset = () => {
    if(poll.current) clearInterval(poll.current);
    setStage("idle"); setFile(null); setJobId(null); setPct(0); setPdfs([]); setErr(""); setLbl("");
  };

  const active = stage==="uploading"||stage==="processing"||stage==="downloading";

  return (
    <div className="root">
      {/* NAV */}
      <nav className="nav">
        <div className="nav-l">
          <img src={LOGO} alt="RS PRO" className="nav-logo" />
        </div>
      </nav>


      {/* CARD */}
      <main className="main">
        <div className="card">

          {/* IDLE */}
          {stage==="idle" && (
            <div>
              <div className={`dz ${drag?"dz-on":""}`}
                onDragOver={e=>{e.preventDefault();setDrag(true);}}
                onDragLeave={()=>setDrag(false)}
                onDrop={e=>{e.preventDefault();setDrag(false);const f=e.dataTransfer.files?.[0];if(f)validate(f);}}
                onClick={()=>ref.current?.click()}>
                <span className="material-symbols-outlined dz-ic">{drag?"file_open":"upload_file"}</span>
                <p className="dz-h">{drag?"Release to start generating":"Drag & drop your Excel file here"}</p>
                <p className="dz-s">Supports .xlsx and .xls &nbsp;·&nbsp; Click anywhere to browse</p>
                <div className="dz-btn">
                  <span className="material-symbols-outlined" style={{fontSize:18,marginRight:7,verticalAlign:"middle"}}>upload</span>
                  Choose File
                </div>
                <input ref={ref} type="file" accept=".xlsx,.xls" onChange={e=>{const f=e.target.files?.[0];if(f)validate(f);e.target.value="";}} style={{display:"none"}}/>
              </div>

              <div className="steps">
                {[
                  {ic:"table_view",  h:"Upload Excel",    d:"Drop or select your XLSX product file"},
                  {ic:"auto_awesome",h:"Auto Generate",   d:"System compiles branded PDF datasheets"},
                  {ic:"download",    h:"Instant Download",d:"ZIP or single PDF saves to your device"},
                ].map((s,i)=>(
                  <div key={i} className="step">
                    <span className="material-symbols-outlined step-ic">{s.ic}</span>
                    <div className="step-h">{s.h}</div>
                    <div className="step-d">{s.d}</div>
                  </div>
                ))}
              </div>

            </div>
          )}

          {/* ACTIVE */}
          {active && (
            <div className="act">
              <div className="ring">
                <svg className="ring-svg" viewBox="0 0 96 96">
                  <circle cx="48" cy="48" r="40" fill="none" stroke="#EEF0F5" strokeWidth="7"/>
                  <circle cx="48" cy="48" r="40" fill="none" stroke="#CC0000" strokeWidth="7" strokeLinecap="round" strokeDasharray="75 177"/>
                </svg>
                <div className="ring-in">
                  {stage==="uploading"
                    ? <span className="material-symbols-outlined" style={{fontSize:26,color:"#CC0000"}}>upload</span>
                    : stage==="downloading"
                    ? <span className="material-symbols-outlined" style={{fontSize:26,color:"#CC0000"}}>download</span>
                    : <span className="ring-pct">{pct}%</span>}
                </div>
              </div>
              <h2 className="act-h">
                {stage==="uploading"?"Uploading Spreadsheet…":stage==="downloading"?"Downloading PDFs…":"Generating PDFs…"}
              </h2>
              <p className="act-s">{lbl}</p>
              {stage==="processing" && (
                <div className="pb-wrap">
                  <div className="pb"><div className="pb-fill" style={{width:`${pct}%`}}/></div>
                  <div className="pb-meta"><span>{pct}% complete</span><span>{done} / {total} datasheets</span></div>
                </div>
              )}
              {(stage==="uploading"||stage==="downloading") && (
                <div className="dots">
                  {[0,1,2].map(i=><div key={i} className="dot" style={{animationDelay:`${i*0.2}s`}}/>)}
                </div>
              )}
            </div>
          )}

          {/* DONE */}
          {stage==="done" && (
            <div>
              <div className="done-banner">
                <span className="material-symbols-outlined" style={{fontSize:46,color:"#fff",flexShrink:0}}>check_circle</span>
                <div>
                  <h2 className="done-h">{pdfs.length} PDF{pdfs.length!==1?"s":""} Generated Successfully!</h2>
                  <p className="done-s">{pdfs.length>1?"A ZIP archive has downloaded automatically.":"Your PDF has downloaded automatically."}</p>
                </div>
              </div>

              {pdfs.length>0 && (
                <div className="tbl-wrap">
                  <div className="tbl-top">
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <span className="material-symbols-outlined" style={{fontSize:18,color:"#CC0000"}}>picture_as_pdf</span>
                      <span>{pdfs.length} Generated Datasheet{pdfs.length!==1?"s":""}</span>
                    </div>
                    {pdfs.length>1 && jobId && (
                      <a href={`${API_BASE}/job/${jobId}/zip`} download className="re-btn">
                        <span className="material-symbols-outlined" style={{fontSize:15}}>download</span>
                        Re-download ZIP
                      </a>
                    )}
                  </div>
                  <table className="tbl">
                    <thead><tr><th>Stock No.</th><th>Product Name</th><th>File</th><th></th></tr></thead>
                    <tbody>
                      {pdfs.map((p,i)=>(
                        <tr key={i}>
                          <td className="td-s">{p.stockNo}</td>
                          <td className="td-n">{p.productName||"—"}</td>
                          <td className="td-f">{p.filename}</td>
                          <td className="td-a">
                            <a href={`${API_BASE}/pdf/download/${encodeURIComponent(p.filename)}`} download className="dl-a">
                              <span className="material-symbols-outlined" style={{fontSize:14}}>download</span>PDF
                            </a>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              <div style={{display:"flex",justifyContent:"center",marginTop:24}}>
                <button className="btn-pr" onClick={reset}>
                  <span className="material-symbols-outlined" style={{fontSize:18}}>upload_file</span>
                  Generate Another Batch
                </button>
              </div>
            </div>
          )}

          {/* ERROR */}
          {stage==="error" && (
            <div className="err-wrap">
              <span className="material-symbols-outlined" style={{fontSize:56,color:"#CC0000",display:"block",margin:"0 auto 14px"}}>error</span>
              <h2 className="err-h">Something went wrong</h2>
              <p className="err-s">{err||"An unexpected error occurred. Please try again."}</p>
              <button className="btn-pr" onClick={reset}>
                <span className="material-symbols-outlined" style={{fontSize:18}}>refresh</span>
                Try Again
              </button>
            </div>
          )}

        </div>
      </main>

      {/* FOOTER */}
      <footer className="foot">
        <span>© RS PRO Datasheet Generator</span>
        <span>Backend · {API_BASE.replace("/api","")}</span>
      </footer>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
        @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200');
        *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
        .material-symbols-outlined{font-family:'Material Symbols Outlined';font-style:normal;display:inline-block;line-height:1;vertical-align:middle;font-variation-settings:'FILL' 0,'wght' 400,'GRAD' 0,'opsz' 24;}
        .root{min-height:100vh;background:#F0F2F8;display:flex;flex-direction:column;font-family:'Inter',system-ui,sans-serif;color:#1A1A2E;}
        /* NAV */
        .nav{background:#ffffff;height:64px;display:flex;align-items:center;justify-content:space-between;padding:0 32px;border-bottom:2px solid #E8EBF3;box-shadow:0 1px 6px rgba(0,0,0,0.06);}
        .nav-l{display:flex;align-items:center;gap:14px;}
        .nav-logo{height:36px;object-fit:contain;}
        /* HERO */
        .hero{background:linear-gradient(135deg,#A80000 0%,#CC0000 55%,#E53935 100%);padding:44px 32px 52px;text-align:center;}
        .hero-h1{font-size:30px;font-weight:800;color:#fff;letter-spacing:-0.5px;margin-bottom:10px;}
        .hero-p{font-size:15px;color:rgba(255,255,255,0.82);max-width:560px;margin:0 auto;line-height:1.7;}
        /* MAIN */
        .main{flex:1;display:flex;justify-content:center;padding:44px 16px 52px;}
        .card{width:100%;max-width:680px;background:#fff;border-radius:18px;border:1px solid #DDE1EC;box-shadow:0 6px 32px rgba(0,0,0,0.08);padding:38px 42px;}
        /* DROP ZONE */
        .dz{border:2.5px dashed #CDD2E1;border-radius:14px;padding:56px 32px;text-align:center;cursor:pointer;transition:all 0.18s;background:#FAFBFE;}
        .dz:hover{border-color:#CC0000;background:#FFF5F5;}
        .dz-on{border-color:#CC0000!important;background:#FFF0F0!important;}
        .dz-ic{font-size:52px;color:#C5CAD9;display:block;margin:0 auto 14px;transition:color 0.18s;}
        .dz:hover .dz-ic,.dz-on .dz-ic{color:#CC0000;}
        .dz-h{font-size:17px;font-weight:700;color:#222;margin-bottom:6px;}
        .dz-on .dz-h{color:#CC0000;}
        .dz-s{font-size:13px;color:#9BA3B8;margin-bottom:22px;}
        .dz-btn{display:inline-flex;align-items:center;background:#CC0000;color:#fff;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;padding:11px 28px;border-radius:8px;letter-spacing:0.2px;}
        /* STEPS */
        .steps{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-top:24px;}
        .step{background:#FAFBFE;border:1.5px solid #E4E8F2;border-radius:11px;padding:20px 14px;text-align:center;transition:border-color 0.15s;}
        .step:hover{border-color:#CC0000;}
        .step-ic{font-size:26px;color:#CC0000;display:block;margin:0 auto 9px;}
        .step-h{font-size:13px;font-weight:700;color:#222;margin-bottom:3px;}
        .step-d{font-size:11.5px;color:#9BA3B8;line-height:1.5;}
        /* INFO */
        .info{margin-top:20px;background:#FFF8F8;border:1.5px solid #FFCECE;border-radius:9px;padding:13px 16px;display:flex;align-items:flex-start;gap:10px;font-size:12.5px;color:#555;line-height:1.6;}
        .info code{background:#FFE9E9;color:#CC0000;padding:1px 5px;border-radius:3px;font-size:11.5px;font-family:monospace;}
        /* ACTIVE */
        .act{text-align:center;padding:24px 0;}
        .ring{position:relative;width:96px;height:96px;margin:0 auto 28px;}
        .ring-svg{position:absolute;inset:0;width:96px;height:96px;animation:spin 1.1s linear infinite;}
        .ring-in{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;}
        .ring-pct{font-weight:900;font-size:17px;color:#CC0000;}
        .act-h{font-size:20px;font-weight:800;color:#1A1A2E;margin-bottom:8px;}
        .act-s{font-size:13.5px;color:#9BA3B8;margin-bottom:24px;}
        .pb-wrap{max-width:440px;margin:0 auto;}
        .pb{height:8px;background:#EEF0F8;border-radius:4px;overflow:hidden;margin-bottom:8px;}
        .pb-fill{height:100%;background:linear-gradient(90deg,#CC0000,#E53935);border-radius:4px;transition:width 0.45s ease;}
        .pb-meta{display:flex;justify-content:space-between;font-size:12px;color:#9BA3B8;}
        .dots{display:flex;justify-content:center;gap:8px;margin-top:8px;}
        .dot{width:9px;height:9px;border-radius:50%;background:#CC0000;animation:bounce 1s infinite;opacity:0.75;}
        /* DONE */
        .done-banner{display:flex;align-items:center;gap:18px;background:linear-gradient(135deg,#A80000,#CC0000);border-radius:13px;padding:26px 28px;margin-bottom:24px;color:#fff;}
        .done-h{font-size:19px;font-weight:800;margin-bottom:4px;}
        .done-s{font-size:13px;opacity:0.85;}
        .tbl-wrap{border:1.5px solid #E4E8F2;border-radius:11px;overflow:hidden;margin-bottom:24px;}
        .tbl-top{display:flex;align-items:center;justify-content:space-between;padding:12px 18px;background:#FAFBFE;border-bottom:1.5px solid #E4E8F2;font-size:12.5px;font-weight:700;color:#555;gap:8px;}
        .re-btn{display:inline-flex;align-items:center;gap:5px;background:#CC0000;color:#fff;border:none;border-radius:6px;padding:6px 14px;font-family:'Inter',sans-serif;font-size:12px;font-weight:700;cursor:pointer;text-decoration:none;}
        .tbl{width:100%;border-collapse:collapse;font-size:13px;}
        .tbl th{background:#FAFBFE;padding:10px 14px;text-align:left;font-weight:700;color:#888;font-size:11px;text-transform:uppercase;letter-spacing:0.5px;border-bottom:1.5px solid #E4E8F2;}
        .tbl td{padding:11px 14px;border-bottom:1px solid #F2F4FA;vertical-align:middle;}
        .tbl tr:last-child td{border-bottom:none;}
        .tbl tr:hover td{background:#FAFBFE;}
        .td-s{font-weight:700;color:#CC0000;white-space:nowrap;}
        .td-n{font-weight:500;color:#333;}
        .td-f{font-size:11.5px;color:#9BA3B8;font-family:monospace;}
        .td-a{text-align:right;}
        .dl-a{display:inline-flex;align-items:center;gap:4px;color:#CC0000;border:1.5px solid #CC0000;border-radius:5px;padding:4px 10px;font-size:11.5px;font-weight:700;text-decoration:none;transition:all 0.14s;}
        .dl-a:hover{background:#CC0000;color:#fff;}
        .btn-pr{display:inline-flex;align-items:center;gap:8px;background:#CC0000;color:#fff;border:none;border-radius:9px;padding:13px 30px;font-family:'Inter',sans-serif;font-size:14px;font-weight:700;cursor:pointer;letter-spacing:0.2px;transition:background 0.15s;}
        .btn-pr:hover{background:#A80000;}
        /* ERROR */
        .err-wrap{text-align:center;padding:24px 0;}
        .err-h{font-size:20px;font-weight:800;color:#CC0000;margin-bottom:8px;}
        .err-s{font-size:14px;color:#666;line-height:1.6;margin-bottom:24px;max-width:420px;margin-left:auto;margin-right:auto;}
        /* FOOTER */
        .foot{background:#fff;border-top:1px solid #E4E8F2;padding:14px 32px;display:flex;justify-content:space-between;align-items:center;font-size:11.5px;color:#C0C6D8;}
        @keyframes spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}
        @keyframes bounce{0%,80%,100%{transform:translateY(0);opacity:0.5;}40%{transform:translateY(-9px);opacity:1;}}
      `}</style>
    </div>
  );
}
'''

out = r'C:\Users\Kalai\Downloads\XSLX TO PDF\frontend\src\app\page.tsx'
with open(out, 'w', encoding='utf-8') as f:
    f.write(tsx)
print("Written:", out)
