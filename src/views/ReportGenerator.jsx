import { useState, useEffect, useCallback } from "react";

const JIRA_PROJECTS = [
  {key:"AI",  name:"AI & Automation Kanban (AI&A)",          hasInitiative:true},
  {key:"AIE", name:"AI Catalog",                              hasInitiative:true},
  {key:"AE",  name:"AI Enablement",                          hasInitiative:true},
  {key:"AILT",name:"AI Leadership Team",                     hasInitiative:false},
  {key:"AOT", name:"AI Ops Transformation",                  hasInitiative:false},
  {key:"ANZMKT",name:"ANZ",                                  hasInitiative:true},
  {key:"ANZECOM",name:"ANZ ECOM",                            hasInitiative:true},
  {key:"AET", name:"ANZ ECom Team",                          hasInitiative:false},
  {key:"ANZP",name:"ANZ Payroll",                            hasInitiative:false},
  {key:"AM",  name:"APAC Merchandising",                     hasInitiative:false},
  {key:"ATHENA",name:"Athena",                               hasInitiative:true},
  {key:"BOM", name:"B2B Ordering Management",                hasInitiative:false},
  {key:"BP25",name:"Budget Process 2025",                    hasInitiative:true},
  {key:"CBWCI",name:"C&C Bi-weekly check-in",               hasInitiative:true},
  {key:"CEPL",name:"Martech / Consumer Engagement Kanban",   hasInitiative:true},
  {key:"AIMACA",name:"D&A - Consumer Analytics",             hasInitiative:true},
  {key:"AIMAGTML",name:"D&A - GTM & Loyalty",               hasInitiative:true},
  {key:"AIMAMDM",name:"D&A - Marketing Data Model",         hasInitiative:true},
  {key:"AIMA",name:"D&A PL Marketing Analytics",            hasInitiative:true},
  {key:"AEP", name:"D&A Reliability",                        hasInitiative:true},
  {key:"AMO", name:"Pandora Agile Center of Excellence",     hasInitiative:false},
  {key:"CBA", name:"Pandora PIM",                            hasInitiative:true},
  {key:"AA",  name:"Moonshot",                               hasInitiative:false},
  {key:"CCO", name:"ERM-CCO",                                hasInitiative:false},
  {key:"CDR", name:"COMPASS Data Readiness",                 hasInitiative:false},
  {key:"C360",name:"Customer Management",                    hasInitiative:false},
];

const RAG = {
  green: { bg:"#F0FDF4", border:"#BBF7D0", text:"#16A34A", label:"On track"  },
  amber: { bg:"#FFFBEB", border:"#FDE68A", text:"#D97706", label:"At risk"   },
  red:   { bg:"#FEF2F2", border:"#FECACA", text:"#DC2626", label:"Off track" },
  gray:  { bg:"#F8FAFC", border:"#E2E8F0", text:"#64748B", label:"No data"   },
};
const ragFromPct = p => p == null ? "gray" : p >= 70 ? "green" : p >= 40 ? "amber" : "red";

const CONFIG_KEY = "pl-report:config";
const dataKey    = k => `pl-report:data:${k}`;
const manualKey  = k => `pl-report:manual:${k}`;
const wellnessKey= k => `pl-report:wellness:${k}`;

const load  = k => { try { return JSON.parse(localStorage.getItem(k)||"null"); } catch { return null; }};
const save  = (k,v) => { try { localStorage.setItem(k,JSON.stringify(v)); } catch {} };

const loadConfig      = ()  => load(CONFIG_KEY);
const saveConfig      = v   => save(CONFIG_KEY, v);
const loadProjectData = k   => load(dataKey(k));
const saveProjectData = (k,v)=> save(dataKey(k), v);
const loadManual      = k   => load(manualKey(k));
const saveManual      = (k,v)=> save(manualKey(k), v);
const loadWellness    = k   => load(wellnessKey(k)) || {};
const saveWellness    = (k,v)=> save(wellnessKey(k), v);
const defaultManual   = ()  => ({ dlName:"", commentary:"", asks:"", capacity:{}, risks:[] });

function addSurveyResp(projectKey, member, scores) {
  const all = loadWellness(projectKey);
  const t = all[member] || { responses:[], prev:null, current:null };
  const responses = [...(t.responses||[]), { ...scores, at:Date.now() }];
  const avg = dim => { const vs=responses.map(r=>r[dim]).filter(Boolean); return vs.length?Math.round(vs.reduce((a,b)=>a+b)/vs.length*10)/10:null; };
  all[member] = { ...t, responses, current:{ happiness:avg("happiness"), teamSpirit:avg("teamSpirit"), belonging:avg("belonging"), respondents:responses.length }};
  saveWellness(projectKey, all);
  return responses.length;
}

function ProjectSwitcher({ current, onSelect, onClose }) {
  const [search, setSearch] = useState("");
  const filtered = JIRA_PROJECTS.filter(p =>
    p.name.toLowerCase().includes(search.toLowerCase()) ||
    p.key.toLowerCase().includes(search.toLowerCase())
  );
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:1000,padding:16}}>
      <div style={{background:"#fff",borderRadius:14,width:"100%",maxWidth:520,maxHeight:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
        <div style={{padding:"20px 20px 12px",borderBottom:"1px solid #E2E8F0"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <h3 style={{fontSize:16,fontWeight:700,color:"#0F172A",margin:0}}>Switch Jira project</h3>
            <button onClick={onClose} style={{background:"none",border:"none",fontSize:22,cursor:"pointer",color:"#64748B",lineHeight:1,padding:"0 4px"}}>×</button>
          </div>
          <input autoFocus value={search} onChange={e=>setSearch(e.target.value)}
            placeholder="Search by name or key..."
            style={{width:"100%",border:"1px solid #E2E8F0",borderRadius:8,padding:"9px 12px",fontSize:14,outline:"none",boxSizing:"border-box"}}/>
        </div>
        <div style={{overflowY:"auto",flex:1}}>
          {filtered.length===0 && <p style={{padding:20,color:"#94A3B8",fontSize:13,textAlign:"center"}}>No projects found</p>}
          {filtered.map(p=>(
            <button key={p.key} onClick={()=>onSelect(p)}
              style={{display:"flex",alignItems:"center",gap:10,width:"100%",padding:"11px 20px",background:p.key===current?.key?"#EFF6FF":"transparent",border:"none",borderBottom:"1px solid #F1F5F9",cursor:"pointer",textAlign:"left"}}>
              <span style={{background:p.hasInitiative?"#EFF6FF":"#F8FAFC",color:p.hasInitiative?"#1E40AF":"#64748B",border:`1px solid ${p.hasInitiative?"#BFDBFE":"#E2E8F0"}`,borderRadius:4,padding:"2px 7px",fontSize:11,fontWeight:700,flexShrink:0,minWidth:58,textAlign:"center"}}>{p.key}</span>
              <span style={{fontSize:13,color:"#1E293B",flex:1}}>{p.name}</span>
              {p.hasInitiative&&<span style={{fontSize:10,color:"#16A34A",fontWeight:600,flexShrink:0}}>Initiatives ✓</span>}
              {p.key===current?.key&&<span style={{fontSize:10,color:"#4FC3F7",fontWeight:700,flexShrink:0}}>Active</span>}
            </button>
          ))}
        </div>
        <div style={{padding:"10px 20px",borderTop:"1px solid #E2E8F0",fontSize:11,color:"#94A3B8"}}>
          "Initiatives ✓" = full Initiative → Epic hierarchy available
        </div>
      </div>
    </div>
  );
}

function SurveyView({ projectKey, member }) {
  const [scores, setScores] = useState({happiness:0,teamSpirit:0,belonging:0});
  const [done, setDone] = useState(false);
  const [count, setCount] = useState(null);
  const allDone = Object.values(scores).every(v=>v>0);
  const QS=[
    {key:"happiness",  q:"How is your overall happiness at work right now?"},
    {key:"teamSpirit", q:"How would you rate the team spirit right now?"},
    {key:"belonging",  q:"How strong is your sense of belonging in the team?"},
  ];
  if(done) return (
    <div style={{minHeight:"100vh",background:"#0D1440",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif"}}>
      <div style={{textAlign:"center",padding:40}}>
        <div style={{fontSize:56,color:"#F59E0B",marginBottom:16}}>★</div>
        <h2 style={{color:"#fff",fontSize:24,marginBottom:8}}>Thank you!</h2>
        <p style={{color:"#CADCFC",opacity:.8}}>{member} — recorded anonymously.</p>
        {count!=null&&count<5&&<p style={{color:"rgba(255,255,255,.5)",marginTop:10,fontSize:13}}>{5-count} more responses needed before scores are visible.</p>}
      </div>
    </div>
  );
  return (
    <div style={{minHeight:"100vh",background:"#0D1440",fontFamily:"system-ui,sans-serif",padding:"0 0 40px"}}>
      <div style={{maxWidth:480,margin:"0 auto"}}>
        <div style={{padding:"32px 24px 20px",textAlign:"center"}}>
          <p style={{fontSize:11,color:"#4FC3F7",fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",marginBottom:8}}>PL Report Tool · Pulse Survey</p>
          <h1 style={{color:"#fff",fontSize:22,marginBottom:4}}>{member}</h1>
          <p style={{color:"#CADCFC",opacity:.8,fontSize:14}}>Anonymous · 3 questions · 30 seconds</p>
        </div>
        {QS.map(({key,q})=>(
          <div key={key} style={{background:"#fff",borderRadius:14,padding:"22px 20px",margin:"0 16px 12px",boxShadow:"0 2px 8px rgba(0,0,0,.15)"}}>
            <p style={{fontSize:17,fontWeight:600,color:"#0F172A",lineHeight:1.4,marginBottom:18}}>{q}</p>
            <div style={{display:"flex",justifyContent:"center",gap:12}}>
              {[1,2,3,4,5].map(n=>(
                <button key={n} onClick={()=>setScores(s=>({...s,[key]:n}))}
                  style={{background:"none",border:"none",cursor:"pointer",fontSize:36,color:n<=scores[key]?"#F59E0B":"#CBD5E1",padding:4,lineHeight:1,WebkitTapHighlightColor:"transparent"}} type="button">
                  {n<=scores[key]?"★":"☆"}
                </button>
              ))}
            </div>
            {scores[key]>0&&<p style={{textAlign:"center",marginTop:10,fontSize:13,color:"#64748B"}}>
              {["","Not great","Below average","Okay","Good","Excellent"][scores[key]]}
            </p>}
          </div>
        ))}
        <button onClick={()=>{const n=addSurveyResp(projectKey,member,scores);setCount(n);setDone(true);}} disabled={!allDone}
          style={{display:"block",width:"calc(100% - 32px)",margin:"4px 16px 0",padding:16,background:allDone?"#1E2761":"#334155",color:"#fff",border:"none",borderRadius:12,fontSize:16,fontWeight:700,cursor:allDone?"pointer":"not-allowed",fontFamily:"system-ui,sans-serif"}}>
          {allDone?"Submit anonymously":"Please answer all 3 questions"}
        </button>
        <p style={{textAlign:"center",marginTop:12,fontSize:12,color:"rgba(255,255,255,.35)",padding:"0 16px"}}>Individual scores never shown. Visible once 5+ respond.</p>
      </div>
    </div>
  );
}

export default function ReportGenerator() {
  const params = new URLSearchParams(typeof window!=="undefined"?window.location.search:"");
  const surveyMember  = params.get("member");
  const surveyProject = params.get("project")||"";
  if (surveyMember) return <SurveyView projectKey={surveyProject} member={surveyMember} />;

  const initProject = loadConfig() || JIRA_PROJECTS.find(p=>p.key==="AMO") || JIRA_PROJECTS[0];
  const [activeProject, setActiveProject] = useState(initProject);
  const [liveData,  setLiveData]  = useState(() => loadProjectData(initProject.key));
  const [manual,    setManual]    = useState(() => loadManual(initProject.key) || defaultManual());
  const [wellness,  setWellness]  = useState(() => loadWellness(initProject.key));
  const [showSwitcher, setShowSwitcher] = useState(false);
  const [phase,  setPhase]  = useState("idle");
  const [tab,    setTab]    = useState("overview");
  const [pptxOk, setPptxOk] = useState(false);
  const [notice, setNotice] = useState(null);

  useEffect(()=>{
    if(window.PptxGenJS){setPptxOk(true);return;}
    const sc=document.createElement("script");
    sc.src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";
    sc.onload=()=>setPptxOk(true);
    document.head.appendChild(sc);
  },[]);

  const setM = useCallback((k,v)=>{
    setManual(m=>{ const next={...m,[k]:v}; saveManual(activeProject.key,next); return next; });
  },[activeProject.key]);

  useEffect(()=>{
    const t=setInterval(()=>setWellness(loadWellness(activeProject.key)),10000);
    return()=>clearInterval(t);
  },[activeProject.key]);

  function switchProject(p) {
    saveConfig(p);
    setActiveProject(p);
    const cached = loadProjectData(p.key);
    setLiveData(cached);
    setManual(loadManual(p.key)||defaultManual());
    setWellness(loadWellness(p.key));
    setShowSwitcher(false);
    setTab("overview");
    setNotice(cached ? `Switched to ${p.name} — loaded cached data.` : `Switched to ${p.name}. No data yet — fetch via Claude.`);
    setTimeout(()=>setNotice(null),4000);
  }

  async function generate() {
    if(!pptxOk)return;
    setPhase("generating");
    const ts=new Date().toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
    const pptx=new window.PptxGenJS();
    pptx.layout="LAYOUT_16x9";
    const PL=activeProject.name; const PK=activeProject.key;
    const sh=()=>({type:"outer",blur:6,offset:2,angle:135,color:"000000",opacity:0.08});
    function hdr(s,title,sub,bg="1E2761"){
      s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:.7,fill:{color:bg},line:{color:bg}});
      s.addText(title,{x:.4,y:0,w:7.8,h:.7,fontSize:19,color:"FFFFFF",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
      if(sub) s.addText(sub,{x:7.8,y:0,w:2,h:.7,fontSize:9,color:"CADCFC",fontFace:"Calibri",align:"right",valign:"middle",margin:0});
    }
    function pb(s,x,y,w,pct,clr){
      s.addShape(pptx.shapes.RECTANGLE,{x,y,w,h:.13,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
      if(pct>0) s.addShape(pptx.shapes.RECTANGLE,{x,y,w:Math.max(.04,w*pct/100),h:.13,fill:{color:clr},line:{color:clr}});
    }
    const epics=liveData?.epics||[]; const deps=liveData?.dependencies||[];
    const ct=liveData?.cycleTime||null; const members=liveData?.members||[];
    const total=epics.length,done=epics.filter(e=>e.pct>=100).length;
    const inProg=epics.filter(e=>e.pct>0&&e.pct<100).length,notStart=epics.filter(e=>e.pct===0).length;
    const ovPct=total?Math.round(epics.reduce((s,e)=>s+e.pct,0)/total):0;

    const s1=pptx.addSlide(); s1.background={color:"0D1440"};
    s1.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:.18,h:5.625,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
    s1.addShape(pptx.shapes.RECTANGLE,{x:0,y:4.7,w:10,h:.925,fill:{color:"1E2761"},line:{color:"1E2761"}});
    s1.addText(PK.toUpperCase(),{x:.42,y:1.05,w:9,h:.44,fontSize:11,color:"4FC3F7",bold:true,charSpacing:5,fontFace:"Calibri",margin:0});
    s1.addText(PL,{x:.42,y:1.55,w:9,h:1.0,fontSize:36,color:"FFFFFF",bold:true,fontFace:"Calibri",margin:0});
    s1.addText("Epics · Flow metrics · Dependencies · Team health · Risks",{x:.42,y:2.68,w:8.5,h:.44,fontSize:13,color:"CADCFC",fontFace:"Calibri",italic:true,margin:0});
    s1.addText(`Generated: ${ts}`,{x:.42,y:4.8,w:6,h:.28,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});
    s1.addText(`Prepared by: ${manual.dlName||"—"}`,{x:.42,y:5.1,w:6,h:.28,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});

    const s2=pptx.addSlide(); s2.background={color:"F7F9FF"};
    hdr(s2,"Portfolio overview",`${PK} · ${liveData?.fetchedAt||ts}`);
    [{l:"Total",v:total,c:"1E2761",bg:"EFF6FF",bd:"BFDBFE"},{l:"Done",v:done,c:"16A34A",bg:"F0FDF4",bd:"BBF7D0"},{l:"In progress",v:inProg,c:"D97706",bg:"FFFBEB",bd:"FDE68A"},{l:"Not started",v:notStart,c:"DC2626",bg:"FEF2F2",bd:"FECACA"}].forEach((st,i)=>{
      const x=.4+i*2.35;
      s2.addShape(pptx.shapes.RECTANGLE,{x,y:.86,w:2.1,h:1.2,fill:{color:st.bg},line:{color:st.bd},shadow:sh()});
      s2.addText(String(st.v),{x,y:.9,w:2.1,h:.68,fontSize:36,color:st.c,bold:true,fontFace:"Calibri",align:"center",valign:"middle",margin:0});
      s2.addText(st.l,{x,y:1.58,w:2.1,h:.32,fontSize:11,color:st.c,fontFace:"Calibri",align:"center",margin:0});
    });
    s2.addText(`Portfolio completion: ${ovPct}%`,{x:.4,y:2.22,w:5,h:.28,fontSize:12,color:"1E293B",bold:true,fontFace:"Calibri",margin:0});
    s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:2.55,w:9.2,h:.26,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
    if(ovPct>0) s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:2.55,w:Math.max(.1,9.2*ovPct/100),h:.26,fill:{color:"1E2761"},line:{color:"1E2761"}});
    if(manual.commentary){
      s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.05,w:9.2,h:2.05,fill:{color:"FFFFFF"},line:{color:"E2E8F0"},shadow:sh()});
      s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.05,w:.055,h:2.05,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
      s2.addText("Delivery Lead commentary",{x:.55,y:3.15,w:8.9,h:.24,fontSize:9,color:"8A9BC4",bold:true,fontFace:"Calibri",margin:0});
      s2.addText(manual.commentary,{x:.55,y:3.41,w:8.9,h:1.6,fontSize:12,color:"1E293B",fontFace:"Calibri",valign:"top",margin:0});
    }

    const s3=pptx.addSlide(); s3.background={color:"F7F9FF"};
    hdr(s3,"Epic progress — active",PK,"92400E");
    epics.filter(e=>e.pct>0&&e.pct<100).slice(0,8).forEach((ep,ei)=>{
      const col=ei%2,row=Math.floor(ei/2),x=.4+col*4.75,y=.86+row*1.1;
      const rc=RAG[ep.rag||ragFromPct(ep.pct)];
      s3.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:.95,fill:{color:rc.bg.replace("#","")},line:{color:rc.border.replace("#","")},shadow:sh()});
      s3.addShape(pptx.shapes.RECTANGLE,{x,y,w:.055,h:.95,fill:{color:rc.text.replace("#","")},line:{color:rc.text.replace("#","")}});
      s3.addText(`${ep.key}  ·  ${ep.assignee||""}`,{x:x+.13,y:y+.07,w:4.2,h:.2,fontSize:9,color:rc.text.replace("#",""),bold:true,fontFace:"Calibri",margin:0});
      s3.addText(ep.name,{x:x+.13,y:y+.27,w:4.2,h:.26,fontSize:12,color:"1E293B",fontFace:"Calibri",margin:0});
      pb(s3,x+.13,y+.63,3.5,ep.pct,rc.text.replace("#",""));
      s3.addText(`${ep.pct}%`,{x:x+3.8,y:y+.55,w:.6,h:.24,fontSize:11,color:rc.text.replace("#",""),bold:true,align:"right",fontFace:"Calibri",margin:0});
    });

    const s4=pptx.addSlide(); s4.background={color:"F7F9FF"};
    hdr(s4,"Flow metrics",`${PK} · Live where available`,"0F766E");
    if(ct){
      s4.addText("Cycle time — live",{x:.4,y:.86,w:5.5,h:.28,fontSize:11,color:"0F766E",bold:true,fontFace:"Calibri",margin:0});
      ct.trend.slice(-4).forEach((p,pi)=>{
        const x=.4+pi*1.36,y=1.38,barH=Math.min(1.6,(p.avgDays/50)*1.6);
        const clr=p.avgDays<=14?"16A34A":p.avgDays<=25?"D97706":"DC2626";
        s4.addShape(pptx.shapes.RECTANGLE,{x:x+.2,y:y+(1.6-barH),w:.8,h:barH,fill:{color:clr},line:{color:clr}});
        s4.addText(`${p.avgDays}d`,{x,y:y+(1.6-barH)-.28,w:1.2,h:.25,fontSize:10,color:clr,bold:true,align:"center",fontFace:"Calibri",margin:0});
        s4.addText(p.period.replace(/ 20\d\d/,""),{x,y:3.06,w:1.2,h:.22,fontSize:8,color:"64748B",align:"center",fontFace:"Calibri",margin:0});
      });
      s4.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.48,w:5.5,h:.22,fill:{color:"EFF6FF"},line:{color:"BFDBFE"}});
      s4.addText(`Avg: ${ct.avg}d  ·  Median: ${ct.median}d  ·  Target: ≤14d`,{x:.4,y:3.48,w:5.5,h:.22,fontSize:9,color:"1E40AF",fontFace:"Calibri",align:"center",valign:"middle",margin:0});
    } else {
      s4.addShape(pptx.shapes.RECTANGLE,{x:.4,y:.86,w:5.5,h:3.0,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s4.addText("Cycle time — no data available",{x:.4,y:.86,w:5.5,h:3.0,fontSize:12,color:"94A3B8",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    }
    s4.addText("Velocity — last 3 sprints",{x:6.1,y:.86,w:3.5,h:.28,fontSize:11,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
    ["Sprint -2","Sprint -1","Current"].forEach((lbl,si)=>{
      const y=1.22+si*.7;
      s4.addShape(pptx.shapes.RECTANGLE,{x:6.1,y,w:3.5,h:.56,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s4.addText(`${lbl}: — pts committed / — pts delivered`,{x:6.2,y,w:3.3,h:.56,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",margin:0});
    });
    s4.addText("Sprint reliability",{x:6.1,y:3.46,w:3.5,h:.26,fontSize:11,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
    ["Sprint -2","Sprint -1","Current"].forEach((lbl,si)=>{
      const y=3.78+si*.42;
      s4.addShape(pptx.shapes.RECTANGLE,{x:6.1,y,w:3.5,h:.34,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s4.addText(`${lbl}: — / — = —%`,{x:6.18,y,w:3.3,h:.34,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",margin:0});
    });

    const s5=pptx.addSlide(); s5.background={color:"F7F9FF"};
    hdr(s5,'Dependencies',PK,"7C3AED");
    if(deps.length>0){
      deps.slice(0,5).forEach((d,di)=>{
        const y=.86+di*.9;
        s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.76,fill:{color:"F5F3FF"},line:{color:"DDD6FE"},shadow:sh()});
        s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:.055,h:.76,fill:{color:"7C3AED"},line:{color:"7C3AED"}});
        s5.addShape(pptx.shapes.RECTANGLE,{x:.55,y:y+.1,w:3.6,h:.52,fill:{color:"EDE9FE"},line:{color:"C4B5FD"}});
        s5.addText(d.from,{x:.6,y:y+.1,w:.9,h:.52,fontSize:9,color:"7C3AED",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
        s5.addText(d.fromName||"",{x:1.5,y:y+.1,w:2.6,h:.52,fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
        s5.addText("depends on →",{x:4.25,y:y+.18,w:1.4,h:.3,fontSize:9,color:"7C3AED",italic:true,fontFace:"Calibri",align:"center",margin:0});
        const dc=d.status==="Done";
        s5.addShape(pptx.shapes.RECTANGLE,{x:5.7,y:y+.1,w:3.7,h:.52,fill:{color:dc?"F0FDF4":"EDE9FE"},line:{color:dc?"BBF7D0":"C4B5FD"}});
        s5.addText(d.to,{x:5.75,y:y+.1,w:.9,h:.52,fontSize:9,color:dc?"16A34A":"7C3AED",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
        s5.addText(d.toName||"",{x:6.65,y:y+.1,w:2.6,h:.52,fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
      });
    } else {
      s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y:.86,w:9.2,h:1.5,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s5.addText("No dependency links found",{x:.4,y:.86,w:9.2,h:1.5,fontSize:13,color:"94A3B8",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    }

    const s6=pptx.addSlide(); s6.background={color:"F7F9FF"};
    hdr(s6,"Team capacity","Enter manually","0F766E");
    members.forEach((m,mi)=>{
      const col=mi%2,row=Math.floor(mi/2),x=.4+col*4.75,y=.86+row*.72;
      const cap=manual.capacity?.[m]||null;
      const clr=!cap?"CBD5E1":cap>=80?"16A34A":cap>=50?"D97706":"DC2626";
      const bg=!cap?"F8FAFC":cap>=80?"F0FDF4":cap>=50?"FFFBEB":"FEF2F2";
      s6.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:.58,fill:{color:bg},line:{color:"E2E8F0"},shadow:sh()});
      s6.addText(m,{x:x+.12,y,w:1.8,h:.58,fontSize:11,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
      if(cap){
        s6.addShape(pptx.shapes.RECTANGLE,{x:x+2.0,y:y+.16,w:2.0,h:.26,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
        s6.addShape(pptx.shapes.RECTANGLE,{x:x+2.0,y:y+.16,w:Math.max(.04,2.0*cap/100),h:.26,fill:{color:clr},line:{color:clr}});
        s6.addText(`${cap}%`,{x:x+4.1,y,w:.35,h:.58,fontSize:11,color:clr,bold:true,fontFace:"Calibri",valign:"middle",margin:0});
      } else {
        s6.addText("— enter capacity %",{x:x+2.0,y,w:2.4,h:.58,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",margin:0});
      }
    });

    const s7=pptx.addSlide(); s7.background={color:"F7F9FF"};
    hdr(s7,"Risks",PK,"B91C1C");
    const sevC={High:{bg:"FEF2F2",bd:"FECACA",tx:"DC2626"},Medium:{bg:"FFFBEB",bd:"FDE68A",tx:"D97706"},Low:{bg:"F0FDF4",bd:"BBF7D0",tx:"16A34A"}};
    for(let ri=0;ri<6;ri++){
      const col=ri%2,row=Math.floor(ri/2),x=.4+col*4.75,y=.86+row*1.55;
      const risk=(manual.risks||[])[ri];
      if(risk){
        const sc=sevC[risk.severity]||sevC.Medium;
        s7.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:1.38,fill:{color:sc.bg},line:{color:sc.bd},shadow:sh()});
        s7.addShape(pptx.shapes.RECTANGLE,{x,y,w:.055,h:1.38,fill:{color:sc.tx},line:{color:sc.tx}});
        s7.addText(risk.severity,{x:x+.1,y:y+.07,w:.8,h:.24,fontSize:8,color:sc.tx,bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
        s7.addText(risk.title||"",{x:x+.13,y:y+.34,w:4.2,h:.3,fontSize:12,color:"1E293B",bold:true,fontFace:"Calibri",margin:0});
        s7.addText(risk.description||"",{x:x+.13,y:y+.64,w:4.2,h:.3,fontSize:10,color:"475569",fontFace:"Calibri",margin:0});
        s7.addShape(pptx.shapes.RECTANGLE,{x:x+.13,y:y+1.0,w:4.2,h:.28,fill:{color:"FFFBEB"},line:{color:"FDE68A"}});
        s7.addText(`Mitigation: ${risk.mitigation||"—"}  ·  Owner: ${risk.owner||"—"}`,{x:x+.2,y:y+1.0,w:4.0,h:.28,fontSize:9,color:"D97706",fontFace:"Calibri",valign:"middle",margin:0});
      } else {
        s7.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:1.38,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
        s7.addText("No risk entered",{x,y,w:4.5,h:1.38,fontSize:10,color:"CBD5E1",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
      }
    }

    const s8=pptx.addSlide(); s8.background={color:"F7F9FF"};
    hdr(s8,"Where you can help",PK);
    const asks=(manual.asks||"").split("\n").filter(Boolean).slice(0,5);
    asks.forEach((a,ai)=>{
      const y=.86+ai*.88;
      s8.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.72,fill:{color:"EFF6FF"},line:{color:"BFDBFE"},shadow:sh()});
      s8.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:.055,h:.72,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
      s8.addText(`${ai+1}`,{x:.55,y,w:.4,h:.72,fontSize:16,color:"4FC3F7",bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
      s8.addText(a,{x:1.1,y,w:8.3,h:.72,fontSize:13,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    });
    if(!asks.length) s8.addText("No specific asks this period.",{x:.4,y:1.6,w:9.2,h:.5,fontSize:14,color:"94A3B8",italic:true,fontFace:"Calibri",margin:0});

    const sE=pptx.addSlide(); sE.background={color:"0D1440"};
    sE.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:.18,h:5.625,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
    sE.addShape(pptx.shapes.RECTANGLE,{x:0,y:4.7,w:10,h:.925,fill:{color:"1E2761"},line:{color:"1E2761"}});
    sE.addText("Thank you",{x:.42,y:1.6,w:9,h:1.0,fontSize:42,color:"FFFFFF",bold:true,fontFace:"Calibri",margin:0});
    sE.addText("Questions & discussion",{x:.42,y:2.7,w:9,h:.5,fontSize:18,color:"CADCFC",italic:true,fontFace:"Calibri",margin:0});
    sE.addText(`${PL}  ·  ${manual.dlName||"Delivery Lead"}`,{x:.42,y:3.5,w:9,h:.34,fontSize:12,color:"4FC3F7",fontFace:"Calibri",margin:0});
    sE.addText(`Generated: ${ts}`,{x:.42,y:4.8,w:8,h:.3,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});

    await pptx.writeFile({fileName:`${PK}-Report-${ts.replace(/[: ,/]/g,"-")}.pptx`});
    setPhase("done");
  }

  const S={
    page:{fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#F1F5F9",minHeight:"100vh",padding:16,maxWidth:980,margin:"0 auto"},
    hdr:{background:"#0D1440",borderRadius:12,padding:"16px 20px",marginBottom:12},
    card:{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"18px 20px",marginBottom:12,boxShadow:"0 1px 3px rgba(0,0,0,.06)"},
    dark:{background:"#0D1440",borderRadius:12,padding:"20px",marginBottom:12,textAlign:"center"},
    sec:{fontSize:11,fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:".07em",marginBottom:12},
    lbl:{fontSize:11,fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:".05em",display:"block",marginBottom:4},
    inp:{width:"100%",border:"1px solid #E2E8F0",borderRadius:7,padding:"9px 12px",fontSize:14,background:"#fff",color:"#0F172A",outline:"none",boxSizing:"border-box",fontFamily:"'Segoe UI',system-ui,sans-serif"},
    ta:{width:"100%",border:"1px solid #E2E8F0",borderRadius:7,padding:"9px 12px",fontSize:13,background:"#fff",color:"#0F172A",outline:"none",resize:"vertical",fontFamily:"'Segoe UI',system-ui,sans-serif",boxSizing:"border-box"},
    btn:{display:"inline-flex",alignItems:"center",gap:7,padding:"9px 18px",border:"none",borderRadius:8,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"'Segoe UI',system-ui,sans-serif"},
    navBtn:(a)=>({padding:"7px 14px",border:"none",borderRadius:7,fontSize:12,fontWeight:600,cursor:"pointer",background:a?"#1E2761":"#F1F5F9",color:a?"#fff":"#64748B",fontFamily:"'Segoe UI',system-ui,sans-serif"}),
    empty:{background:"#F8FAFC",border:"1px dashed #CBD5E1",borderRadius:7,padding:"12px 14px",fontSize:13,color:"#94A3B8",fontStyle:"italic",textAlign:"center"},
  };

  const epics=liveData?.epics||[]; const deps=liveData?.dependencies||[];
  const ct=liveData?.cycleTime||null; const members=liveData?.members||[];
  const total=epics.length,done=epics.filter(e=>e.pct>=100).length;
  const inProg=epics.filter(e=>e.pct>0&&e.pct<100).length,notStart=epics.filter(e=>e.pct===0).length;
  const ovPct=total?Math.round(epics.reduce((s,e)=>s+e.pct,0)/total):0;
  const surveyBase=typeof window!=="undefined"?`${window.location.origin}${window.location.pathname}`:"";
  const TABS=[{id:"overview",l:"Portfolio"},{id:"flow",l:"Flow"},{id:"deps",l:"Dependencies"},{id:"capacity",l:"Capacity"},{id:"wellness",l:"Wellness"},{id:"risks",l:"Risks"},{id:"asks",l:"Asks"}];

  return (
    <div style={S.page}>
      {showSwitcher&&<ProjectSwitcher current={activeProject} onSelect={switchProject} onClose={()=>setShowSwitcher(false)}/>}

      <div style={S.hdr}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
          <div>
            <div style={{fontSize:10,color:"#4FC3F7",fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",marginBottom:4}}>PACE · PL Report Tool v0.2</div>
            <div style={{fontSize:20,fontWeight:700,color:"#fff",marginBottom:2}}>{activeProject.name}</div>
            <div style={{fontSize:12,color:"#CADCFC",opacity:.8}}>
              {activeProject.key} · {liveData?`Data: ${liveData.fetchedAt||"cached"}`:"No data loaded"}
              {activeProject.hasInitiative&&<span style={{marginLeft:8,color:"#4ADE80",fontSize:11}}>Initiatives ✓</span>}
            </div>
          </div>
          <button onClick={()=>setShowSwitcher(true)} style={{...S.btn,background:"rgba(255,255,255,.12)",color:"#fff",border:"1px solid rgba(255,255,255,.2)"}}>
            ⇄ Change project
          </button>
        </div>
      </div>

      {notice&&<div style={{background:"#EFF6FF",border:"1px solid #BFDBFE",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#1E40AF",marginBottom:12}}>{notice}</div>}

      {!liveData&&(
        <div style={{...S.card,textAlign:"center",padding:"32px 24px"}}>
          <div style={{fontSize:32,marginBottom:12}}>📋</div>
          <h3 style={{fontSize:16,fontWeight:700,marginBottom:8}}>No data loaded for {activeProject.key}</h3>
          <p style={{fontSize:13,color:"#64748B",maxWidth:460,margin:"0 auto 8px"}}>Ask Claude in this conversation:</p>
          <code style={{display:"block",background:"#F1F5F9",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#1E2761",margin:"0 auto",maxWidth:460,textAlign:"left"}}>
            "Fetch Jira data for project {activeProject.key} and update the report tool"
          </code>
          <p style={{fontSize:12,color:"#94A3B8",marginTop:12}}>Previously loaded data is remembered per project in your browser.</p>
        </div>
      )}

      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
        {TABS.map(t=><button key={t.id} style={S.navBtn(tab===t.id)} onClick={()=>setTab(t.id)}>{t.l}</button>)}
      </div>

      {tab==="overview"&&(
        <div style={S.card}>
          <p style={S.sec}>Portfolio snapshot</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
            {[{l:"Total",v:total,c:"#1E2761",bg:"#EFF6FF"},{l:"Done",v:done,c:"#16A34A",bg:"#F0FDF4"},{l:"In progress",v:inProg,c:"#D97706",bg:"#FFFBEB"},{l:"Not started",v:notStart,c:"#DC2626",bg:"#FEF2F2"}].map((s,i)=>(
              <div key={i} style={{background:s.bg,borderRadius:10,padding:"12px 8px",textAlign:"center",border:"1px solid #E2E8F0"}}>
                <div style={{fontSize:28,fontWeight:700,color:s.c}}>{s.v}</div>
                <div style={{fontSize:11,color:"#64748B",marginTop:2}}>{s.l}</div>
              </div>
            ))}
          </div>
          {total>0&&<><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Overall completion — {ovPct}%</div>
          <div style={{height:10,background:"#E2E8F0",borderRadius:5,marginBottom:14}}><div style={{width:`${ovPct}%`,height:"100%",background:"#1E2761",borderRadius:5}}/></div></>}
          <label style={S.lbl}>Portfolio commentary</label>
          <textarea style={{...S.ta,minHeight:72}} rows={3} value={manual.commentary||""} onChange={e=>setM("commentary",e.target.value)} placeholder="What's the overall story? Where is progress being made?"/>
          <div style={{marginTop:12}}><label style={S.lbl}>Your name</label><input style={S.inp} value={manual.dlName||""} onChange={e=>setM("dlName",e.target.value)} placeholder="Delivery Lead name"/></div>
        </div>
      )}

      {tab==="flow"&&(
        <div style={S.card}>
          <p style={S.sec}>Flow metrics</p>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <span style={{fontSize:13,fontWeight:700,color:ct?"#0F766E":"#64748B"}}>Cycle time</span>
            <span style={{background:ct?"#F0FDFA":"#F8FAFC",color:ct?"#0F766E":"#94A3B8",borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,border:`1px ${ct?"solid":"dashed"} ${ct?"#99F6E4":"#CBD5E1"}`}}>{ct?"LIVE · Jira":"NOT AVAILABLE"}</span>
          </div>
          {ct?(
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:16}}>
              {ct.trend.slice(-4).map((p,i)=>{const c=p.avgDays<=14?"#16A34A":p.avgDays<=25?"#D97706":"#DC2626";const bg=p.avgDays<=14?"#F0FDF4":p.avgDays<=25?"#FFFBEB":"#FEF2F2";return <div key={i} style={{background:bg,borderRadius:8,padding:"10px 14px",border:"1px solid #E2E8F0",textAlign:"center",minWidth:100}}><div style={{fontSize:22,fontWeight:700,color:c}}>{p.avgDays}d</div><div style={{fontSize:10,color:"#64748B",marginTop:2}}>{p.period.replace(/ 20\d\d/,"")}</div><div style={{fontSize:10,color:"#94A3B8"}}>{p.stories} stories</div></div>;})}
              <div style={{background:"#EFF6FF",borderRadius:8,padding:"10px 14px",border:"1px solid #BFDBFE",textAlign:"center",minWidth:100}}><div style={{fontSize:22,fontWeight:700,color:"#1E40AF"}}>{ct.avg}d</div><div style={{fontSize:10,color:"#1E40AF",marginTop:2}}>avg</div><div style={{fontSize:10,color:"#1E40AF"}}>p50: {ct.median}d</div></div>
            </div>
          ):<div style={{...S.empty,marginBottom:16}}>No cycle time data — requires resolved dates on Jira issues</div>}
          {[{l:"Velocity — last 3 sprints",slots:["Sprint -2: — pts committed / — delivered","Sprint -1: — pts committed / — delivered","Current: — pts committed / — delivered"]},
            {l:"Sprint reliability — last 3 sprints",slots:["Sprint -2: — / — = —%","Sprint -1: — / — = —%","Current: — / — = —%"]}
          ].map((sec,si)=>(
            <div key={si} style={{marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                <span style={{fontSize:13,fontWeight:700,color:"#64748B"}}>{sec.l}</span>
                <span style={{background:"#F8FAFC",color:"#94A3B8",borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,border:"1px dashed #CBD5E1"}}>NOT AVAILABLE</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>{sec.slots.map((s,i)=><div key={i} style={S.empty}>{s}</div>)}</div>
            </div>
          ))}
        </div>
      )}

      {tab==="deps"&&(
        <div style={S.card}>
          <p style={S.sec}>Dependencies</p>
          {deps.length>0?deps.map((d,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px",background:"#F5F3FF",borderRadius:8,marginBottom:8,border:"1px solid #DDD6FE",flexWrap:"wrap"}}>
              <div style={{background:"#EDE9FE",borderRadius:6,padding:"6px 10px",fontSize:12,minWidth:100}}><span style={{fontSize:10,fontWeight:700,color:"#7C3AED",display:"block"}}>{d.from}</span><span style={{color:"#1E293B"}}>{d.fromName}</span></div>
              <div style={{fontSize:11,color:"#7C3AED",fontWeight:600,flexShrink:0}}>depends on →</div>
              <div style={{background:d.status==="Done"?"#F0FDF4":"#EDE9FE",borderRadius:6,padding:"6px 10px",fontSize:12,flex:1,minWidth:100,border:d.status==="Done"?"1px solid #BBF7D0":"1px solid #C4B5FD"}}><span style={{fontSize:10,fontWeight:700,color:d.status==="Done"?"#16A34A":"#7C3AED",display:"block"}}>{d.to}</span><span style={{color:"#1E293B"}}>{d.toName||""}</span></div>
              <span style={{background:d.status==="Done"?"#F0FDF4":"#FFFBEB",color:d.status==="Done"?"#16A34A":"#D97706",border:`1px solid ${d.status==="Done"?"#BBF7D0":"#FDE68A"}`,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>{d.status}</span>
            </div>
          )):<div style={S.empty}>No dependency links found for {activeProject.key}</div>}
        </div>
      )}

      {tab==="capacity"&&(
        <div style={S.card}>
          <p style={S.sec}>Team capacity — manual entry</p>
          {members.length===0?<div style={S.empty}>No team members loaded yet.</div>:(
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}>
              {members.map((m,i)=>{const cap=manual.capacity?.[m]||"";const capNum=parseFloat(cap);const clr=!cap?"#94A3B8":capNum>=80?"#16A34A":capNum>=50?"#D97706":"#DC2626";const bg=!cap?"#F8FAFC":capNum>=80?"#F0FDF4":capNum>=50?"#FFFBEB":"#FEF2F2";return <div key={i} style={{background:bg,borderRadius:10,padding:"12px 14px",border:"1px solid #E2E8F0"}}><div style={{fontWeight:600,fontSize:13,marginBottom:8}}>{m.split(" ")[0]}</div><label style={{...S.lbl,marginBottom:4}}>% available</label><input type="number" min="0" max="100" style={{...S.inp,fontSize:16,fontWeight:700,color:clr,padding:"8px 10px"}} value={cap} placeholder="e.g. 80" onChange={e=>setM("capacity",{...manual.capacity,[m]:e.target.value})}/>{cap&&<div style={{height:5,background:"#E2E8F0",borderRadius:3,marginTop:8}}><div style={{width:`${Math.min(100,capNum)}%`,height:"100%",background:clr,borderRadius:3}}/></div>}</div>;})}
            </div>
          )}
        </div>
      )}

      {tab==="wellness"&&(
        <div style={S.card}>
          <p style={S.sec}>Team pulse survey</p>
          {members.length===0?<div style={S.empty}>No team members loaded yet.</div>:(
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
              {members.map((m,i)=>{const ws=wellness[m]||{};const curr=ws.current;const prev=ws.prev;const resp=(ws.responses||[]).length;const url=`${surveyBase}?member=${encodeURIComponent(m)}&project=${activeProject.key}`;return <div key={i} style={{background:"#F8FAFC",borderRadius:10,padding:14,border:"1px solid #E2E8F0"}}><div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}><strong style={{fontSize:13}}>{m.split(" ")[0]}</strong><span style={{fontSize:11,color:resp>=5?"#16A34A":"#94A3B8",fontWeight:600}}>{resp>=5?`${resp} responses`:`${resp}/5`}</span></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:5,marginBottom:10}}>{[["Happy","happiness"],["Spirit","teamSpirit"],["Belong","belonging"]].map(([lbl,dim])=>{const pv=prev?.[dim];const cv=resp>=5?curr?.[dim]:null;const arrow=cv!=null&&pv!=null?(cv-pv>0.2?"↑":cv-pv<-0.2?"↓":"→"):"";const ac=arrow==="↑"?"#16A34A":arrow==="↓"?"#DC2626":"#94A3B8";return <div key={dim} style={{background:"#fff",borderRadius:6,padding:"5px 7px",border:"1px solid #E2E8F0",textAlign:"center"}}><div style={{fontSize:9,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",marginBottom:2}}>{lbl}</div>{cv!=null?<div style={{fontSize:14,fontWeight:700,color:"#1E2761"}}>{cv}{arrow&&<span style={{color:ac,fontSize:11}}> {arrow}</span>}</div>:<div style={{fontSize:12,color:"#CBD5E1"}}>{resp>0?`${resp}/5`:"—"}</div>}</div>;})}          </div><div style={{display:"flex",gap:6,alignItems:"center"}}><input readOnly value={url} onClick={e=>e.target.select()} style={{...S.inp,fontSize:10,padding:"5px 8px",flex:1,background:"#fff"}}/><button onClick={()=>navigator.clipboard.writeText(url)} style={{...S.btn,padding:"5px 10px",background:"#F1F5F9",border:"1px solid #E2E8F0",fontSize:11}}>Copy</button></div></div>;})}
            </div>
          )}
        </div>
      )}

      {tab==="risks"&&(
        <div style={S.card}>
          <p style={S.sec}>Risks — manual entry</p>
          {(manual.risks||[]).map((r,ri)=>(
            <div key={ri} style={{display:"grid",gridTemplateColumns:"110px 1fr 1fr 1fr auto",gap:8,marginBottom:10,padding:12,background:"#F8FAFC",borderRadius:8,border:"1px solid #E2E8F0",alignItems:"start"}}>
              <div><label style={S.lbl}>Severity</label><select style={{...S.inp,padding:"8px 10px"}} value={r.severity||"Medium"} onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],severity:e.target.value};setM("risks",rs);}}><option>High</option><option>Medium</option><option>Low</option></select></div>
              <div><label style={S.lbl}>Risk</label><input style={S.inp} value={r.title||""} placeholder="Risk title" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],title:e.target.value};setM("risks",rs);}}/></div>
              <div><label style={S.lbl}>Description</label><input style={S.inp} value={r.description||""} placeholder="Impact" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],description:e.target.value};setM("risks",rs);}}/></div>
              <div><label style={S.lbl}>Mitigation · Owner</label><div style={{display:"flex",gap:4}}><input style={{...S.inp,flex:2}} value={r.mitigation||""} placeholder="Action" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],mitigation:e.target.value};setM("risks",rs);}}/><input style={{...S.inp,flex:1}} value={r.owner||""} placeholder="Owner" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],owner:e.target.value};setM("risks",rs);}}/></div></div>
              <button onClick={()=>setM("risks",(manual.risks||[]).filter((_,i)=>i!==ri))} style={{...S.btn,padding:"6px 10px",background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",marginTop:20,alignSelf:"end"}}>✕</button>
            </div>
          ))}
          <button style={{...S.btn,background:"#F1F5F9",border:"1px solid #E2E8F0",color:"#1E2761",marginTop:4}} onClick={()=>setM("risks",[...(manual.risks||[]),{severity:"Medium",title:"",description:"",mitigation:"",owner:""}])}>+ Add risk</button>
        </div>
      )}

      {tab==="asks"&&(
        <div style={S.card}>
          <p style={S.sec}>Where stakeholders can help</p>
          <label style={S.lbl}>One ask per line</label>
          <textarea style={{...S.ta,minHeight:120}} rows={5} value={manual.asks||""} onChange={e=>setM("asks",e.target.value)} placeholder={"Prioritisation: need leadership alignment\nCapacity: additional resources needed\nUnblock: vendor contract approval stalling Team X"}/>
        </div>
      )}

      {phase==="done"?(
        <div style={{...S.card,textAlign:"center",padding:"36px 24px"}}>
          <div style={{fontSize:36,marginBottom:10}}>✓</div>
          <h2 style={{color:"#16A34A",marginBottom:6,fontWeight:700}}>Report downloaded</h2>
          <p style={{fontSize:13,color:"#64748B",marginBottom:20}}>Check Downloads for {activeProject.key}-Report-*.pptx</p>
          <button style={{...S.btn,background:"#F1F5F9",color:"#0F172A",border:"1px solid #E2E8F0"}} onClick={()=>setPhase("idle")}>Generate another version</button>
        </div>
      ):(
        <div style={S.dark}>
          <p style={{fontSize:13,color:"#CADCFC",marginBottom:4,opacity:.8}}>Project: <strong style={{color:"#4FC3F7"}}>{activeProject.key} — {activeProject.name}</strong></p>
          <p style={{fontSize:11,color:"#4FC3F7",marginBottom:14}}>9 slides · Epics · Flow · Dependencies · Capacity · Wellness · Risks · Asks</p>
          {phase==="generating"?<div style={{color:"#CADCFC",fontSize:15}}>⏳ Building PowerPoint…</div>:
          <button style={{...S.btn,background:"#15803D",color:"#fff",fontSize:15,padding:"13px 28px"}} onClick={generate} disabled={!pptxOk}>
            {pptxOk?"⬇  Generate PowerPoint":"⏳  Loading library…"}
          </button>}
        </div>
      )}
    </div>
  );
}
