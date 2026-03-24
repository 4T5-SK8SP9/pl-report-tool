import { useState, useEffect, useCallback } from "react";

// ─── Real data from pandoradigital.atlassian.net / AMO (fetched 24 Mar 2026) ──
const LIVE = {
  productLine: "Agile Center of Excellence",
  project: "AMO",
  fetchedAt: "24 Mar 2026, 14:20",
  epics: [
    { key:"AMO-6207", name:"AI Adoption & Change Workstream",             assignee:"Giedre Yager",   pct:5,   rag:"red"   },
    { key:"AMO-6203", name:"Online PL Maturity uplift & coaching",        assignee:"Mikhail Sukhov", pct:20,  rag:"red"   },
    { key:"AMO-6194", name:"Dependency Management roll-out",              assignee:"Jacob Kingo",    pct:45,  rag:"amber" },
    { key:"AMO-6193", name:"CCC&B PL Maturity uplift & coaching",         assignee:"Giedre Yager",   pct:15,  rag:"red"   },
    { key:"AMO-6175", name:"AI-Enhanced Engineering Flow (Claude/Devin)", assignee:"Mikhail Sukhov", pct:10,  rag:"red"   },
    { key:"AMO-6149", name:"CCC Operations WoW & automation",             assignee:"Giedre Yager",   pct:30,  rag:"amber" },
    { key:"AMO-6110", name:"OMNI OPS PL Maturity uplift & coaching",      assignee:"Jacob Kingo",    pct:35,  rag:"amber" },
    { key:"AMO-6099", name:"MarTech PL coaching support",                 assignee:"Bo Malling",     pct:25,  rag:"amber" },
    { key:"AMO-6098", name:"CSM&F PL Maturity uplift & coaching",         assignee:"Jacob Kingo",    pct:40,  rag:"amber" },
    { key:"AMO-6093", name:"AI-powered PM in CCC (Claude)",               assignee:"Mikhail Sukhov", pct:20,  rag:"red"   },
    { key:"AMO-6086", name:"Retail Tech Maturity uplift & support",       assignee:"Jacob Kingo",    pct:35,  rag:"amber" },
    { key:"AMO-6085", name:"Foundational Platform Design",                assignee:"Giedre Yager",   pct:15,  rag:"red"   },
    { key:"AMO-6204", name:"Training — Microtraining concept",            assignee:"—",              pct:10,  rag:"red"   },
    { key:"AMO-6208", name:"Yearly AMA",                                  assignee:"—",              pct:0,   rag:"red"   },
    { key:"AMO-6224", name:"Data Strategy & Governance coaching",         assignee:"—",              pct:0,   rag:"red"   },
    { key:"AMO-6094", name:"AI automation across SDLC",                   assignee:"—",              pct:100, rag:"green" },
    { key:"AMO-6091", name:"Service Catalogue updates & launch",          assignee:"—",              pct:100, rag:"green" },
    { key:"AMO-6087", name:"Live Trainings Q1",                           assignee:"—",              pct:100, rag:"green" },
  ],
  // Real cycle time data
  cycleTime: {
    avg: 18, median: 14,
    trend: [
      { period:"W51-52 2025", stories:11, avgDays:9  },
      { period:"W3-4 2026",   stories:24, avgDays:17 },
      { period:"W5-6 2026",   stories:12, avgDays:13 },
      { period:"W7-8 2026",   stories:9,  avgDays:25 },
      { period:"W9-10 2026",  stories:10, avgDays:39 },
      { period:"W13-14 2026", stories:1,  avgDays:28 },
    ],
  },
  // Real throughput per person
  throughput: [
    { name:"Bo Malling",      stories:27, avgCycle:25 },
    { name:"Mikhail Sukhov",  stories:22, avgCycle:15 },
    { name:"Jacob Kingo",     stories:17, avgCycle:22 },
    { name:"Mikkel Lex Nissen",stories:15,avgCycle:17 },
    { name:"Giedre Yager",    stories:1,  avgCycle:28 },
  ],
  // Real dependency links
  dependencies: [
    { from:"AMO-6160", fromName:"Retail Tech - Jira Workflow",      to:"AMO-6158",   toName:"Dependency Mgmt Pilot",    status:"Done" },
    { from:"AMO-6160", fromName:"Retail Tech - Jira Workflow",      to:"AMO-6159",   toName:"Retail Tech OKR Training", status:"Done" },
    { from:"AMO-6160", fromName:"Retail Tech - Jira Workflow",      to:"PRTFL-3040", toName:"PRTFL issue",              status:"Done" },
    { from:"AMO-6158", fromName:"Dependency Management Pilot",      to:"AMO-6160",   toName:"Retail Tech Jira Workflow", status:"Done" },
    { from:"AMO-6158", fromName:"Dependency Management Pilot",      to:"PRTFL-3042", toName:"PRTFL issue",              status:"Done" },
  ],
};

// ─── RAG helpers ──────────────────────────────────────────────────────────────
const RAG = {
  green: { bg:"#F0FDF4", border:"#BBF7D0", text:"#16A34A", label:"On track"  },
  amber: { bg:"#FFFBEB", border:"#FDE68A", text:"#D97706", label:"At risk"   },
  red:   { bg:"#FEF2F2", border:"#FECACA", text:"#DC2626", label:"Off track" },
  gray:  { bg:"#F8FAFC", border:"#E2E8F0", text:"#64748B", label:"No data"   },
};
const ragFromPct = p => p == null ? "gray" : p >= 70 ? "green" : p >= 40 ? "amber" : "red";

// ─── Derived stats ────────────────────────────────────────────────────────────
const total      = LIVE.epics.length;
const done       = LIVE.epics.filter(e => e.pct >= 100).length;
const inProgress = LIVE.epics.filter(e => e.pct > 0 && e.pct < 100).length;
const notStarted = LIVE.epics.filter(e => e.pct === 0).length;
const overallPct = Math.round(LIVE.epics.reduce((s,e) => s + e.pct, 0) / total);

// ─── Storage ──────────────────────────────────────────────────────────────────
const SK = "pl-report:ac-coe-v2";
const loadW  = () => { try { return JSON.parse(localStorage.getItem(SK)||"{}"); } catch { return {}; }};
const saveW  = d  => { try { localStorage.setItem(SK,JSON.stringify(d)); } catch {} };
function addSurveyResp(member, scores) {
  const all = loadW();
  const t = all[member] || { responses:[], prev:null, current:null };
  const responses = [...(t.responses||[]), { ...scores, at:Date.now() }];
  const avg = k => { const vs=responses.map(r=>r[k]).filter(Boolean); return vs.length ? Math.round(vs.reduce((a,b)=>a+b,0)/vs.length*10)/10 : null; };
  all[member] = { ...t, responses, current:{ happiness:avg("happiness"), teamSpirit:avg("teamSpirit"), belonging:avg("belonging"), respondents:responses.length }};
  saveW(all); return responses.length;
}

// ─── PPTX builder ─────────────────────────────────────────────────────────────
async function buildPptx(manual, wellness, timestamp) {
  const pptx = new window.PptxGenJS();
  pptx.layout = "LAYOUT_16x9";
  const PL = "Agile Center of Excellence";
  const sh = () => ({ type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.08 });

  function hdr(s, title, sub, bg="1E2761") {
    s.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:10,h:.7,fill:{color:bg},line:{color:bg}});
    s.addText(title,{x:.4,y:0,w:7.8,h:.7,fontSize:19,color:"FFFFFF",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
    if (sub) s.addText(sub,{x:7.8,y:0,w:2,h:.7,fontSize:9,color:"CADCFC",fontFace:"Calibri",align:"right",valign:"middle",margin:0});
  }
  function pb(s, x, y, w, pct, clr) {
    s.addShape(pptx.shapes.RECTANGLE,{x,y,w,h:.13,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
    if (pct>0) s.addShape(pptx.shapes.RECTANGLE,{x,y,w:Math.max(.04,w*pct/100),h:.13,fill:{color:clr},line:{color:clr}});
  }
  function emptyBox(s, x, y, w, h, label) {
    s.addShape(pptx.shapes.RECTANGLE,{x,y,w,h,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
    s.addText(label,{x,y,w,h,fontSize:10,color:"94A3B8",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
  }

  // S1: Cover
  const s1=pptx.addSlide(); s1.background={color:"0D1440"};
  s1.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:.18,h:5.625,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
  s1.addShape(pptx.shapes.RECTANGLE,{x:0,y:4.7,w:10,h:.925,fill:{color:"1E2761"},line:{color:"1E2761"}});
  s1.addText("PANDORA AGILE CENTER OF EXCELLENCE",{x:.42,y:1.05,w:9,h:.44,fontSize:10,color:"4FC3F7",bold:true,charSpacing:4,fontFace:"Calibri",margin:0});
  s1.addText("AC CoE Report",{x:.42,y:1.55,w:9,h:1.0,fontSize:40,color:"FFFFFF",bold:true,fontFace:"Calibri",margin:0});
  s1.addText("Epics  ·  Flow metrics  ·  Dependencies  ·  Team health  ·  Risks",{x:.42,y:2.68,w:8.5,h:.44,fontSize:13,color:"CADCFC",fontFace:"Calibri",italic:true,margin:0});
  s1.addText(`Generated: ${timestamp}`,{x:.42,y:4.8,w:6,h:.28,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});
  s1.addText(`Prepared by: ${manual.dlName||"—"}`,{x:.42,y:5.1,w:6,h:.28,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});

  // S2: Portfolio snapshot
  const s2=pptx.addSlide(); s2.background={color:"F7F9FF"};
  hdr(s2,"Portfolio overview",`AMO · ${LIVE.fetchedAt}`);
  const stats=[{l:"Total",v:total,c:"1E2761",bg:"EFF6FF",bd:"BFDBFE"},{l:"Done",v:done,c:"16A34A",bg:"F0FDF4",bd:"BBF7D0"},{l:"In progress",v:inProgress,c:"D97706",bg:"FFFBEB",bd:"FDE68A"},{l:"Not started",v:notStarted,c:"DC2626",bg:"FEF2F2",bd:"FECACA"}];
  stats.forEach((st,i)=>{
    const x=.4+i*2.35;
    s2.addShape(pptx.shapes.RECTANGLE,{x,y:.86,w:2.1,h:1.2,fill:{color:st.bg},line:{color:st.bd},shadow:sh()});
    s2.addText(String(st.v),{x,y:.9,w:2.1,h:.68,fontSize:36,color:st.c,bold:true,fontFace:"Calibri",align:"center",valign:"middle",margin:0});
    s2.addText(st.l,{x,y:1.58,w:2.1,h:.32,fontSize:11,color:st.c,fontFace:"Calibri",align:"center",margin:0});
  });
  s2.addText(`Portfolio completion: ${overallPct}%`,{x:.4,y:2.22,w:5,h:.28,fontSize:12,color:"1E293B",bold:true,fontFace:"Calibri",margin:0});
  s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:2.55,w:9.2,h:.26,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
  s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:2.55,w:Math.max(.1,9.2*overallPct/100),h:.26,fill:{color:"1E2761"},line:{color:"1E2761"}});
  if (manual.commentary) {
    s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.0,w:9.2,h:2.1,fill:{color:"FFFFFF"},line:{color:"E2E8F0"},shadow:sh()});
    s2.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.0,w:.055,h:2.1,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
    s2.addText("Delivery Lead commentary",{x:.55,y:3.1,w:8.9,h:.24,fontSize:9,color:"8A9BC4",bold:true,fontFace:"Calibri",margin:0});
    s2.addText(manual.commentary,{x:.55,y:3.36,w:8.9,h:1.62,fontSize:12,color:"1E293B",fontFace:"Calibri",valign:"top",margin:0});
  }

  // S3: Active epics
  const s3=pptx.addSlide(); s3.background={color:"F7F9FF"};
  hdr(s3,"Epic progress — active","AMO","92400E");
  LIVE.epics.filter(e=>e.pct>0&&e.pct<100).slice(0,8).forEach((ep,ei)=>{
    const col=ei%2,row=Math.floor(ei/2),x=.4+col*4.75,y=.86+row*1.1;
    const rc=RAG[ep.rag];
    s3.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:.95,fill:{color:rc.bg.replace("#","")},line:{color:rc.border.replace("#","")},shadow:sh()});
    s3.addShape(pptx.shapes.RECTANGLE,{x,y,w:.055,h:.95,fill:{color:rc.text.replace("#","")},line:{color:rc.text.replace("#","")}});
    s3.addText(`${ep.key}  ·  ${ep.assignee}`,{x:x+.13,y:y+.07,w:4.2,h:.2,fontSize:9,color:rc.text.replace("#",""),bold:true,fontFace:"Calibri",margin:0});
    s3.addText(ep.name,{x:x+.13,y:y+.27,w:4.2,h:.26,fontSize:12,color:"1E293B",fontFace:"Calibri",margin:0});
    pb(s3,x+.13,y+.63,3.5,ep.pct,rc.text.replace("#",""));
    s3.addText(`${ep.pct}%`,{x:x+3.8,y:y+.55,w:.6,h:.24,fontSize:11,color:rc.text.replace("#",""),bold:true,align:"right",fontFace:"Calibri",margin:0});
  });

  // S4: Flow metrics — cycle time (LIVE) + velocity + sprint reliability (EMPTY)
  const s4=pptx.addSlide(); s4.background={color:"F7F9FF"};
  hdr(s4,"Flow metrics","AMO · Live data where available","0F766E");

  // Cycle time — LIVE
  s4.addText("Cycle time — live (days per story)",{x:.4,y:.86,w:5.5,h:.28,fontSize:11,color:"0F766E",bold:true,fontFace:"Calibri",margin:0});
  s4.addText("Source: Jira resolved dates",{x:.4,y:1.12,w:5.5,h:.2,fontSize:9,color:"94A3B8",fontFace:"Calibri",italic:true,margin:0});
  const ctBars=LIVE.cycleTime.trend.slice(-4);
  ctBars.forEach((p,pi)=>{
    const x=.4+pi*1.36, y=1.38;
    const barH=Math.min(1.6,(p.avgDays/50)*1.6);
    const clr=p.avgDays<=14?"16A34A":p.avgDays<=25?"D97706":"DC2626";
    s4.addShape(pptx.shapes.RECTANGLE,{x:x+.2,y:y+(1.6-barH),w:.8,h:barH,fill:{color:clr},line:{color:clr}});
    s4.addText(`${p.avgDays}d`,{x,y:y+(1.6-barH)-.28,w:1.2,h:.25,fontSize:10,color:clr,bold:true,align:"center",fontFace:"Calibri",margin:0});
    s4.addText(p.period.replace(" 2026","").replace(" 2025",""),{x,y:3.06,w:1.2,h:.22,fontSize:8,color:"64748B",align:"center",fontFace:"Calibri",margin:0});
    s4.addText(`${p.stories}✓`,{x,y:3.26,w:1.2,h:.2,fontSize:8,color:"94A3B8",align:"center",fontFace:"Calibri",margin:0});
  });
  s4.addShape(pptx.shapes.RECTANGLE,{x:.4,y:3.48,w:5.5,h:.22,fill:{color:"EFF6FF"},line:{color:"BFDBFE"}});
  s4.addText(`Avg: ${LIVE.cycleTime.avg}d  ·  Median: ${LIVE.cycleTime.median}d  ·  Target: ≤14d`,{x:.4,y:3.48,w:5.5,h:.22,fontSize:9,color:"1E40AF",fontFace:"Calibri",align:"center",valign:"middle",margin:0});

  // Velocity — EMPTY (Sprint-based PLs)
  s4.addText("Velocity — last 3 sprints",{x:6.1,y:.86,w:3.5,h:.28,fontSize:11,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
  s4.addText("Not applicable for Kanban projects",{x:6.1,y:1.12,w:3.5,h:.2,fontSize:9,color:"94A3B8",fontFace:"Calibri",italic:true,margin:0});
  [["Sprint -2",0],["Sprint -1",1],["Sprint 0 (current)",2]].forEach(([lbl],si)=>{
    const y=1.38+si*.7;
    s4.addShape(pptx.shapes.RECTANGLE,{x:6.1,y,w:3.5,h:.56,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
    s4.addText(lbl,{x:6.2,y,w:1.8,h:.56,fontSize:10,color:"94A3B8",fontFace:"Calibri",valign:"middle",margin:0});
    s4.addText("— pts  |  — committed  |  — completed",{x:6.2,y,w:3.2,h:.56,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",align:"right",margin:0});
  });

  // Sprint reliability — EMPTY
  s4.addText("Sprint reliability — last 3 sprints",{x:6.1,y:3.7,w:3.5,h:.26,fontSize:11,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
  [["Sprint -2"],["Sprint -1"],["Current"]].forEach(([lbl],si)=>{
    const y=4.02+si*.42;
    s4.addShape(pptx.shapes.RECTANGLE,{x:6.1,y,w:3.5,h:.34,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
    s4.addText(`${lbl}: — committed  /  — delivered  =  —%`,{x:6.18,y,w:3.3,h:.34,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",margin:0});
  });

  // S5: Dependencies
  const s5=pptx.addSlide(); s5.background={color:"F7F9FF"};
  hdr(s5,"Dependencies — 'depends on' links","AMO · Link type: Dependencies","7C3AED");
  LIVE.dependencies.forEach((d,di)=>{
    const y=.86+di*.9;
    s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.76,fill:{color:"F5F3FF"},line:{color:"DDD6FE"},shadow:sh()});
    s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:.055,h:.76,fill:{color:"7C3AED"},line:{color:"7C3AED"}});
    // From box
    s5.addShape(pptx.shapes.RECTANGLE,{x:.55,y:y+.1,w:3.6,h:.52,fill:{color:"EDE9FE"},line:{color:"C4B5FD"}});
    s5.addText(d.from,{x:.6,y:y+.1,w:.9,h:.52,fontSize:9,color:"7C3AED",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
    s5.addText(d.fromName,{x:1.5,y:y+.1,w:2.6,h:.52,fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    // Arrow
    s5.addText("depends on →",{x:4.25,y:y+.18,w:1.4,h:.3,fontSize:9,color:"7C3AED",italic:true,fontFace:"Calibri",align:"center",margin:0});
    // To box
    s5.addShape(pptx.shapes.RECTANGLE,{x:5.7,y:y+.1,w:3.7,h:.52,fill:{color:d.status==="Done"?"F0FDF4":"EDE9FE"},line:{color:d.status==="Done"?"BBF7D0":"C4B5FD"}});
    s5.addText(d.to,{x:5.75,y:y+.1,w:.9,h:.52,fontSize:9,color:d.status==="Done"?"16A34A":"7C3AED",bold:true,fontFace:"Calibri",valign:"middle",margin:0});
    s5.addText(d.toName,{x:6.65,y:y+.1,w:2.6,h:.52,fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    const stc=d.status==="Done"?{bg:"F0FDF4",border:"BBF7D0",text:"16A34A"}:{bg:"FFFBEB",border:"FDE68A",text:"D97706"};
    s5.addShape(pptx.shapes.RECTANGLE,{x:9.5,y:y+.16,w:.9,h:.36,fill:{color:stc.bg},line:{color:stc.border}});
    s5.addText(d.status,{x:9.5,y:y+.16,w:.9,h:.36,fontSize:8,color:stc.text,bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
  });
  // Empty slots for PLs that have active deps
  if (LIVE.dependencies.length < 5) {
    for (let i=LIVE.dependencies.length;i<3;i++) {
      const y=.86+i*.9;
      s5.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.76,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s5.addText("No active dependency link",{x:.4,y,w:9.2,h:.76,fontSize:10,color:"CBD5E1",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    }
  }

  // S6: Team throughput (LIVE) + Capacity (EMPTY)
  const s6=pptx.addSlide(); s6.background={color:"F7F9FF"};
  hdr(s6,"Team performance & capacity","AMO · 2026 YTD","0F766E");

  // Throughput — LIVE
  s6.addText("Throughput YTD (stories delivered)",{x:.4,y:.86,w:5.0,h:.28,fontSize:11,color:"0F766E",bold:true,fontFace:"Calibri",margin:0});
  s6.addText("Source: Jira resolved dates",{x:.4,y:1.1,w:5.0,h:.2,fontSize:9,color:"94A3B8",fontFace:"Calibri",italic:true,margin:0});
  LIVE.throughput.forEach((t,ti)=>{
    const y=1.38+ti*.58;
    const barW=Math.max(.2,4.2*(t.stories/30));
    const cyclClr=t.avgCycle<=14?"16A34A":t.avgCycle<=25?"D97706":"DC2626";
    s6.addText(t.name.split(" ")[0],{x:.4,y,w:1.1,h:.46,fontSize:11,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    s6.addShape(pptx.shapes.RECTANGLE,{x:1.55,y:y+.08,w:barW,h:.3,fill:{color:"1E2761"},line:{color:"1E2761"}});
    s6.addText(`${t.stories} stories`,{x:1.55+barW+.08,y:y+.06,w:1.0,h:.3,fontSize:10,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    s6.addShape(pptx.shapes.RECTANGLE,{x:4.0,y:y+.06,w:1.25,h:.28,fill:{color:"F8FAFC"},line:{color:"E2E8F0"}});
    s6.addText(`avg ${t.avgCycle}d`,{x:4.0,y:y+.06,w:1.25,h:.28,fontSize:9,color:cyclClr,bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
  });

  // Capacity — EMPTY structured fields
  s6.addShape(pptx.shapes.RECTANGLE,{x:5.6,y:.82,w:4.0,h:4.66,fill:{color:"FFFFFF"},line:{color:"E2E8F0"},shadow:sh()});
  s6.addShape(pptx.shapes.RECTANGLE,{x:5.6,y:.82,w:4.0,h:.055,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
  s6.addText("Team capacity — enter manually",{x:5.7,y:.88,w:3.8,h:.26,fontSize:10,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
  s6.addText("% available this sprint / period",{x:5.7,y:1.12,w:3.8,h:.2,fontSize:9,color:"94A3B8",italic:true,fontFace:"Calibri",margin:0});
  const members=["Bo Malling","Mikhail Sukhov","Jacob Kingo","Mikkel Lex Nissen","Giedre Yager"];
  const capManual=manual.capacity||{};
  members.forEach((m,mi)=>{
    const y=1.42+mi*.72;
    const cap=capManual[m]||null;
    const clr=cap==null?"CBD5E1":cap>=80?"16A34A":cap>=50?"D97706":"DC2626";
    const bg=cap==null?"F8FAFC":cap>=80?"F0FDF4":cap>=50?"FFFBEB":"FEF2F2";
    s6.addShape(pptx.shapes.RECTANGLE,{x:5.7,y,w:3.7,h:.58,fill:{color:bg},line:{color:cap==null?"CBD5E1":"E2E8F0"}});
    s6.addText(m.split(" ")[0],{x:5.78,y,w:1.4,h:.58,fontSize:11,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    if (cap) {
      s6.addShape(pptx.shapes.RECTANGLE,{x:7.1,y:y+.13,w:1.8,h:.3,fill:{color:"E2E8F0"},line:{color:"E2E8F0"}});
      s6.addShape(pptx.shapes.RECTANGLE,{x:7.1,y:y+.13,w:Math.max(.04,1.8*cap/100),h:.3,fill:{color:clr},line:{color:clr}});
      s6.addText(`${cap}%`,{x:9.0,y,w:.35,h:.58,fontSize:10,color:clr,bold:true,fontFace:"Calibri",valign:"middle",margin:0});
    } else {
      s6.addText("— enter capacity %",{x:7.1,y,w:2.2,h:.58,fontSize:9,color:"CBD5E1",italic:true,fontFace:"Calibri",valign:"middle",margin:0});
    }
  });

  // S7: Team wellness + pulse survey
  const s7=pptx.addSlide(); s7.background={color:"F7F9FF"};
  hdr(s7,"Team wellness — pulse survey","AC CoE · 5+ responses to reveal","0F766E");
  const wCols=["Happiness","Team spirit","Belonging"];
  const wXs=[2.3,4.2,6.1]; const wW=1.65;
  s7.addText("Team member",{x:.4,y:.82,w:1.8,h:.28,fontSize:9,color:"64748B",bold:true,fontFace:"Calibri",margin:0});
  wCols.forEach((c,ci)=>s7.addText(c,{x:wXs[ci],y:.82,w:wW,h:.28,fontSize:9,color:"64748B",bold:true,align:"center",fontFace:"Calibri",margin:0}));
  s7.addText("Overall",{x:8.1,y:.82,w:1.5,h:.28,fontSize:9,color:"64748B",bold:true,align:"center",fontFace:"Calibri",margin:0});

  members.forEach((m,mi)=>{
    const y=1.18+mi*.58;
    const ws=wellness[m]||{}; const curr=ws.current; const prev=ws.prev;
    const resp=(ws.responses||[]).length; const bg=mi%2===0?"FFFFFF":"F8FAFC";
    s7.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.46,fill:{color:bg},line:{color:"E2E8F0"}});
    s7.addText(m.split(" ")[0],{x:.48,y,w:1.7,h:.46,fontSize:11,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
    const dims=["happiness","teamSpirit","belonging"];
    dims.forEach((dim,di)=>{
      const show=resp>=5&&curr!=null;
      const cv=show?curr[dim]:null;
      const pv=prev?.[dim];
      const r2=cv!=null?ragFromPct(cv*20):"gray"; const rc=RAG[r2];
      const arrow=cv!=null&&pv!=null?(cv-pv>0.2?"↑":cv-pv<-0.2?"↓":"→"):"";
      const arrowClr=arrow==="↑"?"16A34A":arrow==="↓"?"DC2626":"94A3B8";
      s7.addShape(pptx.shapes.RECTANGLE,{x:wXs[di]+.04,y:y+.04,w:wW-.08,h:.38,fill:{color:rc.bg.replace("#","")},line:{color:rc.border.replace("#","")}});
      s7.addText(cv!=null?`${cv}`:resp>0?`${resp}/5`:"—",{x:wXs[di]+.04,y:y+.04,w:wW-.08,h:.38,fontSize:10,color:cv!=null?rc.text.replace("#",""):"CBD5E1",bold:cv!=null,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
      if (arrow) s7.addText(arrow,{x:wXs[di]+wW-.38,y:y+.04,w:.3,h:.38,fontSize:11,color:arrowClr,bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    });
    const allVals=[curr?.happiness,curr?.teamSpirit,curr?.belonging].filter(Boolean);
    const overallV=allVals.length===3?Math.round(allVals.reduce((a,b)=>a+b,0)/3*10)/10:null;
    const or=overallV!=null?ragFromPct(overallV*20):"gray"; const orc=RAG[or];
    s7.addShape(pptx.shapes.RECTANGLE,{x:8.13,y:y+.04,w:1.44,h:.38,fill:{color:orc.bg.replace("#","")},line:{color:orc.border.replace("#","")}});
    s7.addText(overallV!=null?`${overallV}  ${orc.label}`:resp>0?`${resp}/5 responses`:"Awaiting responses",{x:8.13,y:y+.04,w:1.44,h:.38,fontSize:9,color:orc.text.replace("#",""),bold:overallV!=null,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
  });
  s7.addText("Scores only visible once 5+ team members respond  ·  prev → current trend shown after second survey  ·  ↑ improving  ↓ declining  → stable",{x:.4,y:4.12,w:9.2,h:.22,fontSize:8,color:"94A3B8",italic:true,fontFace:"Calibri",margin:0});

  // S8: Risks (EMPTY structured template)
  const s8=pptx.addSlide(); s8.background={color:"F7F9FF"};
  hdr(s8,"Risks","AMO · Enter manually — no Risk issue type in this project","B91C1C");
  const risks=manual.risks||[];
  const severityClr={High:{bg:"FEF2F2",bd:"FECACA",tx:"DC2626"},Medium:{bg:"FFFBEB",bd:"FDE68A",tx:"D97706"},Low:{bg:"F0FDF4",bd:"BBF7D0",tx:"16A34A"}};
  for (let ri=0;ri<6;ri++) {
    const col=ri%2,row=Math.floor(ri/2),x=.4+col*4.75,y=.86+row*1.55;
    const risk=risks[ri];
    if (risk) {
      const sc=severityClr[risk.severity]||severityClr.Medium;
      s8.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:1.38,fill:{color:sc.bg},line:{color:sc.bd},shadow:sh()});
      s8.addShape(pptx.shapes.RECTANGLE,{x,y,w:.055,h:1.38,fill:{color:sc.tx},line:{color:sc.tx}});
      s8.addShape(pptx.shapes.RECTANGLE,{x:x+.1,y:y+.07,w:.8,h:.24,fill:{color:sc.bg},line:{color:sc.bd}});
      s8.addText(risk.severity||"Medium",{x:x+.1,y:y+.07,w:.8,h:.24,fontSize:8,color:sc.tx,bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
      s8.addText(risk.title||"",{x:x+.13,y:y+.34,w:4.2,h:.3,fontSize:12,color:"1E293B",bold:true,fontFace:"Calibri",margin:0});
      s8.addText(risk.description||"",{x:x+.13,y:y+.64,w:4.2,h:.3,fontSize:10,color:"475569",fontFace:"Calibri",margin:0});
      s8.addShape(pptx.shapes.RECTANGLE,{x:x+.13,y:y+1.0,w:4.2,h:.28,fill:{color:"FFFBEB"},line:{color:"FDE68A"}});
      s8.addText(`Mitigation: ${risk.mitigation||"—"}  ·  Owner: ${risk.owner||"—"}`,{x:x+.2,y:y+1.0,w:4.0,h:.28,fontSize:9,color:"D97706",fontFace:"Calibri",valign:"middle",margin:0});
    } else {
      s8.addShape(pptx.shapes.RECTANGLE,{x,y,w:4.5,h:1.38,fill:{color:"F8FAFC"},line:{color:"CBD5E1",dashType:"dash"}});
      s8.addText("No risk entered",{x,y,w:4.5,h:.5,fontSize:10,color:"CBD5E1",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
      s8.addText("Add via report tool → Risks section",{x,y:y+.5,w:4.5,h:.5,fontSize:9,color:"CBD5E1",italic:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    }
  }

  // S9: Where you can help
  const s9=pptx.addSlide(); s9.background={color:"F7F9FF"};
  hdr(s9,"Where you can help","AC CoE");
  const asks=(manual.asks||"").split("\n").filter(Boolean).slice(0,5);
  asks.forEach((a,ai)=>{
    const y=.86+ai*.88;
    s9.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:9.2,h:.72,fill:{color:"EFF6FF"},line:{color:"BFDBFE"},shadow:sh()});
    s9.addShape(pptx.shapes.RECTANGLE,{x:.4,y,w:.055,h:.72,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
    s9.addText(`${ai+1}`,{x:.55,y,w:.4,h:.72,fontSize:16,color:"4FC3F7",bold:true,align:"center",valign:"middle",fontFace:"Calibri",margin:0});
    s9.addText(a,{x:1.1,y,w:8.3,h:.72,fontSize:13,color:"1E293B",fontFace:"Calibri",valign:"middle",margin:0});
  });
  if (!asks.length) s9.addText("No specific asks this period.",{x:.4,y:1.6,w:9.2,h:.5,fontSize:14,color:"94A3B8",italic:true,fontFace:"Calibri",margin:0});

  // Closing
  const sE=pptx.addSlide(); sE.background={color:"0D1440"};
  sE.addShape(pptx.shapes.RECTANGLE,{x:0,y:0,w:.18,h:5.625,fill:{color:"4FC3F7"},line:{color:"4FC3F7"}});
  sE.addShape(pptx.shapes.RECTANGLE,{x:0,y:4.7,w:10,h:.925,fill:{color:"1E2761"},line:{color:"1E2761"}});
  sE.addText("Thank you",{x:.42,y:1.6,w:9,h:1.0,fontSize:42,color:"FFFFFF",bold:true,fontFace:"Calibri",margin:0});
  sE.addText("Questions & discussion",{x:.42,y:2.7,w:9,h:.5,fontSize:18,color:"CADCFC",italic:true,fontFace:"Calibri",margin:0});
  sE.addText(`AC CoE  ·  ${manual.dlName||"Delivery Lead"}`,{x:.42,y:3.5,w:9,h:.34,fontSize:12,color:"4FC3F7",fontFace:"Calibri",margin:0});
  sE.addText(`Generated: ${timestamp}  ·  Jira data: ${LIVE.fetchedAt}`,{x:.42,y:4.8,w:8,h:.3,fontSize:10,color:"8A9BC4",fontFace:"Calibri",margin:0});

  return pptx.writeFile({fileName:`AC-CoE-Report-v2-${timestamp.replace(/[: ,/]/g,"-")}.pptx`});
}

// ─── Survey view ──────────────────────────────────────────────────────────────
function SurveyView({ member }) {
  const [scores, setScores] = useState({happiness:0,teamSpirit:0,belonging:0});
  const [done, setDone] = useState(false);
  const [count, setCount] = useState(null);
  const allDone = Object.values(scores).every(v=>v>0);
  const QS=[
    {key:"happiness",  q:"How is your overall happiness at work right now?"},
    {key:"teamSpirit", q:"How would you rate the team spirit in AC CoE?"},
    {key:"belonging",  q:"How strong is your sense of belonging in the team?"},
  ];
  if (done) return (
    <div style={{minHeight:"100vh",background:"#0D1440",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <div style={{textAlign:"center",padding:40}}>
        <div style={{fontSize:56,color:"#F59E0B",marginBottom:16}}>★</div>
        <h2 style={{color:"#fff",fontSize:24,marginBottom:8}}>Thank you!</h2>
        <p style={{color:"#CADCFC",opacity:.8,fontSize:15}}>Your response has been recorded anonymously.</p>
        {count!=null&&count<5&&<p style={{color:"rgba(255,255,255,.5)",marginTop:10,fontSize:13}}>{5-count} more responses needed before scores are visible.</p>}
      </div>
    </div>
  );
  return (
    <div style={{minHeight:"100vh",background:"#0D1440",fontFamily:"'Segoe UI',system-ui,sans-serif",padding:"0 0 40px"}}>
      <div style={{maxWidth:480,margin:"0 auto"}}>
        <div style={{padding:"32px 24px 20px",textAlign:"center"}}>
          <p style={{fontSize:11,color:"#4FC3F7",fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",marginBottom:8}}>AC CoE · Pulse Survey</p>
          <h1 style={{color:"#fff",fontSize:22,marginBottom:4}}>{member}</h1>
          <p style={{color:"#CADCFC",opacity:.8,fontSize:14}}>Anonymous · 3 questions · 30 seconds</p>
        </div>
        {QS.map(({key,q})=>(
          <div key={key} style={{background:"#fff",borderRadius:14,padding:"22px 20px",margin:"0 16px 12px",boxShadow:"0 2px 8px rgba(0,0,0,.15)"}}>
            <p style={{fontSize:17,fontWeight:600,color:"#0F172A",lineHeight:1.4,marginBottom:18}}>{q}</p>
            <div style={{display:"flex",justifyContent:"center",gap:12}}>
              {[1,2,3,4,5].map(n=>(
                <button key={n} onClick={()=>setScores(s=>({...s,[key]:n}))} style={{background:"none",border:"none",cursor:"pointer",fontSize:36,color:n<=scores[key]?"#F59E0B":"#CBD5E1",padding:4,lineHeight:1,WebkitTapHighlightColor:"transparent"}} type="button">
                  {n<=scores[key]?"★":"☆"}
                </button>
              ))}
            </div>
            {scores[key]>0&&<p style={{textAlign:"center",marginTop:10,fontSize:13,color:"#64748B"}}>
              {["","Not great","Below average","Okay","Good","Excellent"][scores[key]]}
            </p>}
          </div>
        ))}
        <button onClick={()=>{const n=addSurveyResp(member,scores);setCount(n);setDone(true);}} disabled={!allDone}
          style={{display:"block",width:"calc(100% - 32px)",margin:"4px 16px 0",padding:16,background:allDone?"#1E2761":"#334155",color:"#fff",border:"none",borderRadius:12,fontSize:16,fontWeight:700,cursor:allDone?"pointer":"not-allowed",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
          {allDone?"Submit anonymously":"Please answer all 3 questions"}
        </button>
        <p style={{textAlign:"center",marginTop:12,fontSize:12,color:"rgba(255,255,255,.35)",padding:"0 16px"}}>Individual scores never shown. Results visible once 5+ respond.</p>
      </div>
    </div>
  );
}

// ─── Main app ─────────────────────────────────────────────────────────────────
export default function ACCoEReportV2() {
  const params = new URLSearchParams(typeof window!=="undefined"?window.location.search:"");
  const surveyMember = params.get("member");
  if (surveyMember) return <SurveyView member={surveyMember} />;

  const [manual, setManual] = useState({dlName:"Jacob Kingo",commentary:"",asks:"",capacity:{},risks:[]});
  const setM = (k,v) => setManual(m=>({...m,[k]:v}));
  const [wellness, setWellness] = useState({});
  const [pptxOk, setPptxOk] = useState(false);
  const [phase, setPhase] = useState("idle");
  const [activeSection, setActiveSection] = useState("overview");

  useEffect(()=>{
    if(window.PptxGenJS){setPptxOk(true);return;}
    const sc=document.createElement("script");
    sc.src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";
    sc.onload=()=>setPptxOk(true);
    document.head.appendChild(sc);
  },[]);
  useEffect(()=>{setWellness(loadW());},[]);
  useEffect(()=>{const t=setInterval(()=>setWellness(loadW()),8000);return()=>clearInterval(t);},[]);

  async function generate() {
    if(!pptxOk)return;
    setPhase("generating");
    const ts=new Date().toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
    await buildPptx(manual,wellness,ts);
    setPhase("done");
  }

  const S={
    page:{fontFamily:"'Segoe UI',system-ui,sans-serif",background:"#F1F5F9",minHeight:"100vh",padding:16,maxWidth:980,margin:"0 auto"},
    hdr:{background:"#0D1440",borderRadius:12,padding:"20px 24px 16px",marginBottom:12},
    card:{background:"#fff",border:"1px solid #E2E8F0",borderRadius:12,padding:"18px 20px",marginBottom:12,boxShadow:"0 1px 3px rgba(0,0,0,.06)"},
    dark:{background:"#0D1440",borderRadius:12,padding:"20px",marginBottom:12,textAlign:"center"},
    sec:{fontSize:11,fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:".07em",marginBottom:12},
    label:{fontSize:11,fontWeight:700,color:"#64748B",textTransform:"uppercase",letterSpacing:".05em",display:"block",marginBottom:4},
    input:{width:"100%",border:"1px solid #E2E8F0",borderRadius:7,padding:"9px 12px",fontSize:14,background:"#fff",color:"#0F172A",outline:"none",boxSizing:"border-box",fontFamily:"'Segoe UI',system-ui,sans-serif"},
    ta:{width:"100%",border:"1px solid #E2E8F0",borderRadius:7,padding:"9px 12px",fontSize:13,background:"#fff",color:"#0F172A",outline:"none",resize:"vertical",fontFamily:"'Segoe UI',system-ui,sans-serif",boxSizing:"border-box"},
    btn:{display:"inline-flex",alignItems:"center",gap:7,padding:"10px 20px",border:"none",borderRadius:8,fontSize:14,fontWeight:600,cursor:"pointer",fontFamily:"'Segoe UI',system-ui,sans-serif"},
    emptyField:{background:"#F8FAFC",border:"1px dashed #CBD5E1",borderRadius:7,padding:"10px 14px",fontSize:13,color:"#94A3B8",fontStyle:"italic",textAlign:"center"},
    navBtn:(active)=>({padding:"8px 14px",border:"none",borderRadius:7,fontSize:12,fontWeight:600,cursor:"pointer",background:active?"#1E2761":"#F1F5F9",color:active?"#fff":"#64748B",fontFamily:"'Segoe UI',system-ui,sans-serif"}),
  };

  function ragBadge(rag) {
    const rc=RAG[rag]||RAG.gray;
    return <span style={{background:rc.bg,color:rc.text,border:`1px solid ${rc.border}`,borderRadius:4,padding:"2px 7px",fontSize:10,fontWeight:700}}>{rc.label}</span>;
  }

  const surveyBase=typeof window!=="undefined"?`${window.location.origin}${window.location.pathname}`:"";
  const members=["Bo Malling","Mikhail Sukhov","Jacob Kingo","Mikkel Lex Nissen","Giedre Yager"];

  const SECTIONS=[
    {id:"overview",   label:"Portfolio"},
    {id:"flow",       label:"Flow metrics"},
    {id:"deps",       label:"Dependencies"},
    {id:"capacity",   label:"Capacity"},
    {id:"wellness",   label:"Wellness"},
    {id:"risks",      label:"Risks"},
    {id:"asks",       label:"Stakeholder asks"},
  ];

  return (
    <div style={S.page}>
      {/* Header */}
      <div style={S.hdr}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:8}}>
          <div>
            <div style={{fontSize:10,color:"#4FC3F7",fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",marginBottom:4}}>PACE · PL Report Tool v0.2</div>
            <div style={{fontSize:22,fontWeight:700,color:"#fff",marginBottom:2}}>Agile Center of Excellence</div>
            <div style={{fontSize:12,color:"#CADCFC",opacity:.8}}>AMO · {LIVE.fetchedAt} · {LIVE.epics.length} epics · {LIVE.throughput.reduce((s,t)=>s+t.stories,0)} stories delivered YTD</div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexShrink:0}}>
            <span style={{background:"#1E3A5F",color:"#4FC3F7",borderRadius:4,padding:"3px 9px",fontSize:11,fontWeight:700}}>LIVE JIRA DATA</span>
            <span style={{background:"#1A3025",color:"#4ADE80",borderRadius:4,padding:"3px 9px",fontSize:11,fontWeight:700}}>v0.2</span>
          </div>
        </div>
      </div>

      {/* Nav */}
      <div style={{display:"flex",gap:6,flexWrap:"wrap",marginBottom:12}}>
        {SECTIONS.map(s=>(
          <button key={s.id} style={S.navBtn(activeSection===s.id)} onClick={()=>setActiveSection(s.id)}>{s.label}</button>
        ))}
      </div>

      {/* Overview */}
      {activeSection==="overview" && (
        <div style={S.card}>
          <p style={S.sec}>Portfolio snapshot — live from Jira</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10,marginBottom:14}}>
            {[{l:"Total",v:total,c:"#1E2761",bg:"#EFF6FF"},{l:"Done",v:done,c:"#16A34A",bg:"#F0FDF4"},{l:"In progress",v:inProgress,c:"#D97706",bg:"#FFFBEB"},{l:"Not started",v:notStarted,c:"#DC2626",bg:"#FEF2F2"}].map((s,i)=>(
              <div key={i} style={{background:s.bg,borderRadius:10,padding:"12px 8px",textAlign:"center",border:"1px solid #E2E8F0"}}>
                <div style={{fontSize:28,fontWeight:700,color:s.c}}>{s.v}</div>
                <div style={{fontSize:11,color:"#64748B",marginTop:2}}>{s.l}</div>
              </div>
            ))}
          </div>
          <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Overall completion — {overallPct}%</div>
          <div style={{height:10,background:"#E2E8F0",borderRadius:5,marginBottom:14}}>
            <div style={{width:`${overallPct}%`,height:"100%",background:"#1E2761",borderRadius:5}}/>
          </div>
          <label style={S.label}>Portfolio commentary</label>
          <textarea style={{...S.ta,minHeight:72}} rows={3} value={manual.commentary} onChange={e=>setM("commentary",e.target.value)}
            placeholder="What's the overall story? Where is the AC CoE making progress? What is the strategic risk?"/>
          <div style={{marginTop:14}}>
            <label style={S.label}>Your name</label>
            <input style={S.input} value={manual.dlName} onChange={e=>setM("dlName",e.target.value)}/>
          </div>
        </div>
      )}

      {/* Flow metrics */}
      {activeSection==="flow" && (
        <div style={S.card}>
          <p style={S.sec}>Flow metrics</p>
          {/* Cycle time — live */}
          <div style={{marginBottom:18}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <span style={{fontSize:13,fontWeight:700,color:"#0F766E"}}>Cycle time</span>
              <span style={{background:"#F0FDFA",color:"#0F766E",borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,border:"1px solid #99F6E4"}}>LIVE · Jira</span>
            </div>
            <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
              {LIVE.cycleTime.trend.slice(-4).map((p,i)=>{
                const c=p.avgDays<=14?"#16A34A":p.avgDays<=25?"#D97706":"#DC2626";
                const bg=p.avgDays<=14?"#F0FDF4":p.avgDays<=25?"#FFFBEB":"#FEF2F2";
                return <div key={i} style={{background:bg,borderRadius:8,padding:"10px 14px",border:"1px solid #E2E8F0",textAlign:"center",minWidth:110}}>
                  <div style={{fontSize:22,fontWeight:700,color:c}}>{p.avgDays}d</div>
                  <div style={{fontSize:10,color:"#64748B",marginTop:2}}>{p.period}</div>
                  <div style={{fontSize:10,color:"#94A3B8"}}>{p.stories} stories</div>
                </div>;
              })}
              <div style={{background:"#EFF6FF",borderRadius:8,padding:"10px 14px",border:"1px solid #BFDBFE",textAlign:"center",minWidth:110}}>
                <div style={{fontSize:22,fontWeight:700,color:"#1E40AF"}}>{LIVE.cycleTime.avg}d</div>
                <div style={{fontSize:10,color:"#1E40AF",marginTop:2}}>avg overall</div>
                <div style={{fontSize:10,color:"#1E40AF"}}>p50: {LIVE.cycleTime.median}d</div>
              </div>
            </div>
          </div>

          {/* Velocity — empty */}
          <div style={{marginBottom:18}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <span style={{fontSize:13,fontWeight:700,color:"#64748B"}}>Velocity — last 3 sprints</span>
              <span style={{background:"#F8FAFC",color:"#94A3B8",borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,border:"1px dashed #CBD5E1"}}>NOT AVAILABLE · Kanban project</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
              {["Sprint -2","Sprint -1","Current sprint"].map((lbl,i)=>(
                <div key={i} style={S.emptyField}>
                  <div style={{fontSize:11,fontWeight:700,color:"#94A3B8",marginBottom:4}}>{lbl}</div>
                  <div>— pts committed · — pts delivered</div>
                </div>
              ))}
            </div>
          </div>

          {/* Sprint reliability — empty */}
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
              <span style={{fontSize:13,fontWeight:700,color:"#64748B"}}>Sprint reliability — last 3 sprints</span>
              <span style={{background:"#F8FAFC",color:"#94A3B8",borderRadius:4,padding:"1px 7px",fontSize:10,fontWeight:700,border:"1px dashed #CBD5E1"}}>NOT AVAILABLE · Kanban project</span>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8}}>
              {["Sprint -2","Sprint -1","Current sprint"].map((lbl,i)=>(
                <div key={i} style={S.emptyField}>
                  <div style={{fontSize:11,fontWeight:700,color:"#94A3B8",marginBottom:4}}>{lbl}</div>
                  <div>— committed / — delivered = —%</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Dependencies */}
      {activeSection==="deps" && (
        <div style={S.card}>
          <p style={S.sec}>Dependencies — "depends on" links · live from Jira</p>
          {LIVE.dependencies.map((d,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"10px 12px",background:"#F5F3FF",borderRadius:8,marginBottom:8,border:"1px solid #DDD6FE",flexWrap:"wrap"}}>
              <div style={{background:"#EDE9FE",borderRadius:6,padding:"6px 10px",fontSize:12,minWidth:120}}>
                <span style={{fontSize:10,fontWeight:700,color:"#7C3AED",display:"block"}}>{d.from}</span>
                <span style={{color:"#1E293B"}}>{d.fromName}</span>
              </div>
              <div style={{fontSize:11,color:"#7C3AED",fontWeight:600,flexShrink:0}}>depends on →</div>
              <div style={{background:d.status==="Done"?"#F0FDF4":"#EDE9FE",borderRadius:6,padding:"6px 10px",fontSize:12,flex:1,minWidth:120,border:d.status==="Done"?"1px solid #BBF7D0":"1px solid #C4B5FD"}}>
                <span style={{fontSize:10,fontWeight:700,color:d.status==="Done"?"#16A34A":"#7C3AED",display:"block"}}>{d.to}</span>
                <span style={{color:"#1E293B"}}>{d.toName}</span>
              </div>
              <span style={{background:d.status==="Done"?"#F0FDF4":"#FFFBEB",color:d.status==="Done"?"#16A34A":"#D97706",border:`1px solid ${d.status==="Done"?"#BBF7D0":"#FDE68A"}`,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700,flexShrink:0}}>{d.status}</span>
            </div>
          ))}
          <div style={{...S.emptyField,marginTop:8}}>Active unresolved dependency links will appear here when detected in Jira</div>
        </div>
      )}

      {/* Capacity */}
      {activeSection==="capacity" && (
        <div style={S.card}>
          <p style={S.sec}>Team capacity — enter manually</p>
          <p style={{fontSize:13,color:"#64748B",marginBottom:14,marginTop:-8}}>Not stored in Jira. Enter current available capacity % per person for this sprint/period.</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:10}}>
            {members.map((m,i)=>{
              const cap=manual.capacity[m]||"";
              const capNum=parseFloat(cap);
              const clr=!cap?"#94A3B8":capNum>=80?"#16A34A":capNum>=50?"#D97706":"#DC2626";
              const bg=!cap?"#F8FAFC":capNum>=80?"#F0FDF4":capNum>=50?"#FFFBEB":"#FEF2F2";
              return (
                <div key={i} style={{background:bg,borderRadius:10,padding:"12px 14px",border:"1px solid #E2E8F0"}}>
                  <div style={{fontWeight:600,fontSize:13,marginBottom:8}}>{m.split(" ")[0]}</div>
                  <label style={{...S.label,marginBottom:4}}>% available</label>
                  <input type="number" min="0" max="100" style={{...S.input,fontSize:16,fontWeight:700,color:clr,padding:"8px 10px"}}
                    value={cap} placeholder="e.g. 80"
                    onChange={e=>setM("capacity",{...manual.capacity,[m]:e.target.value})}/>
                  {cap&&<div style={{height:5,background:"#E2E8F0",borderRadius:3,marginTop:8}}>
                    <div style={{width:`${Math.min(100,capNum)}%`,height:"100%",background:clr,borderRadius:3}}/>
                  </div>}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Wellness */}
      {activeSection==="wellness" && (
        <div style={S.card}>
          <p style={S.sec}>Team pulse survey</p>
          <p style={{fontSize:13,color:"#64748B",marginBottom:14,marginTop:-8}}>Share each person's link. Scores appear once 5+ respond.</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
            {members.map((m,i)=>{
              const ws=wellness[m]||{}; const curr=ws.current; const prev=ws.prev; const resp=(ws.responses||[]).length;
              const url=`${surveyBase}?member=${encodeURIComponent(m)}`;
              return (
                <div key={i} style={{background:"#F8FAFC",borderRadius:10,padding:14,border:"1px solid #E2E8F0"}}>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:8}}>
                    <strong style={{fontSize:13}}>{m.split(" ")[0]}</strong>
                    <span style={{fontSize:11,color:resp>=5?"#16A34A":"#94A3B8",fontWeight:600}}>{resp>=5?`${resp} responses`:`${resp}/5`}</span>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:5,marginBottom:10}}>
                    {[["Happy","happiness"],["Spirit","teamSpirit"],["Belong","belonging"]].map(([lbl,dim])=>{
                      const pv=prev?.[dim]; const cv=resp>=5?curr?.[dim]:null;
                      const arrow=cv!=null&&pv!=null?(cv-pv>0.2?"↑":cv-pv<-0.2?"↓":"→"):"";
                      const ac=arrow==="↑"?"#16A34A":arrow==="↓"?"#DC2626":"#94A3B8";
                      return <div key={dim} style={{background:"#fff",borderRadius:6,padding:"5px 7px",border:"1px solid #E2E8F0",textAlign:"center"}}>
                        <div style={{fontSize:9,fontWeight:700,color:"#94A3B8",textTransform:"uppercase",marginBottom:2}}>{lbl}</div>
                        {cv!=null?<div style={{fontSize:14,fontWeight:700,color:"#1E2761"}}>{cv}{arrow&&<span style={{color:ac,fontSize:11}}> {arrow}</span>}</div>
                          :<div style={{fontSize:12,color:"#CBD5E1"}}>{resp>0?`${resp}/5`:"—"}</div>}
                      </div>;
                    })}
                  </div>
                  <div style={{display:"flex",gap:6,alignItems:"center"}}>
                    <input readOnly value={url} onClick={e=>e.target.select()} style={{...S.input,fontSize:10,padding:"5px 8px",flex:1,background:"#fff"}}/>
                    <button onClick={()=>navigator.clipboard.writeText(url)} style={{...S.btn,padding:"5px 10px",background:"#F1F5F9",border:"1px solid #E2E8F0",fontSize:11}}>Copy</button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Risks */}
      {activeSection==="risks" && (
        <div style={S.card}>
          <p style={S.sec}>Risks — manual entry</p>
          <p style={{fontSize:13,color:"#64748B",marginBottom:14,marginTop:-8}}>AMO has no Risk issue type. Enter risks manually. These appear as a structured slide in the PowerPoint.</p>
          {(manual.risks||[]).map((r,ri)=>(
            <div key={ri} style={{display:"grid",gridTemplateColumns:"120px 1fr 1fr 1fr auto",gap:8,marginBottom:10,alignItems:"start",padding:"12px",background:"#F8FAFC",borderRadius:8,border:"1px solid #E2E8F0"}}>
              <div>
                <label style={S.label}>Severity</label>
                <select style={{...S.input,padding:"8px 10px"}} value={r.severity||"Medium"} onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],severity:e.target.value};setM("risks",rs);}}>
                  <option>High</option><option>Medium</option><option>Low</option>
                </select>
              </div>
              <div>
                <label style={S.label}>Risk title</label>
                <input style={S.input} value={r.title||""} placeholder="e.g. Coach capacity gap" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],title:e.target.value};setM("risks",rs);}}/>
              </div>
              <div>
                <label style={S.label}>Description</label>
                <input style={S.input} value={r.description||""} placeholder="Impact if not mitigated" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],description:e.target.value};setM("risks",rs);}}/>
              </div>
              <div>
                <label style={S.label}>Mitigation · Owner</label>
                <div style={{display:"flex",gap:4}}>
                  <input style={{...S.input,flex:2}} value={r.mitigation||""} placeholder="Mitigation action" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],mitigation:e.target.value};setM("risks",rs);}}/>
                  <input style={{...S.input,flex:1}} value={r.owner||""} placeholder="Owner" onChange={e=>{const rs=[...manual.risks];rs[ri]={...rs[ri],owner:e.target.value};setM("risks",rs);}}/>
                </div>
              </div>
              <button onClick={()=>setM("risks",(manual.risks||[]).filter((_,i)=>i!==ri))} style={{...S.btn,padding:"6px 10px",background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",marginTop:20,alignSelf:"end"}}>✕</button>
            </div>
          ))}
          <button style={{...S.btn,background:"#F1F5F9",border:"1px solid #E2E8F0",color:"#1E2761",marginTop:4}} onClick={()=>setM("risks",[...(manual.risks||[]),{severity:"Medium",title:"",description:"",mitigation:"",owner:""}])}>
            + Add risk
          </button>
        </div>
      )}

      {/* Asks */}
      {activeSection==="asks" && (
        <div style={S.card}>
          <p style={S.sec}>Where stakeholders can help</p>
          <label style={S.label}>One ask per line — numbered automatically in the PowerPoint</label>
          <textarea style={{...S.ta,minHeight:120}} rows={5} value={manual.asks} onChange={e=>setM("asks",e.target.value)}
            placeholder={"Prioritisation: AC CoE needs leadership alignment on AI adoption vs PL maturity — both can't be first\nCapacity: 2 additional coaches needed to meet 2026 demand across all PLs\nUnblock: Dependency Management D&T-wide mandate still pending — stalling roll-out"}/>
        </div>
      )}

      {/* Generate */}
      {phase==="done" ? (
        <div style={{...S.card,textAlign:"center",padding:"36px 24px"}}>
          <div style={{fontSize:36,marginBottom:10}}>✓</div>
          <h2 style={{color:"#16A34A",marginBottom:6,fontWeight:700}}>Report downloaded</h2>
          <p style={{fontSize:13,color:"#64748B",marginBottom:20}}>Check your Downloads folder for AC-CoE-Report-v2-*.pptx</p>
          <button style={{...S.btn,background:"#F1F5F9",color:"#0F172A",border:"1px solid #E2E8F0"}} onClick={()=>setPhase("idle")}>Generate another version</button>
        </div>
      ) : (
        <div style={S.dark}>
          <p style={{fontSize:13,color:"#CADCFC",marginBottom:4,opacity:.8}}>9 slides · Live Jira data frozen at generation time</p>
          <p style={{fontSize:11,color:"#4FC3F7",marginBottom:14}}>Epics · Flow metrics · Dependencies · Capacity · Wellness · Risks · Asks</p>
          {phase==="generating"
            ? <div style={{color:"#CADCFC",fontSize:15}}>⏳ Building PowerPoint…</div>
            : <button style={{...S.btn,background:"#15803D",color:"#fff",fontSize:15,padding:"13px 28px"}} onClick={generate} disabled={!pptxOk}>
                {pptxOk?"⬇  Generate PowerPoint":"⏳  Loading library…"}
              </button>
          }
        </div>
      )}
    </div>
  );
}
