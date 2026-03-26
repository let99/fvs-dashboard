import { useMemo, useState, useCallback } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const C = {
  bg:"#0f172a", card:"#1e293b", border:"#334155", accent:"#38bdf8",
  ok:"#16a34a", warn:"#f59e0b", bad:"#dc2626", na:"#94a3b8",
  blue:"#2563eb", purple:"#7c3aed", teal:"#0f766e", white:"#f8fafc",
  muted:"#64748b", row:"#ffffff08",
};

const san = v => String(v||"").replace(/\s+/g," ").trim();
const pav = a => { const m=String(a).match(/\d{3,4}/); return m?`${Math.floor(Number(m[0])/100)}º`:""; };

function fix(s){
  return String(s||"")
    .replace(/Ã‡/g,"Ç").replace(/Ã§/g,"ç").replace(/Ã£/g,"ã").replace(/Ãƒ/g,"Ã")
    .replace(/Ã¢/g,"â").replace(/Ã©/g,"é").replace(/Ãª/g,"ê").replace(/Ã­/g,"í")
    .replace(/Ãµ/g,"õ").replace(/Ãº/g,"ú").replace(/Ã"/g,"Ó").replace(/Â°/g,"°")
    .replace(/Ã‰/g,"É").replace(/Ã"/g,"Ô").replace(/ÃƒO/g,"ÃO");
}

function fixEnc(s){ try{ return decodeURIComponent(escape(s)); }catch{ return s; } }

function parseCSVLines(text){
  return text.split(/\r?\n/).map(l=>l.split(",").map(c=>san(fix(fixEnc(c)))));
}

function detectTipo(fileName, rows){
  const head = rows.slice(0,6).flat().join(" ").toLowerCase();
  const fn   = fileName.toLowerCase();
  if (/shaft/i.test(fn)||/shaft/i.test(head))           return "shaft";
  if (/capiaç|capiac/i.test(fn)||/capiaç|capiac/i.test(head)) return "capiacos";
  if (/passante/i.test(fn)||/passante/i.test(head))     return "passantes";
  if (/esquadria/i.test(fn)||/esquadria/i.test(head))   return "esquadrias";
  if (/cerâmica|ceramica|varanda/i.test(fn))             return "ceramica";
  return "generico";
}

const SHAFT_AMBIENTES = ["VARANDA","COZINHA","BWC SERVIÇO","ÁREA SERVIÇO","BWC SUÍTE 01 E 02","BWC SUÍTE 03","BWC SUÍTE MASTER"];

function parseShafts(rows, fileName){
  const result=[];
  const isCasarao = /casar/i.test(fileName)||rows.slice(0,5).flat().join(" ").toLowerCase().includes("casar");
  if(isCasarao){
    const hiIdx = rows.findIndex(r=>r[0]&&/casarão|casarao|casarã/i.test(r[0]));
    if(hiIdx===-1) return result;
    const ambientes = rows[hiIdx].slice(1).filter(Boolean);
    for(let i=hiIdx+1;i<rows.length;i++){
      const r=rows[i];
      if(!r[0]||/^[,\s]*$/.test(r[0])||/^[A-Z]$/.test(r[0])||/total|legenda/i.test(r[0])) continue;
      ambientes.forEach((amb,j)=>{
        const val=san(r[j+1]);
        if(!val) return;
        result.push({ tipo:"shaft", torre:"CASARÃO", apto:san(r[0]), ambiente:fix(amb), status:val, fonte:fileName });
      });
    }
    return result;
  }
  const dataStart = rows.findIndex(r=>r[0]&&/^\d{3,4}$/.test(r[0]));
  if(dataStart===-1) return result;
  for(let i=dataStart;i<rows.length;i++){
    const r=rows[i];
    if(!r[0]||!/^\d{3,4}$/.test(r[0])) continue;
    const apto1=san(r[0]), torre1=san(r[1]);
    SHAFT_AMBIENTES.forEach((amb,j)=>{
      const val=san(r[2+j]);
      if(!val||val===torre1) return;
      result.push({ tipo:"shaft", torre:torre1, apto:apto1, pav:pav(apto1), ambiente:amb, status:val, fonte:fileName });
    });
    const apto2=san(r[10]), torre2=san(r[11]);
    if(apto2&&torre2&&/^\d{3,4}$/.test(apto2)){
      SHAFT_AMBIENTES.forEach((amb,j)=>{
        const val=san(r[12+j]);
        if(!val||val===torre2) return;
        result.push({ tipo:"shaft", torre:torre2, apto:apto2, pav:pav(apto2), ambiente:amb, status:val, fonte:fileName });
      });
    }
  }
  return result;
}

function parseCapiacos(rows, fileName){
  const result=[];
  const headerIdx=rows.findIndex(r=>r[1]&&/apto/i.test(r[1])&&r[2]&&/torre/i.test(r[2]));
  if(headerIdx===-1) return result;
  const headers=rows[headerIdx].slice(3).map(fix).filter(Boolean);
  for(let i=headerIdx+1;i<rows.length;i++){
    const r=rows[i];
    if(!r[1]||!/^\d{3,4}$/.test(r[1])) continue;
    const apto=san(r[1]), torre=san(r[2]);
    headers.forEach((amb,j)=>{
      const val=san(r[3+j]);
      if(!val) return;
      result.push({ tipo:"capiacos", torre, apto, pav:pav(apto), ambiente:amb, status:val, fonte:fileName });
    });
  }
  return result;
}

const PASS_AMBIENTES=["VARANDA","COZINHA","BWC SERVIÇO","BWC SUÍTE 01 E 02","BWC SUÍTE 03","BWC SUÍTE MASTER","ÁREA DE SERVIÇO","LAVABO"];

function parsePassantes(rows, fileName){
  const result=[];
  const headerIdxs=rows.reduce((acc,r,i)=>{ if(r[1]&&/apto/i.test(r[1])&&r[2]&&/torre/i.test(r[2])) acc.push(i); return acc; },[]);
  for(const hi of headerIdxs){
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i];
      if(!r[1]||!/^\d{3,4}$/.test(r[1])) continue;
      const apto1=san(r[1]), torre1=san(r[2]);
      PASS_AMBIENTES.forEach((amb,j)=>{
        const val=san(r[3+j]);
        if(!val) return;
        result.push({ tipo:"passantes", torre:torre1, apto:apto1, pav:pav(apto1), ambiente:amb, status:val, fonte:fileName });
      });
      const apto2=san(r[11]), torre2=san(r[12]);
      if(apto2&&torre2&&/^\d{3,4}$/.test(apto2)){
        PASS_AMBIENTES.forEach((amb,j)=>{
          const val=san(r[13+j]);
          if(!val) return;
          result.push({ tipo:"passantes", torre:torre2, apto:apto2, pav:pav(apto2), ambiente:amb, status:val, fonte:fileName });
        });
      }
    }
  }
  return result;
}

function parseEsquadrias(rows, fileName){
  const result=[];
  const headerIdxs=rows.reduce((acc,r,i)=>{ if(r[0]&&/apto/i.test(r[0])&&r[1]&&/torre/i.test(r[1])) acc.push(i); return acc; },[]);
  for(const hi of headerIdxs){
    const headers=rows[hi].slice(2).map(h=>san(fix(h)));
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i];
      if(!r[0]||!/^\d{3,4}$/.test(r[0])) continue;
      const apto=san(r[0]), torre=san(r[1]);
      headers.forEach((amb,j)=>{
        const val=san(r[2+j]);
        if(!val) return;
        result.push({ tipo:"esquadrias", torre, apto, pav:pav(apto), ambiente:amb, status:val, fonte:fileName });
      });
    }
  }
  return result;
}

function parseFile(fileName, csvText){
  const rows=parseCSVLines(csvText);
  const tipo=detectTipo(fileName, rows);
  switch(tipo){
    case "shaft":      return parseShafts(rows, fileName);
    case "capiacos":   return parseCapiacos(rows, fileName);
    case "passantes":  return parsePassantes(rows, fileName);
    case "esquadrias": return parseEsquadrias(rows, fileName);
    default:           return [];
  }
}

const CLASS = {
  shaft:{
    ok:["A"], warn:["FS"], pend:["FC"], na:["N/A","N/V","?"],
    labels:{ A:"Aberto", FS:"Fech. s/ cerâmica", FC:"Fech. c/ cerâmica", "N/A":"N/A", "N/V":"N/V" },
    colors:{ A:C.ok, FS:C.warn, FC:C.blue, "N/A":C.na, "N/V":C.muted },
  },
  capiacos:{
    ok:["OK"], warn:["Q","N","Q.I","F"], na:["N/V","N/A"],
    labels:{ OK:"Correto", Q:"Sem queda", N:"Desnivelado", "Q.I":"Queda invertida", F:"Falta fachada", "N/V":"N/V" },
    colors:{ OK:C.ok, Q:C.bad, N:C.warn, "Q.I":C.bad, F:C.bad, "N/V":C.na },
  },
  passantes:{
    ok:["OK"], warn:["R","C","Q","P","F","S"], na:["N/V","N/A"],
    labels:{ OK:"Correto", R:"Rente ao piso", C:"PEX chumbado", Q:"Quebrado", P:"Mal-fixado", F:"Falta tubo", S:"Sujeira", "N/V":"N/V" },
    colors:{ OK:C.ok, R:C.bad, C:C.warn, Q:C.bad, P:C.warn, F:C.bad, S:C.muted, "N/V":C.na },
  },
  esquadrias:{
    ok:["E"], warn:["I","F","C","P.U","S"], na:[],
    labels:{ E:"Instalada", I:"Pronto p/ instalar", F:"Serviço faltante", C:"Falta contramarco", "P.U":"P.U incompleto", S:"Contramarco sujo" },
    colors:{ E:C.ok, I:C.blue, F:C.bad, C:C.warn, "P.U":C.warn, S:C.muted },
  },
};

function calcTipo(rows, tipo){
  const cl=CLASS[tipo];
  if(!cl) return { counts:{}, byTorre:[], byApto:[] };
  const counts={}, byTorre={}, byApto={};
  rows.filter(r=>r.tipo===tipo).forEach(r=>{
    const s=r.status||"?";
    counts[s]=(counts[s]||0)+1;
    const t=r.torre||"?";
    if(!byTorre[t]) byTorre[t]={ torre:t, counts:{} };
    byTorre[t].counts[s]=(byTorre[t].counts[s]||0)+1;
    const key=`${t}-${r.apto}`;
    if(!byApto[key]) byApto[key]={ torre:t, apto:r.apto, pav:r.pav||pav(r.apto), total:0, prob:0, statuses:new Set() };
    byApto[key].total++;
    if(cl.warn.includes(s)){ byApto[key].prob++; byApto[key].statuses.add(s); }
  });
  return {
    counts,
    byTorre: Object.values(byTorre).sort((a,b)=>a.torre.localeCompare(b.torre)),
    byApto:  Object.values(byApto).map(x=>({
      ...x, verificacoes:x.total,
      pct: x.total?`${Math.round(x.prob/x.total*100)}%`:"0%",
      statuses:[...x.statuses].join(", "),
    })).sort((a,b)=>a.torre.localeCompare(b.torre)||Number(String(a.apto).match(/\d+/)?.[0]||0)-Number(String(b.apto).match(/\d+/)?.[0]||0)),
  };
}

function BarChart({ counts, labels, colors }){
  const entries=Object.entries(counts).filter(([,v])=>v>0);
  const total=entries.reduce((a,[,v])=>a+v,0)||1;
  const sorted=[...entries].sort((a,b)=>b[1]-a[1]);
  const maxVal=sorted[0]?.[1]||1;
  return(
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      {sorted.map(([k,v])=>(
        <div key={k} style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:140,fontSize:12,color:C.white,textAlign:"right",flexShrink:0}}>{labels?.[k]||k}</div>
          <div style={{flex:1,background:"#0f172a",borderRadius:4,height:24,position:"relative"}}>
            <div style={{width:`${(v/maxVal)*100}%`,background:colors?.[k]||C.blue,height:"100%",borderRadius:4,transition:"width .4s"}}/>
            <span style={{position:"absolute",right:8,top:4,fontSize:11,color:C.white,fontWeight:"bold"}}>{v} ({Math.round(v/total*100)}%)</span>
          </div>
        </div>
      ))}
    </div>
  );
}

const KPI = ({label,value,sub,color}) => (
  <div style={{background:C.card,borderRadius:14,padding:"18px 22px",borderLeft:`4px solid ${color||C.accent}`}}>
    <div style={{fontSize:10,textTransform:"uppercase",color:C.muted,fontWeight:"bold",marginBottom:6}}>{label}</div>
    <div style={{fontSize:30,fontWeight:"bold",color:C.white,marginBottom:3}}>{value}</div>
    <div style={{fontSize:12,color:C.muted}}>{sub}</div>
  </div>
);

const Box = ({title,children,action}) => (
  <div style={{background:C.card,borderRadius:18,padding:22,marginTop:20}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
      <h2 style={{margin:0,color:C.white,fontSize:16}}>{title}</h2>
      {action}
    </div>
    {children}
  </div>
);

const TH = ({c}) => <th style={{borderBottom:`1px solid ${C.border}`,padding:"9px 11px",textAlign:"left",color:C.muted,fontSize:11,textTransform:"uppercase",background:C.card}}>{c}</th>;
const TD = ({c,bold,color,children}) => <td style={{borderBottom:`1px solid #0f172a`,padding:"9px 11px",color:color||C.white,fontWeight:bold?"bold":"normal",fontSize:12}}>{children!==undefined?children:c}</td>;
const Btn = ({children,onClick,color}) => (
  <button onClick={onClick} style={{background:color||C.blue,color:"white",border:"none",borderRadius:8,padding:"9px 15px",cursor:"pointer",fontWeight:"bold",fontSize:13}}>{children}</button>
);

function TabelaApto({ data, tipo }){
  if(!data.byApto.length) return <p style={{color:C.muted}}>Sem dados para este filtro.</p>;
  return(
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",minWidth:700}}>
        <thead><tr>
          <TH c="Torre"/><TH c="Apto"/><TH c="Pav"/><TH c="Total"/><TH c="Problemas"/><TH c="% Prob."/><TH c="Status encontrados"/>
        </tr></thead>
        <tbody>
          {data.byApto.map((r,i)=>(
            <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
              <TD c={r.torre} color={C.accent} bold/>
              <TD c={r.apto} bold/>
              <TD c={r.pav}/>
              <TD c={r.total}/>
              <TD c={r.prob} color={r.prob>0?C.bad:C.ok}/>
              <TD c={r.pct}/>
              <TD c={r.statuses||"—"} color={r.statuses?C.warn:C.muted}/>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function TabelaTorre({ data }){
  if(!data.byTorre.length) return <p style={{color:C.muted}}>Sem dados.</p>;
  const statusKeys=[...new Set(data.byTorre.flatMap(t=>Object.keys(t.counts)))];
  return(
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse"}}>
        <thead><tr>
          <TH c="Torre"/>
          {statusKeys.map(k=><TH key={k} c={k}/>)}
          <TH c="Total"/>
        </tr></thead>
        <tbody>
          {data.byTorre.map((t,i)=>{
            const total=Object.values(t.counts).reduce((a,b)=>a+b,0);
            return(
              <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
                <TD c={`Torre ${t.torre}`} color={C.accent} bold/>
                {statusKeys.map(k=><TD key={k} c={t.counts[k]||0}/>)}
                <TD c={total} bold/>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

export default function App(){
  const [allRows, setAllRows]         = useState([]);
  const [status, setStatus]           = useState("Envie os arquivos CSV, XLSX, DOCX ou PDF.");
  const [fileNames, setFileNames]     = useState([]);
  const [errors, setErrors]           = useState([]);
  const [tab, setTab]                 = useState("shafts");
  const [torreFilter, setTorreFilter] = useState("TODAS");
  const [aiText, setAiText]           = useState("");
  const [aiLoading, setAiLoading]     = useState(false);

  const handleFile = useCallback(async e=>{
    const files=Array.from(e.target.files||[]);
    if(!files.length) return;
    setAllRows([]); setErrors([]); setAiText("");
    setFileNames(files.map(f=>f.name));
    setStatus("Processando...");
    let all=[], errs=[];

    for(const file of files){
      try{
        const ext=file.name.split(".").pop()?.toLowerCase();
        if(ext==="csv"){
          const text=await file.text();
          const parsed=parseFile(file.name, text);
          if(parsed.length) all=[...all,...parsed];
          else errs.push(`${file.name}: nenhuma linha reconhecida.`);
          continue;
        }
        if(ext==="xlsx"||ext==="xls"){
          const buf=await file.arrayBuffer();
          const wb=XLSX.read(buf,{type:"array"});
          for(const sn of wb.SheetNames){
            const csv=XLSX.utils.sheet_to_csv(wb.Sheets[sn]);
            all=[...all,...parseFile(`${file.name} ${sn}`, csv)];
          }
          continue;
        }
        if(ext==="docx"){
          const buf=await file.arrayBuffer();
          const res=await mammoth.extractRawText({arrayBuffer:buf});
          all=[...all,...parseFile(file.name, res.value||"")];
          continue;
        }
        if(ext==="pdf"){
          const buf=await file.arrayBuffer();
          const pdf=await pdfjsLib.getDocument({data:buf}).promise;
          let txt="";
          for(let p=1;p<=pdf.numPages;p++){
            const page=await pdf.getPage(p);
            const ct=await page.getTextContent();
            txt+=" "+ct.items.map(i=>i.str).join(" ");
          }
          all=[...all,...parseFile(file.name, txt)];
          continue;
        }
        errs.push(`${file.name}: formato não suportado.`);
      }catch(err){
        console.error(err);
        errs.push(`${file.name}: erro — ${err.message}`);
      }
    }
    setAllRows(all); setErrors(errs);
    setStatus(`${files.length} arquivo(s) processado(s). ${all.length} registros carregados.`);
  },[]);

  const torres=useMemo(()=>["TODAS",...[...new Set(allRows.map(r=>r.torre).filter(Boolean))].sort()],[allRows]);
  const scoped=useMemo(()=>torreFilter==="TODAS"?allRows:allRows.filter(r=>r.torre===torreFilter),[allRows,torreFilter]);

  const shaftData = useMemo(()=>calcTipo(scoped,"shaft"),      [scoped]);
  const capData   = useMemo(()=>calcTipo(scoped,"capiacos"),   [scoped]);
  const passData  = useMemo(()=>calcTipo(scoped,"passantes"),  [scoped]);
  const esqData   = useMemo(()=>calcTipo(scoped,"esquadrias"), [scoped]);

  const shaftTotal  = Object.values(shaftData.counts).reduce((a,b)=>a+b,0);
  const shaftAberto = shaftData.counts["A"]||0;
  const capTotal    = Object.values(capData.counts).reduce((a,b)=>a+b,0);
  const capProb     = (capData.counts["Q"]||0)+(capData.counts["N"]||0)+(capData.counts["Q.I"]||0)+(capData.counts["F"]||0);
  const passTotal   = Object.values(passData.counts).reduce((a,b)=>a+b,0);
  const passProb    = passTotal-(passData.counts["OK"]||0)-(passData.counts["N/V"]||0);
  const esqTotal    = Object.values(esqData.counts).reduce((a,b)=>a+b,0);
  const esqInst     = esqData.counts["E"]||0;

  function exportCSV(){
    const lines=["tipo,torre,apto,pav,ambiente,status,fonte"];
    allRows.forEach(r=>lines.push(`${r.tipo},${r.torre},${r.apto||""},${r.pav||""},${r.ambiente||""},${r.status||""},${r.fonte||""}`));
    const blob=new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"});
    const url=URL.createObjectURL(blob); const a=document.createElement("a");
    a.href=url; a.download="fvs_completo.csv"; a.click(); URL.revokeObjectURL(url);
  }

  async function runAI(){
    setAiLoading(true); setAiText("");
    const sumario=`SHAFTS: ${shaftAberto}/${shaftTotal} abertos\nPor status: ${JSON.stringify(shaftData.counts)}\nCAPIAÇOS: ${capProb} problemas de ${capTotal}\nPor status: ${JSON.stringify(capData.counts)}\nPASSANTES: ${passProb} problemas de ${passTotal}\nPor status: ${JSON.stringify(passData.counts)}\nESQUADRIAS: ${esqInst}/${esqTotal} instaladas\nPor status: ${JSON.stringify(esqData.counts)}\nPor torre Shafts: ${JSON.stringify(shaftData.byTorre.map(t=>({torre:t.torre,...t.counts})))}\nPor torre Capiaços: ${JSON.stringify(capData.byTorre.map(t=>({torre:t.torre,...t.counts})))}\nPor torre Passantes: ${JSON.stringify(passData.byTorre.map(t=>({torre:t.torre,...t.counts})))}\nPor torre Esquadrias: ${JSON.stringify(esqData.byTorre.map(t=>({torre:t.torre,...t.counts})))}`;
    try{
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({
          model:"claude-sonnet-4-20250514", max_tokens:1200,
          system:"Você é um especialista em qualidade de obras residenciais. Analise os dados de verificação de serviços (shafts, capiaços, passantes e esquadrias) e produza um relatório executivo em português com: 1) Situação geral de cada serviço, 2) Torres mais críticas, 3) Principais pendências, 4) Recomendações prioritárias. Use linguagem técnica e objetiva.",
          messages:[{role:"user",content:`Analise os dados e gere o relatório:\n\n${sumario}`}]
        })
      });
      const data=await res.json();
      setAiText(data.content?.filter(c=>c.type==="text").map(c=>c.text).join("\n")||"Sem resposta.");
    }catch(err){ setAiText(`Erro: ${err.message}`); }
    setAiLoading(false);
  }

  const TABS=[
    {id:"shafts",label:"🔲 Shafts"},
    {id:"capiacos",label:"🏗 Capiaços"},
    {id:"passantes",label:"🔧 Passantes"},
    {id:"esquadrias",label:"🪟 Esquadrias"},
    {id:"ia",label:"🤖 Análise IA"},
  ];

  return(
    <div style={{minHeight:"100vh",background:C.bg,color:C.white,padding:"28px 20px",fontFamily:"'Inter',Arial,sans-serif"}}>
      <div style={{maxWidth:1280,margin:"0 auto"}}>

        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
          <span style={{background:"#082f49",color:C.accent,borderRadius:999,padding:"3px 12px",fontSize:12,fontWeight:"bold"}}>FVS Qualidade</span>
          <span style={{color:C.muted,fontSize:12}}>v3.0 — Multi-serviço</span>
        </div>
        <h1 style={{fontSize:32,margin:"0 0 4px",color:C.white}}>Dashboard de Verificação de Serviços</h1>
        <p style={{color:C.muted,margin:"0 0 22px",fontSize:13}}>Shafts · Capiaços · Passantes · Esquadrias · Cerâmica</p>

        <div style={{background:C.card,borderRadius:18,padding:20,marginBottom:20,border:`1px solid ${C.border}`}}>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
            <label style={{background:C.blue,color:"white",borderRadius:9,padding:"9px 16px",cursor:"pointer",fontWeight:"bold",fontSize:13}}>
              📂 Carregar Arquivos
              <input type="file" multiple accept=".csv,.xlsx,.xls,.docx,.pdf" onChange={handleFile} style={{display:"none"}}/>
            </label>
            {allRows.length>0 && <Btn onClick={exportCSV} color={C.purple}>⬇ Exportar CSV</Btn>}
            {allRows.length>0 && (
              <div style={{marginLeft:"auto"}}>
                <span style={{fontSize:11,color:C.muted,marginRight:8}}>TORRE</span>
                <select value={torreFilter} onChange={e=>setTorreFilter(e.target.value)} style={{background:"#0f172a",color:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 11px",fontSize:12}}>
                  {torres.map(t=><option key={t} value={t}>{t==="TODAS"?"Todas as torres":`Torre ${t}`}</option>)}
                </select>
              </div>
            )}
          </div>
          {fileNames.length>0 && <div style={{marginTop:10,color:C.muted,fontSize:12}}>{fileNames.join(" · ")}</div>}
          <div style={{marginTop:6,fontSize:12,color:errors.length?C.bad:C.muted}}>{status}</div>
          {errors.map((e,i)=><div key={i} style={{color:C.bad,fontSize:11,marginTop:3}}>⚠ {e}</div>)}
        </div>

        {allRows.length>0 && <>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:14,marginBottom:20}}>
            <KPI label="Shafts Abertos"        value={`${shaftAberto}/${shaftTotal}`} sub={`${shaftTotal?Math.round(shaftAberto/shaftTotal*100):0}% do total`}  color={C.ok}/>
            <KPI label="Capiaços c/ Problema"  value={capProb}                        sub={`de ${capTotal} verificados`}                                         color={capProb>0?C.bad:C.ok}/>
            <KPI label="Passantes c/ Problema" value={passProb}                       sub={`de ${passTotal} verificados`}                                        color={passProb>0?C.bad:C.ok}/>
            <KPI label="Esquadrias Instaladas" value={`${esqInst}/${esqTotal}`}       sub={`${esqTotal?Math.round(esqInst/esqTotal*100):0}% concluído`}          color={C.blue}/>
          </div>

          <div style={{display:"flex",gap:3,borderBottom:`1px solid ${C.border}`,marginBottom:0}}>
            {TABS.map(t=>(
              <button key={t.id} onClick={()=>setTab(t.id)} style={{background:tab===t.id?C.blue:"transparent",color:tab===t.id?C.white:C.muted,border:"none",borderRadius:"8px 8px 0 0",padding:"10px 18px",cursor:"pointer",fontWeight:tab===t.id?"bold":"normal",fontSize:13}}>
                {t.label}
              </button>
            ))}
          </div>

          {tab==="shafts" && <>
            <Box title="Distribuição de status — Shafts"><BarChart counts={shaftData.counts} labels={CLASS.shaft.labels} colors={CLASS.shaft.colors}/></Box>
            <Box title="Shafts por Torre"><TabelaTorre data={shaftData}/></Box>
            <Box title="Detalhe por Apartamento"><TabelaApto data={shaftData} tipo="shaft"/></Box>
          </>}
          {tab==="capiacos" && <>
            <Box title="Distribuição de status — Capiaços"><BarChart counts={capData.counts} labels={CLASS.capiacos.labels} colors={CLASS.capiacos.colors}/></Box>
            <Box title="Capiaços por Torre"><TabelaTorre data={capData}/></Box>
            <Box title="Detalhe por Apartamento"><TabelaApto data={capData} tipo="capiacos"/></Box>
          </>}
          {tab==="passantes" && <>
            <Box title="Distribuição de status — Passantes"><BarChart counts={passData.counts} labels={CLASS.passantes.labels} colors={CLASS.passantes.colors}/></Box>
            <Box title="Passantes por Torre"><TabelaTorre data={passData}/></Box>
            <Box title="Detalhe por Apartamento"><TabelaApto data={passData} tipo="passantes"/></Box>
          </>}
          {tab==="esquadrias" && <>
            <Box title="Distribuição de status — Esquadrias"><BarChart counts={esqData.counts} labels={CLASS.esquadrias.labels} colors={CLASS.esquadrias.colors}/></Box>
            <Box title="Esquadrias por Torre"><TabelaTorre data={esqData}/></Box>
            <Box title="Detalhe por Apartamento"><TabelaApto data={esqData} tipo="esquadrias"/></Box>
          </>}
          {tab==="ia" && (
            <Box title="🤖 Análise Automática por IA" action={<Btn onClick={runAI}>{aiLoading?"Analisando...":"Gerar Análise"}</Btn>}>
              {!aiText&&!aiLoading&&<p style={{color:C.muted,fontSize:13}}>Clique em "Gerar Análise" para obter um relatório executivo com diagnóstico de todos os serviços e recomendações por torre.</p>}
              {aiLoading&&<div style={{color:C.accent,padding:"16px 0",fontSize:13}}>⏳ Gerando análise inteligente...</div>}
              {aiText&&<pre style={{whiteSpace:"pre-wrap",wordBreak:"break-word",fontSize:13,lineHeight:1.7,color:"#e2e8f0",margin:0}}>{aiText}</pre>}
            </Box>
          )}
        </>}

        {allRows.length===0&&(
          <div style={{textAlign:"center",padding:"60px 20px",color:C.muted}}>
            <div style={{fontSize:48,marginBottom:14}}>📋</div>
            <div style={{fontSize:17,marginBottom:6,color:C.white}}>Nenhum dado carregado</div>
            <div style={{fontSize:13}}>Carregue os arquivos CSV, XLSX, DOCX ou PDF dos relatórios de FVS.</div>
            <div style={{marginTop:12,fontSize:12,color:"#475569"}}>Suporta: Shafts · Capiaços · Passantes · Esquadrias</div>
          </div>
        )}
      </div>
    </div>
  );
}
