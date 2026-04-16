import { useMemo, useState, useCallback } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";

const C = {
  bg:"#0f172a", card:"#1e293b", border:"#334155", accent:"#38bdf8",
  ok:"#16a34a", warn:"#f59e0b", bad:"#dc2626", na:"#94a3b8",
  blue:"#2563eb", purple:"#7c3aed", white:"#f8fafc", muted:"#64748b", row:"#ffffff08",
  orange:"#ea580c",
};
const PALETTE = ["#16a34a","#2563eb","#dc2626","#f59e0b","#ea580c","#7c3aed","#0891b2","#0f766e","#b91c1c","#64748b"];

const san     = v => String(v||"").replace(/\s+/g," ").trim();
const pav     = a => { const m=String(a).match(/\d{3,4}/); return m?`${Math.floor(Number(m[0])/100)}º`:""; };
const isApto  = v => /^\d{3,4}$|^t[eé]rreo$/i.test(san(v));
const isTorre = v => /^[A-D]$/i.test(san(v));

function fix(s){
  return String(s||"")
    .replace(/Ã‡/g,"Ç").replace(/Ã§/g,"ç").replace(/Ã£/g,"ã").replace(/Ãƒ/g,"Ã")
    .replace(/Ã¢/g,"â").replace(/Ã©/g,"é").replace(/Ãª/g,"ê").replace(/Ã­/g,"í")
    .replace(/Ãµ/g,"õ").replace(/Ãº/g,"ú").replace(/Ã"/g,"Ó").replace(/Â°/g,"°")
    .replace(/Ã‰/g,"É").replace(/Ã"/g,"Ô").replace(/ÃƒO/g,"ÃO");
}
function fixEnc(s){ try{ return decodeURIComponent(escape(s)); }catch{ return s; } }
function parseCSVLines(text){
  const sep = text.indexOf("\t")!==-1 && (text.indexOf(",")===-1||text.indexOf("\t")<text.indexOf(",")) ? "\t" : ",";
  return text.split(/\r?\n/).map(l=>l.split(sep).map(c=>san(fix(fixEnc(c)))));
}

// ─── Extrai torre e apto do nome do arquivo ───────────────────────────────
function extractTorreApto(fileName){
  const fn = fileName.replace(/_R\d+_?/gi,"");
  let apto="", torre="";
  let m = fn.match(/[_\s\.](\d{3,4})[_\s\.]*([A-D])[_\s\.\-]/i)
        || fn.match(/[_\s\.](\d{3,4})[_\s\.]*([A-D])$/i)
        || fn.match(/(\d{3,4})[_\s\.]*([A-D])[_\s\.\-]/i)
        || fn.match(/(\d{3,4})[_\s\.]*([A-D])(?:\.|$)/i);
  if(m){ apto=m[1]; torre=m[2].toUpperCase(); }
  else{ m=fn.match(/apto\s*(\d{3,4})\s*([A-D])/i); if(m){apto=m[1];torre=m[2].toUpperCase();} }
  return { apto, torre };
}

// ─── Detecta serviço FVS ──────────────────────────────────────────────────
function detectFvsServico(fileName, text){
  const s = (fileName+" "+text).toLowerCase();
  if (/contrapiso/i.test(s))                         return "contrapiso";
  if (/porta.*madeira|madeira.*porta/i.test(s))      return "porta";
  if (/alumin/i.test(s)||/esquadria.*alum/i.test(s)) return "esquadria_alum";
  if (/cerâmic|revestimento/i.test(s))               return "ceramica";
  return null;
}

const FVS_CRITERIOS = {
  ceramica:      ["Planicidade","Peças sem trincas e lascas","Declividade em direção aos ralos","Rejunte","Sem excesso de argamassa","Dupla colagem","Terminalidade","Presença de som cavo","Limpeza"],
  contrapiso:    ["Planicidade","Homogeneidade","Declividade em direção aos ralos","Presença de som cavo","Integridade e fixação dos tubos passantes","Bacia de 30cm","Terminalidade","Limpeza"],
  porta:         ["Prumo e esquadro","Peças sem manchas ou arranhões","Verificação do encontro em 45 graus","Encaixe da porta na grade sem aberturas","Teste de abre e fecha","Chaves em poder da administração","Integridade das dobradiças e maçanetas","Fixação de alisar","Terminalidade","Limpeza"],
  esquadria_alum:["Limpeza do contramarco","Limpeza da Junta","Tratamento dos cantos contramarco","Tratamento das juntas","Silicone de instalação","Proteção da esquadria","Silicone de vedação externa","Terminalidade","Limpeza"],
};
const FVS_SERVICO_LABELS = { ceramica:"Cerâmica", contrapiso:"Contrapiso", porta:"Porta de Madeira", esquadria_alum:"Esquadria de Alumínio" };
const AMBIENTES_FVS = ["SALA","VARANDA","COZINHA","ÁREA SERV","DEPÓSITO","BWC SERV","LAVABO","SUÍTE 01","BWC SUÍTE 01 E 02","SUÍTE 02","SUÍTE 03","BWC SUÍTE 03","SUÍTE 04","BWC SUÍTE 04"];

// ─── Parser DOCX/PDF FVS ──────────────────────────────────────────────────
function parseFvsDocx(fileName, text){
  const { apto, torre } = extractTorreApto(fileName);
  const servico = detectFvsServico(fileName, text);
  if(!servico) return [];
  const dataMatch = text.match(/DATA\s*:?\s*(\d{2}\/\d{2}\/\d{4})/i);
  const data = dataMatch?.[1]||"";
  const criterios = FVS_CRITERIOS[servico];
  const VALID = new Set(["A","R","N/V","N/A"]);
  const tokens = text.replace(/\*\*/g,"").split(/\n/).map(l=>l.trim()).filter(Boolean);
  let ambientes = [...AMBIENTES_FVS];
  const salaIdx = tokens.findIndex(t=>/^SALA$/i.test(t));
  if(salaIdx !== -1){
    const amb = [];
    for(let i=salaIdx; i<tokens.length; i++){
      const t = tokens[i].toUpperCase();
      if(/^\d+$/.test(t)) break;
      if(t.length > 1) amb.push(t);
    }
    if(amb.length >= 3) ambientes = amb;
  }
  const rows = [];
  for(let i=0; i<tokens.length; i++){
    const tok = tokens[i].trim();
    if(!/^\d+$/.test(tok)) continue;
    const numIdx = parseInt(tok)-1;
    if(numIdx < 0 || numIdx >= criterios.length) continue;
    const criterio = criterios[numIdx];
    const resultados = [];
    for(let j=i+1; j<tokens.length && resultados.length < ambientes.length; j++){
      const v = tokens[j].toUpperCase().replace(/\s+/g,"");
      if(VALID.has(v)) resultados.push(v);
      else if(/^\d+$/.test(v)) break;
      else if(resultados.length > 0 && v.length > 6) break;
    }
    for(let k=0; k<resultados.length && k<ambientes.length; k++){
      rows.push({ tipo_doc:"fvs", servico, torre, apto, pav:pav(apto), data, ambiente:ambientes[k], criterio, resultado:resultados[k], fonte:fileName });
    }
  }
  return rows;
}

// ─── Parser Cerâmica Varanda ──────────────────────────────────────────────
function parseCeramicaVaranda(rows, fileName){
  const result = [];
  const COL_MAP = [
    {col:1,torre:"A",suf:"02"},{col:2,torre:"A",suf:"01"},
    {col:3,torre:"B",suf:"02"},{col:4,torre:"B",suf:"01"},
    {col:5,torre:"C",suf:"02"},{col:6,torre:"C",suf:"01"},
    {col:7,torre:"D",suf:"02"},{col:8,torre:"D",suf:"01"},
  ];
  const VALID=new Set(["S","N","C","R"]);
  const dataStart = rows.findIndex(r => /^[1-9]\d?$/.test(san(r[0])));
  if(dataStart === -1) return result;
  for(let i=dataStart; i<rows.length; i++){
    const r = rows[i];
    const andarCell = san(r[0]);
    if(!/^\d+$/.test(andarCell)) break;
    const andar = parseInt(andarCell);
    if(andar < 1 || andar > 30) break;
    COL_MAP.forEach(({col,torre,suf})=>{
      const val = san(r[col]).toUpperCase();
      if(!VALID.has(val)) return;
      result.push({ tipo_doc:"varanda", servico:"varanda", torre, apto:`${andar}${suf}`, pav:`${andar}º`, status:val, fonte:fileName });
    });
  }
  return result;
}

// ─── Parser Som Cavo ─────────────────────────────────────────────────────
const SOMCAVO_AMBIENTES=["VARANDA","SALA","COZINHA","ÁREA DE SERVIÇO","DEPÓSITO","BWC SERVIÇO","LAVABO","BWC SUÍTE 01 E 02","BWC SUÍTE 03","BWC SUÍTE MASTER"];

function parseSomCavo(rows, fileName){
  const result=[];
  const headerIdxs=rows.reduce((acc,r,i)=>{
    if(r.some(c=>/^apto$/i.test(san(c)))&&r.some(c=>/^torre$/i.test(san(c)))) acc.push(i);
    return acc;
  },[]);
  for(const hi of headerIdxs){
    const hRow=rows[hi];
    const blocos=[];
    hRow.forEach((cell,ci)=>{
      if(!/^apto$/i.test(san(cell))) return;
      for(let ti=ci+1;ti<=ci+2&&ti<hRow.length;ti++){
        if(/^torre$/i.test(san(hRow[ti]))){
          blocos.push({aptoCol:ci, torreCol:ti, ambStart:ti+1});
          break;
        }
      }
    });
    if(!blocos.length) continue;
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i];
      if(r.every(c=>!san(c))) continue;
      for(const b of blocos){
        const apto=san(r[b.aptoCol]);
        const torre=san(r[b.torreCol]).toUpperCase();
        if(!/^\d{3,4}$/.test(apto)) continue;
        if(!/^[A-D]$/.test(torre)) continue;
        SOMCAVO_AMBIENTES.forEach((amb,j)=>{
          const raw=san(r[b.ambStart+j]);
          if(!raw||raw==="") return;
          if(/n\/v/i.test(raw)){
            result.push({tipo_doc:"somcavo",torre,apto,pav:pav(apto),ambiente:amb,pedras:null,status:"N/V",fonte:fileName});
          } else {
            const n=parseInt(raw);
            if(isNaN(n)) return;
            result.push({tipo_doc:"somcavo",torre,apto,pav:pav(apto),ambiente:amb,pedras:n,status:n>0?"R":"A",fonte:fileName});
          }
        });
      }
    }
  }
  return result;
}

// ─── Detecção de tipo ─────────────────────────────────────────────────────
function detectTipo(fileName, rows){
  const head = rows.slice(0,8).flat().join(" ").toLowerCase();
  const fn   = fileName.toLowerCase();
  if (/shaft/i.test(fn))         return "shaft";
  if (/capiaç|capiac/i.test(fn)) return "capiacos";
  if (/passante/i.test(fn))      return "passantes";
  if (/esquadria/i.test(fn))     return "esquadrias";
  if (/cerâmica.*varanda|varanda.*cerâmica|mapeamento.*varanda/i.test(fn)) return "varanda";
  if (/som.cavo|mapeamento.*fvs.*cer|mapeamento.*cer.*apto/i.test(fn)) return "somcavo";
  if (/\b(ab|cd)\b/i.test(fn))   return "shaft";
  if (/casar/i.test(fn))         return "shaft";
  if (/casarão|casarao/i.test(head))                              return "shaft";
  if (/copa.*\d|hall\s*\d|academia|brinquedo|festas/i.test(head)) return "shaft";
  if (/shaft\s*\d/i.test(head))                                   return "shaft";
  if (/shafts|mapeamento.*shaft/i.test(head))  return "shaft";
  if (/som.cavo|mapeamento.*fvs.*cer|mapeamento.*cer.*apto/i.test(head)) return "somcavo";
  if (/verifica.*capiaç|capiaç.*verifica/i.test(head)) return "capiacos";
  if (/verifica.*passante|passante.*verifica/i.test(head)) return "passantes";
  if (/serviço.*esquadria|precedente.*esquadria/i.test(head)) return "esquadrias";
  const hasShaftData = rows.slice(0,10).some(r=>{
    const v = r.map(c=>san(c).toUpperCase()).filter(Boolean);
    return /^\d{3,4}$/.test(v[0]) && /^[A-D]$/.test(v[1]) &&
           v.slice(2,6).some(x=>["A","FS","FC","N/A","N/V"].includes(x));
  });
  if(hasShaftData) return "shaft";
  if (/passante/i.test(head))      return "passantes";
  if (/esquadria/i.test(head))     return "esquadrias";
  if (/capiaç|capiac/i.test(head)) return "capiacos";
  return "generico";
}

// ─── Helper resumo genérico ───────────────────────────────────────────────
function parseSummaryByTorre(rows, fileName, tipo, statusMap, headerPattern){
  const result=[];
  const headIdx=rows.findIndex(r=>headerPattern.test(r.join(" ")));
  if(headIdx===-1) return null;
  const headRow=rows[headIdx];
  const torresCols=[];
  headRow.forEach((cell,ci)=>{
    const m=san(fix(cell)).match(/(?:TOTAL\s+)?TORRE\s+([A-D])$/i);
    if(m) torresCols.push({torre:m[1].toUpperCase(),col:ci});
  });
  if(!torresCols.length) return null;
  for(let i=headIdx+1;i<Math.min(headIdx+12,rows.length);i++){
    const r=rows[i];
    const statusCell=san(r[0]);
    if(/^apto$|^total$/i.test(statusCell)) break;
    const status=statusMap[statusCell];
    if(!status) continue;
    torresCols.forEach(({torre,col},ti)=>{
      const nextCol=ti+1<torresCols.length?torresCols[ti+1].col:col+4;
      const windowEnd=Math.min(col+3,nextCol);
      let val=0;
      for(let x=col;x<windowEnd;x++){
        const n=parseInt(san(r[x]));
        if(!isNaN(n)&&n>=0){val=n;break;}
      }
      for(let k=0;k<val;k++) result.push({tipo,torre,apto:"",pav:"",ambiente:"",status,fonte:fileName});
    });
  }
  return result.length>0?result:null;
}

// ─── Parser Shafts ────────────────────────────────────────────────────────
const SHAFT_AMBIENTES=["VARANDA","COZINHA","BWC SERVIÇO","ÁREA SERVIÇO","BWC SUÍTE 01 E 02","BWC SUÍTE 03","BWC SUÍTE MASTER"];
const SHAFT_VALID=new Set(["A","FS","FC","N/A","N/V"]);

function parseShafts(rows, fileName){
  const result=[];
  const flatHead=rows.slice(0,5).flat().join(" ");
  const isCasarao=/casar/i.test(fileName)||/casar/i.test(flatHead);
  if(isCasarao){
    const hiIdx=rows.findIndex(r=>
      r.some(c=>/casarão|casarao|casarã/i.test(san(c)))||
      r.some(c=>/copa|hall|academia|brinquedo|festas/i.test(san(c)))
    );
    if(hiIdx===-1) return result;
    const ambientes=rows[hiIdx].slice(1).map(c=>san(fix(c))).filter(Boolean);
    function normCasarao(v){
      const s=san(v).toUpperCase().replace(/\s/g,"");
      if(s==="NF"||s==="N/F") return "NF";
      if(["A","FS","FC","N/A","N/V","?"].includes(s)) return s;
      return s||null;
    }
    for(let i=hiIdx+1;i<rows.length;i++){
      const r=rows[i];
      const shaftName=san(r[0]);
      if(!shaftName||!/shaft\s*\d+/i.test(shaftName)) continue;
      ambientes.forEach((amb,j)=>{
        const val=normCasarao(r[j+1]);
        if(!val) return;
        result.push({tipo:"shaft",torre:"CASARÃO",apto:shaftName,pav:"",ambiente:amb,status:val,fonte:fileName});
      });
    }
    return result;
  }
  const headerIdx=rows.findIndex(r=>
    r.some(c=>/^apto$/i.test(san(c)))&&r.some(c=>/^torre$/i.test(san(c)))
  );
  if(headerIdx===-1) return result;
  const hRow=rows[headerIdx];
  const blocos=[];
  hRow.forEach((cell,ci)=>{
    if(!/^apto$/i.test(san(cell))) return;
    for(let ti=ci+1;ti<=ci+2&&ti<hRow.length;ti++){
      if(/^torre$/i.test(san(hRow[ti]))){
        blocos.push({aptoCol:ci,torreCol:ti,ambStart:ti+1});
        break;
      }
    }
  });
  if(!blocos.length) return result;
  for(let i=headerIdx+1;i<rows.length;i++){
    const r=rows[i];
    for(const b of blocos){
      const apto=san(r[b.aptoCol]);
      const torre=san(r[b.torreCol]).toUpperCase();
      if(!/^\d{3,4}$/.test(apto)) continue;
      if(!/^[A-D]$/.test(torre)) continue;
      SHAFT_AMBIENTES.forEach((amb,j)=>{
        const val=san(r[b.ambStart+j]).toUpperCase();
        if(!SHAFT_VALID.has(val)) return;
        result.push({tipo:"shaft",torre,apto,pav:pav(apto),ambiente:amb,status:val,fonte:fileName});
      });
    }
  }
  return result;
}

// ─── Parser Capiaços ─────────────────────────────────────────────────────
const CAP_STATUS_MAP={Q:"Q",N:"N","Q.I":"Q.I",F:"F",OK:"OK","N/V":"N/V"};
function parseCapiacos(rows,fileName){
  const summary=parseSummaryByTorre(rows,fileName,"capiacos",CAP_STATUS_MAP,/TORRE\s+[A-D]/i);
  if(summary) return summary;
  const result=[];
  const hi=rows.findIndex(r=>r.findIndex(c=>/^apto$/i.test(c))!==-1&&r.findIndex(c=>/^torre$/i.test(c))!==-1);
  if(hi===-1) return result;
  const h=rows[hi],ac=h.findIndex(c=>/^apto$/i.test(c)),tc=h.findIndex(c=>/^torre$/i.test(c)),as=Math.max(ac,tc)+1;
  const headers=h.slice(as).map(fix).filter(Boolean);
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]; const a=san(r[ac]),t=san(r[tc]);
    if(!a||!isApto(a)||!t) continue;
    headers.forEach((amb,j)=>{ const val=san(r[as+j]); if(!val) return; result.push({tipo:"capiacos",torre:t,apto:a,pav:pav(a),ambiente:amb,status:val,fonte:fileName}); });
  }
  return result;
}

// ─── Parser Passantes ────────────────────────────────────────────────────
const PASS_STATUS_MAP={R:"R",C:"C",OK:"OK",Q:"Q","N/V":"N/V",P:"P",F:"F",S:"S"};
const PASS_AMBIENTES=["VARANDA","COZINHA","BWC SERVIÇO","BWC SUÍTE 01 E 02","BWC SUÍTE 03","BWC SUÍTE MASTER","ÁREA DE SERVIÇO","LAVABO"];
function parsePassantes(rows,fileName){
  const summary=parseSummaryByTorre(rows,fileName,"passantes",PASS_STATUS_MAP,/TORRE\s+[AB]/i);
  if(summary) return summary;
  const result=[];
  const hIdxs=rows.reduce((acc,r,i)=>{ if(r[1]&&/apto/i.test(r[1])&&r[2]&&/torre/i.test(r[2])) acc.push(i); return acc; },[]);
  for(const hi of hIdxs){
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i]; if(!r[1]||!/^\d{3,4}$/.test(r[1])) continue;
      const a1=san(r[1]),t1=san(r[2]);
      PASS_AMBIENTES.forEach((amb,j)=>{ const val=san(r[3+j]); if(!val) return; result.push({tipo:"passantes",torre:t1,apto:a1,pav:pav(a1),ambiente:amb,status:val,fonte:fileName}); });
      const a2=san(r[11]),t2=san(r[12]);
      if(a2&&t2&&/^\d{3,4}$/.test(a2))
        PASS_AMBIENTES.forEach((amb,j)=>{ const val=san(r[13+j]); if(!val) return; result.push({tipo:"passantes",torre:t2,apto:a2,pav:pav(a2),ambiente:amb,status:val,fonte:fileName}); });
    }
  }
  return result;
}

// ─── Parser Esquadrias ───────────────────────────────────────────────────
const ESQ_STATUS_MAP={S:"S",C:"C","P.U":"P.U",F:"F",I:"I",E:"E"};
function parseEsquadrias(rows,fileName){
  const summary=parseSummaryByTorre(rows,fileName,"esquadrias",ESQ_STATUS_MAP,/TOTAL TORRE/i);
  if(summary) return summary;
  const result=[];
  const STATUS_VALIDOS=new Set(["E","I","F","C","P.U","S"]);
  function detectBlocos(hRow){
    const blocos=[]; let ci=0;
    while(ci<hRow.length){
      const cell=san(hRow[ci]);
      const isAp=/^apto$/i.test(cell),isEm=cell===""&&ci+1<hRow.length&&/^torre$/i.test(san(hRow[ci+1]));
      if(isAp||isEm){
        const ac=ci,tc=ci+1,as=tc+1; let ae=hRow.length;
        for(let x=as+1;x<hRow.length;x++){
          const cx=san(hRow[x]),cx1=x+1<hRow.length?san(hRow[x+1]):"";
          if(/^apto$/i.test(cx)||(cx===""&&/^torre$/i.test(cx1))){ ae=x; break; }
        }
        blocos.push({aptoCol:ac,torreCol:tc,ambStart:as,ambEnd:ae,headers:hRow.slice(as,ae).map(h=>san(fix(h)))}); ci=ae;
      } else ci++;
    }
    return blocos;
  }
  const hLines=rows.reduce((acc,r,i)=>{
    if(r.some(c=>/^torre$/i.test(san(c)))&&r.filter(c=>san(c).length>2&&!/^apto$|^torre$/i.test(san(c))).length>=3) acc.push(i);
    return acc;
  },[]);
  for(const hi of hLines){
    const blocos=detectBlocos(rows[hi]); if(!blocos.length) continue;
    for(let i=hi+1;i<rows.length;i++){
      const r=rows[i];
      for(const b of blocos){
        const a=san(r[b.aptoCol]),t=san(r[b.torreCol]);
        if(!a||!isApto(a)||!t||!isTorre(t)) continue;
        b.headers.forEach((amb,j)=>{ const val=san(r[b.ambStart+j]); if(!val||!STATUS_VALIDOS.has(val)) return; result.push({tipo:"esquadrias",torre:t.toUpperCase(),apto:a,pav:pav(a),ambiente:fix(amb),status:val,fonte:fileName}); });
      }
    }
  }
  return result;
}

function parseFile(fileName,csvText){
  const rows=parseCSVLines(csvText);
  const tipo=detectTipo(fileName,rows);
  switch(tipo){
    case "shaft":      return parseShafts(rows,fileName);
    case "capiacos":   return parseCapiacos(rows,fileName);
    case "passantes":  return parsePassantes(rows,fileName);
    case "varanda":    return parseCeramicaVaranda(rows,fileName);
    case "somcavo":    return parseSomCavo(rows,fileName);
    case "esquadrias": { const s=parseSummaryByTorre(rows,fileName,"esquadrias",ESQ_STATUS_MAP,/TOTAL TORRE/i); return s||parseEsquadrias(rows,fileName); }
    default: return [];
  }
}

// ─── Classificadores ─────────────────────────────────────────────────────
const CLASS={
  shaft:{ ok:["A"],warn:["FS"],na:["N/A","N/V","?","NF"], labels:{A:"Aberto",FS:"Fech. s/ cerâmica",FC:"Fech. c/ cerâmica","N/A":"N/A","N/V":"N/V",NF:"Não fechado","?":"Não verificado"}, colors:{A:"#16a34a",FS:"#f59e0b",FC:"#2563eb","N/A":"#94a3b8","N/V":"#64748b",NF:"#dc2626","?":"#475569"} },
  capiacos:{ ok:["OK"],warn:["Q","N","Q.I","F"],na:["N/V","N/A"], labels:{OK:"Correto",Q:"Sem queda",N:"Desnivelado","Q.I":"Queda invertida",F:"Falta fachada","N/V":"N/V"}, colors:{OK:"#16a34a",Q:"#dc2626",N:"#f59e0b","Q.I":"#b91c1c",F:"#ea580c","N/V":"#94a3b8"} },
  passantes:{ ok:["OK"],warn:["R","C","Q","P","F","S"],na:["N/V","N/A"], labels:{OK:"Correto",R:"Rente ao piso",C:"PEX chumbado",Q:"Quebrado",P:"Mal-fixado",F:"Falta tubo",S:"Sujeira","N/V":"N/V"}, colors:{OK:"#16a34a",R:"#dc2626",C:"#f59e0b",Q:"#b91c1c",P:"#ea580c",F:"#7c3aed",S:"#0891b2","N/V":"#94a3b8"} },
  esquadrias:{ ok:["E"],warn:["I","F","C","P.U","S"],na:[], labels:{E:"Instalada",I:"Pronto p/ instalar",F:"Falta rejunte/fachada",C:"Falta contramarco","P.U":"P.U incompleto",S:"Contramarco sujo"}, colors:{E:"#16a34a",I:"#2563eb",F:"#dc2626",C:"#f59e0b","P.U":"#ea580c",S:"#64748b"} },
};

function calcTipo(rows,tipo){
  const cl=CLASS[tipo]; if(!cl) return{counts:{},byTorre:[],byApto:[]};
  const counts={},byTorre={},byApto={};
  rows.filter(r=>r.tipo===tipo).forEach(r=>{
    const s=r.status||"?";
    counts[s]=(counts[s]||0)+1;
    const t=r.torre||"?"; if(!byTorre[t]) byTorre[t]={torre:t,counts:{}}; byTorre[t].counts[s]=(byTorre[t].counts[s]||0)+1;
    if(!r.apto) return;
    const key=`${t}-${r.apto}`; if(!byApto[key]) byApto[key]={torre:t,apto:r.apto,pav:r.pav||pav(r.apto),total:0,prob:0,statuses:new Set()};
    byApto[key].total++; if(cl.warn.includes(s)){byApto[key].prob++;byApto[key].statuses.add(s);}
  });
  return{
    counts,
    byTorre:Object.values(byTorre).sort((a,b)=>a.torre.localeCompare(b.torre)),
    byApto:Object.values(byApto).map(x=>({...x,verificacoes:x.total,pct:x.total?`${Math.round(x.prob/x.total*100)}%`:"0%",statuses:[...x.statuses].join(", ")}))
      .sort((a,b)=>a.torre.localeCompare(b.torre)||Number(String(a.apto).match(/\d+/)?.[0]||0)-Number(String(b.apto).match(/\d+/)?.[0]||0)),
  };
}

function calcFvs(rows,servico){
  const filtered=rows.filter(r=>r.tipo_doc==="fvs"&&r.servico===servico);
  const byApto={};
  filtered.forEach(r=>{
    const key=`${r.torre}-${r.apto}`;
    if(!byApto[key]) byApto[key]={torre:r.torre,apto:r.apto,pav:r.pav,data:r.data,A:0,R:0,NV:0,NA:0,crits:new Set()};
    if(r.resultado==="A") byApto[key].A++;
    else if(r.resultado==="R"){byApto[key].R++;byApto[key].crits.add(r.criterio);}
    else if(r.resultado==="N/V") byApto[key].NV++;
    else if(r.resultado==="N/A") byApto[key].NA++;
  });
  const aptoTable=Object.values(byApto).map(x=>{
    const total=x.A+x.R; const tapi=total?Math.round(x.A/total*100):0;
    return{...x,total:x.A+x.R+x.NV+x.NA,tapi,crits:[...x.crits].slice(0,4).join(", "),status:x.R>0?"REPROVADO":"APROVADO"};
  }).sort((a,b)=>a.torre.localeCompare(b.torre)||Number(a.apto)-Number(b.apto));
  const critMap={};
  filtered.forEach(r=>{
    if(!["A","R"].includes(r.resultado)) return;
    const c=r.criterio; if(!critMap[c]) critMap[c]={t:0,r:0};
    critMap[c].t++; if(r.resultado==="R") critMap[c].r++;
  });
  const pareto=Object.entries(critMap).map(([c,v])=>({criterio:c,total:v.t,r:v.r,pct:v.t?Math.round(v.r/v.t*100):0})).sort((a,b)=>b.r-a.r);
  const totA=filtered.filter(r=>r.resultado==="A").length;
  const totR=filtered.filter(r=>r.resultado==="R").length;
  const tapi=(totA+totR)?Math.round(totA/(totA+totR)*100):0;
  return{aptoTable,pareto,totA,totR,tapi};
}

function calcVaranda(rows){
  const filtered=rows.filter(r=>r.tipo_doc==="varanda");
  const byTorre={};
  filtered.forEach(r=>{
    const t=r.torre||"?";
    if(!byTorre[t]) byTorre[t]={torre:t,S:0,C:0,N:0,R:0,total:0};
    byTorre[t][r.status]=(byTorre[t][r.status]||0)+1;
    byTorre[t].total++;
  });
  const torreTable=Object.values(byTorre).map(x=>({...x,executado:x.S+x.C,pendente:x.N+x.R,pct:x.total?Math.round((x.S+x.C)/x.total*100):0})).sort((a,b)=>a.torre.localeCompare(b.torre));
  const byApto={};
  filtered.forEach(r=>{
    const key=`${r.torre}-${r.apto}`;
    if(!byApto[key]) byApto[key]={torre:r.torre,apto:r.apto,pav:r.pav,status:r.status};
  });
  const aptoTable=Object.values(byApto).sort((a,b)=>a.torre.localeCompare(b.torre)||Number(a.apto)-Number(b.apto));
  const total=filtered.length, exec=filtered.filter(r=>["S","C"].includes(r.status)).length;
  return{torreTable,aptoTable,total,exec,pctGeral:total?Math.round(exec/total*100):0};
}

function calcSomCavo(rows){
  const filtered=rows.filter(r=>r.tipo_doc==="somcavo");
  const byAmb={};
  filtered.forEach(r=>{
    if(!byAmb[r.ambiente]) byAmb[r.ambiente]={ambiente:r.ambiente,pedras:0,ambReprov:0,ambTotal:0,nvCount:0};
    if(r.status==="N/V"){byAmb[r.ambiente].nvCount++;return;}
    byAmb[r.ambiente].ambTotal++;
    byAmb[r.ambiente].pedras+=r.pedras||0;
    if(r.status==="R") byAmb[r.ambiente].ambReprov++;
  });
  const paretoAmb=Object.values(byAmb).map(x=>({...x,pct:x.ambTotal?Math.round(x.ambReprov/x.ambTotal*100):0})).sort((a,b)=>b.pedras-a.pedras);
  const byTorre={};
  filtered.forEach(r=>{
    const t=r.torre||"?";
    if(!byTorre[t]) byTorre[t]={torre:t,pedras:0,ambReprov:0,ambTotal:0,nv:0};
    if(r.status==="N/V"){byTorre[t].nv++;return;}
    byTorre[t].ambTotal++;
    byTorre[t].pedras+=r.pedras||0;
    if(r.status==="R") byTorre[t].ambReprov++;
  });
  const torreTable=Object.values(byTorre).map(x=>({...x,pct:x.ambTotal?Math.round(x.ambReprov/x.ambTotal*100):0})).sort((a,b)=>a.torre.localeCompare(b.torre));
  const byApto={};
  filtered.forEach(r=>{
    const key=`${r.torre}-${r.apto}`;
    if(!byApto[key]) byApto[key]={torre:r.torre,apto:r.apto,pav:r.pav,pedras:0,ambReprov:0,ambTotal:0,ambientes:new Set()};
    if(r.status==="N/V") return;
    byApto[key].ambTotal++;
    byApto[key].pedras+=r.pedras||0;
    if(r.status==="R"){byApto[key].ambReprov++;byApto[key].ambientes.add(r.ambiente);}
  });
  const aptoTable=Object.values(byApto).filter(x=>x.pedras>0||x.ambReprov>0).map(x=>({...x,ambientes:[...x.ambientes].join(", ")})).sort((a,b)=>b.pedras-a.pedras);
  const totalPedras=filtered.filter(r=>r.status==="R").reduce((a,r)=>a+(r.pedras||0),0);
  const totalAmbReprov=filtered.filter(r=>r.status==="R").length;
  const totalAmbVerif=filtered.filter(r=>r.status!=="N/V").length;
  return{paretoAmb,torreTable,aptoTable,totalPedras,totalAmbReprov,totalAmbVerif};
}

// ─── Componentes UI ───────────────────────────────────────────────────────
function BarChart({counts,labels,colors}){
  const entries=Object.entries(counts).filter(([k,v])=>v>0&&labels?.[k]).sort((a,b)=>b[1]-a[1]);
  const total=entries.reduce((a,[,v])=>a+v,0)||1,maxVal=entries[0]?.[1]||1;
  return(<div style={{display:"flex",flexDirection:"column",gap:8}}>
    {entries.map(([k,v],idx)=>{ const color=colors?.[k]||PALETTE[idx%PALETTE.length]; return(
      <div key={k} style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:160,fontSize:12,color:C.white,textAlign:"right",flexShrink:0}}>{labels[k]}</div>
        <div style={{flex:1,background:"#0f172a",borderRadius:4,height:24,position:"relative"}}>
          <div style={{width:`${(v/maxVal)*100}%`,background:color,height:"100%",borderRadius:4,transition:"width .4s"}}/>
          <span style={{position:"absolute",right:8,top:4,fontSize:11,color:C.white,fontWeight:"bold"}}>{v} ({Math.round(v/total*100)}%)</span>
        </div>
      </div>);})}
  </div>);
}

function ParetoFvs({pareto}){
  const max=Math.max(...pareto.map(p=>p.r),1);
  return(<div style={{display:"flex",flexDirection:"column",gap:6}}>
    {pareto.map((p,i)=>(
      <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
        <div style={{width:220,fontSize:12,color:C.white,textAlign:"right",flexShrink:0}}>{p.criterio}</div>
        <div style={{flex:1,background:"#0f172a",borderRadius:4,height:22,position:"relative"}}>
          <div style={{width:`${(p.r/max)*100}%`,background:p.pct>=50?C.bad:p.pct>=25?C.orange:C.ok,height:"100%",borderRadius:4}}/>
          <span style={{position:"absolute",right:8,top:3,fontSize:11,color:C.white,fontWeight:"bold"}}>{p.r} R ({p.pct}%)</span>
        </div>
      </div>))}
  </div>);
}

function VarandaProgressBars({torreTable}){
  return(<div style={{display:"flex",flexDirection:"column",gap:10}}>
    {torreTable.map((t,i)=>(
      <div key={i} style={{display:"flex",alignItems:"center",gap:12}}>
        <div style={{width:80,fontSize:13,color:C.white,fontWeight:"bold",flexShrink:0}}>Torre {t.torre}</div>
        <div style={{flex:1,background:"#0f172a",borderRadius:6,height:28,position:"relative",overflow:"hidden"}}>
          <div style={{width:`${t.pct}%`,background:t.pct>=80?C.ok:t.pct>=50?C.warn:C.bad,height:"100%",borderRadius:6,transition:"width .4s"}}/>
          <span style={{position:"absolute",right:10,top:5,fontSize:12,color:C.white,fontWeight:"bold"}}>{t.executado}/{t.total} ({t.pct}%)</span>
        </div>
        <div style={{width:130,fontSize:11,color:C.muted,flexShrink:0}}>S:{t.S} C:{t.C} N:{t.N} R:{t.R}</div>
      </div>))}
  </div>);
}

const KPI=({label,value,sub,color})=>(
  <div style={{background:C.card,borderRadius:14,padding:"18px 22px",borderLeft:`4px solid ${color||C.accent}`}}>
    <div style={{fontSize:10,textTransform:"uppercase",color:C.muted,fontWeight:"bold",marginBottom:6}}>{label}</div>
    <div style={{fontSize:30,fontWeight:"bold",color:C.white,marginBottom:3}}>{value}</div>
    <div style={{fontSize:12,color:C.muted}}>{sub}</div>
  </div>
);
const Box=({title,children,action})=>(
  <div style={{background:C.card,borderRadius:18,padding:22,marginTop:20}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
      <h2 style={{margin:0,color:C.white,fontSize:16}}>{title}</h2>{action}
    </div>{children}
  </div>
);
const TH=({c})=><th style={{borderBottom:`1px solid ${C.border}`,padding:"9px 11px",textAlign:"left",color:C.muted,fontSize:11,textTransform:"uppercase",background:C.card}}>{c}</th>;
const TD=({c,bold,color,children})=><td style={{borderBottom:`1px solid #0f172a`,padding:"9px 11px",color:color||C.white,fontWeight:bold?"bold":"normal",fontSize:12}}>{children!==undefined?children:c}</td>;
const Btn=({children,onClick,color})=>(<button onClick={onClick} style={{background:color||C.blue,color:"white",border:"none",borderRadius:8,padding:"9px 15px",cursor:"pointer",fontWeight:"bold",fontSize:13}}>{children}</button>);

function TabelaApto({data}){
  if(!data.byApto.length) return <p style={{color:C.muted}}>Sem dados por apartamento.</p>;
  return(<div style={{overflowX:"auto"}}>
    <table style={{width:"100%",borderCollapse:"collapse",minWidth:700}}>
      <thead><tr><TH c="Torre"/><TH c="Apto"/><TH c="Pav"/><TH c="Total"/><TH c="Problemas"/><TH c="% Prob."/><TH c="Status encontrados"/></tr></thead>
      <tbody>{data.byApto.map((r,i)=>(
        <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
          <TD c={r.torre} color={C.accent} bold/><TD c={r.apto} bold/><TD c={r.pav}/><TD c={r.total}/>
          <TD c={r.prob} color={r.prob>0?C.bad:C.ok}/><TD c={r.pct}/><TD c={r.statuses||"—"} color={r.statuses?C.warn:C.muted}/>
        </tr>))}</tbody>
    </table>
  </div>);
}

function TabelaTorre({data}){
  if(!data.byTorre.length) return <p style={{color:C.muted}}>Sem dados.</p>;
  const keys=[...new Set(data.byTorre.flatMap(t=>Object.keys(t.counts)))];
  return(<div style={{overflowX:"auto"}}>
    <table style={{width:"100%",borderCollapse:"collapse"}}>
      <thead><tr><TH c="Torre"/>{keys.map(k=><TH key={k} c={k}/>)}<TH c="Total"/></tr></thead>
      <tbody>{data.byTorre.map((t,i)=>{
        const total=Object.values(t.counts).reduce((a,b)=>a+b,0);
        return(<tr key={i} style={{background:i%2===0?"transparent":C.row}}>
          <TD c={`Torre ${t.torre}`} color={C.accent} bold/>
          {keys.map(k=><TD key={k} c={t.counts[k]||0}/>)}<TD c={total} bold/>
        </tr>);})}</tbody>
    </table>
  </div>);
}

function TabelaFvsApto({aptoTable}){
  if(!aptoTable.length) return <p style={{color:C.muted}}>Nenhum apto carregado. Envie os arquivos DOCX das FVS.</p>;
  return(<div style={{overflowX:"auto"}}>
    <table style={{width:"100%",borderCollapse:"collapse",minWidth:800}}>
      <thead><tr><TH c="Torre"/><TH c="Apto"/><TH c="Pav"/><TH c="Data"/><TH c="A"/><TH c="R"/><TH c="N/V"/><TH c="N/A"/><TH c="TAPI"/><TH c="Critérios reprovados"/><TH c="Status"/></tr></thead>
      <tbody>{aptoTable.map((r,i)=>(
        <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
          <TD c={r.torre} color={C.accent} bold/><TD c={r.apto} bold/><TD c={r.pav}/><TD c={r.data||"—"}/>
          <TD c={r.A} color={C.ok}/><TD c={r.R} color={r.R>0?C.bad:C.muted}/>
          <TD c={r.NV} color={C.warn}/><TD c={r.NA} color={C.muted}/>
          <TD c={`${r.tapi}%`} bold color={r.tapi>=85?C.ok:C.bad}/>
          <TD c={r.crits||"—"} color={r.crits?C.warn:C.muted}/>
          <TD bold color={r.status==="REPROVADO"?C.bad:C.ok}>{r.status}</TD>
        </tr>))}</tbody>
    </table>
  </div>);
}

function TabelaVarandaApto({aptoTable}){
  if(!aptoTable.length) return <p style={{color:C.muted}}>Sem dados.</p>;
  const sc={S:C.ok,C:C.blue,N:C.bad,R:C.orange};
  const sl={S:"Finalizado",C:"Crédito (finalizado)",N:"Pendente",R:"Reforma (pendente)"};
  return(<div style={{overflowX:"auto"}}>
    <table style={{width:"100%",borderCollapse:"collapse",minWidth:500}}>
      <thead><tr><TH c="Torre"/><TH c="Apto"/><TH c="Pav"/><TH c="Status"/></tr></thead>
      <tbody>{aptoTable.map((r,i)=>(
        <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
          <TD c={r.torre} color={C.accent} bold/><TD c={r.apto} bold/><TD c={r.pav}/>
          <TD bold color={sc[r.status]||C.muted}>{sl[r.status]||r.status}</TD>
        </tr>))}</tbody>
    </table>
  </div>);
}

function exportCSV(allRows,fvsRows,varandaRows,somCavoRows){
  const lines=["tipo,servico,torre,apto,pav,ambiente,criterio,resultado,status,pedras,fonte"];
  [...allRows,...fvsRows,...varandaRows,...somCavoRows].forEach(r=>lines.push(
    `${r.tipo_doc||r.tipo||""},${r.servico||""},${r.torre||""},${r.apto||""},${r.pav||""},${r.ambiente||""},${r.criterio||""},${r.resultado||""},${r.status||""},${r.pedras??""},${r.fonte||""}`
  ));
  const blob=new Blob([lines.join("\n")],{type:"text/csv;charset=utf-8"});
  const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download="fvs_completo.csv"; a.click(); URL.revokeObjectURL(url);
}

// ─── App ──────────────────────────────────────────────────────────────────
export default function App(){
  const [allRows,setAllRows]                       = useState([]);
  const [fvsRows,setFvsRows]                       = useState([]);
  const [varandaRows,setVarandaRows]               = useState([]);
  const [somCavoRows,setSomCavoRows]               = useState([]);
  const [status,setStatus]                         = useState("Envie CSV/XLSX (planilhas) ou DOCX/PDF (FVS).");
  const [fileNames,setFileNames]                   = useState([]);
  const [errors,setErrors]                         = useState([]);
  const [mainTab,setMainTab]                       = useState("planilhas");
  const [planTab,setPlanTab]                       = useState("shafts");
  const [fvsTab,setFvsTab]                         = useState("ceramica");
  const [torreFilter,setTorreFilter]               = useState("TODAS");
  const [fvsTorreFilter,setFvsTorreFilter]         = useState("TODAS");
  const [varandaTorreFilter,setVarandaTorreFilter] = useState("TODAS");
  const [somCavoTorreFilter,setSomCavoTorreFilter] = useState("TODAS");

  const handleFile = useCallback(async e=>{
    const files=Array.from(e.target.files||[]);
    if(!files.length) return;
    setAllRows([]); setFvsRows([]); setVarandaRows([]); setSomCavoRows([]); setErrors([]);
    setFileNames(files.map(f=>f.name));
    setStatus("Processando...");
    let planilhas=[], fvs=[], varanda=[], somcavo=[], errs=[];

    for(const file of files){
      try{
        const ext=file.name.split(".").pop()?.toLowerCase();
        if(ext==="csv"){
          const text=await file.text();
          const parsed=parseFile(file.name,text);
          const sc=parsed.filter(r=>r.tipo_doc==="somcavo");
          const va=parsed.filter(r=>r.tipo_doc==="varanda");
          const pl=parsed.filter(r=>r.tipo_doc!=="somcavo"&&r.tipo_doc!=="varanda");
          somcavo=[...somcavo,...sc];
          varanda=[...varanda,...va];
          if(pl.length) planilhas=[...planilhas,...pl];
          else if(!sc.length&&!va.length) errs.push(`${file.name}: nenhuma linha reconhecida.`);
        } else if(ext==="xlsx"||ext==="xls"){
          const buf=await file.arrayBuffer();
          const wb=XLSX.read(buf,{type:"array"});
          for(const sn of wb.SheetNames){
            const csv=XLSX.utils.sheet_to_csv(wb.Sheets[sn]);
            const parsed=parseFile(`${file.name} ${sn}`,csv);
            somcavo=[...somcavo,...parsed.filter(r=>r.tipo_doc==="somcavo")];
            varanda=[...varanda,...parsed.filter(r=>r.tipo_doc==="varanda")];
            planilhas=[...planilhas,...parsed.filter(r=>r.tipo_doc!=="somcavo"&&r.tipo_doc!=="varanda")];
          }
        } else if(ext==="docx"){
          const buf=await file.arrayBuffer();
          const res=await mammoth.extractRawText({arrayBuffer:buf});
          const text=res.value||"";
          const fvsParsed=parseFvsDocx(file.name,text);
          if(fvsParsed.length) fvs=[...fvs,...fvsParsed];
          else{ const parsed=parseFile(file.name,text); if(parsed.length) planilhas=[...planilhas,...parsed]; else errs.push(`${file.name}: nenhum dado reconhecido.`); }
        } else if(ext==="pdf"){
          const buf=await file.arrayBuffer();
          const pdf=await pdfjsLib.getDocument({data:buf}).promise;
          let txt="";
          for(let p=1;p<=pdf.numPages;p++){
            const page=await pdf.getPage(p); const ct=await page.getTextContent();
            let lastY=null;
            for(const item of ct.items){
              const y=item.transform?.[5];
              if(lastY!==null&&Math.abs(y-lastY)>5) txt+="\n";
              txt+=item.str; lastY=y;
            }
            txt+="\n";
          }
          const fvsParsed=parseFvsDocx(file.name,txt);
          if(fvsParsed.length){ fvs=[...fvs,...fvsParsed]; }
          else{
            const pdfRows=parseCSVLines(txt.replace(/[ \t]+/g,","));
            const vParsed=parseCeramicaVaranda(pdfRows,file.name);
            if(vParsed.length) varanda=[...varanda,...vParsed];
            else{ const parsed=parseFile(file.name,txt); if(parsed.length) planilhas=[...planilhas,...parsed]; else errs.push(`${file.name}: nenhum dado reconhecido.`); }
          }
        } else errs.push(`${file.name}: formato não suportado.`);
      }catch(err){ console.error(err); errs.push(`${file.name}: erro — ${err.message}`); }
    }
    setAllRows(planilhas); setFvsRows(fvs); setVarandaRows(varanda); setSomCavoRows(somcavo); setErrors(errs);
    setStatus(`${files.length} arquivo(s) processado(s). ${planilhas.length} planilha + ${fvs.length} FVS + ${varanda.length} varanda + ${somcavo.length} som cavo.`);
    if(fvs.length>0&&planilhas.length===0) setMainTab("fvs");
    else if(planilhas.length>0||varanda.length>0||somcavo.length>0) setMainTab("planilhas");
  },[]);

  const torres=useMemo(()=>["TODAS",...[...new Set(allRows.map(r=>r.torre).filter(Boolean))].sort()],[allRows]);
  const scoped=useMemo(()=>torreFilter==="TODAS"?allRows:allRows.filter(r=>r.torre===torreFilter),[allRows,torreFilter]);
  const shaftData =useMemo(()=>calcTipo(scoped,"shaft"),     [scoped]);
  const capData   =useMemo(()=>calcTipo(scoped,"capiacos"),  [scoped]);
  const passData  =useMemo(()=>calcTipo(scoped,"passantes"), [scoped]);
  const esqData   =useMemo(()=>calcTipo(scoped,"esquadrias"),[scoped]);
  const shaftTotal=Object.values(shaftData.counts).reduce((a,b)=>a+b,0)-(shaftData.counts["N/A"]||0);
  const shaftAberto=shaftData.counts["A"]||0;
  const capTotal=Object.values(capData.counts).reduce((a,b)=>a+b,0);
  const capProb=(capData.counts["Q"]||0)+(capData.counts["N"]||0)+(capData.counts["Q.I"]||0)+(capData.counts["F"]||0);
  const passTotal=Object.values(passData.counts).reduce((a,b)=>a+b,0);
  const passProb=passTotal-(passData.counts["OK"]||0)-(passData.counts["N/V"]||0);
  const esqTotal=Object.values(esqData.counts).reduce((a,b)=>a+b,0);
  const esqInst=esqData.counts["E"]||0;

  const fvsServicos=["ceramica","contrapiso","porta","esquadria_alum"];
  const fvsTorres=useMemo(()=>["TODAS",...[...new Set(fvsRows.map(r=>r.torre).filter(Boolean))].sort()],[fvsRows]);
  const fvsScopedRows=useMemo(()=>fvsTorreFilter==="TODAS"?fvsRows:fvsRows.filter(r=>r.torre===fvsTorreFilter),[fvsRows,fvsTorreFilter]);
  const fvsCurrent=useMemo(()=>calcFvs(fvsScopedRows,fvsTab),[fvsScopedRows,fvsTab]);

  const varandaTorres=useMemo(()=>["TODAS",...[...new Set(varandaRows.map(r=>r.torre).filter(Boolean))].sort()],[varandaRows]);
  const varandaScoped=useMemo(()=>varandaTorreFilter==="TODAS"?varandaRows:varandaRows.filter(r=>r.torre===varandaTorreFilter),[varandaRows,varandaTorreFilter]);
  const varandaData=useMemo(()=>calcVaranda(varandaScoped),[varandaScoped]);

  const somCavoTorres=useMemo(()=>["TODAS",...[...new Set(somCavoRows.map(r=>r.torre).filter(Boolean))].sort()],[somCavoRows]);
  const somCavoScoped=useMemo(()=>somCavoTorreFilter==="TODAS"?somCavoRows:somCavoRows.filter(r=>r.torre===somCavoTorreFilter),[somCavoRows,somCavoTorreFilter]);
  const somCavoData=useMemo(()=>calcSomCavo(somCavoScoped),[somCavoScoped]);

  const MAIN_TABS=[{id:"planilhas",label:"📊 Planilhas"},{id:"fvs",label:"📋 FVS"}];
  const PLAN_TABS=[
    {id:"shafts",    label:"🔲 Shafts"},
    {id:"capiacos",  label:"🏗 Capiaços"},
    {id:"passantes", label:"🔧 Passantes"},
    {id:"esquadrias",label:"🪟 Esquadrias"},
    {id:"varanda",   label:"🟫 Cerâmica Varanda"},
    {id:"somcavo",   label:"🔊 Som Cavo"},
  ];
  const selStyle={background:"#0f172a",color:C.white,border:`1px solid ${C.border}`,borderRadius:7,padding:"7px 11px",fontSize:12};
  const torreLabel=t=>t==="TODAS"?"Todas as torres":`Torre ${t}`;

  return (
    <div style={{minHeight:"100vh",background:C.bg,color:C.white,padding:"28px 20px",fontFamily:"'Inter',Arial,sans-serif"}}>
      <div style={{maxWidth:1280,margin:"0 auto"}}>

        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:6}}>
          <span style={{background:"#082f49",color:C.accent,borderRadius:999,padding:"3px 12px",fontSize:12,fontWeight:"bold"}}>FVS Qualidade</span>
          <span style={{color:C.muted,fontSize:12}}>v4.4</span>
        </div>
        <h1 style={{fontSize:32,margin:"0 0 4px",color:C.white}}>Dashboard de Verificação de Serviços</h1>
        <p style={{color:C.muted,margin:"0 0 22px",fontSize:13}}>Shafts · Capiaços · Passantes · Esquadrias · Cerâmica Varanda · Som Cavo · Contrapiso · Portas</p>

        {/* Upload */}
        <div style={{background:C.card,borderRadius:18,padding:20,marginBottom:20,border:`1px solid ${C.border}`}}>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
            <label style={{background:C.blue,color:"white",borderRadius:9,padding:"9px 16px",cursor:"pointer",fontWeight:"bold",fontSize:13}}>
              📂 Carregar Arquivos
              <input type="file" multiple accept=".csv,.xlsx,.xls,.docx,.pdf" onChange={handleFile} style={{display:"none"}}/>
            </label>
            {(allRows.length>0||fvsRows.length>0||varandaRows.length>0||somCavoRows.length>0)&&
              <Btn onClick={()=>exportCSV(allRows,fvsRows,varandaRows,somCavoRows)} color={C.purple}>⬇ Exportar CSV</Btn>}
            {allRows.length>0&&(
              <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:11,color:C.muted}}>TORRE</span>
                <select value={torreFilter} onChange={e=>setTorreFilter(e.target.value)} style={selStyle}>
                  {torres.map(t=><option key={t} value={t}>{torreLabel(t)}</option>)}
                </select>
              </div>
            )}
          </div>
          {fileNames.length>0&&<div style={{marginTop:10,color:C.muted,fontSize:12}}>{fileNames.join(" · ")}</div>}
          <div style={{marginTop:6,fontSize:12,color:errors.length?C.bad:C.muted}}>{status}</div>
          {errors.map((e,i)=><div key={i} style={{color:C.bad,fontSize:11,marginTop:3}}>⚠ {e}</div>)}
        </div>

        {/* Main tabs */}
        <div style={{display:"flex",gap:3,borderBottom:`1px solid ${C.border}`}}>
          {MAIN_TABS.map(t=>(
            <button key={t.id} onClick={()=>setMainTab(t.id)} style={{background:mainTab===t.id?C.blue:"transparent",color:mainTab===t.id?C.white:C.muted,border:"none",borderRadius:"8px 8px 0 0",padding:"10px 20px",cursor:"pointer",fontWeight:mainTab===t.id?"bold":"normal",fontSize:14}}>
              {t.label}
            </button>))}
        </div>

        {/* ── PLANILHAS ── */}
        {mainTab==="planilhas" && (
          <div>
            {(allRows.length>0||varandaRows.length>0||somCavoRows.length>0) ? (
              <div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:14,marginTop:20,marginBottom:20}}>
                  <KPI label="Shafts Abertos"        value={`${shaftAberto}/${shaftTotal}`} sub={`${shaftTotal?Math.round(shaftAberto/shaftTotal*100):0}%`} color={C.ok}/>
                  <KPI label="Capiaços c/ Problema"  value={capProb}  sub={`de ${capTotal}`} color={capProb>0?C.bad:C.ok}/>
                  <KPI label="Passantes c/ Problema" value={passProb} sub={`de ${passTotal}`} color={passProb>0?C.bad:C.ok}/>
                  <KPI label="Esquadrias Instaladas" value={`${esqInst}/${esqTotal}`} sub={`${esqTotal?Math.round(esqInst/esqTotal*100):0}%`} color={C.blue}/>
                  <KPI label="Varanda c/ Cerâmica"   value={`${varandaData.exec}/${varandaData.total}`} sub={`${varandaData.pctGeral}% executado`} color={varandaData.pctGeral>=80?C.ok:C.warn}/>
                  <KPI label="Som Cavo — Pedras"     value={somCavoData.totalPedras} sub={`${somCavoData.totalAmbReprov} ambientes`} color={C.bad}/>
                </div>
                <div style={{display:"flex",gap:3,borderBottom:`1px solid ${C.border}`,flexWrap:"wrap"}}>
                  {PLAN_TABS.map(t=>(<button key={t.id} onClick={()=>setPlanTab(t.id)} style={{background:planTab===t.id?C.blue:"transparent",color:planTab===t.id?C.white:C.muted,border:"none",borderRadius:"8px 8px 0 0",padding:"9px 16px",cursor:"pointer",fontWeight:planTab===t.id?"bold":"normal",fontSize:13}}>{t.label}</button>))}
                </div>

                {planTab==="shafts" && (
                  <div>
                    <Box title={`Distribuição — Shafts — ${torreLabel(torreFilter)}`}><BarChart counts={shaftData.counts} labels={CLASS.shaft.labels} colors={CLASS.shaft.colors}/></Box>
                    <Box title="Por Torre — Shafts"><TabelaTorre data={shaftData}/></Box>
                    <Box title={`Por Apartamento — Shafts — ${torreLabel(torreFilter)}`}><TabelaApto data={shaftData}/></Box>
                  </div>
                )}
                {planTab==="capiacos" && (
                  <div>
                    <Box title={`Distribuição — Capiaços — ${torreLabel(torreFilter)}`}><BarChart counts={capData.counts} labels={CLASS.capiacos.labels} colors={CLASS.capiacos.colors}/></Box>
                    <Box title="Por Torre — Capiaços"><TabelaTorre data={capData}/></Box>
                    <Box title={`Por Apartamento — Capiaços — ${torreLabel(torreFilter)}`}><TabelaApto data={capData}/></Box>
                  </div>
                )}
                {planTab==="passantes" && (
                  <div>
                    <Box title={`Distribuição — Passantes — ${torreLabel(torreFilter)}`}><BarChart counts={passData.counts} labels={CLASS.passantes.labels} colors={CLASS.passantes.colors}/></Box>
                    <Box title="Por Torre — Passantes"><TabelaTorre data={passData}/></Box>
                    <Box title={`Por Apartamento — Passantes — ${torreLabel(torreFilter)}`}><TabelaApto data={passData}/></Box>
                  </div>
                )}
                {planTab==="esquadrias" && (
                  <div>
                    <Box title={`Distribuição — Esquadrias — ${torreLabel(torreFilter)}`}><BarChart counts={esqData.counts} labels={CLASS.esquadrias.labels} colors={CLASS.esquadrias.colors}/></Box>
                    <Box title="Por Torre — Esquadrias"><TabelaTorre data={esqData}/></Box>
                    <Box title={`Por Apartamento — Esquadrias — ${torreLabel(torreFilter)}`}><TabelaApto data={esqData}/></Box>
                  </div>
                )}
                {planTab==="varanda" && (
                  <div>
                    <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",gap:8,marginTop:16,marginBottom:4}}>
                      {varandaRows.length>0 && <>
                        <span style={{fontSize:11,color:C.muted}}>TORRE</span>
                        <select value={varandaTorreFilter} onChange={e=>setVarandaTorreFilter(e.target.value)} style={selStyle}>
                          {varandaTorres.map(t=><option key={t} value={t}>{torreLabel(t)}</option>)}
                        </select>
                      </>}
                    </div>
                    <Box title={`Progresso — Cerâmica Varanda — ${torreLabel(varandaTorreFilter)}`}>
                      {varandaData.torreTable.length ? <VarandaProgressBars torreTable={varandaData.torreTable}/> : <p style={{color:C.muted}}>Sem dados. Carregue o arquivo de mapeamento de cerâmica varanda.</p>}
                    </Box>
                    {varandaData.aptoTable.length>0 && (
                      <Box title={`Por Apartamento — Cerâmica Varanda — ${torreLabel(varandaTorreFilter)}`}>
                        <TabelaVarandaApto aptoTable={varandaData.aptoTable}/>
                      </Box>
                    )}
                  </div>
                )}
                {planTab==="somcavo" && (
                  <div>
                    <div style={{display:"flex",justifyContent:"flex-end",alignItems:"center",gap:8,marginTop:16,marginBottom:4}}>
                      {somCavoRows.length>0 && <>
                        <span style={{fontSize:11,color:C.muted}}>TORRE</span>
                        <select value={somCavoTorreFilter} onChange={e=>setSomCavoTorreFilter(e.target.value)} style={selStyle}>
                          {somCavoTorres.map(t=><option key={t} value={t}>{torreLabel(t)}</option>)}
                        </select>
                      </>}
                    </div>
                    {somCavoRows.length>0 ? (
                      <div>
                        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:14,marginBottom:4}}>
                          <KPI label="Total pedras reprovadas" value={somCavoData.totalPedras} sub="incidências de som cavo" color={C.bad}/>
                          <KPI label="Ambientes reprovados" value={somCavoData.totalAmbReprov} sub={`de ${somCavoData.totalAmbVerif} verificados`} color={C.orange}/>
                          <KPI label="% ambientes c/ som cavo" value={`${somCavoData.totalAmbVerif?Math.round(somCavoData.totalAmbReprov/somCavoData.totalAmbVerif*100):0}%`} sub="sobre verificados" color={C.warn}/>
                        </div>
                        <Box title={`Pedras reprovadas por ambiente — ${torreLabel(somCavoTorreFilter)}`}>
                          <div style={{display:"flex",flexDirection:"column",gap:8}}>
                            {somCavoData.paretoAmb.filter(p=>p.pedras>0).map((p,i)=>{
                              const max=somCavoData.paretoAmb.filter(x=>x.pedras>0)[0]?.pedras||1;
                              return(
                                <div key={i} style={{display:"flex",alignItems:"center",gap:10}}>
                                  <div style={{width:160,fontSize:12,color:C.white,textAlign:"right",flexShrink:0}}>{p.ambiente}</div>
                                  <div style={{flex:1,background:"#0f172a",borderRadius:4,height:24,position:"relative"}}>
                                    <div style={{width:`${(p.pedras/max)*100}%`,background:p.pct>=50?C.bad:p.pct>=25?C.orange:C.warn,height:"100%",borderRadius:4}}/>
                                    <span style={{position:"absolute",right:8,top:4,fontSize:11,color:C.white,fontWeight:"bold"}}>{p.pedras} pedras · {p.ambReprov}/{p.ambTotal} amb. ({p.pct}%)</span>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </Box>
                        <Box title={`Por Torre — Som Cavo — ${torreLabel(somCavoTorreFilter)}`}>
                          <div style={{overflowX:"auto"}}>
                            <table style={{width:"100%",borderCollapse:"collapse"}}>
                              <thead><tr><TH c="Torre"/><TH c="Pedras reprov."/><TH c="Amb. reprov."/><TH c="Amb. verif."/><TH c="% amb. c/ som cavo"/><TH c="N/V"/></tr></thead>
                              <tbody>{somCavoData.torreTable.map((t,i)=>(
                                <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
                                  <TD c={`Torre ${t.torre}`} color={C.accent} bold/>
                                  <TD c={t.pedras} color={t.pedras>0?C.bad:C.ok} bold/>
                                  <TD c={t.ambReprov} color={t.ambReprov>0?C.orange:C.ok}/>
                                  <TD c={t.ambTotal}/>
                                  <TD c={`${t.pct}%`} bold color={t.pct>=50?C.bad:t.pct>=25?C.orange:C.ok}/>
                                  <TD c={t.nv} color={C.muted}/>
                                </tr>))}
                              </tbody>
                            </table>
                          </div>
                        </Box>
                        <Box title={`Aptos com maior incidência — ${torreLabel(somCavoTorreFilter)}`}>
                          <div style={{overflowX:"auto"}}>
                            <table style={{width:"100%",borderCollapse:"collapse",minWidth:700}}>
                              <thead><tr><TH c="Torre"/><TH c="Apto"/><TH c="Pav"/><TH c="Pedras reprov."/><TH c="Amb. reprov."/><TH c="Ambientes c/ problema"/></tr></thead>
                              <tbody>{somCavoData.aptoTable.map((r,i)=>(
                                <tr key={i} style={{background:i%2===0?"transparent":C.row}}>
                                  <TD c={r.torre} color={C.accent} bold/>
                                  <TD c={r.apto} bold/>
                                  <TD c={r.pav}/>
                                  <TD c={r.pedras} color={C.bad} bold/>
                                  <TD c={r.ambReprov} color={C.orange}/>
                                  <TD c={r.ambientes||"—"} color={C.warn}/>
                                </tr>))}
                              </tbody>
                            </table>
                          </div>
                        </Box>
                      </div>
                    ) : (
                      <div style={{textAlign:"center",padding:"40px",color:C.muted}}>
                        <div style={{fontSize:40,marginBottom:12}}>🔊</div>
                        <div>Carregue o arquivo de mapeamento FVS cerâmica (som cavo).</div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            ) : (
              <div style={{textAlign:"center",padding:"40px",color:C.muted}}>
                <div style={{fontSize:40,marginBottom:12}}>📊</div>
                <div>Carregue os arquivos CSV/XLSX das planilhas.</div>
              </div>
            )}
          </div>
        )}

        {/* ── FVS ── */}
        {mainTab==="fvs" && (
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,paddingTop:4}}>
              <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
                {fvsServicos.map(s=>(<button key={s} onClick={()=>setFvsTab(s)} style={{background:fvsTab===s?C.blue:"transparent",color:fvsTab===s?C.white:C.muted,border:"none",borderRadius:"8px 8px 0 0",padding:"9px 16px",cursor:"pointer",fontWeight:fvsTab===s?"bold":"normal",fontSize:13}}>{FVS_SERVICO_LABELS[s]}</button>))}
              </div>
              {fvsRows.length>0 && (
                <div style={{display:"flex",alignItems:"center",gap:8,paddingBottom:4}}>
                  <span style={{fontSize:11,color:C.muted}}>TORRE</span>
                  <select value={fvsTorreFilter} onChange={e=>setFvsTorreFilter(e.target.value)} style={selStyle}>
                    {fvsTorres.map(t=><option key={t} value={t}>{torreLabel(t)}</option>)}
                  </select>
                </div>
              )}
            </div>
            <div style={{borderBottom:`1px solid ${C.border}`,marginBottom:0}}/>
            {fvsRows.length>0 ? (
              <div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:14,marginTop:20,marginBottom:4}}>
                  <KPI label="Aptos inspecionados" value={fvsCurrent.aptoTable.length} sub={FVS_SERVICO_LABELS[fvsTab]} color={C.accent}/>
                  <KPI label="TAPI" value={`${fvsCurrent.tapi}%`} sub={fvsCurrent.tapi>=85?"✅ Meta atingida":"❌ Abaixo de 85%"} color={fvsCurrent.tapi>=85?C.ok:C.bad}/>
                  <KPI label="Aprovações (A)" value={fvsCurrent.totA} sub="verificações aprovadas" color={C.ok}/>
                  <KPI label="Reprovações (R)" value={fvsCurrent.totR} sub="verificações reprovadas" color={fvsCurrent.totR>0?C.bad:C.ok}/>
                </div>
                <Box title={`Pareto — ${FVS_SERVICO_LABELS[fvsTab]} — ${torreLabel(fvsTorreFilter)}`}>
                  {fvsCurrent.pareto.length ? <ParetoFvs pareto={fvsCurrent.pareto}/> : <p style={{color:C.muted}}>Sem dados suficientes.</p>}
                </Box>
                <Box title={`Por Apartamento — ${FVS_SERVICO_LABELS[fvsTab]} — ${torreLabel(fvsTorreFilter)}`}>
                  <TabelaFvsApto aptoTable={fvsCurrent.aptoTable}/>
                </Box>
              </div>
            ) : (
              <div style={{textAlign:"center",padding:"40px",color:C.muted}}>
                <div style={{fontSize:40,marginBottom:12}}>📋</div>
                <div>Carregue os arquivos DOCX/PDF das FVS.</div>
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}
