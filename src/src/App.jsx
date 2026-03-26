// ─── DEPENDÊNCIAS ──────────────────────────────────────────────────────────
// npm install xlsx mammoth pdfjs-dist
// Para pdfjs, adicione no vite.config.js:
//   optimizeDeps: { exclude: ["pdfjs-dist"] }
// E copie o worker para /public:
//   node_modules/pdfjs-dist/build/pdf.worker.min.mjs → public/pdf.worker.min.mjs
// ──────────────────────────────────────────────────────────────────────────

import { useMemo, useState, useCallback } from "react";
import * as XLSX from "xlsx";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

// ─── Constantes ────────────────────────────────────────────────────────────
const AMBIENTES_PADRAO = [
  "SALA","VARANDA","COZINHA","ÁREA SERV","DEPÓSITO","BWC SERV","LAVABO",
  "SUÍTE 01","BWC SUÍTE 01 E 02","SUÍTE 02","SUÍTE 03","BWC SUÍTE 03","SUÍTE 04","BWC SUÍTE 04",
];

const CRITERIOS_POR_SERVICO = {
  ceramica:   ["Planicidade","Peças sem trincas e lascas","Declividade em direção aos ralos","Rejunte","Sem excesso de argamassa","Dupla colagem","Terminalidade","Presença de som cavo","Limpeza"],
  esquadrias: ["Folga perimetral","Vedação/silicone","Fixação/parafusos","Nivelamento","Funcionamento de abertura","Acabamento/pintura","Limpeza"],
  capiacos:   ["Caimento para fora","Fixação","Rejunte","Trincas/lascas","Impermeabilização","Limpeza"],
  tubos:      ["Posicionamento","Prumo/nível","Fixação","Proteção/tampão","Limpeza"],
};

const SERVICOS_LABELS = {
  ceramica:   "Cerâmica",
  esquadrias: "Esquadrias",
  capiacos:   "Capiaços",
  tubos:      "Tubos Passantes",
};

const CORES = {
  bg:       "#0f172a",
  card:     "#1e293b",
  approved: "#16a34a",
  reproved: "#dc2626",
  nv:       "#f59e0b",
  na:       "#94a3b8",
  blue:     "#2563eb",
  textLight:"#f8fafc",
  border:   "#334155",
  accent:   "#38bdf8",
};

// ─── Helpers ───────────────────────────────────────────────────────────────
const normResult = (v) => {
  const s = String(v || "").trim().toUpperCase();
  if (s === "A") return "A";
  if (s === "R") return "R";
  if (["N/V","NV"].includes(s)) return "NV";
  if (["N/A","NA","-"].includes(s)) return "NA";
  return s;
};
const sanitize  = (v) => String(v || "").replace(/\s+/g, " ").trim();
const inferPav  = (apto) => { const m = String(apto).match(/\d{3,4}/); return m ? `${Math.floor(Number(m[0]) / 100)}º` : ""; };
const slugify   = (v) => String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"");
const first     = (text, re, g=1) => { const m = String(text||"").match(re); return m?.[g]?.trim()||""; };
const normalizeSpaces = (v) => String(v||"").replace(/[\t\r]+/g," ").replace(/\n+/g," ").replace(/ {2,}/g," ").trim();

// ─── Extração torre / apto / serviço ──────────────────────────────────────
function extractTower(fileName, text) {
  const s = `${fileName} ${text}`;
  return (
    first(s, /torre\s*:?\s*([A-Za-z0-9]+)/i) ||
    first(s, /bloco\s*:?\s*([A-Za-z0-9]+)/i) ||
    first(fileName, /[_\-\s]([A-Za-z])\s*[_\-.\s]/i) ||
    first(fileName, /([A-Za-z])\.(?:docx|pdf|csv|xlsx)$/i) ||
    first(fileName, /\d{3,4}\s*([A-Za-z])/i) ||
    ""
  ).toUpperCase();
}

function extractApto(fileName, text) {
  return (
    first(text,     /(?:apto|apartamento|unidade|local[^:]*)\s*:?\s*(\d{3,4})/i) ||
    first(fileName, /apto\s*(\d{3,4})/i) ||
    first(fileName, /(?:^|[_\-\s])(\d{3,4})(?:[_\-\s]|[A-Za-z]?\.\w+$)/i) ||
    ""
  );
}

function detectServico(fileName, text) {
  const s = `${fileName} ${text}`.toLowerCase();
  if (/esquadria/.test(s))             return "esquadrias";
  if (/capi[aã][cç]o|capiaço/.test(s)) return "capiacos";
  if (/tubo\s*pass/.test(s))           return "tubos";
  return "ceramica";
}

// ─── Parsers ───────────────────────────────────────────────────────────────
function parseFvsText(fileName, rawText) {
  const text    = normalizeSpaces(rawText);
  const torre   = extractTower(fileName, text) || "SEM TORRE";
  const apto    = extractApto(fileName, text);
  const servico = detectServico(fileName, text);
  const crits   = CRITERIOS_POR_SERVICO[servico];
  const data    = first(text, /DATA\s*:?\s*(\d{2}\/\d{2}\/\d{4})/i) || "";
  const resp    = first(text, /respons[aá]vel\s*:?\s*([A-Za-zÀ-ÿ ]{3,60})/i) || "";

  const rows = [];
  for (const crit of crits) {
    const idx = text.toLowerCase().indexOf(crit.toLowerCase());
    if (idx === -1) continue;
    const seg    = text.slice(idx, idx + 400);
    const tokens = (seg.match(/N\/V|N\/A|\bA\b|\bR\b|-/gi) || []).map(t => t.toUpperCase());
    if (!tokens.length) continue;
    for (let i = 0; i < Math.min(tokens.length, AMBIENTES_PADRAO.length); i++) {
      rows.push({
        torre: sanitize(torre), apto: sanitize(apto), pav: inferPav(apto),
        data, ambiente: AMBIENTES_PADRAO[i], criterio: crit,
        resultado: normResult(tokens[i]), servico, equipe: resp, fonte: fileName,
      });
    }
  }
  return rows;
}

// Parser XLSX — tenta colunas estruturadas; se falhar, faz extração por texto
function parseXLSX(buffer, fileName) {
  const wb    = XLSX.read(buffer, { type: "array" });
  const rows  = [];

  wb.SheetNames.forEach(sheetName => {
    const ws  = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    if (!json.length) return;

    // Detecta se a planilha tem colunas reconhecíveis
    const keys    = Object.keys(json[0]).map(k => k.toLowerCase().trim());
    const hasStructure =
      keys.some(k => /torre|bloco/.test(k)) ||
      keys.some(k => /apto|apt|unid/.test(k)) ||
      keys.some(k => /result/.test(k));

    if (hasStructure) {
      // Leitura estruturada
      json.forEach(r => {
        const raw = Object.fromEntries(
          Object.entries(r).map(([k, v]) => [k.toLowerCase().trim(), v])
        );
        rows.push({
          torre:     sanitize(raw.torre || raw.bloco || extractTower(fileName, JSON.stringify(raw)) || "SEM TORRE"),
          apto:      sanitize(raw.apto || raw.apartamento || raw.unidade || ""),
          pav:       raw.pav || raw.pavimento || inferPav(raw.apto || raw.apartamento || ""),
          data:      raw.data || raw.date || "",
          ambiente:  raw.ambiente || raw.local || "",
          criterio:  raw.criterio || raw["critério"] || raw.item || "",
          resultado: normResult(raw.resultado || raw.result || raw.status || ""),
          servico:   raw.servico || raw["serviço"] || detectServico(fileName, JSON.stringify(raw)),
          equipe:    raw.equipe || raw.responsavel || raw["responsável"] || "",
          fonte:     `${fileName} — ${sheetName}`,
        });
      });
    } else {
      // Extração por texto bruto da aba
      const txt = XLSX.utils.sheet_to_csv(ws);
      rows.push(...parseFvsText(`${fileName} — ${sheetName}`, txt));
    }
  });

  return rows;
}

// ─── Cálculo de métricas ──────────────────────────────────────────────────
function calcMetrics(rows) {
  const valid = rows.filter(r => ["A","R","NV","NA"].includes(normResult(r.resultado)));
  const A  = valid.filter(r => normResult(r.resultado) === "A").length;
  const R  = valid.filter(r => normResult(r.resultado) === "R").length;
  const NV = valid.filter(r => normResult(r.resultado) === "NV").length;
  const NA = valid.filter(r => normResult(r.resultado) === "NA").length;
  const tapi = (A + R) ? Math.round(A / (A + R) * 100) : 0;

  const byApto = {};
  valid.forEach(r => {
    const key = `${sanitize(r.torre||"SEM TORRE")}-${sanitize(r.apto||"")}`;
    if (!byApto[key]) byApto[key] = { key, apto: r.apto, pav: r.pav || inferPav(r.apto), torre: sanitize(r.torre||"SEM TORRE"), data: r.data||"", v: 0, ncs: 0, crits: new Set() };
    if (normResult(r.resultado) !== "NA") byApto[key].v++;
    if (normResult(r.resultado) === "R")  { byApto[key].ncs++; byApto[key].crits.add(r.criterio); }
  });

  const aptoTable = Object.values(byApto).map(x => ({
    ...x, verificacoes: x.v,
    criterios: [...x.crits].slice(0, 3).join(", "),
    pct:    x.v ? `${Math.round(x.ncs / x.v * 100)}%` : "0%",
    status: x.ncs > 0 ? "REPROVADO" : "APROVADO",
  })).sort((a,b) => a.torre.localeCompare(b.torre) || Number(String(a.apto).match(/\d+/)?.[0]||0) - Number(String(b.apto).match(/\d+/)?.[0]||0));

  const critMap = {};
  valid.forEach(r => {
    const c = r.criterio || "N/I";
    if (!critMap[c]) critMap[c] = { t: 0, r: 0 };
    critMap[c].t++;
    if (normResult(r.resultado) === "R") critMap[c].r++;
  });
  const pareto = Object.entries(critMap)
    .map(([c,v]) => ({ criterio: c, reprovacoes: v.r, taxa: v.t ? Math.round(v.r/v.t*100) : 0 }))
    .sort((a,b) => b.reprovacoes - a.reprovacoes).slice(0, 8);

  const apartments = [...new Set(valid.map(r => `${sanitize(r.torre||"SEM TORRE")}-${r.apto}`).filter(x => x !== "SEM TORRE-"))].length;
  return { A, R, NV, NA, tapi, aptoTable, pareto, apartments, reprovedApts: aptoTable.filter(x => x.status==="REPROVADO").length, approvedApts: aptoTable.filter(x => x.status==="APROVADO").length };
}

// ─── SVG Charts ───────────────────────────────────────────────────────────
function ParetoSVG({ data }) {
  const W=700, H=400, L=200, T=40, B=60, R=20;
  const cW = W-L-R, cH = H-T-B;
  const max = Math.max(...data.map(d => d.reprovacoes), 1);
  const bH  = Math.min(36, cH / Math.max(data.length, 1) - 8);
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: "100%" }}>
      <rect width="100%" height="100%" fill="white"/>
      {[0,.25,.5,.75,1].map(t => {
        const x = L + t*cW;
        return <g key={t}><line x1={x} y1={T} x2={x} y2={H-B} stroke="#e2e8f0" strokeDasharray="4 4"/><text x={x} y={H-B+20} textAnchor="middle" fontSize={12} fill="#64748b">{Math.round(t*max)}</text></g>;
      })}
      {data.map((d,i) => {
        const y = T + i*(bH+10), w = (d.reprovacoes/max)*cW;
        return (
          <g key={i}>
            <text x={L-8} y={y+bH/2+5} fontSize={11} textAnchor="end" fill="#0f172a">{d.criterio}</text>
            <rect x={L} y={y} width={w} height={bH} rx={5} fill="#2563eb"/>
            <text x={L+w+6} y={y+bH/2+5} fontSize={11} fill="#0f172a">{d.reprovacoes} ({d.taxa}%)</text>
          </g>
        );
      })}
    </svg>
  );
}

function PieSVG({ data }) {
  const cx=160, cy=160, r=120;
  const total = Math.max(data.reduce((a,d) => a+d.value, 0), 1);
  const colors = { A: CORES.approved, R: CORES.reproved, NV: CORES.nv, NA: CORES.na };
  let ang = -90;
  const slices = data.map(d => {
    const a = (d.value/total)*360;
    const s = pxy(cx,cy,r,ang), e = pxy(cx,cy,r,ang+a);
    const lg = a>180?1:0;
    const path = `M${cx},${cy} L${s.x},${s.y} A${r},${r} 0 ${lg} 1 ${e.x},${e.y} Z`;
    ang += a;
    return { ...d, path };
  });
  return (
    <svg viewBox="0 0 480 320" style={{ width: "100%" }}>
      <rect width="100%" height="100%" fill="white"/>
      {slices.map((d,i) => <path key={i} d={d.path} fill={colors[d.key]||"#ccc"} stroke="white" strokeWidth={2}/>)}
      <circle cx={cx} cy={cy} r={50} fill="white"/>
      <text x={cx} y={cy} textAnchor="middle" fontSize={18} fontWeight="bold" fill="#0f172a" dy={6}>{total}</text>
      {data.map((d,i) => (
        <g key={i}>
          <rect x={340} y={80+i*50} width={18} height={18} rx={4} fill={colors[d.key]||"#ccc"}/>
          <text x={366} y={94+i*50} fontSize={13} fill="#0f172a">{d.label}: {d.value} ({total ? Math.round(d.value/total*100) : 0}%)</text>
        </g>
      ))}
    </svg>
  );
}
function pxy(cx,cy,r,deg) { const rad=(deg-90)*Math.PI/180; return { x: cx+r*Math.cos(rad), y: cy+r*Math.sin(rad) }; }

// ─── Exportação CSV ───────────────────────────────────────────────────────
function exportCSV(rows, towerSummary) {
  const lines = ["RESUMO POR TORRE","Torre,Apartamentos,Aprovadas,Reprovadas,NV,NA,TAPI"];
  towerSummary.forEach(t => lines.push(`${t.torre},${t.apts},${t.A},${t.R},${t.NV},${t.NA},${t.tapi}%`));
  lines.push("","DADOS DETALHADOS","Torre,Apto,Pav,Data,Serviço,Critério,Resultado,Ambiente");
  rows.forEach(r => lines.push(`${r.torre},${r.apto},${r.pav||""},${r.data||""},${r.servico||""},${r.criterio||""},${r.resultado||""},${r.ambiente||""}`));
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a"); a.href=url; a.download="fvs_export.csv"; a.click();
  URL.revokeObjectURL(url);
}

// ─── Componentes UI ───────────────────────────────────────────────────────
const KpiCard = ({ label, value, sub, color }) => (
  <div style={{ background: CORES.card, borderRadius: 16, padding: "20px 24px", borderLeft: `4px solid ${color||CORES.accent}` }}>
    <div style={{ fontSize: 11, textTransform: "uppercase", color: "#94a3b8", fontWeight: "bold", marginBottom: 8 }}>{label}</div>
    <div style={{ fontSize: 34, fontWeight: "bold", color: "white", marginBottom: 4 }}>{value}</div>
    <div style={{ fontSize: 13, color: "#64748b" }}>{sub}</div>
  </div>
);

const Section = ({ title, children, action }) => (
  <div style={{ background: CORES.card, borderRadius: 20, padding: 24, marginTop: 24 }}>
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
      <h2 style={{ margin: 0, color: "white", fontSize: 18 }}>{title}</h2>
      {action}
    </div>
    {children}
  </div>
);

const TH = ({ children }) => <th style={{ borderBottom: "1px solid #334155", padding: "10px 12px", textAlign: "left", color: "#94a3b8", fontSize: 12, textTransform: "uppercase", fontWeight: "bold", background: "#1e293b" }}>{children}</th>;
const TD = ({ children, bold, color }) => <td style={{ borderBottom: "1px solid #1e293b", padding: "10px 12px", color: color||"#cbd5e1", fontWeight: bold?"bold":"normal", fontSize: 13 }}>{children}</td>;
const Btn = ({ children, onClick, color, small }) => (
  <button onClick={onClick} style={{ background: color||CORES.blue, color: "white", border: "none", borderRadius: 8, padding: small?"6px 12px":"10px 16px", cursor: "pointer", fontWeight: "bold", fontSize: small?12:14 }}>{children}</button>
);

// ─── App Principal ────────────────────────────────────────────────────────
export default function App() {
  const [rows, setRows]               = useState([]);
  const [status, setStatus]           = useState("Envie arquivos CSV, XLSX, DOCX ou PDF.");
  const [fileNames, setFileNames]     = useState([]);
  const [errors, setErrors]           = useState([]);
  const [selectedTower, setSelectedTower]   = useState("CONSOLIDADO");
  const [selectedServico, setSelectedServico] = useState("TODOS");
  const [selectedPav, setSelectedPav]       = useState("TODOS");
  const [aiAnalysis, setAiAnalysis]         = useState("");
  const [aiLoading, setAiLoading]           = useState(false);
  const [activeTab, setActiveTab]           = useState("dashboard");

  const handleFile = useCallback(async (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setRows([]); setErrors([]); setAiAnalysis("");
    setFileNames(files.map(f => f.name));
    setStatus("Processando arquivos...");
    let all = [], errs = [];

    for (const file of files) {
      try {
        const ext = file.name.split(".").pop()?.toLowerCase();

        if (ext === "csv") {
          const text = await file.text();
          const lines = text.split(/\r?\n/).filter(l => l.trim());
          if (!lines.length) continue;
          const headers = lines[0].split(",").map(h => h.replace(/^"|"$/g,"").trim().toLowerCase());
          const parsed  = lines.slice(1).map(line => {
            const vals = line.split(",").map(v => v.replace(/^"|"$/g,"").trim());
            const obj  = {};
            headers.forEach((h,i) => obj[h] = vals[i]||"");
            return {
              torre:    sanitize(obj.torre || obj.bloco || extractTower(file.name, "") || "SEM TORRE"),
              apto:     sanitize(obj.apto  || obj.apartamento || ""),
              pav:      obj.pav || inferPav(obj.apto||""),
              data:     obj.data || "",
              ambiente: obj.ambiente || "",
              criterio: obj.criterio || obj["critério"] || "",
              resultado:normResult(obj.resultado || obj.result || ""),
              servico:  obj.servico || detectServico(file.name, ""),
              equipe:   obj.equipe  || obj.responsavel || "",
              fonte:    file.name,
            };
          });
          all = [...all, ...parsed];
          continue;
        }

        if (ext === "xlsx" || ext === "xls") {
          const buffer = await file.arrayBuffer();
          all = [...all, ...parseXLSX(buffer, file.name)];
          continue;
        }

        if (ext === "docx") {
          const buffer = await file.arrayBuffer();
          const result = await mammoth.extractRawText({ arrayBuffer: buffer });
          all = [...all, ...parseFvsText(file.name, result.value || "")];
          continue;
        }

        if (ext === "pdf") {
          const buffer = await file.arrayBuffer();
          const pdf    = await pdfjsLib.getDocument({ data: buffer }).promise;
          let pdfText  = "";
          for (let p = 1; p <= pdf.numPages; p++) {
            const page    = await pdf.getPage(p);
            const content = await page.getTextContent();
            pdfText      += " " + content.items.map(i => i.str).join(" ");
          }
          all = [...all, ...parseFvsText(file.name, pdfText)];
          continue;
        }

        errs.push(`${file.name}: formato não suportado.`);
      } catch (err) {
        console.error(err);
        errs.push(`${file.name}: erro — ${err.message}`);
      }
    }

    setRows(all); setErrors(errs);
    setStatus(`${files.length} arquivo(s) processado(s). ${all.length} linhas carregadas.`);
  }, []);

  const towers   = useMemo(() => ["CONSOLIDADO", ...[...new Set(rows.map(r => sanitize(r.torre)).filter(Boolean))].sort()], [rows]);
  const servicos = useMemo(() => ["TODOS",       ...[...new Set(rows.map(r => r.servico).filter(Boolean))].sort()],         [rows]);
  const pavs     = useMemo(() => ["TODOS",       ...[...new Set(rows.map(r => r.pav || inferPav(r.apto)).filter(Boolean))].sort((a,b)=>parseInt(a)-parseInt(b))], [rows]);

  const scoped = useMemo(() => {
    let r = rows;
    if (selectedTower   !== "CONSOLIDADO") r = r.filter(x => sanitize(x.torre) === selectedTower);
    if (selectedServico !== "TODOS")       r = r.filter(x => x.servico === selectedServico);
    if (selectedPav     !== "TODOS")       r = r.filter(x => (x.pav || inferPav(x.apto)) === selectedPav);
    return r;
  }, [rows, selectedTower, selectedServico, selectedPav]);

  const metrics     = useMemo(() => calcMetrics(scoped), [scoped]);
  const towerSummary = useMemo(() => {
    const map = {};
    rows.forEach(r => {
      const t = sanitize(r.torre || "SEM TORRE");
      if (!map[t]) map[t] = { torre: t, A:0, R:0, NV:0, NA:0, apts: new Set() };
      const res = normResult(r.resultado);
      if (res==="A")  map[t].A++;
      if (res==="R")  map[t].R++;
      if (res==="NV") map[t].NV++;
      if (res==="NA") map[t].NA++;
      if (r.apto) map[t].apts.add(`${t}-${r.apto}`);
    });
    return Object.values(map).map(x => ({ ...x, apts: x.apts.size, tapi: (x.A+x.R) ? Math.round(x.A/(x.A+x.R)*100) : 0 })).sort((a,b) => a.torre.localeCompare(b.torre));
  }, [rows]);

  const pieData = [
    { key:"A",  label:"Aprovado",  value: metrics.A  },
    { key:"R",  label:"Reprovado", value: metrics.R  },
    { key:"NV", label:"N/V",       value: metrics.NV },
    { key:"NA", label:"N/A",       value: metrics.NA },
  ];
  const scopeLabel = selectedTower === "CONSOLIDADO" ? "Todas as torres" : `Torre ${selectedTower}`;

  async function runAI() {
    setAiLoading(true); setAiAnalysis("");
    const summary = `Dados FVS — ${scopeLabel}\nAPTOS: ${metrics.apartments}\nTAPI: ${metrics.tapi}%\nAprovadas: ${metrics.A} | Reprovadas: ${metrics.R} | NV: ${metrics.NV} | NA: ${metrics.NA}\nPareto critérios:\n${metrics.pareto.map(p => `- ${p.criterio}: ${p.reprovacoes} R (${p.taxa}%)`).join("\n")}\nResumo por torre:\n${towerSummary.map(t => `- Torre ${t.torre}: ${t.apts} aptos, TAPI ${t.tapi}%`).join("\n")}`;
    try {
      const res  = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model:      "claude-sonnet-4-20250514",
          max_tokens: 1000,
          system:     "Você é um especialista em qualidade de construção civil. Analise os dados de FVS (Ficha de Verificação de Serviço) fornecidos e gere um relatório executivo em português com: 1) Avaliação geral do TAPI, 2) Principais não-conformidades (Pareto), 3) Pontos de atenção por torre, 4) Recomendações práticas. Seja objetivo e use linguagem técnica de qualidade.",
          messages:   [{ role: "user", content: `Analise estes dados de FVS e gere o relatório:\n\n${summary}` }],
        }),
      });
      const data = await res.json();
      const text = data.content?.filter(c => c.type==="text").map(c => c.text).join("\n") || "Sem resposta.";
      setAiAnalysis(text);
    } catch(err) {
      setAiAnalysis(`Erro ao chamar IA: ${err.message}`);
    }
    setAiLoading(false);
  }

  const tabs = [
    { id:"dashboard", label:"Dashboard" },
    { id:"tabela",    label:"Por Apartamento" },
    { id:"torres",    label:"Por Torre" },
    { id:"ia",        label:"🤖 Análise IA" },
  ];

  return (
    <div style={{ minHeight:"100vh", background: CORES.bg, color: CORES.textLight, padding:"32px 24px", fontFamily:"'Inter',Arial,sans-serif" }}>
      <div style={{ maxWidth:1200, margin:"0 auto" }}>

        {/* Header */}
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:8 }}>
          <span style={{ background:"#082f49", color: CORES.accent, borderRadius:999, padding:"4px 14px", fontSize:13, fontWeight:"bold" }}>FVS Qualidade</span>
          <span style={{ color:"#475569", fontSize:13 }}>v2.0 — Multi-serviço por Torre</span>
        </div>
        <h1 style={{ fontSize:36, margin:"0 0 6px", color:"white" }}>Dashboard de Inspeção FVS</h1>
        <p  style={{ color:"#64748b", margin:"0 0 24px" }}>Cerâmica · Esquadrias · Capiaços · Tubos Passantes</p>

        {/* Upload */}
        <div style={{ background: CORES.card, borderRadius:20, padding:24, marginBottom:24, border:`1px solid ${CORES.border}` }}>
          <div style={{ display:"flex", gap:12, flexWrap:"wrap", alignItems:"center" }}>
            <label style={{ background: CORES.blue, color:"white", borderRadius:10, padding:"10px 18px", cursor:"pointer", fontWeight:"bold", fontSize:14 }}>
              📂 Carregar Arquivos
              <input type="file" multiple accept=".csv,.xlsx,.xls,.docx,.pdf" onChange={handleFile} style={{ display:"none" }}/>
            </label>
            {rows.length > 0 && <Btn onClick={() => exportCSV(rows, towerSummary)} color="#7c3aed">⬇ Exportar CSV</Btn>}
          </div>
          {fileNames.length > 0 && <div style={{ marginTop:12, color:"#94a3b8", fontSize:13 }}>{fileNames.join(" · ")}</div>}
          <div style={{ marginTop:8, color: errors.length?"#f87171":"#64748b", fontSize:13 }}>{status}</div>
          {errors.map((e,i) => <div key={i} style={{ color:"#f87171", fontSize:12, marginTop:4 }}>⚠ {e}</div>)}
        </div>

        {rows.length > 0 && <>
          {/* Filtros */}
          <div style={{ display:"flex", gap:12, flexWrap:"wrap", marginBottom:20 }}>
            {[
              ["Torre",    towers,   selectedTower,    setSelectedTower,    v => v==="CONSOLIDADO"?"Consolidado":`Torre ${v}`],
              ["Serviço",  servicos, selectedServico,  setSelectedServico,  v => v==="TODOS"?"Todos":SERVICOS_LABELS[v]||v],
              ["Pavimento",pavs,     selectedPav,      setSelectedPav,      v => v==="TODOS"?"Todos":v],
            ].map(([label,opts,val,setter,fmt]) => (
              <div key={label}>
                <div style={{ fontSize:11, color:"#64748b", marginBottom:4, textTransform:"uppercase" }}>{label}</div>
                <select value={val} onChange={e => setter(e.target.value)} style={{ background:"#1e293b", color:"white", border:`1px solid ${CORES.border}`, borderRadius:8, padding:"8px 12px", fontSize:13 }}>
                  {opts.map(o => <option key={o} value={o}>{fmt(o)}</option>)}
                </select>
              </div>
            ))}
          </div>

          {/* Tabs */}
          <div style={{ display:"flex", gap:4, marginBottom:20, borderBottom:`1px solid ${CORES.border}` }}>
            {tabs.map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)} style={{ background: activeTab===t.id?CORES.blue:"transparent", color: activeTab===t.id?"white":"#64748b", border:"none", borderRadius:"8px 8px 0 0", padding:"10px 20px", cursor:"pointer", fontWeight: activeTab===t.id?"bold":"normal", fontSize:14 }}>
                {t.label}
              </button>
            ))}
          </div>

          {/* ── TAB DASHBOARD ── */}
          {activeTab==="dashboard" && <>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))", gap:16, marginBottom:24 }}>
              <KpiCard label="Aptos Inspecionados" value={metrics.apartments}    sub={scopeLabel}                                                color={CORES.accent}/>
              <KpiCard label="TAPI"                value={`${metrics.tapi}%`}    sub={metrics.tapi>=85?"✅ Meta atingida":"❌ Abaixo da meta (85%)"}  color={metrics.tapi>=85?CORES.approved:CORES.reproved}/>
              <KpiCard label="Reprovados"          value={metrics.reprovedApts}  sub={`${metrics.approvedApts} aprovados`}                        color={CORES.reproved}/>
              <KpiCard label="Total de NCs"        value={metrics.R}             sub="Ocorrências reprovadas"                                     color={CORES.nv}/>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20 }}>
              <Section title="Pareto de Critérios Reprovados">
                {metrics.pareto.length ? <ParetoSVG data={metrics.pareto}/> : <p style={{ color:"#64748b" }}>Sem dados.</p>}
              </Section>
              <Section title="Distribuição de Resultados">
                <PieSVG data={pieData}/>
              </Section>
            </div>
          </>}

          {/* ── TAB TABELA ── */}
          {activeTab==="tabela" && (
            <Section title={`Resultado por Apartamento — ${scopeLabel}`}>
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse", minWidth:900 }}>
                  <thead><tr>{["Torre","Apto","Pav","Data","Verif.","NCs","% R","Critérios críticos","Status"].map(h=><TH key={h}>{h}</TH>)}</tr></thead>
                  <tbody>
                    {metrics.aptoTable.map((r,i) => (
                      <tr key={i} style={{ background: i%2===0?"transparent":"#ffffff08" }}>
                        <TD>{r.torre||"—"}</TD>
                        <TD bold>{r.apto}</TD>
                        <TD>{r.pav}</TD>
                        <TD>{r.data||"—"}</TD>
                        <TD>{r.verificacoes}</TD>
                        <TD color={r.ncs>0?CORES.reproved:"#94a3b8"}>{r.ncs}</TD>
                        <TD>{r.pct}</TD>
                        <TD>{r.criterios||"—"}</TD>
                        <TD color={r.status==="REPROVADO"?CORES.reproved:CORES.approved} bold>{r.status}</TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>
          )}

          {/* ── TAB TORRES ── */}
          {activeTab==="torres" && (
            <Section title="Resumo Consolidado por Torre">
              <div style={{ overflowX:"auto" }}>
                <table style={{ width:"100%", borderCollapse:"collapse" }}>
                  <thead><tr>{["Torre","Apts","Aprovadas","Reprovadas","N/V","N/A","TAPI"].map(h=><TH key={h}>{h}</TH>)}</tr></thead>
                  <tbody>
                    {towerSummary.map((t,i) => (
                      <tr key={i} style={{ background: i%2===0?"transparent":"#ffffff08" }}>
                        <TD bold color={CORES.accent}>Torre {t.torre}</TD>
                        <TD>{t.apts}</TD>
                        <TD color={CORES.approved}>{t.A}</TD>
                        <TD color={CORES.reproved}>{t.R}</TD>
                        <TD color={CORES.nv}>{t.NV}</TD>
                        <TD>{t.NA}</TD>
                        <TD bold color={t.tapi>=85?CORES.approved:CORES.reproved}>{t.tapi}%</TD>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Section>
          )}

          {/* ── TAB IA ── */}
          {activeTab==="ia" && (
            <Section title="🤖 Análise Automática por IA" action={<Btn onClick={runAI}>{aiLoading?"Analisando...":"Gerar Análise"}</Btn>}>
              {!aiAnalysis && !aiLoading && <p style={{ color:"#64748b" }}>Clique em "Gerar Análise" para gerar um relatório executivo com recomendações baseado nos dados carregados.</p>}
              {aiLoading  && <div style={{ color: CORES.accent, padding:"20px 0" }}>⏳ Gerando análise inteligente dos dados...</div>}
              {aiAnalysis && <pre style={{ whiteSpace:"pre-wrap", wordBreak:"break-word", fontSize:14, lineHeight:1.7, color:"#e2e8f0", margin:0 }}>{aiAnalysis}</pre>}
            </Section>
          )}
        </>}

        {rows.length===0 && (
          <div style={{ textAlign:"center", padding:"60px 20px", color:"#475569" }}>
            <div style={{ fontSize:48, marginBottom:16 }}>📋</div>
            <div style={{ fontSize:18, marginBottom:8 }}>Nenhum dado carregado</div>
            <div style={{ fontSize:14 }}>Carregue arquivos CSV, XLSX, DOCX ou PDF para ver o dashboard.</div>
          </div>
        )}
      </div>
    </div>
  );
}
