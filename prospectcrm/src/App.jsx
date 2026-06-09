import { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip,
  ResponsiveContainer, Cell, PieChart, Pie,
} from "recharts";

/* CONSTANTES */
const LABELS = {
  "Não atendeu":         { color: "#94a3b8", bg: "#1e293b", icon: "📵" },
  "Sem resposta":        { color: "#a3a3a3", bg: "#27272a", icon: "🔕" },
  "Sem interesse":       { color: "#f87171", bg: "#450a0a", icon: "✖️" },
  "Chamar no WhatsApp":  { color: "#4ade80", bg: "#052e16", icon: "💬" },
  "Ligação agendada":    { color: "#fbbf24", bg: "#2d1f02", icon: "📅" },
  "Cotação enviada":     { color: "#22d3ee", bg: "#083344", icon: "📄" },
  "Cliente interessado": { color: "#60a5fa", bg: "#0c1a3d", icon: "⭐" },
  "Venda realizada":     { color: "#c084fc", bg: "#1a0533", icon: "🏆" },
};

const SEED = [
  { id:"s1", name:"João Silva",     phone:"85999991111", label:"Cliente interessado", calls:2, updatedAt:"29/05 14:32" },
  { id:"s2", name:"Maria Santos",   phone:"85988882222", label:"Venda realizada",     calls:3, updatedAt:"30/05 09:10" },
  { id:"s3", name:"Pedro Alves",    phone:"85977773333", label:"Ligação agendada",    calls:1, updatedAt:"30/05 11:45" },
  { id:"s4", name:"Fernanda Lima",  phone:"85966664444", label:"Não atendeu",         calls:4, updatedAt:"28/05 16:00" },
  { id:"s5", name:"Lucas Ferreira", phone:"85955555555", label:"Sem interesse",       calls:1, updatedAt:"27/05 10:20" },
  { id:"s6", name:"Carla Dias",     phone:"85944446666", label:"Chamar no WhatsApp",  calls:2, updatedAt:"29/05 13:00" },
  { id:"s7", name:"Bruno Costa",    phone:"85933337777", label:null, calls:0, updatedAt:"31/05 08:00" },
  { id:"s8", name:"Juliana Neves",  phone:"85922228888", label:null, calls:0, updatedAt:"31/05 08:00" },
];

const fmt015 = p => "015" + String(p).replace(/\D/g, "");
const fmtWA  = p => "55"  + String(p).replace(/\D/g, "");

/* primeiro nome do lead */
const firstName = full => String(full || "").trim().split(/\s+/)[0] || "";

/* mensagem padrao de abertura no WhatsApp */
function waMessage(leadName, vendorName) {
  const nome = firstName(leadName);
  const vend = (vendorName || "").trim() || "seu Consultor";
  return (
    `Olá, ${nome}! Aqui é o ${vend}, Consultor Comercial da Hapvida NotreDame. 😊` + "\n\n" +
    "Estou com condições especiais este mês. Você tem interesse?" + "\n\n" +
    "1️⃣ Sim, quero saber mais" + "\n" +
    "2️⃣ Não tenho interesse" + "\n\n" +
    "_Para não receber mais mensagens, responda 2._"
  );
}

/* URL do WhatsApp ja com a mensagem */
function waLink(phone, leadName, vendorName) {
  return `https://wa.me/${fmtWA(phone)}?text=${encodeURIComponent(waMessage(leadName, vendorName))}`;
}
const nowStr = () => new Date()
  .toLocaleString("pt-BR", { day:"2-digit", month:"2-digit", hour:"2-digit", minute:"2-digit" })
  .replace(",", "");

/* Carrega jsPDF via CDN sob demanda (o ambiente do artifact nao permite import direto) */
function loadJsPDF() {
  return new Promise((resolve, reject) => {
    if (window.jspdf && window.jspdf.jsPDF) { resolve(window.jspdf.jsPDF); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    s.onload = () => {
      if (window.jspdf && window.jspdf.jsPDF) resolve(window.jspdf.jsPDF);
      else reject(new Error("jsPDF carregou mas nao foi encontrado"));
    };
    s.onerror = () => reject(new Error("Falha ao carregar jsPDF (rede bloqueada?)"));
    document.head.appendChild(s);
  });
}

/* PERSISTENCIA: salva os leads no proprio dispositivo (localStorage) */
const STORAGE_KEY = "prospectcrm_leads_v1";

function loadLeads() {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch (e) { /* localStorage indisponivel (ex: dentro do preview) */ }
  return null;   // null = primeira vez, usa SEED
}

function saveLeads(leads) {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(leads));
  } catch (e) { /* ignora se nao puder salvar */ }
}

const VENDOR_KEY = "prospectcrm_vendor_v1";
function loadVendor() { try { return window.localStorage.getItem(VENDOR_KEY) || ""; } catch (e) { return ""; } }
function saveVendor(name) { try { window.localStorage.setItem(VENDOR_KEY, name); } catch (e) {} }

/* PARSE: matrix [[c,c],[c,c]] -> {leads, info, error} */
function parseMatrix(matrix) {
  const info = { rows: matrix ? matrix.length : 0, header: -1, nameCol: -1, phoneCol: -1, found: [] };
  if (!matrix || matrix.length === 0) return { leads: [], info, error: "Arquivo vazio." };

  const norm = v => String(v == null ? "" : v)
    .replace(/[\u00a0\u200b\u200c\u200d\ufeff]/g, "")
    .trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");

  const PHONE = new Set(["telefone","fone","tel","phone","celular","cel","whatsapp","numero","contato","telef"]);

  for (let r = 0; r < Math.min(6, matrix.length); r++) {
    const row = (matrix[r] || []).map(norm);
    const ni  = row.findIndex(c => c === "nome");
    const pi  = row.findIndex(c => PHONE.has(c));
    if (ni !== -1 && pi !== -1) { info.header = r; info.nameCol = ni; info.phoneCol = pi; break; }
    if ((ni !== -1 || pi !== -1) && info.header === -1) { info.header = r; info.nameCol = ni; info.phoneCol = pi; }
  }

  if (info.header === -1) {
    info.found = (matrix[0] || []).map(c => String(c).trim());
    return { leads: [], info, error: "Cabeçalho não encontrado." };
  }
  info.found = (matrix[info.header] || []).map(c => String(c).trim());
  if (info.nameCol === -1 || info.phoneCol === -1) {
    return { leads: [], info, error: `Falta coluna ${info.nameCol === -1 ? "Nome" : "Telefone"}.` };
  }

  const leads = [];
  for (let r = info.header + 1; r < matrix.length; r++) {
    const row  = matrix[r] || [];
    const name = String(row[info.nameCol] == null ? "" : row[info.nameCol]).trim();
    const rawP = row[info.phoneCol];
    const phone = (
      rawP == null            ? "" :
      typeof rawP === "number" ? String(Math.round(rawP)) :
      String(rawP)
    ).replace(/\D/g, "");
    if (name && phone.length >= 7) {
      leads.push({ id:`imp_${Date.now()}_${r}`, name, phone, label:null, calls:0, updatedAt:nowStr() });
    }
  }
  return { leads, info, error: leads.length ? null : "Nenhum dado válido abaixo do cabeçalho." };
}

/* APP */
export default function App() {
  const [tab,      setTab]   = useState("leads");
  const [leads,    setLeads] = useState(() => loadLeads() ?? SEED);
  const [dark,     setDark]  = useState(true);
  const [vendorName, setVendorName] = useState(() => loadVendor());
  const [toast,    setToast] = useState(null);
  const [sheet,    setSheet] = useState(false);
  const [sheetData,setSData] = useState(null);

  const D    = dark;
  const bg   = D ? "#080c14" : "#f0f4f8";
  const card = D ? "#0e1525" : "#ffffff";
  const bdr  = D ? "#1a2540" : "#e2e8f0";
  const txt  = D ? "#e2e8f0" : "#0f172a";
  const sub  = D ? "#64748b" : "#94a3b8";

  /* salva automaticamente no dispositivo a cada alteracao */
  useEffect(() => { saveLeads(leads); }, [leads]);
  useEffect(() => { saveVendor(vendorName); }, [vendorName]);

  const fire = (msg, type = "ok") => { setToast({ msg, type }); setTimeout(() => setToast(null), 3000); };
  const openSheet  = d  => { setSData(d); setSheet(true); };
  const closeSheet = () => { setSheet(false); setSData(null); };

  const updateLabel = (id, label) => {
    setLeads(p => p.map(l => l.id === id ? { ...l, label, calls: l.calls + 1, updatedAt: nowStr() } : l));
    closeSheet(); fire("Etiqueta salva ✓");
  };

  const addLeads = incoming => {
    const seen  = new Set(leads.map(l => l.phone));
    const fresh = incoming.filter(l => !seen.has(l.phone));
    setLeads(p => [...p, ...fresh]);
    const dup = incoming.length - fresh.length;
    fire(`${fresh.length} importados${dup ? ` · ${dup} duplicados` : ""}`);
  };

  const clearLeads = () => {
    const count = leads.length;
    setLeads([]);
    closeSheet();
    fire(`${count} leads apagados`, "ok");
  };

  return (
    <div style={{ fontFamily:"'Plus Jakarta Sans',sans-serif", background:bg, color:txt,
      minHeight:"100dvh", display:"flex", flexDirection:"column", maxWidth:480, margin:"0 auto", position:"relative" }}>
      <div style={{ height:"env(safe-area-inset-top,0px)" }} />

      <header style={{ padding:"14px 18px 12px", background:D?"rgba(8,12,20,.95)":"rgba(240,244,248,.95)",
        backdropFilter:"blur(16px)", borderBottom:`1px solid ${bdr}`, display:"flex", alignItems:"center",
        justifyContent:"space-between", position:"sticky", top:0, zIndex:40 }}>
        <div style={{ display:"flex", alignItems:"center", gap:9 }}>
          <div style={{ width:34, height:34, borderRadius:10, background:"linear-gradient(135deg,#2563eb,#7c3aed)",
            display:"flex", alignItems:"center", justifyContent:"center", fontSize:17 }}>📞</div>
          <div>
            <div style={{ fontWeight:800, fontSize:15 }}>ProspectCRM</div>
            <div style={{ fontSize:10, color:sub }}>{leads.length} leads · {leads.filter(l=>l.label).length} trabalhados</div>
          </div>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={()=>openSheet({type:"config"})} title="Configurar vendedor"
            style={{ width:36, height:36, borderRadius:10, border:`1px solid ${bdr}`,
              background:D?"#1a2540":"#e2e8f0", color:sub, fontSize:16, cursor:"pointer" }}>⚙️</button>
          <button onClick={()=>setDark(d=>!d)} style={{ width:36, height:36, borderRadius:10,
            border:`1px solid ${bdr}`, background:D?"#1a2540":"#e2e8f0", color:sub, fontSize:16, cursor:"pointer" }}>
            {D?"☀️":"🌙"}
          </button>
        </div>
      </header>

      <div style={{ flex:1, overflowY:"auto", paddingBottom:80, WebkitOverflowScrolling:"touch" }}>
        {tab==="leads" && (
          <LeadsTab leads={leads} vendorName={vendorName} D={D} card={card} bdr={bdr} txt={txt} sub={sub}
            onLabel={l=>openSheet({type:"label",lead:l})}
            onImport={()=>openSheet({type:"import"})}
            onExport={()=>openSheet({type:"export",leads})}
            onClear={()=>openSheet({type:"clear",count:leads.length})} fire={fire} />
        )}
        {tab==="dash" && <DashTab leads={leads} D={D} card={card} bdr={bdr} txt={txt} sub={sub} />}
      </div>

      <nav style={{ position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)", width:"100%", maxWidth:480,
        background:D?"rgba(8,12,20,.97)":"rgba(255,255,255,.97)", backdropFilter:"blur(20px)",
        borderTop:`1px solid ${bdr}`, display:"grid", gridTemplateColumns:"1fr 1fr",
        paddingBottom:"env(safe-area-inset-bottom,12px)", zIndex:50 }}>
        {[{id:"leads",icon:"👥",label:"Leads"},{id:"dash",icon:"📊",label:"Dashboard"}].map(n=>(
          <button key={n.id} onClick={()=>setTab(n.id)} style={{ padding:"12px 0 6px", border:"none",
            background:"transparent", cursor:"pointer", display:"flex", flexDirection:"column", alignItems:"center", gap:3 }}>
            <div style={{ width:40, height:40, borderRadius:12,
              background:tab===n.id?"linear-gradient(135deg,#2563eb22,#7c3aed22)":"transparent",
              display:"flex", alignItems:"center", justifyContent:"center", fontSize:20 }}>{n.icon}</div>
            <span style={{ fontSize:10, fontWeight:tab===n.id?700:500, color:tab===n.id?"#818cf8":sub }}>{n.label}</span>
          </button>
        ))}
      </nav>

      {sheet && sheetData && (
        <BottomSheet onClose={closeSheet} D={D} card={card} bdr={bdr}>
          {sheetData.type==="label"  && <LabelSheet  lead={sheetData.lead} onSave={updateLabel} onClose={closeSheet} D={D} bdr={bdr} sub={sub} />}
          {sheetData.type==="import" && <ImportSheet onDone={addLeads} onClose={closeSheet} D={D} bdr={bdr} sub={sub} />}
          {sheetData.type==="export" && <ExportSheet leads={sheetData.leads} onClose={closeSheet} fire={fire} D={D} bdr={bdr} sub={sub} />}
          {sheetData.type==="clear"  && <ClearSheet count={sheetData.count} onConfirm={clearLeads} onClose={closeSheet} D={D} bdr={bdr} sub={sub} />}
          {sheetData.type==="config" && <ConfigSheet vendorName={vendorName} setVendorName={setVendorName} onClose={closeSheet} fire={fire} D={D} bdr={bdr} sub={sub} />}
        </BottomSheet>
      )}

      {toast && (
        <div style={{ position:"fixed", bottom:90, left:"50%", transform:"translateX(-50%)",
          background:toast.type==="ok"?"#22c55e":"#ef4444", color:"#fff", padding:"11px 22px",
          borderRadius:50, fontSize:13, fontWeight:600, whiteSpace:"nowrap",
          boxShadow:"0 8px 24px rgba(0,0,0,.4)", zIndex:200, animation:"popIn .25s cubic-bezier(.34,1.56,.64,1)" }}>
          {toast.msg}
        </div>
      )}

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
        ::-webkit-scrollbar{width:3px}::-webkit-scrollbar-thumb{background:#334155;border-radius:99px}
        @keyframes popIn{from{transform:translateX(-50%) scale(.8);opacity:0}to{transform:translateX(-50%) scale(1);opacity:1}}
        @keyframes slideUp{from{transform:translateY(100%)}to{transform:translateY(0)}}
        @keyframes fadeIn{from{opacity:0}to{opacity:1}}
        @keyframes cardIn{from{transform:translateY(12px);opacity:0}to{transform:translateY(0);opacity:1}}
        input,select,textarea,button{font-family:inherit;}
      `}</style>
    </div>
  );
}

/* LEADS TAB */
function LeadsTab({ leads, vendorName, D, card, bdr, txt, sub, onLabel, onImport, onExport, onClear, fire }) {
  const [q,setQ]=useState(""); const [filter,setFilter]=useState("all");
  const [sort,setSort]=useState("name"); const [expanded,setExpanded]=useState(null);

  const list = useMemo(() => {
    let a = [...leads];
    if (q) { const s=q.toLowerCase(); a=a.filter(l=>l.name.toLowerCase().includes(s)||l.phone.includes(s)); }
    if (filter==="pending") a=a.filter(l=>!l.label);
    else if (filter!=="all") a=a.filter(l=>l.label===filter);
    if (sort==="name") a.sort((a,b)=>a.name.localeCompare(b.name));
    if (sort==="recent") a.sort((a,b)=>(b.updatedAt||"").localeCompare(a.updatedAt||""));
    return a;
  }, [leads,q,filter,sort]);

  return (
    <div style={{ padding:"16px 14px", display:"flex", flexDirection:"column", gap:12 }}>
      <div style={{ display:"flex", gap:8 }}>
        <Btn2 onClick={onImport} accent="#2563eb" D={D}>📥 Importar</Btn2>
        <Btn2 onClick={onExport} accent="#059669" D={D}>📤 Exportar</Btn2>
        <Btn2 onClick={onClear}  accent="#dc2626" D={D}>🗑️ Limpar</Btn2>
      </div>
      <div style={{ position:"relative" }}>
        <span style={{ position:"absolute", left:12, top:"50%", transform:"translateY(-50%)", fontSize:15, pointerEvents:"none" }}>🔍</span>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Buscar nome ou telefone..."
          style={{ width:"100%", padding:"11px 12px 11px 38px", borderRadius:12, border:`1.5px solid ${bdr}`,
            background:D?"#0e1525":"#fff", color:txt, fontSize:14, outline:"none" }} />
      </div>
      <div style={{ display:"flex", gap:6, overflowX:"auto", paddingBottom:2, scrollbarWidth:"none" }}>
        {[["all","Todos"],["pending","⏳ Pendentes"],...Object.entries(LABELS).map(([l,c])=>[l,c.icon+" "+l])].map(([v,label])=>(
          <button key={v} onClick={()=>setFilter(v)} style={{ flexShrink:0, padding:"6px 12px", borderRadius:50,
            fontSize:11, fontWeight:600, border:`1.5px solid ${filter===v?(LABELS[v]?.color||"#6366f1"):bdr}`,
            background:filter===v?(D?"rgba(99,102,241,.15)":"rgba(99,102,241,.08)"):"transparent",
            color:filter===v?(LABELS[v]?.color||"#818cf8"):(D?"#64748b":"#94a3b8"), cursor:"pointer", whiteSpace:"nowrap" }}>
            {label}
          </button>
        ))}
      </div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
        <span style={{ fontSize:12, color:sub }}>{list.length} leads</span>
        <select value={sort} onChange={e=>setSort(e.target.value)} style={{ background:D?"#0e1525":"#fff",
          border:`1px solid ${bdr}`, borderRadius:8, padding:"6px 10px", color:sub, fontSize:11, outline:"none", cursor:"pointer" }}>
          <option value="name">A-Z Nome</option>
          <option value="recent">Mais recentes</option>
        </select>
      </div>
      {list.length===0
        ? <div style={{ textAlign:"center", padding:"48px 0", color:sub }}>
            <div style={{ fontSize:44, marginBottom:12 }}>📭</div>
            <p style={{ fontWeight:600 }}>Nenhum lead</p>
            <p style={{ fontSize:12, marginTop:4 }}>Importe uma planilha para começar</p>
          </div>
        : list.map((lead,i)=>(
            <LeadCard key={lead.id} lead={lead} i={i} vendorName={vendorName} D={D} card={card} bdr={bdr} sub={sub}
              expanded={expanded===lead.id} onToggle={()=>setExpanded(p=>p===lead.id?null:lead.id)}
              onLabel={()=>onLabel(lead)} fire={fire} />
          ))
      }
    </div>
  );
}

/* LEAD CARD */
function LeadCard({ lead, i, vendorName, D, card, bdr, sub, expanded, onToggle, onLabel, fire }) {
  const lc = lead.label ? LABELS[lead.label] : null;
  return (
    <div style={{ background:card, borderRadius:16, overflow:"hidden",
      border:`1.5px solid ${expanded?"#2563eb55":bdr}`, animation:`cardIn .3s ease ${i*25}ms both`,
      boxShadow:expanded?"0 8px 32px rgba(37,99,235,.15)":"none", transition:"box-shadow .2s,border-color .2s" }}>
      <div onClick={onToggle} style={{ padding:"14px 16px", display:"flex", alignItems:"center", gap:12, cursor:"pointer" }}>
        <div style={{ width:42, height:42, borderRadius:12, flexShrink:0,
          background:`linear-gradient(135deg,${lc?lc.color+"33":"#1e2d4a"},${lc?lc.color+"11":"#0e1525"})`,
          border:`1.5px solid ${lc?lc.color+"44":bdr}`, display:"flex", alignItems:"center", justifyContent:"center",
          fontSize:18, fontWeight:800, color:lc?lc.color:"#64748b" }}>{lead.name.charAt(0).toUpperCase()}</div>
        <div style={{ flex:1, minWidth:0 }}>
          <div style={{ fontWeight:700, fontSize:14, marginBottom:2, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{lead.name}</div>
          <div style={{ fontSize:11, color:sub, fontFamily:"monospace" }}>{lead.phone}</div>
        </div>
        <div style={{ display:"flex", flexDirection:"column", alignItems:"flex-end", gap:4, flexShrink:0 }}>
          {lc
            ? <span style={{ fontSize:10, fontWeight:700, padding:"3px 8px", borderRadius:50, background:lc.bg,
                color:lc.color, border:`1px solid ${lc.color}44`, whiteSpace:"nowrap" }}>{lc.icon} {lead.label}</span>
            : <span style={{ fontSize:10, color:"#475569", fontWeight:500 }}>— pendente</span>}
          <span style={{ fontSize:10, color:sub, display:"inline-block", transition:"transform .2s",
            transform:expanded?"rotate(180deg)":"rotate(0)" }}>▾</span>
        </div>
      </div>
      {expanded && (
        <div style={{ borderTop:`1px solid ${bdr}`, padding:"12px 14px", display:"flex", flexDirection:"column",
          gap:10, background:D?"rgba(14,21,37,.6)":"rgba(248,250,252,.8)", animation:"fadeIn .15s ease" }}>
          <div style={{ background:D?"#0a0f1e":"#f1f5f9", borderRadius:10, padding:"9px 12px", fontSize:12,
            color:sub, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <span>Número formatado</span>
            <span style={{ fontFamily:"monospace", fontWeight:700, color:"#818cf8" }}>{fmt015(lead.phone)}</span>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
            <a href={`tel:${fmt015(lead.phone)}`} onClick={()=>fire(`Ligando para ${lead.name}…`)}
              style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:7, padding:"14px",
                borderRadius:12, textDecoration:"none", background:"linear-gradient(135deg,#1d4ed8,#3b82f6)",
                color:"#fff", fontWeight:700, fontSize:13, boxShadow:"0 4px 16px rgba(59,130,246,.45)" }}>📞 Ligar</a>
            <a href={waLink(lead.phone, lead.name, vendorName)} target="_blank" rel="noopener noreferrer"
              style={{ display:"flex", alignItems:"center", justifyContent:"center", gap:7, padding:"14px",
                borderRadius:12, textDecoration:"none", background:"linear-gradient(135deg,#15803d,#22c55e)",
                color:"#fff", fontWeight:700, fontSize:13, boxShadow:"0 4px 16px rgba(34,197,94,.4)" }}>💬 WhatsApp</a>
          </div>
          <button onClick={onLabel} style={{ padding:"12px", borderRadius:12, border:"1.5px solid #7c3aed44",
            background:D?"rgba(124,58,237,.1)":"rgba(124,58,237,.06)", color:"#a78bfa", fontWeight:700,
            fontSize:13, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}>
            🏷️ Registrar resultado · {lead.calls} ligação{lead.calls!==1?"ões":""}
          </button>
          <div style={{ fontSize:10, color:sub, textAlign:"center" }}>Atualizado: {lead.updatedAt}</div>
        </div>
      )}
    </div>
  );
}

/* DASHBOARD */
function DashTab({ leads, D, card, bdr, txt, sub }) {
  const s = useMemo(() => {
    const total=leads.length, worked=leads.filter(l=>l.label).length;
    const sold=leads.filter(l=>l.label==="Venda realizada").length;
    const calls=leads.reduce((a,l)=>a+(l.calls||0),0);
    const rate=worked>0?((sold/worked)*100).toFixed(1):"0.0";
    const byLabel=Object.entries(LABELS).map(([label,cfg])=>{
      const count=leads.filter(l=>l.label===label).length;
      const pct=worked>0?((count/worked)*100).toFixed(1):"0.0";
      return {label,count,pct,...cfg};
    });
    const pieData=byLabel.filter(b=>b.count>0).map(b=>({name:b.label,value:b.count,color:b.color}));
    return {total,worked,pending:total-worked,sold,calls,rate,byLabel,pieData};
  }, [leads]);

  return (
    <div style={{ padding:"16px 14px", display:"flex", flexDirection:"column", gap:14 }}>
      <h2 style={{ fontWeight:800, fontSize:18 }}>Dashboard</h2>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        {[{label:"Total",v:s.total,icon:"👥",c:"#3b82f6"},{label:"Trabalhados",v:s.worked,icon:"✅",c:"#22c55e"},
          {label:"Pendentes",v:s.pending,icon:"⏳",c:"#f59e0b"},{label:"Ligações",v:s.calls,icon:"📞",c:"#818cf8"},
          {label:"Vendas",v:s.sold,icon:"🏆",c:"#c084fc"},{label:"Conversão",v:s.rate+"%",icon:"📈",c:"#f472b6"}].map(k=>(
          <div key={k.label} style={{ background:card, border:`1.5px solid ${bdr}`, borderRadius:14,
            padding:"14px", borderTop:`3px solid ${k.c}`, animation:"cardIn .35s ease both" }}>
            <div style={{ fontSize:22, marginBottom:4 }}>{k.icon}</div>
            <div style={{ fontSize:24, fontWeight:800, color:k.c, lineHeight:1 }}>{k.v}</div>
            <div style={{ fontSize:10, color:sub, marginTop:3, fontWeight:600 }}>{k.label}</div>
          </div>
        ))}
      </div>
      <div style={{ background:card, border:`1.5px solid ${bdr}`, borderRadius:14, padding:"16px" }}>
        <div style={{ fontWeight:700, fontSize:13, marginBottom:12 }}>Resultado por etiqueta</div>
        <ResponsiveContainer width="100%" height={150}>
          <BarChart data={s.byLabel} barSize={14} margin={{top:0,right:4,left:-28,bottom:0}}>
            <XAxis dataKey="icon" tick={{fontSize:16}} axisLine={false} tickLine={false} />
            <YAxis tick={{fontSize:10,fill:sub}} axisLine={false} tickLine={false} />
            <Tooltip cursor={{fill:"rgba(255,255,255,.04)"}}
              contentStyle={{background:D?"#0e1525":"#fff",border:`1px solid ${bdr}`,borderRadius:10,fontSize:12,color:txt}}
              formatter={(v,_,p)=>[v+" leads",p.payload.label]} />
            <Bar dataKey="count" radius={[6,6,0,0]}>{s.byLabel.map((e,i)=><Cell key={i} fill={e.color} />)}</Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
      {s.pieData.length>0 && (
        <div style={{ background:card, border:`1.5px solid ${bdr}`, borderRadius:14, padding:"16px" }}>
          <div style={{ fontWeight:700, fontSize:13, marginBottom:4 }}>Distribuição</div>
          <ResponsiveContainer width="100%" height={180}>
            <PieChart>
              <Pie data={s.pieData} cx="50%" cy="50%" outerRadius={72} dataKey="value"
                label={({percent})=>`${(percent*100).toFixed(0)}%`} labelLine={false} fontSize={10}>
                {s.pieData.map((e,i)=><Cell key={i} fill={e.color} />)}
              </Pie>
              <Tooltip contentStyle={{background:D?"#0e1525":"#fff",border:`1px solid ${bdr}`,borderRadius:10,fontSize:12,color:txt}} />
            </PieChart>
          </ResponsiveContainer>
        </div>
      )}
      <div style={{ background:card, border:`1.5px solid ${bdr}`, borderRadius:14, overflow:"hidden" }}>
        <div style={{ padding:"14px 16px", borderBottom:`1px solid ${bdr}`, fontWeight:700, fontSize:13 }}>Detalhamento</div>
        {s.byLabel.map((b,i)=>(
          <div key={b.label} style={{ padding:"12px 16px", borderBottom:i<s.byLabel.length-1?`1px solid ${bdr}`:"none",
            display:"flex", flexDirection:"column", gap:6 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <span style={{ fontSize:13, fontWeight:600 }}>{b.icon} {b.label}</span>
              <span style={{ fontWeight:800, color:b.color, fontSize:15 }}>{b.count}
                <span style={{ fontWeight:500, fontSize:11, color:sub }}> ({b.pct}%)</span></span>
            </div>
            <div style={{ background:D?"#1a2540":"#f1f5f9", borderRadius:99, height:5, overflow:"hidden" }}>
              <div style={{ width:b.pct+"%", height:"100%", background:b.color, borderRadius:99, transition:"width .6s ease" }} />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* BOTTOM SHEET */
function BottomSheet({ onClose, children, D, card, bdr }) {
  return (
    <div style={{ position:"fixed", inset:0, zIndex:100, display:"flex", flexDirection:"column", justifyContent:"flex-end" }}>
      <div onClick={onClose} style={{ position:"absolute", inset:0, background:"rgba(0,0,0,.55)",
        backdropFilter:"blur(4px)", animation:"fadeIn .2s ease" }} />
      <div style={{ position:"relative", background:card, borderRadius:"20px 20px 0 0", border:`1px solid ${bdr}`,
        maxHeight:"85dvh", overflowY:"auto", animation:"slideUp .28s cubic-bezier(.32,0,.67,0)",
        paddingBottom:"env(safe-area-inset-bottom,16px)" }}>
        <div style={{ display:"flex", justifyContent:"center", padding:"12px 0 4px" }}>
          <div style={{ width:36, height:4, borderRadius:99, background:D?"#1e2d4a":"#e2e8f0" }} />
        </div>
        {children}
      </div>
    </div>
  );
}

/* LABEL SHEET */
function LabelSheet({ lead, onSave, onClose, D, bdr, sub }) {
  const [sel,setSel]=useState(lead.label||null);
  return (
    <div style={{ padding:"4px 16px 20px" }}>
      <div style={{ fontWeight:800, fontSize:16, marginBottom:2 }}>{lead.name}</div>
      <div style={{ fontSize:12, color:sub, marginBottom:16 }}>📞 {lead.phone} · {lead.calls} ligação(ões)</div>
      <div style={{ display:"flex", flexDirection:"column", gap:8, marginBottom:18 }}>
        {Object.entries(LABELS).map(([label,cfg])=>(
          <button key={label} onClick={()=>setSel(label)} style={{ display:"flex", alignItems:"center", gap:12,
            padding:"14px 16px", borderRadius:12, border:`2px solid ${sel===label?cfg.color:(D?"#1a2540":"#e2e8f0")}`,
            background:sel===label?cfg.bg:"transparent", color:sel===label?cfg.color:(D?"#94a3b8":"#475569"),
            cursor:"pointer", fontWeight:sel===label?700:500, fontSize:14, transition:"all .15s", textAlign:"left" }}>
            <span style={{ fontSize:20 }}>{cfg.icon}</span>
            <span style={{ flex:1 }}>{label}</span>
            {sel===label && <span>✓</span>}
          </button>
        ))}
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 2fr", gap:10 }}>
        <GhostBtn onClick={onClose} D={D} bdr={bdr}>Cancelar</GhostBtn>
        <PrimaryBtn onClick={()=>sel&&onSave(lead.id,sel)} disabled={!sel}>Salvar Etiqueta</PrimaryBtn>
      </div>
    </div>
  );
}

/* IMPORT SHEET */
function ImportSheet({ onDone, onClose, D, bdr, sub }) {
  const [preview,setPreview]=useState(null);
  const [err,setErr]=useState("");
  const [debug,setDebug]=useState("");
  const [loading,setLoading]=useState(false);
  const fileRef=useRef();

  function handleFile(file) {
    if (!file) return;
    setErr(""); setPreview(null); setDebug(""); setLoading(true);

    if (typeof XLSX === "undefined" || !XLSX || !XLSX.read) {
      setErr("Modulo XLSX indisponivel (typeof = " + (typeof XLSX) + ")");
      setLoading(false);
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      let dbg = "Arquivo: " + file.name + " | " + e.target.result.byteLength + " bytes\n";
      try {
        const bytes = new Uint8Array(e.target.result);
        const wb = XLSX.read(bytes, { type:"array" });
        dbg += "Abas: " + wb.SheetNames.join(", ") + "\n";
        const ws = wb.Sheets[wb.SheetNames[0]];
        const mat = XLSX.utils.sheet_to_json(ws, { header:1, defval:"", raw:false, blankrows:false });
        dbg += "Linhas: " + mat.length + "\n";
        dbg += "L1: " + JSON.stringify(mat[0]) + "\n";
        dbg += "L2: " + JSON.stringify(mat[1]);

        const { leads, info, error } = parseMatrix(mat);
        dbg += "\nHeader linha: " + info.header + " | colNome: " + info.nameCol + " | colTel: " + info.phoneCol;
        dbg += "\nColunas vistas: " + JSON.stringify(info.found);

        if (error) { setErr(error + "\n\n" + dbg); setLoading(false); return; }
        setDebug(dbg);
        setPreview(leads);
      } catch (ex) {
        setErr("Erro: " + ex.message + "\n\n" + dbg);
      }
      setLoading(false);
    };
    reader.onerror = () => { setErr("Falha no FileReader."); setLoading(false); };
    reader.readAsArrayBuffer(file);
  }

  return (
    <div style={{ padding:"4px 16px 20px" }}>
      <div style={{ fontWeight:800, fontSize:16, marginBottom:14 }}>📥 Importar Planilha</div>
      <div onClick={()=>fileRef.current.click()} onDragOver={e=>e.preventDefault()}
        onDrop={e=>{e.preventDefault();handleFile(e.dataTransfer.files[0]);}}
        style={{ border:`2px dashed ${preview?"#22c55e":err?"#f87171":bdr}`, borderRadius:14, padding:"28px 16px",
          textAlign:"center", cursor:"pointer", background:D?"rgba(255,255,255,.02)":"rgba(0,0,0,.02)", marginBottom:12 }}>
        <div style={{ fontSize:36, marginBottom:8 }}>{loading?"⏳":preview?"✅":"📊"}</div>
        <p style={{ fontWeight:600, fontSize:13, marginBottom:4 }}>
          {loading?"Processando...":preview?`${preview.length} leads encontrados`:"Toque para selecionar arquivo"}
        </p>
        <p style={{ fontSize:11, color:sub }}>.xlsx, .xls ou .csv · 1ª linha: <strong>Nome</strong> | <strong>Telefone</strong></p>
      </div>
      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }}
        onChange={e=>handleFile(e.target.files[0])} />

      {err && (
        <div style={{ background:"rgba(248,113,113,.08)", border:"1px solid #f8717144", borderRadius:10,
          padding:"10px 12px", marginBottom:12, color:"#f87171", fontSize:11, lineHeight:1.6,
          whiteSpace:"pre-wrap", fontFamily:"monospace", wordBreak:"break-word" }}>⚠️ {err}</div>
      )}

      {debug && !err && (
        <div style={{ background:D?"#0a1628":"#f0f9ff", border:"1px solid #3b82f655", borderRadius:10,
          padding:"10px 12px", marginBottom:12, color:"#60a5fa", fontSize:11, lineHeight:1.6,
          whiteSpace:"pre-wrap", fontFamily:"monospace", wordBreak:"break-word" }}>🔍 {debug}</div>
      )}

      {preview && (
        <div style={{ maxHeight:160, overflowY:"auto", background:D?"#080c14":"#f8fafc", border:`1px solid ${bdr}`,
          borderRadius:10, padding:"10px", marginBottom:14 }}>
          {preview.slice(0,6).map((r,i)=>(
            <div key={i} style={{ fontSize:11, padding:"5px 0", borderBottom:`1px solid ${bdr}`, color:sub }}>
              <strong style={{ color:D?"#e2e8f0":"#0f172a" }}>{r.name}</strong>
              {" · "}<span style={{ fontFamily:"monospace" }}>{r.phone}</span>
              {" → "}<span style={{ color:"#818cf8", fontFamily:"monospace" }}>{fmt015(r.phone)}</span>
            </div>
          ))}
          {preview.length>6 && <p style={{ fontSize:10, color:sub, marginTop:6 }}>...e mais {preview.length-6} leads</p>}
        </div>
      )}

      <div style={{ display:"grid", gridTemplateColumns:"1fr 2fr", gap:10 }}>
        <GhostBtn onClick={onClose} D={D} bdr={bdr}>Cancelar</GhostBtn>
        <PrimaryBtn onClick={()=>preview&&(onDone(preview),onClose())} disabled={!preview||loading}>
          {loading?"⏳ Processando...":`Importar ${preview?preview.length:0} leads`}
        </PrimaryBtn>
      </div>
    </div>
  );
}

/* EXPORT SHEET */
function ExportSheet({ leads, onClose, fire, D, bdr, sub }) {
  const [fl, setFl]         = useState("all");
  const [vendor, setVendor] = useState("");

  const getList = () =>
    fl === "all"     ? leads :
    fl === "pending" ? leads.filter(l => !l.label) :
    leads.filter(l => l.label === fl);

  const doPDF = async () => {
    const list = getList();
    if (list.length === 0) { fire("Nenhum lead para exportar.", "err"); return; }

    let jsPDF;
    try {
      jsPDF = await loadJsPDF();
    } catch (ex) {
      fire("Erro ao carregar gerador de PDF: " + ex.message, "err");
      return;
    }

    const doc   = new jsPDF({ unit: "pt", format: "a4" });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const M     = 40;                       // margem
    const dataStr = new Date().toLocaleString("pt-BR", {
      day:"2-digit", month:"2-digit", year:"numeric", hour:"2-digit", minute:"2-digit",
    });

    /* ---- Estatisticas (dashboard) ---- */
    const total   = list.length;
    const worked  = list.filter(l => l.label).length;
    const pending = total - worked;
    const calls   = list.reduce((a, l) => a + (l.calls || 0), 0);
    const sold    = list.filter(l => l.label === "Venda realizada").length;
    const rate    = worked > 0 ? ((sold / worked) * 100).toFixed(1) : "0.0";

    const LABEL_ORDER = Object.keys(LABELS);
    const grouped = {};
    LABEL_ORDER.forEach(lab => { grouped[lab] = list.filter(l => l.label === lab); });
    const pendentes = list.filter(l => !l.label);

    let y = 0;

    /* helper: nova pagina */
    const ensure = (need) => {
      if (y + need > pageH - 40) { doc.addPage(); y = 40; }
    };

    /* ---- CABECALHO ---- */
    doc.setFillColor(37, 99, 235);
    doc.rect(0, 0, pageW, 70, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(20);
    doc.text("ProspectCRM", M, 32);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);
    doc.text("Relatorio de Leads", M, 50);
    y = 92;

    /* ---- VENDEDOR + DATA ---- */
    doc.setTextColor(60, 60, 60);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");   doc.text("Vendedor:", M, y);
    doc.setFont("helvetica", "normal"); doc.text(vendor.trim() || "-", M + 70, y);
    doc.setFont("helvetica", "bold");   doc.text("Data:", 320, y);
    doc.setFont("helvetica", "normal"); doc.text(dataStr, 360, y);
    y += 26;

    /* ---- RESUMO GERAL (cards do dashboard) ---- */
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(20, 20, 20);
    doc.text("Resumo Geral", M, y);
    y += 12;

    const kpis = [
      ["Total de leads", String(total)],
      ["Trabalhados",    String(worked)],
      ["Pendentes",      String(pending)],
      ["Ligacoes",       String(calls)],
      ["Vendas",         String(sold)],
      ["Conversao",      rate + "%"],
    ];
    const cardW = (pageW - M * 2 - 20) / 3;
    const cardH = 46;
    kpis.forEach((k, i) => {
      const col = i % 3, rowi = Math.floor(i / 3);
      const cx = M + col * (cardW + 10);
      const cy = y + rowi * (cardH + 10);
      doc.setFillColor(244, 247, 250);
      doc.roundedRect(cx, cy, cardW, cardH, 6, 6, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(18);
      doc.setTextColor(37, 99, 235);
      doc.text(k[1], cx + 12, cy + 24);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(110, 110, 110);
      doc.text(k[0], cx + 12, cy + 38);
    });
    y += cardH * 2 + 10 + 24;

    /* ---- RESULTADO POR ETIQUETA (contagem + %) ---- */
    ensure(40);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(20, 20, 20);
    doc.text("Resultado por Etiqueta", M, y);
    y += 16;

    doc.setFontSize(10);
    LABEL_ORDER.forEach(lab => {
      const cnt = grouped[lab].length;
      const pct = worked > 0 ? ((cnt / worked) * 100).toFixed(1) : "0.0";
      ensure(20);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(40, 40, 40);
      doc.text(lab, M, y);
      doc.setFont("helvetica", "bold");
      doc.text(`${cnt}  (${pct}%)`, pageW - M - 70, y);
      y += 16;
    });
    y += 14;

    /* ---- LEADS AGRUPADOS POR ETIQUETA ---- */
    const drawGroupHeader = (titulo, cor, qtd) => {
      ensure(34);
      const [r, g, b] = cor;
      doc.setFillColor(r, g, b);
      doc.roundedRect(M, y - 4, pageW - M * 2, 24, 4, 4, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(255, 255, 255);
      doc.text(`${titulo}  (${qtd})`, M + 10, y + 12);
      y += 30;
    };

    const drawRows = (arr) => {
      doc.setFontSize(9);
      arr.forEach((l, i) => {
        ensure(20);
        if (i % 2 === 0) {
          doc.setFillColor(247, 249, 252);
          doc.rect(M, y - 11, pageW - M * 2, 18, "F");
        }
        doc.setFont("helvetica", "normal");
        doc.setTextColor(30, 30, 30);
        const nome = l.name.length > 30 ? l.name.slice(0, 29) + "…" : l.name;
        doc.text(nome, M + 6, y);
        doc.text(fmt015(l.phone), M + 250, y);
        doc.text(String(l.calls) + " lig.", pageW - M - 50, y);
        y += 18;
      });
      y += 10;
    };

    const hexToRgb = (hex) => {
      const h = hex.replace("#", "");
      return [parseInt(h.slice(0,2),16), parseInt(h.slice(2,4),16), parseInt(h.slice(4,6),16)];
    };

    ensure(30);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(20, 20, 20);
    doc.text("Leads por Etiqueta", M, y);
    y += 20;

    LABEL_ORDER.forEach(lab => {
      const arr = grouped[lab];
      if (arr.length === 0) return;
      drawGroupHeader(lab, hexToRgb(LABELS[lab].color), arr.length);
      drawRows(arr);
    });

    if (pendentes.length > 0) {
      drawGroupHeader("Pendentes (sem etiqueta)", [100, 116, 139], pendentes.length);
      drawRows(pendentes);
    }

    /* ---- RODAPE ---- */
    const pages = doc.internal.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      doc.setFontSize(8);
      doc.setTextColor(150, 150, 150);
      doc.text(`Pagina ${p} de ${pages}`, pageW - 90, pageH - 20);
      doc.text("Gerado por ProspectCRM", M, pageH - 20);
    }

    const fileName = "relatorio_" + (vendor.trim() ? vendor.trim().replace(/\s+/g, "_") + "_" : "") +
      new Date().toISOString().slice(0, 10) + ".pdf";
    doc.save(fileName);
    fire("PDF exportado ✓");
    onClose();
  };

  const n = getList().length;
  return (
    <div style={{ padding:"4px 16px 20px" }}>
      <div style={{ fontWeight:800, fontSize:16, marginBottom:14 }}>📤 Exportar Relatorio (PDF)</div>

      <label style={{ fontSize:12, fontWeight:600, color:sub, display:"block", marginBottom:6 }}>Nome do vendedor</label>
      <input value={vendor} onChange={e=>setVendor(e.target.value)} placeholder="Ex: Gabriel Lemos"
        style={{ width:"100%", padding:"12px", borderRadius:12, marginBottom:14, border:`1.5px solid ${bdr}`,
          background:D?"#0e1525":"#fff", color:D?"#e2e8f0":"#0f172a", fontSize:13, outline:"none" }} />

      <label style={{ fontSize:12, fontWeight:600, color:sub, display:"block", marginBottom:6 }}>Filtrar por etiqueta</label>
      <select value={fl} onChange={e=>setFl(e.target.value)} style={{ width:"100%", padding:"12px", borderRadius:12,
        marginBottom:14, border:`1.5px solid ${bdr}`, background:D?"#0e1525":"#fff", color:D?"#e2e8f0":"#0f172a",
        fontSize:13, outline:"none", cursor:"pointer" }}>
        <option value="all">Todos os leads</option>
        <option value="pending">⏳ Pendentes</option>
        {Object.entries(LABELS).map(([l,c])=><option key={l} value={l}>{c.icon} {l}</option>)}
      </select>

      <div style={{ background:D?"#0a1628":"#f0f9ff", border:"1px solid #3b82f644", borderRadius:10,
        padding:"10px 12px", marginBottom:16, fontSize:12, color:"#60a5fa", lineHeight:1.5 }}>
        📄 O PDF terá: dados do dashboard (resumo geral + resultado por etiqueta) e os {n} lead{n!==1?"s":""} agrupados por etiqueta, com vendedor e data.
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 2fr", gap:10 }}>
        <GhostBtn onClick={onClose} D={D} bdr={bdr}>Cancelar</GhostBtn>
        <PrimaryBtn onClick={doPDF} disabled={n===0}>📄 Gerar PDF ({n})</PrimaryBtn>
      </div>
    </div>
  );
}

/* CLEAR SHEET */
function ClearSheet({ count, onConfirm, onClose, D, bdr, sub }) {
  return (
    <div style={{ padding:"4px 16px 24px" }}>
      <div style={{ textAlign:"center", marginBottom:8 }}>
        <div style={{ fontSize:44, marginBottom:8 }}>⚠️</div>
        <div style={{ fontWeight:800, fontSize:17, marginBottom:6 }}>Limpar todos os leads?</div>
        <div style={{ fontSize:13, color:sub, lineHeight:1.5 }}>
          Esta ação vai apagar <strong style={{ color:"#f87171" }}>{count} lead{count!==1?"s":""}</strong> permanentemente.
          <br/>Não é possível desfazer.
        </div>
      </div>

      <div style={{ background:D?"rgba(220,38,38,.08)":"rgba(220,38,38,.05)", border:"1px solid #dc262644",
        borderRadius:10, padding:"12px 14px", margin:"18px 0", fontSize:12, color:D?"#fca5a5":"#dc2626", lineHeight:1.5 }}>
        💡 Dica: se quiser guardar uma cópia antes, use <strong>Exportar</strong> para salvar os leads em Excel ou CSV.
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
        <GhostBtn onClick={onClose} D={D} bdr={bdr}>Cancelar</GhostBtn>
        <button onClick={onConfirm} style={{ padding:"14px", borderRadius:12, border:"none",
          background:"linear-gradient(135deg,#dc2626,#ef4444)", color:"#fff", cursor:"pointer",
          fontWeight:700, fontSize:14, boxShadow:"0 4px 18px rgba(220,38,38,.45)" }}>
          🗑️ Apagar tudo
        </button>
      </div>
    </div>
  );
}

/* CONFIG SHEET */
function ConfigSheet({ vendorName, setVendorName, onClose, fire, D, bdr, sub }) {
  const [name, setName] = useState(vendorName || "");
  const exemplo = (name.trim().split(/\s+/)[0] || "seu Consultor");
  return (
    <div style={{ padding:"4px 16px 22px" }}>
      <div style={{ fontWeight:800, fontSize:16, marginBottom:6 }}>⚙️ Configurações</div>
      <p style={{ fontSize:12, color:sub, marginBottom:16, lineHeight:1.5 }}>
        Seu nome aparece na mensagem automática do WhatsApp e no relatório em PDF.
      </p>
      <label style={{ fontSize:12, fontWeight:600, color:sub, display:"block", marginBottom:6 }}>Seu nome (vendedor)</label>
      <input value={name} onChange={e=>setName(e.target.value)} placeholder="Ex: Gabriel Lemos" autoFocus
        style={{ width:"100%", padding:"12px", borderRadius:12, marginBottom:16, border:`1.5px solid ${bdr}`,
          background:D?"#0e1525":"#fff", color:D?"#e2e8f0":"#0f172a", fontSize:14, outline:"none" }} />
      <div style={{ background:D?"#0a1628":"#f0f9ff", border:"1px solid #3b82f644", borderRadius:10,
        padding:"12px 14px", marginBottom:18, fontSize:12, color:D?"#93c5fd":"#2563eb", lineHeight:1.6 }}>
        <div style={{ fontWeight:700, marginBottom:4, color:D?"#bfdbfe":"#1d4ed8" }}>Prévia da mensagem:</div>
        Olá, <strong>Maria</strong>! Aqui é o <strong>{exemplo}</strong>, Consultor Comercial da Hapvida NotreDame. 😊
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"1fr 2fr", gap:10 }}>
        <GhostBtn onClick={onClose} D={D} bdr={bdr}>Cancelar</GhostBtn>
        <PrimaryBtn onClick={()=>{ setVendorName(name.trim()); onClose(); fire("Nome salvo ✓"); }}>Salvar</PrimaryBtn>
      </div>
    </div>
  );
}

/* SHARED */
function Btn2({ onClick, accent, D, children }) {
  return (
    <button onClick={onClick} style={{ flex:1, padding:"11px 0", borderRadius:12, border:`1.5px solid ${accent}44`,
      background:D?`${accent}14`:`${accent}0d`, color:accent, cursor:"pointer", display:"flex",
      alignItems:"center", justifyContent:"center", gap:5, fontWeight:600, fontSize:12 }}>{children}</button>
  );
}
function PrimaryBtn({ onClick, disabled, children }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{ padding:"14px", borderRadius:12, border:"none",
      background:disabled?"#1e2d4a":"linear-gradient(135deg,#2563eb,#7c3aed)", color:disabled?"#334155":"#fff",
      cursor:disabled?"not-allowed":"pointer", fontWeight:700, fontSize:14,
      boxShadow:disabled?"none":"0 4px 18px rgba(37,99,235,.4)" }}>{children}</button>
  );
}
function GhostBtn({ onClick, D, bdr, children }) {
  return (
    <button onClick={onClick} style={{ padding:"14px", borderRadius:12, border:`1px solid ${bdr}`,
      background:"transparent", color:D?"#64748b":"#94a3b8", cursor:"pointer", fontWeight:600, fontSize:13 }}>{children}</button>
  );
}
