import { useState, useEffect, useMemo } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";
import { supabase } from "./supabaseClient";

// ── Helpers ──────────────────────────────────────────────────────────────────
const n = (v) => parseFloat(String(v).replace(",", ".")) || 0;

function calc(d) {
  const revenue = n(d.revenue), cog = n(d.cog);
  const ads = n(d.ads_fb) + n(d.ads2) + n(d.ads3);
  const refunds = n(d.refunds);
  const tx = revenue * 0.05;
  const profit = revenue - cog - ads - refunds - tx;
  const roas = ads > 0 ? revenue / ads : null;
  const margin = revenue > 0 ? profit / revenue : null;
  return { ...d, revenue, cog, ads, refunds, tx, profit, roas, margin };
}

const eur = (v, d = 2) => {
  if (typeof v !== "number" || !isFinite(v)) return "—";
  return (v < 0 ? "-€" : "€") + Math.abs(v).toFixed(d);
};
const pct = (v) => typeof v === "number" && isFinite(v) ? (v * 100).toFixed(1) + "%" : "—";
const fmtDate = (s) => {
  const d = new Date(s + "T00:00:00");
  return d.toLocaleDateString("pt-PT", { day: "numeric", month: "short" });
};

const TODAY = new Date().toISOString().slice(0, 10);
const MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const QUARTERS = { Q1:[0,1,2], Q2:[3,4,5], Q3:[6,7,8], Q4:[9,10,11] };

function ChartTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  const v = payload[0]?.value;
  return (
    <div style={{ background:"#1a1a2e", border:`1px solid ${v>=0?"#4ade8044":"#f8717144"}`, borderRadius:10, padding:"10px 14px" }}>
      <div style={{ color:"#6b7280", fontSize:11, marginBottom:4 }}>{label}</div>
      <div style={{ color:v>=0?"#4ade80":"#f87171", fontSize:15, fontWeight:800 }}>{eur(v)}</div>
    </div>
  );
}

// ── AUTH SCREENS ─────────────────────────────────────────────────────────────
function AuthScreen({ onAuth }) {
  const [modo, setModo] = useState("login"); // login | registo
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState("");
  const [sucesso, setSucesso] = useState("");

  const inputSt = { width:"100%", background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:12, padding:"13px 16px", color:"#e2e8f0", fontSize:16, outline:"none", fontFamily:"inherit", boxSizing:"border-box" };

  async function handleSubmit() {
    if (!email || !password) { setErro("Preenche todos os campos."); return; }
    setLoading(true); setErro(""); setSucesso("");
    if (modo === "login") {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) setErro("Email ou password incorrectos.");
      else onAuth();
    } else {
      const { error } = await supabase.auth.signUp({ email, password });
      if (error) setErro("Erro ao criar conta. Tenta outro email.");
      else setSucesso("Conta criada! Verifica o teu email para confirmar.");
    }
    setLoading(false);
  }

  return (
    <div style={{ minHeight:"100vh", background:"#07070f", display:"flex", alignItems:"center", justifyContent:"center", padding:"20px", fontFamily:"'DM Sans',sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap'); * { box-sizing:border-box; }`}</style>
      <div style={{ width:"100%", maxWidth:400 }}>
        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", width:52, height:52, background:"linear-gradient(135deg,#6366f1,#4f46e5)", borderRadius:16, marginBottom:12 }}>
            <span style={{ color:"#fff", fontSize:24, fontWeight:800 }}>S</span>
          </div>
          <div style={{ fontSize:24, fontWeight:800, color:"#e2e8f0", letterSpacing:"-0.02em" }}>StorePNL</div>
          <div style={{ color:"#374151", fontSize:13, marginTop:4 }}>P&L diário para lojas Shopify</div>
        </div>

        {/* Card */}
        <div style={{ background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:20, padding:"28px 24px" }}>
          {/* Tabs */}
          <div style={{ display:"flex", background:"#07070f", borderRadius:12, padding:4, marginBottom:24 }}>
            {[["login","Entrar"],["registo","Criar conta"]].map(([m,l]) => (
              <button key={m} onClick={() => { setModo(m); setErro(""); setSucesso(""); }}
                style={{ flex:1, background:modo===m?"#1a1a2e":"transparent", border:"none", borderRadius:9, padding:"9px", color:modo===m?"#e2e8f0":"#374151", fontSize:14, fontWeight:700, cursor:"pointer", transition:"all 0.15s" }}>
                {l}
              </button>
            ))}
          </div>

          <div style={{ marginBottom:14 }}>
            <div style={{ color:"#4a5568", fontSize:11, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase", marginBottom:6 }}>Email</div>
            <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="teu@email.com" style={inputSt} />
          </div>
          <div style={{ marginBottom:20 }}>
            <div style={{ color:"#4a5568", fontSize:11, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase", marginBottom:6 }}>Password</div>
            <input type="password" value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSubmit()} placeholder="••••••••" style={inputSt} />
          </div>

          {erro && <div style={{ background:"#1a0a0a", border:"1px solid #f8717133", borderRadius:10, padding:"10px 14px", color:"#f87171", fontSize:13, marginBottom:14 }}>{erro}</div>}
          {sucesso && <div style={{ background:"#0a1a0a", border:"1px solid #4ade8033", borderRadius:10, padding:"10px 14px", color:"#4ade80", fontSize:13, marginBottom:14 }}>{sucesso}</div>}

          <button onClick={handleSubmit} disabled={loading}
            style={{ width:"100%", background:"linear-gradient(135deg,#6366f1,#4f46e5)", border:"none", borderRadius:12, padding:"14px", color:"#fff", fontSize:15, fontWeight:800, cursor:"pointer" }}>
            {loading ? "A carregar..." : modo==="login" ? "Entrar" : "Criar conta grátis"}
          </button>

          {modo==="registo" && (
            <div style={{ color:"#374151", fontSize:11, textAlign:"center", marginTop:12, lineHeight:1.5 }}>
              Ao criar conta, aceitas os nossos Termos de Serviço e Política de Privacidade.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(null);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [entries, setEntries] = useState([]);
  const [loadingData, setLoadingData] = useState(false);
  const [tab, setTab] = useState("home");
  const [selected, setSelected] = useState(null);
  const [editId, setEditId] = useState(null);
  const [statFilter, setStatFilter] = useState("mes");
  const [statValue, setStatValue] = useState(() => {
    const now = new Date();
    return { mes: now.getMonth(), quarter:"Q2", ano: now.getFullYear() };
  });
  const [form, setForm] = useState({ date:TODAY, revenue:"", cog:"", ads_fb:"", ads2:"", ads3:"", refunds:"" });
  const [saving, setSaving] = useState(false);
  const [deleteConfirm, setDeleteConfirm] = useState(false);

  // Auth listener
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoadingAuth(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
    });
    return () => subscription.unsubscribe();
  }, []);

  // Load entries when session changes
  useEffect(() => {
    if (session) loadEntries();
  }, [session]);

  async function loadEntries() {
    setLoadingData(true);
    const { data, error } = await supabase
      .from("entries")
      .select("*")
      .eq("user_id", session.user.id)
      .order("date", { ascending: false });
    if (!error) setEntries(data || []);
    setLoadingData(false);
  }

  async function handleSave() {
    setSaving(true);
    const payload = {
      user_id: session.user.id,
      date: form.date,
      revenue: n(form.revenue),
      cog: n(form.cog),
      ads_fb: n(form.ads_fb),
      ads2: n(form.ads2),
      ads3: n(form.ads3),
      refunds: n(form.refunds),
    };
    if (editId) {
      await supabase.from("entries").update(payload).eq("id", editId);
    } else {
      await supabase.from("entries").insert(payload);
    }
    await loadEntries();
    resetForm();
    setSaving(false);
    setTab("home");
  }

  async function handleDelete(id) {
    await supabase.from("entries").delete().eq("id", id);
    await loadEntries();
    setTab("home");
    setDeleteConfirm(false);
  }

  async function handleLogout() {
    await supabase.auth.signOut();
    setEntries([]);
    setTab("home");
  }

  function resetForm() {
    setForm({ date:TODAY, revenue:"", cog:"", ads_fb:"", ads2:"", ads3:"", refunds:"" });
    setEditId(null);
  }

  function openEdit(r) {
    setForm({ date:r.date, revenue:r.revenue, cog:r.cog, ads_fb:r.ads_fb, ads2:r.ads2||"", ads3:r.ads3||"", refunds:r.refunds });
    setEditId(r.id);
    setTab("add");
  }

  const f = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  const all = useMemo(() => entries.map(calc), [entries]);

  const statsData = useMemo(() => {
    const now = new Date();
    return all.filter(r => {
      const d = new Date(r.date + "T00:00:00");
      if (statFilter==="mes") return d.getMonth()===statValue.mes && d.getFullYear()===now.getFullYear();
      if (statFilter==="quarter") return QUARTERS[statValue.quarter].includes(d.getMonth()) && d.getFullYear()===now.getFullYear();
      if (statFilter==="ano") return d.getFullYear()===statValue.ano;
      return true;
    }).sort((a,b) => new Date(a.date)-new Date(b.date));
  }, [all, statFilter, statValue]);

  const totals = useMemo(() => {
    const t = statsData.reduce((a,r) => ({
      revenue:a.revenue+r.revenue, profit:a.profit+r.profit,
      ads:a.ads+r.ads, cog:a.cog+r.cog, refunds:a.refunds+r.refunds,
    }), {revenue:0,profit:0,ads:0,cog:0,refunds:0});
    t.roas = t.ads>0 ? t.revenue/t.ads : null;
    t.margin = t.revenue>0 ? t.profit/t.revenue : null;
    t.dias = statsData.length;
    return t;
  }, [statsData]);

  const homeChart = useMemo(() => [...all].sort((a,b)=>new Date(a.date)-new Date(b.date)).slice(-30).map(r => ({
    date: fmtDate(r.date), profit: r.profit,
  })), [all]);

  const preview = useMemo(() => (form.revenue||form.ads_fb) ? calc({...form, ads_fb:form.ads_fb}) : null, [form]);

  const card = { background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:16 };
  const mono = "'DM Mono','Courier New',monospace";

  if (loadingAuth) return (
    <div style={{ minHeight:"100vh", background:"#07070f", display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ color:"#374151", fontFamily:"sans-serif" }}>A carregar...</div>
    </div>
  );

  if (!session) return <AuthScreen onAuth={() => {}} />;

  return (
    <div style={{ minHeight:"100vh", background:"#07070f", fontFamily:"'DM Sans',sans-serif", color:"#e2e8f0", maxWidth:430, margin:"0 auto" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap');
        * { box-sizing:border-box; -webkit-tap-highlight-color:transparent; margin:0; padding:0; }
        input { font-family:'DM Mono',monospace !important; }
        input[type=number]::-webkit-inner-spin-button { -webkit-appearance:none; }
        ::-webkit-scrollbar { display:none; }
        body { background:#07070f; }
      `}</style>

      {/* ── HOME ── */}
      {tab==="home" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ padding:"56px 20px 20px", display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
            <div>
              <div style={{ color:"#374151", fontSize:12, fontWeight:600, letterSpacing:"0.08em", textTransform:"uppercase" }}>StorePNL</div>
              <div style={{ fontSize:26, fontWeight:800, letterSpacing:"-0.02em" }}>Dashboard</div>
            </div>
            <div style={{ display:"flex", gap:8, alignItems:"center" }}>
              <button onClick={handleLogout} style={{ background:"transparent", border:"1px solid #1a1a2e", borderRadius:10, padding:"7px 12px", color:"#374151", fontSize:12, cursor:"pointer" }}>Sair</button>
              <button onClick={() => { resetForm(); setTab("add"); }} style={{ background:"#6366f1", border:"none", borderRadius:12, padding:"10px 18px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>+ Dia</button>
            </div>
          </div>

          {loadingData ? (
            <div style={{ textAlign:"center", padding:"80px 20px", color:"#374151", fontSize:13 }}>A carregar dados...</div>
          ) : all.length===0 ? (
            <div style={{ textAlign:"center", padding:"80px 20px", color:"#374151" }}>
              <div style={{ fontSize:48, marginBottom:12 }}>📊</div>
              <div style={{ fontSize:16, fontWeight:700, color:"#4a5568", marginBottom:8 }}>Sem dados ainda</div>
              <div style={{ fontSize:13, marginBottom:24 }}>Adiciona o teu primeiro dia</div>
              <button onClick={() => { resetForm(); setTab("add"); }} style={{ background:"#6366f1", border:"none", borderRadius:12, padding:"12px 24px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>+ Adicionar dia</button>
            </div>
          ) : (
            <div style={{ padding:"0 16px" }}>
              {/* Este mês */}
              {(() => {
                const mes = all.filter(r => new Date(r.date+"T00:00:00").getMonth()===new Date().getMonth());
                const rev = mes.reduce((a,r)=>a+r.revenue,0);
                const prof = mes.reduce((a,r)=>a+r.profit,0);
                return (
                  <div style={{ ...card, padding:"18px 20px", marginBottom:12 }}>
                    <div style={{ color:"#374151", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:12 }}>Este Mês</div>
                    <div style={{ display:"flex", justifyContent:"space-between" }}>
                      <div><div style={{ color:"#6b7280", fontSize:11, marginBottom:4 }}>Faturação</div><div style={{ color:"#818cf8", fontSize:22, fontWeight:800, fontFamily:mono }}>{eur(rev)}</div></div>
                      <div style={{ textAlign:"right" }}><div style={{ color:"#6b7280", fontSize:11, marginBottom:4 }}>Lucro Líquido</div><div style={{ color:prof>=0?"#4ade80":"#f87171", fontSize:22, fontWeight:800, fontFamily:mono }}>{eur(prof)}</div></div>
                    </div>
                  </div>
                );
              })()}

              {/* Gráfico */}
              <div style={{ ...card, padding:"18px 16px 12px", marginBottom:12 }}>
                <div style={{ color:"#374151", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:12 }}>Lucro / Prejuízo — Últimos 30 dias</div>
                <ResponsiveContainer width="100%" height={150}>
                  <BarChart data={homeChart} barSize={homeChart.length>20?5:10} margin={{top:4,right:4,left:0,bottom:0}}>
                    <XAxis dataKey="date" tick={{fill:"#374151",fontSize:9}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                    <YAxis tick={{fill:"#374151",fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={44} />
                    <Tooltip content={<ChartTip />} cursor={{fill:"#ffffff08"}} />
                    <ReferenceLine y={0} stroke="#1a1a2e" />
                    <Bar dataKey="profit" radius={[3,3,0,0]}>
                      {homeChart.map((e,i) => <Cell key={i} fill={e.profit>=0?"#4ade8088":"#f8717188"} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Lista de dias */}
              <div style={{ color:"#374151", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:8, paddingLeft:2 }}>Dias Recentes</div>
              {all.slice(0,60).map(r => (
                <div key={r.id} onClick={() => { setSelected(r); setTab("detalhe"); }}
                  style={{ ...card, padding:"14px 18px", marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center", cursor:"pointer" }}>
                  <div>
                    <div style={{ fontSize:14, fontWeight:700, marginBottom:3 }}>{fmtDate(r.date)}</div>
                    <div style={{ color:"#374151", fontSize:12 }}>{eur(r.revenue)} faturação · {eur(r.ads)} ads</div>
                  </div>
                  <div style={{ textAlign:"right" }}>
                    <div style={{ color:r.profit>=0?"#4ade80":"#f87171", fontSize:16, fontWeight:800, fontFamily:mono }}>{eur(r.profit)}</div>
                    <div style={{ color:"#374151", fontSize:11, marginTop:2 }}>{pct(r.margin)} margem</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* ── DETALHE ── */}
      {tab==="detalhe" && selected && (() => {
        const r = calc(selected);
        return (
          <div style={{ paddingBottom:90 }}>
            <div style={{ padding:"56px 20px 20px", display:"flex", alignItems:"center", gap:14 }}>
              <button onClick={() => setTab("home")} style={{ background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", color:"#e2e8f0", fontSize:16 }}>←</button>
              <div><div style={{ color:"#374151", fontSize:12 }}>Relatório do Dia</div><div style={{ fontSize:20, fontWeight:800 }}>{fmtDate(r.date)}</div></div>
            </div>
            <div style={{ padding:"0 16px" }}>
              <div style={{ ...card, padding:"24px 20px", marginBottom:10, textAlign:"center" }}>
                <div style={{ color:"#6b7280", fontSize:12, marginBottom:8 }}>Lucro Líquido</div>
                <div style={{ color:r.profit>=0?"#4ade80":"#f87171", fontSize:40, fontWeight:800, fontFamily:mono, letterSpacing:"-0.03em" }}>{eur(r.profit)}</div>
                <div style={{ color:r.margin>=0?"#4ade80":"#f87171", fontSize:13, marginTop:6 }}>{pct(r.margin)} margem</div>
              </div>
              {[["Faturação",eur(r.revenue),"#818cf8"],["Meta Adspend",eur(r.ads),"#fbbf24"],["Custo de Produto",eur(r.cog),"#a78bfa"],["ROAS",r.roas?r.roas.toFixed(2)+"x":"—","#34d399"],["Devoluções",eur(r.refunds),"#f87171"],["Taxa Transação (5%)",eur(r.tx),"#6b7280"]].map(([l,v,c]) => (
                <div key={l} style={{ ...card, padding:"13px 18px", marginBottom:6, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ color:"#6b7280", fontSize:13 }}>{l}</span>
                  <span style={{ color:c, fontSize:15, fontWeight:700, fontFamily:mono }}>{v}</span>
                </div>
              ))}
              <div style={{ display:"flex", gap:10, marginTop:14 }}>
                <button onClick={() => openEdit(r)} style={{ flex:1, background:"#1a1a2e", border:"1px solid #2a2a4e", borderRadius:12, padding:"13px", color:"#818cf8", fontSize:14, fontWeight:700, cursor:"pointer" }}>Editar</button>
                {!deleteConfirm
                  ? <button onClick={() => setDeleteConfirm(true)} style={{ flex:1, background:"#1a0a0a", border:"1px solid #f8717133", borderRadius:12, padding:"13px", color:"#f87171", fontSize:14, fontWeight:700, cursor:"pointer" }}>Apagar</button>
                  : <button onClick={() => handleDelete(r.id)} style={{ flex:1, background:"#f87171", border:"none", borderRadius:12, padding:"13px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>Confirmar?</button>
                }
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── ANALYTICS ── */}
      {tab==="analytics" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ padding:"56px 20px 20px" }}>
            <div style={{ color:"#374151", fontSize:12, fontWeight:600, letterSpacing:"0.08em", textTransform:"uppercase" }}>StorePNL</div>
            <div style={{ fontSize:26, fontWeight:800, letterSpacing:"-0.02em" }}>Analytics</div>
          </div>
          <div style={{ padding:"0 16px" }}>
            {/* Filtros */}
            <div style={{ display:"flex", gap:6, marginBottom:12, overflowX:"auto", paddingBottom:2 }}>
              {[["mes","Mês"],["quarter","Trimestre"],["ano","Ano"],["tudo","Tudo"]].map(([v,l]) => (
                <button key={v} onClick={() => setStatFilter(v)} style={{ background:statFilter===v?"#6366f1":"#0f0f1c", border:`1px solid ${statFilter===v?"#6366f1":"#1a1a2e"}`, borderRadius:20, padding:"6px 14px", color:statFilter===v?"#fff":"#6b7280", fontSize:12, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap" }}>{l}</button>
              ))}
            </div>
            {statFilter==="mes" && (
              <div style={{ display:"flex", gap:5, marginBottom:12, overflowX:"auto", paddingBottom:2 }}>
                {MESES.map((m,i) => (
                  <button key={m} onClick={() => setStatValue(p=>({...p,mes:i}))} style={{ background:statValue.mes===i?"#1a1a2e":"transparent", border:`1px solid ${statValue.mes===i?"#2a2a4e":"#1a1a2e"}`, borderRadius:20, padding:"4px 11px", color:statValue.mes===i?"#818cf8":"#4a5568", fontSize:11, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap" }}>{m}</button>
                ))}
              </div>
            )}
            {statFilter==="quarter" && (
              <div style={{ display:"flex", gap:6, marginBottom:12 }}>
                {["Q1","Q2","Q3","Q4"].map(q => (
                  <button key={q} onClick={() => setStatValue(p=>({...p,quarter:q}))} style={{ background:statValue.quarter===q?"#1a1a2e":"transparent", border:`1px solid ${statValue.quarter===q?"#2a2a4e":"#1a1a2e"}`, borderRadius:20, padding:"5px 14px", color:statValue.quarter===q?"#818cf8":"#4a5568", fontSize:12, fontWeight:600, cursor:"pointer" }}>{q}</button>
                ))}
              </div>
            )}
            {statFilter==="ano" && (
              <div style={{ display:"flex", gap:6, marginBottom:12 }}>
                {[2024,2025,2026].map(y => (
                  <button key={y} onClick={() => setStatValue(p=>({...p,ano:y}))} style={{ background:statValue.ano===y?"#1a1a2e":"transparent", border:`1px solid ${statValue.ano===y?"#2a2a4e":"#1a1a2e"}`, borderRadius:20, padding:"5px 14px", color:statValue.ano===y?"#818cf8":"#4a5568", fontSize:12, fontWeight:600, cursor:"pointer" }}>{y}</button>
                ))}
              </div>
            )}

            {statsData.length===0 ? (
              <div style={{ textAlign:"center", padding:"60px 20px", color:"#374151", fontSize:13 }}>Sem dados para este período</div>
            ) : (
              <>
                <div style={{ ...card, padding:"20px", marginBottom:10 }}>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
                    {[["Faturação",eur(totals.revenue),"#818cf8"],["Lucro Líquido",eur(totals.profit),totals.profit>=0?"#4ade80":"#f87171"],["Adspend",eur(totals.ads),"#fbbf24"],["ROAS",totals.roas?totals.roas.toFixed(2)+"x":"—","#34d399"],["Margem",pct(totals.margin),totals.margin>=0?"#4ade80":"#f87171"],["Dias",totals.dias,"#6b7280"]].map(([l,v,c]) => (
                      <div key={l}><div style={{ color:"#4a5568", fontSize:11, marginBottom:3 }}>{l}</div><div style={{ color:c, fontSize:18, fontWeight:800, fontFamily:mono }}>{v}</div></div>
                    ))}
                  </div>
                </div>
                <div style={{ ...card, padding:"18px 16px 12px", marginBottom:10 }}>
                  <div style={{ color:"#374151", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:12 }}>Lucro / Prejuízo</div>
                  <ResponsiveContainer width="100%" height={140}>
                    <BarChart data={statsData.map(r=>({date:fmtDate(r.date),profit:r.profit}))} barSize={statsData.length>20?5:10} margin={{top:4,right:4,left:0,bottom:0}}>
                      <XAxis dataKey="date" tick={{fill:"#374151",fontSize:9}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                      <YAxis tick={{fill:"#374151",fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={44} />
                      <Tooltip content={<ChartTip />} cursor={{fill:"#ffffff08"}} />
                      <ReferenceLine y={0} stroke="#1a1a2e" />
                      <Bar dataKey="profit" radius={[3,3,0,0]}>
                        {statsData.map((e,i) => <Cell key={i} fill={e.profit>=0?"#4ade8088":"#f8717188"} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
                <div style={{ ...card, padding:"18px 20px" }}>
                  <div style={{ color:"#374151", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:12 }}>Detalhe P&L</div>
                  {[["Faturação",totals.revenue,"#818cf8"],["− Custo de Produto",totals.cog,"#a78bfa"],["− Adspend",totals.ads,"#fbbf24"],["− Devoluções",totals.refunds,"#f87171"],["− Taxas (5%)",totals.revenue*0.05,"#6b7280"],["= Lucro Líquido",totals.profit,totals.profit>=0?"#4ade80":"#f87171"]].map(([l,v,c],i) => (
                    <div key={l} style={{ display:"flex", justifyContent:"space-between", padding:"9px 0", borderTop:i===5?"1px solid #1a1a2e":i>0?"1px solid #0f0f1c":"none" }}>
                      <span style={{ color:i===5?"#e2e8f0":"#6b7280", fontSize:13, fontWeight:i===5?700:400 }}>{l}</span>
                      <span style={{ color:c, fontSize:13, fontWeight:i===5?800:600, fontFamily:mono }}>{eur(v)}</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ── ADD / EDIT ── */}
      {tab==="add" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ padding:"56px 20px 20px", display:"flex", alignItems:"center", gap:14 }}>
            <button onClick={() => { setTab("home"); resetForm(); }} style={{ background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", color:"#e2e8f0", fontSize:16 }}>←</button>
            <div style={{ fontSize:20, fontWeight:800 }}>{editId?"Editar Dia":"Adicionar Dia"}</div>
          </div>
          <div style={{ padding:"0 16px" }}>
            {[["date","Data","date"],["revenue","Faturação (€)","decimal"],["cog","Custo de Produto (€)","decimal"],["ads_fb","Meta Adspend (€)","decimal"],["ads2","Adspend 2 (€)","decimal"],["ads3","Adspend 3 (€)","decimal"],["refunds","Devoluções (€)","decimal"]].map(([key,label,mode]) => (
              <div key={key} style={{ marginBottom:12 }}>
                <div style={{ color:"#4a5568", fontSize:11, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase", marginBottom:6 }}>{label}</div>
                <input type={mode==="date"?"date":"text"} inputMode={mode} value={form[key]} onChange={f(key)} placeholder="0,00"
                  style={{ width:"100%", background:"#0f0f1c", border:"1px solid #1a1a2e", borderRadius:12, padding:"13px 16px", color:"#e2e8f0", fontSize:16, outline:"none" }} />
              </div>
            ))}
            {preview && (
              <div style={{ ...card, padding:"16px 20px", marginBottom:14 }}>
                <div style={{ color:"#374151", fontSize:11, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.07em", marginBottom:10 }}>Pré-visualização</div>
                <div style={{ display:"flex", justifyContent:"space-between" }}>
                  {[["Lucro",eur(preview.profit),preview.profit>=0?"#4ade80":"#f87171"],["ROAS",preview.roas?preview.roas.toFixed(2)+"x":"—","#34d399"],["Margem",pct(preview.margin),preview.margin>=0?"#4ade80":"#f87171"]].map(([l,v,c]) => (
                    <div key={l} style={{ textAlign:"center" }}>
                      <div style={{ color:"#4a5568", fontSize:11, marginBottom:3 }}>{l}</div>
                      <div style={{ color:c, fontSize:16, fontWeight:800, fontFamily:mono }}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
            <button onClick={handleSave} disabled={saving}
              style={{ width:"100%", background:"#6366f1", border:"none", borderRadius:14, padding:"15px", color:"#fff", fontSize:16, fontWeight:800, cursor:"pointer" }}>
              {saving?"A guardar...":editId?"Actualizar Dia":"Guardar Dia"}
            </button>
          </div>
        </div>
      )}

      {/* ── NAV ── */}
      {tab!=="add" && tab!=="detalhe" && (
        <div style={{ position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)", width:"100%", maxWidth:430, background:"#07070f", borderTop:"1px solid #1a1a2e", padding:"12px 0 28px", display:"flex" }}>
          {[["home","📊","Início"],["analytics","📈","Analytics"]].map(([t,icon,label]) => (
            <button key={t} onClick={() => setTab(t)} style={{ background:"none", border:"none", cursor:"pointer", display:"flex", flexDirection:"column", alignItems:"center", gap:4, flex:1 }}>
              <span style={{ fontSize:22 }}>{icon}</span>
              <span style={{ fontSize:10, fontWeight:700, color:tab===t?"#818cf8":"#374151", letterSpacing:"0.06em", textTransform:"uppercase" }}>{label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
