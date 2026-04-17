import { useState, useEffect, useMemo, useRef } from "react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";
import { supabase } from "./supabaseClient";

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

// ── THEME ────────────────────────────────────────────────────────────────────
const T = {
  bg: "#F8F7F5",
  surface: "#FFFFFF",
  border: "#EBEBEB",
  borderStrong: "#D4D4D4",
  text: "#1A1A1A",
  textMuted: "#8A8A8A",
  textLight: "#BABABA",
  accent: "#1A1A1A",
  accentLight: "#F0F0F0",
  green: "#16A34A",
  greenBg: "#F0FDF4",
  red: "#DC2626",
  redBg: "#FEF2F2",
  amber: "#D97706",
  amberBg: "#FFFBEB",
  purple: "#7C3AED",
  purpleBg: "#F5F3FF",
};

function ChartTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  const v = payload[0]?.value;
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px", boxShadow: "0 4px 16px rgba(0,0,0,0.08)" }}>
      <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4 }}>{label}</div>
      <div style={{ color: v >= 0 ? T.green : T.red, fontSize: 15, fontWeight: 700 }}>{eur(v)}</div>
    </div>
  );
}

// ── AUTH ─────────────────────────────────────────────────────────────────────
function AuthScreen() {
  const [modo, setModo] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState("");
  const [sucesso, setSucesso] = useState("");

  const inp = { width: "100%", background: T.bg, border: `1.5px solid ${T.border}`, borderRadius: 12, padding: "13px 16px", color: T.text, fontSize: 15, outline: "none", fontFamily: "inherit", boxSizing: "border-box", transition: "border 0.15s" };

  async function handleSubmit() {
    if (!email || !password) { setErro("Preenche todos os campos."); return; }
    setLoading(true); setErro(""); setSucesso("");
    if (modo === "login") {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) setErro("Email ou password incorrectos.");
    } else {
      const { error } = await supabase.auth.signUp({ email, password });
      if (error) setErro("Erro ao criar conta.");
      else setSucesso("Conta criada! Verifica o teu email para confirmar.");
    }
    setLoading(false);
  }

  return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "'Geist', 'DM Sans', sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap'); * { box-sizing:border-box; }`}</style>
      <div style={{ width: "100%", maxWidth: 380 }}>
        <div style={{ textAlign: "center", marginBottom: 36 }}>
          <div style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 48, height: 48, background: T.text, borderRadius: 14, marginBottom: 14 }}>
            <span style={{ color: "#fff", fontSize: 20, fontWeight: 800 }}>S</span>
          </div>
          <div style={{ fontSize: 22, fontWeight: 800, color: T.text, letterSpacing: "-0.03em" }}>StorePNL</div>
          <div style={{ color: T.textMuted, fontSize: 13, marginTop: 4 }}>P&L diário para lojas Shopify</div>
        </div>

        <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 20, padding: "28px 24px", boxShadow: "0 2px 20px rgba(0,0,0,0.06)" }}>
          <div style={{ display: "flex", background: T.bg, borderRadius: 12, padding: 4, marginBottom: 24, gap: 4 }}>
            {[["login","Entrar"],["registo","Criar conta"]].map(([m,l]) => (
              <button key={m} onClick={() => { setModo(m); setErro(""); setSucesso(""); }}
                style={{ flex: 1, background: modo === m ? T.surface : "transparent", border: modo === m ? `1px solid ${T.border}` : "1px solid transparent", borderRadius: 9, padding: "9px", color: modo === m ? T.text : T.textMuted, fontSize: 13, fontWeight: 600, cursor: "pointer", transition: "all 0.15s", boxShadow: modo === m ? "0 1px 4px rgba(0,0,0,0.06)" : "none" }}>
                {l}
              </button>
            ))}
          </div>

          <div style={{ marginBottom: 12 }}>
            <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>Email</div>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="teu@email.com" style={inp} />
          </div>
          <div style={{ marginBottom: 20 }}>
            <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>Password</div>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} onKeyDown={e => e.key === "Enter" && handleSubmit()} placeholder="••••••••" style={inp} />
          </div>

          {erro && <div style={{ background: T.redBg, border: `1px solid #FCA5A5`, borderRadius: 10, padding: "10px 14px", color: T.red, fontSize: 13, marginBottom: 14 }}>{erro}</div>}
          {sucesso && <div style={{ background: T.greenBg, border: `1px solid #86EFAC`, borderRadius: 10, padding: "10px 14px", color: T.green, fontSize: 13, marginBottom: 14 }}>{sucesso}</div>}

          <button onClick={handleSubmit} disabled={loading}
            style={{ width: "100%", background: T.text, border: "none", borderRadius: 12, padding: "14px", color: "#fff", fontSize: 15, fontWeight: 700, cursor: "pointer", letterSpacing: "-0.01em" }}>
            {loading ? "A carregar..." : modo === "login" ? "Entrar" : "Criar conta grátis"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── ONBOARDING ────────────────────────────────────────────────────────────────
function OnboardingScreen({ userId, onComplete }) {
  const [storeName, setStoreName] = useState("");
  const [logoFile, setLogoFile] = useState(null);
  const [logoPreview, setLogoPreview] = useState(null);
  const [loading, setLoading] = useState(false);
  const fileRef = useRef();

  function handleLogo(e) {
    const file = e.target.files[0];
    if (!file) return;
    setLogoFile(file);
    setLogoPreview(URL.createObjectURL(file));
  }

  async function handleSave() {
    if (!storeName.trim()) return;
    setLoading(true);
    let logo_url = null;

    if (logoFile) {
      const ext = logoFile.name.split(".").pop();
      const path = `${userId}/logo.${ext}`;
      const { error: upErr } = await supabase.storage.from("logos").upload(path, logoFile, { upsert: true });
      if (!upErr) {
        const { data } = supabase.storage.from("logos").getPublicUrl(path);
        logo_url = data.publicUrl;
      }
    }

    await supabase.from("profiles").upsert({ id: userId, store_name: storeName, logo_url });
    setLoading(false);
    onComplete({ store_name: storeName, logo_url });
  }

  return (
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "'DM Sans', sans-serif" }}>
      <div style={{ width: "100%", maxWidth: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ fontSize: 28, marginBottom: 8 }}>👋</div>
          <div style={{ fontSize: 22, fontWeight: 800, color: T.text, letterSpacing: "-0.03em" }}>Bem-vindo ao StorePNL</div>
          <div style={{ color: T.textMuted, fontSize: 14, marginTop: 6 }}>Personaliza a tua dashboard</div>
        </div>

        <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 20, padding: "28px 24px", boxShadow: "0 2px 20px rgba(0,0,0,0.06)" }}>
          {/* Logo upload */}
          <div style={{ textAlign: "center", marginBottom: 24 }}>
            <div onClick={() => fileRef.current.click()}
              style={{ width: 80, height: 80, borderRadius: 20, background: logoPreview ? "transparent" : T.bg, border: `2px dashed ${T.border}`, display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer", overflow: "hidden", transition: "border 0.15s" }}>
              {logoPreview
                ? <img src={logoPreview} alt="logo" style={{ width: "100%", height: "100%", objectFit: "cover" }} />
                : <span style={{ fontSize: 28 }}>🏪</span>
              }
            </div>
            <div style={{ color: T.textMuted, fontSize: 12, marginTop: 8 }}>Clica para adicionar o teu logo</div>
            <input ref={fileRef} type="file" accept="image/*" onChange={handleLogo} style={{ display: "none" }} />
          </div>

          <div style={{ marginBottom: 20 }}>
            <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>Nome da loja</div>
            <input type="text" value={storeName} onChange={e => setStoreName(e.target.value)} placeholder="Ex: Isadora Jewellery"
              style={{ width: "100%", background: T.bg, border: `1.5px solid ${T.border}`, borderRadius: 12, padding: "13px 16px", color: T.text, fontSize: 15, outline: "none", fontFamily: "inherit", boxSizing: "border-box" }} />
          </div>

          <button onClick={handleSave} disabled={loading || !storeName.trim()}
            style={{ width: "100%", background: storeName.trim() ? T.text : T.border, border: "none", borderRadius: 12, padding: "14px", color: storeName.trim() ? "#fff" : T.textMuted, fontSize: 15, fontWeight: 700, cursor: storeName.trim() ? "pointer" : "not-allowed", transition: "all 0.2s" }}>
            {loading ? "A guardar..." : "Começar →"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(null);
  const [loadingAuth, setLoadingAuth] = useState(true);
  const [profile, setProfile] = useState(null);
  const [entries, setEntries] = useState([]);
  const [loadingData, setLoadingData] = useState(false);
  const [tab, setTab] = useState("home");
  const [selected, setSelected] = useState(null);
  const [editId, setEditId] = useState(null);
  const [statFilter, setStatFilter] = useState("mes");
  const [statValue, setStatValue] = useState(() => {
    const now = new Date();
    return { mes: now.getMonth(), quarter: "Q2", ano: now.getFullYear() };
  });
  const [form, setForm] = useState({ date: TODAY, revenue: "", cog: "", ads_fb: "", ads2: "", ads3: "", refunds: "" });
  const [saving, setSaving] = useState(false);
  const [deleteConfirm, setDeleteConfirm] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setLoadingAuth(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, session) => setSession(session));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (session) { loadProfile(); loadEntries(); }
  }, [session]);

  async function loadProfile() {
    const { data } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
    setProfile(data);
  }

  async function loadEntries() {
    setLoadingData(true);
    const { data } = await supabase.from("Storepnl").select("*").eq("user_id", session.user.id).order("date", { ascending: false });
    setEntries(data || []);
    setLoadingData(false);
  }

  async function handleSave() {
    setSaving(true);
    const payload = { user_id: session.user.id, date: form.date, revenue: n(form.revenue), cog: n(form.cog), ads_fb: n(form.ads_fb), ads2: n(form.ads2), ads3: n(form.ads3), refunds: n(form.refunds) };
    if (editId) await supabase.from("Storepnl").update(payload).eq("id", editId);
    else await supabase.from("Storepnl").insert(payload);
    await loadEntries();
    resetForm();
    setSaving(false);
    setTab("home");
  }

  async function handleDelete(id) {
    await supabase.from("Storepnl").delete().eq("id", id);
    await loadEntries();
    setTab("home");
    setDeleteConfirm(false);
  }

  function openEdit(r) {
    setForm({ date: r.date, revenue: r.revenue, cog: r.cog, ads_fb: r.ads_fb, ads2: r.ads2 || "", ads3: r.ads3 || "", refunds: r.refunds });
    setEditId(r.id);
    setTab("add");
  }

  function resetForm() { setForm({ date: TODAY, revenue: "", cog: "", ads_fb: "", ads2: "", ads3: "", refunds: "" }); setEditId(null); }
  const f = k => e => setForm(p => ({ ...p, [k]: e.target.value }));

  const all = useMemo(() => entries.map(calc), [entries]);

  const statsData = useMemo(() => {
    const now = new Date();
    return all.filter(r => {
      const d = new Date(r.date + "T00:00:00");
      if (statFilter === "mes") return d.getMonth() === statValue.mes && d.getFullYear() === now.getFullYear();
      if (statFilter === "quarter") return QUARTERS[statValue.quarter].includes(d.getMonth()) && d.getFullYear() === now.getFullYear();
      if (statFilter === "ano") return d.getFullYear() === statValue.ano;
      return true;
    }).sort((a, b) => new Date(a.date) - new Date(b.date));
  }, [all, statFilter, statValue]);

  const totals = useMemo(() => {
    const t = statsData.reduce((a, r) => ({ revenue: a.revenue + r.revenue, profit: a.profit + r.profit, ads: a.ads + r.ads, cog: a.cog + r.cog, refunds: a.refunds + r.refunds }), { revenue: 0, profit: 0, ads: 0, cog: 0, refunds: 0 });
    t.roas = t.ads > 0 ? t.revenue / t.ads : null;
    t.margin = t.revenue > 0 ? t.profit / t.revenue : null;
    t.dias = statsData.length;
    return t;
  }, [statsData]);

  const homeChart = useMemo(() => [...all].sort((a, b) => new Date(a.date) - new Date(b.date)).slice(-30).map(r => ({ date: fmtDate(r.date), profit: r.profit })), [all]);
  const preview = useMemo(() => (form.revenue || form.ads_fb) ? calc(form) : null, [form]);

  if (loadingAuth) return <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center" }}><div style={{ color: T.textMuted, fontFamily: "sans-serif", fontSize: 14 }}>A carregar...</div></div>;
  if (!session) return <AuthScreen />;
  if (profile === null) return <div style={{ minHeight: "100vh", background: T.bg }} />;
  if (!profile) return <OnboardingScreen userId={session.user.id} onComplete={p => setProfile(p)} />;

  const storeName = profile.store_name || "A minha loja";
  const logoUrl = profile.logo_url;

  const card = { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16 };
  const mono = "'DM Mono','Courier New',monospace";

  return (
    <div style={{ minHeight: "100vh", background: T.bg, fontFamily: "'DM Sans',sans-serif", color: T.text, maxWidth: 430, margin: "0 auto" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap');
        * { box-sizing:border-box; -webkit-tap-highlight-color:transparent; margin:0; padding:0; }
        input { font-family:'DM Mono',monospace !important; }
        input[type=number]::-webkit-inner-spin-button { -webkit-appearance:none; }
        ::-webkit-scrollbar { display:none; }
        body { background:${T.bg}; }
      `}</style>

      {/* ── HOME ── */}
      {tab === "home" && (
        <div style={{ paddingBottom: 90 }}>
          {/* Header */}
          <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "52px 20px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {logoUrl
                ? <img src={logoUrl} alt="logo" style={{ width: 36, height: 36, borderRadius: 10, objectFit: "cover", border: `1px solid ${T.border}` }} />
                : <div style={{ width: 36, height: 36, background: T.text, borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center" }}><span style={{ color: "#fff", fontSize: 14, fontWeight: 800 }}>{storeName[0]}</span></div>
              }
              <div>
                <div style={{ fontSize: 16, fontWeight: 700, letterSpacing: "-0.02em", color: T.text }}>{storeName}</div>
                <div style={{ fontSize: 11, color: T.textMuted }}>Dashboard P&L</div>
              </div>
            </div>
            <button onClick={() => { resetForm(); setTab("add"); }}
              style={{ background: T.text, border: "none", borderRadius: 10, padding: "9px 16px", color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", letterSpacing: "-0.01em" }}>
              + Adicionar
            </button>
          </div>

          {loadingData ? (
            <div style={{ textAlign: "center", padding: "80px 20px", color: T.textMuted, fontSize: 14 }}>A carregar...</div>
          ) : all.length === 0 ? (
            <div style={{ textAlign: "center", padding: "80px 20px" }}>
              <div style={{ fontSize: 48, marginBottom: 12 }}>📊</div>
              <div style={{ fontSize: 18, fontWeight: 700, color: T.text, marginBottom: 6 }}>Sem dados ainda</div>
              <div style={{ fontSize: 14, color: T.textMuted, marginBottom: 24 }}>Adiciona o teu primeiro dia para começar</div>
              <button onClick={() => { resetForm(); setTab("add"); }} style={{ background: T.text, border: "none", borderRadius: 12, padding: "12px 24px", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>+ Adicionar dia</button>
            </div>
          ) : (
            <div style={{ padding: "16px" }}>
              {/* Este mês */}
              {(() => {
                const mes = all.filter(r => new Date(r.date + "T00:00:00").getMonth() === new Date().getMonth());
                const rev = mes.reduce((a, r) => a + r.revenue, 0);
                const prof = mes.reduce((a, r) => a + r.profit, 0);
                return (
                  <div style={{ ...card, padding: "20px", marginBottom: 12 }}>
                    <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 14 }}>Este Mês</div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end" }}>
                      <div>
                        <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 4 }}>Faturação</div>
                        <div style={{ color: T.text, fontSize: 26, fontWeight: 800, fontFamily: mono, letterSpacing: "-0.03em" }}>{eur(rev)}</div>
                      </div>
                      <div style={{ textAlign: "right" }}>
                        <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 4 }}>Lucro Líquido</div>
                        <div style={{ color: prof >= 0 ? T.green : T.red, fontSize: 26, fontWeight: 800, fontFamily: mono, letterSpacing: "-0.03em" }}>{eur(prof)}</div>
                      </div>
                    </div>
                  </div>
                );
              })()}

              {/* Gráfico */}
              <div style={{ ...card, padding: "18px 16px 10px", marginBottom: 12 }}>
                <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 14 }}>Lucro / Prejuízo — 30 dias</div>
                <ResponsiveContainer width="100%" height={140}>
                  <BarChart data={homeChart} barSize={homeChart.length > 20 ? 5 : 10} margin={{ top: 4, right: 4, left: 0, bottom: 0 }}>
                    <XAxis dataKey="date" tick={{ fill: T.textLight, fontSize: 9 }} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                    <YAxis tick={{ fill: T.textLight, fontSize: 9 }} axisLine={false} tickLine={false} tickFormatter={v => "€" + v} width={44} />
                    <Tooltip content={<ChartTip />} cursor={{ fill: "rgba(0,0,0,0.03)" }} />
                    <ReferenceLine y={0} stroke={T.border} />
                    <Bar dataKey="profit" radius={[4, 4, 0, 0]}>
                      {homeChart.map((e, i) => <Cell key={i} fill={e.profit >= 0 ? "#86EFAC" : "#FCA5A5"} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Lista */}
              <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 10, paddingLeft: 2 }}>Dias Recentes</div>
              {all.slice(0, 60).map(r => (
                <div key={r.id} onClick={() => { setSelected(r); setTab("detalhe"); }}
                  style={{ ...card, padding: "14px 18px", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer", transition: "box-shadow 0.15s" }}>
                  <div>
                    <div style={{ fontSize: 15, fontWeight: 600, color: T.text, marginBottom: 3 }}>{fmtDate(r.date)}</div>
                    <div style={{ color: T.textMuted, fontSize: 12 }}>{eur(r.revenue)} · {eur(r.ads)} ads</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ color: r.profit >= 0 ? T.green : T.red, fontSize: 16, fontWeight: 700, fontFamily: mono }}>{eur(r.profit)}</div>
                    <div style={{ color: T.textMuted, fontSize: 11, marginTop: 2 }}>{pct(r.margin)}</div>
                  </div>
                </div>
              ))}

              {/* Logout */}
              <button onClick={() => supabase.auth.signOut()} style={{ width: "100%", background: "transparent", border: `1px solid ${T.border}`, borderRadius: 12, padding: "12px", color: T.textMuted, fontSize: 13, cursor: "pointer", marginTop: 8 }}>Sair da conta</button>
            </div>
          )}
        </div>
      )}

      {/* ── DETALHE ── */}
      {tab === "detalhe" && selected && (() => {
        const r = calc(selected);
        return (
          <div style={{ paddingBottom: 90 }}>
            <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "52px 20px 16px", display: "flex", alignItems: "center", gap: 14 }}>
              <button onClick={() => { setTab("home"); setDeleteConfirm(false); }} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 10, width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", fontSize: 16, color: T.text }}>←</button>
              <div>
                <div style={{ color: T.textMuted, fontSize: 12 }}>Relatório do Dia</div>
                <div style={{ fontSize: 20, fontWeight: 800, letterSpacing: "-0.02em" }}>{fmtDate(r.date)}</div>
              </div>
            </div>
            <div style={{ padding: "16px" }}>
              {/* Hero */}
              <div style={{ ...card, padding: "28px 20px", marginBottom: 10, textAlign: "center" }}>
                <div style={{ color: T.textMuted, fontSize: 12, marginBottom: 8 }}>Lucro Líquido</div>
                <div style={{ color: r.profit >= 0 ? T.green : T.red, fontSize: 44, fontWeight: 800, fontFamily: mono, letterSpacing: "-0.04em" }}>{eur(r.profit)}</div>
                <div style={{ display: "inline-block", background: r.margin >= 0 ? T.greenBg : T.redBg, color: r.margin >= 0 ? T.green : T.red, fontSize: 12, fontWeight: 600, borderRadius: 20, padding: "4px 12px", marginTop: 8 }}>{pct(r.margin)} margem</div>
              </div>

              {[["Faturação", eur(r.revenue), T.text], ["Meta Adspend", eur(r.ads), T.amber], ["Custo de Produto", eur(r.cog), T.purple], ["ROAS", r.roas ? r.roas.toFixed(2) + "x" : "—", T.green], ["Devoluções", eur(r.refunds), T.red], ["Taxa (5%)", eur(r.tx), T.textMuted]].map(([l, v, c]) => (
                <div key={l} style={{ ...card, padding: "13px 18px", marginBottom: 6, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ color: T.textMuted, fontSize: 13 }}>{l}</span>
                  <span style={{ color: c, fontSize: 15, fontWeight: 700, fontFamily: mono }}>{v}</span>
                </div>
              ))}

              <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
                <button onClick={() => openEdit(r)} style={{ flex: 1, background: T.bg, border: `1px solid ${T.border}`, borderRadius: 12, padding: "13px", color: T.text, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Editar</button>
                {!deleteConfirm
                  ? <button onClick={() => setDeleteConfirm(true)} style={{ flex: 1, background: T.redBg, border: `1px solid #FCA5A5`, borderRadius: 12, padding: "13px", color: T.red, fontSize: 14, fontWeight: 600, cursor: "pointer" }}>Apagar</button>
                  : <button onClick={() => handleDelete(r.id)} style={{ flex: 1, background: T.red, border: "none", borderRadius: 12, padding: "13px", color: "#fff", fontSize: 14, fontWeight: 700, cursor: "pointer" }}>Confirmar?</button>
                }
              </div>
            </div>
          </div>
        );
      })()}

      {/* ── ANALYTICS ── */}
      {tab === "analytics" && (
        <div style={{ paddingBottom: 90 }}>
          <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "52px 20px 16px" }}>
            <div style={{ color: T.textMuted, fontSize: 12, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase" }}>StorePNL</div>
            <div style={{ fontSize: 24, fontWeight: 800, letterSpacing: "-0.03em" }}>Analytics</div>
          </div>
          <div style={{ padding: "16px" }}>
            {/* Filtros */}
            <div style={{ display: "flex", gap: 6, marginBottom: 12, overflowX: "auto", paddingBottom: 2 }}>
              {[["mes", "Mês"], ["quarter", "Trimestre"], ["ano", "Ano"], ["tudo", "Tudo"]].map(([v, l]) => (
                <button key={v} onClick={() => setStatFilter(v)} style={{ background: statFilter === v ? T.text : T.surface, border: `1px solid ${statFilter === v ? T.text : T.border}`, borderRadius: 20, padding: "6px 14px", color: statFilter === v ? "#fff" : T.textMuted, fontSize: 12, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>{l}</button>
              ))}
            </div>
            {statFilter === "mes" && (
              <div style={{ display: "flex", gap: 5, marginBottom: 12, overflowX: "auto", paddingBottom: 2 }}>
                {MESES.map((m, i) => (
                  <button key={m} onClick={() => setStatValue(p => ({ ...p, mes: i }))} style={{ background: statValue.mes === i ? T.accentLight : "transparent", border: `1px solid ${statValue.mes === i ? T.borderStrong : T.border}`, borderRadius: 20, padding: "4px 11px", color: statValue.mes === i ? T.text : T.textMuted, fontSize: 11, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>{m}</button>
                ))}
              </div>
            )}
            {statFilter === "quarter" && (
              <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
                {["Q1", "Q2", "Q3", "Q4"].map(q => (
                  <button key={q} onClick={() => setStatValue(p => ({ ...p, quarter: q }))} style={{ background: statValue.quarter === q ? T.accentLight : "transparent", border: `1px solid ${statValue.quarter === q ? T.borderStrong : T.border}`, borderRadius: 20, padding: "5px 14px", color: statValue.quarter === q ? T.text : T.textMuted, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{q}</button>
                ))}
              </div>
            )}
            {statFilter === "ano" && (
              <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
                {[2024, 2025, 2026].map(y => (
                  <button key={y} onClick={() => setStatValue(p => ({ ...p, ano: y }))} style={{ background: statValue.ano === y ? T.accentLight : "transparent", border: `1px solid ${statValue.ano === y ? T.borderStrong : T.border}`, borderRadius: 20, padding: "5px 14px", color: statValue.ano === y ? T.text : T.textMuted, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>{y}</button>
                ))}
              </div>
            )}

            {statsData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px 20px", color: T.textMuted, fontSize: 14 }}>Sem dados para este período</div>
            ) : (
              <>
                <div style={{ ...card, padding: "20px", marginBottom: 10 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    {[["Faturação", eur(totals.revenue), T.text], ["Lucro Líquido", eur(totals.profit), totals.profit >= 0 ? T.green : T.red], ["Adspend", eur(totals.ads), T.amber], ["ROAS", totals.roas ? totals.roas.toFixed(2) + "x" : "—", T.green], ["Margem", pct(totals.margin), totals.margin >= 0 ? T.green : T.red], ["Dias", totals.dias, T.textMuted]].map(([l, v, c]) => (
                      <div key={l}>
                        <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4 }}>{l}</div>
                        <div style={{ color: c, fontSize: 20, fontWeight: 800, fontFamily: mono }}>{v}</div>
                      </div>
                    ))}
                  </div>
                </div>

                <div style={{ ...card, padding: "18px 16px 10px", marginBottom: 10 }}>
                  <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 12 }}>Lucro / Prejuízo</div>
                  <ResponsiveContainer width="100%" height={130}>
                    <BarChart data={statsData.map(r => ({ date: fmtDate(r.date), profit: r.profit }))} barSize={statsData.length > 20 ? 5 : 10} margin={{ top: 4, right: 4, left: 0, bottom: 0 }}>
                      <XAxis dataKey="date" tick={{ fill: T.textLight, fontSize: 9 }} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                      <YAxis tick={{ fill: T.textLight, fontSize: 9 }} axisLine={false} tickLine={false} tickFormatter={v => "€" + v} width={44} />
                      <Tooltip content={<ChartTip />} cursor={{ fill: "rgba(0,0,0,0.03)" }} />
                      <ReferenceLine y={0} stroke={T.border} />
                      <Bar dataKey="profit" radius={[4, 4, 0, 0]}>
                        {statsData.map((e, i) => <Cell key={i} fill={e.profit >= 0 ? "#86EFAC" : "#FCA5A5"} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div style={{ ...card, padding: "18px 20px" }}>
                  <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 12 }}>Detalhe P&L</div>
                  {[["Faturação", totals.revenue, T.text], ["− Custo de Produto", totals.cog, T.purple], ["− Adspend", totals.ads, T.amber], ["− Devoluções", totals.refunds, T.red], ["− Taxas (5%)", totals.revenue * 0.05, T.textMuted], ["= Lucro Líquido", totals.profit, totals.profit >= 0 ? T.green : T.red]].map(([l, v, c], i) => (
                    <div key={l} style={{ display: "flex", justifyContent: "space-between", padding: "9px 0", borderTop: i > 0 ? `1px solid ${T.border}` : "none" }}>
                      <span style={{ color: i === 5 ? T.text : T.textMuted, fontSize: 13, fontWeight: i === 5 ? 700 : 400 }}>{l}</span>
                      <span style={{ color: c, fontSize: 13, fontWeight: i === 5 ? 800 : 600, fontFamily: mono }}>{eur(v)}</span>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      )}

      {/* ── ADD / EDIT ── */}
      {tab === "add" && (
        <div style={{ paddingBottom: 90 }}>
          <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "52px 20px 16px", display: "flex", alignItems: "center", gap: 14 }}>
            <button onClick={() => { setTab("home"); resetForm(); }} style={{ background: T.bg, border: `1px solid ${T.border}`, borderRadius: 10, width: 36, height: 36, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", fontSize: 16, color: T.text }}>←</button>
            <div style={{ fontSize: 20, fontWeight: 800, letterSpacing: "-0.02em" }}>{editId ? "Editar Dia" : "Adicionar Dia"}</div>
          </div>
          <div style={{ padding: "16px" }}>
            {[["date", "Data", "date"], ["revenue", "Faturação (€)", "decimal"], ["cog", "Custo de Produto (€)", "decimal"], ["ads_fb", "Meta Adspend (€)", "decimal"], ["ads2", "Adspend 2 (€)", "decimal"], ["ads3", "Adspend 3 (€)", "decimal"], ["refunds", "Devoluções (€)", "decimal"]].map(([key, label, mode]) => (
              <div key={key} style={{ marginBottom: 12 }}>
                <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>{label}</div>
                <input type={mode === "date" ? "date" : "text"} inputMode={mode} value={form[key]} onChange={f(key)} placeholder="0,00"
                  style={{ width: "100%", background: T.surface, border: `1.5px solid ${T.border}`, borderRadius: 12, padding: "13px 16px", color: T.text, fontSize: 15, outline: "none" }} />
              </div>
            ))}

            {preview && (
              <div style={{ ...card, padding: "16px 20px", marginBottom: 14, background: T.bg }}>
                <div style={{ color: T.textMuted, fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 12 }}>Pré-visualização</div>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  {[["Lucro", eur(preview.profit), preview.profit >= 0 ? T.green : T.red], ["ROAS", preview.roas ? preview.roas.toFixed(2) + "x" : "—", T.green], ["Margem", pct(preview.margin), preview.margin >= 0 ? T.green : T.red]].map(([l, v, c]) => (
                    <div key={l} style={{ textAlign: "center" }}>
                      <div style={{ color: T.textMuted, fontSize: 11, marginBottom: 4 }}>{l}</div>
                      <div style={{ color: c, fontSize: 16, fontWeight: 800, fontFamily: mono }}>{v}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <button onClick={handleSave} disabled={saving}
              style={{ width: "100%", background: T.text, border: "none", borderRadius: 14, padding: "15px", color: "#fff", fontSize: 16, fontWeight: 700, cursor: "pointer", letterSpacing: "-0.01em" }}>
              {saving ? "A guardar..." : editId ? "Actualizar Dia" : "Guardar Dia"}
            </button>
          </div>
        </div>
      )}

      {/* ── NAV ── */}
      {tab !== "add" && tab !== "detalhe" && (
        <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 430, background: T.surface, borderTop: `1px solid ${T.border}`, padding: "10px 0 28px", display: "flex" }}>
          {[["home", "📊", "Início"], ["analytics", "📈", "Analytics"]].map(([t, icon, label]) => (
            <button key={t} onClick={() => setTab(t)} style={{ background: "none", border: "none", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, flex: 1 }}>
              <span style={{ fontSize: 20 }}>{icon}</span>
              <span style={{ fontSize: 10, fontWeight: 700, color: tab === t ? T.text : T.textLight, letterSpacing: "0.06em", textTransform: "uppercase" }}>{label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
