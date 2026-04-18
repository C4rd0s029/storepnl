import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";

// ── Helpers ───────────────────────────────────────────────────────────────────
const n = (v) => parseFloat(String(v).replace(",", ".")) || 0;

function calc(d, txRate = 0.05) {
  const revenue = n(d.revenue), cog = n(d.cog);
  const ads = n(d.ads_fb) + n(d.ads2) + n(d.ads3);
  const refunds = n(d.refunds);
  const tx = revenue * txRate;
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
const fmtDateFull = (s) => {
  const d = new Date(s + "T00:00:00");
  return d.toLocaleDateString("pt-PT", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
};

const TODAY = new Date().toISOString().slice(0, 10);
const MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"];
const QUARTERS = { Q1:[0,1,2], Q2:[3,4,5], Q3:[6,7,8], Q4:[9,10,11] };

// ── Theme ─────────────────────────────────────────────────────────────────────
const T = {
  bg: "#F8F7F5", surface: "#FFFFFF", border: "#EBEBEB", borderStrong: "#D4D4D4",
  text: "#1A1A1A", textMuted: "#8A8A8A", textLight: "#C4C4C4",
  accent: "#1A1A1A", accentLight: "#F0F0F0",
  green: "#16A34A", greenBg: "#F0FDF4", greenBorder: "#86EFAC",
  red: "#DC2626", redBg: "#FEF2F2", redBorder: "#FCA5A5",
  amber: "#D97706", amberBg: "#FFFBEB",
  purple: "#7C3AED", purpleBg: "#F5F3FF",
  blue: "#2563EB", blueBg: "#EFF6FF",
  sidebar: "#FAFAF8",
};

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768);
  useEffect(() => {
    const h = () => setIsMobile(window.innerWidth < 768);
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, []);
  return isMobile;
}

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



// ── EXCEL IMPORT ─────────────────────────────────────────────────────────────
function ExcelImport({ userId, onImportDone, T }) {
  const [preview, setPreview] = useState(null);
  const [importing, setImporting] = useState(false);
  const [done, setDone] = useState(null);
  const [error, setError] = useState("");
  const fileRef = useRef();

  function excelDateToISO(serial) {
    if (!serial || typeof serial !== "number") return null;
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return date.toISOString().slice(0, 10);
  }

  function parseFile(e) {
    setError(""); setPreview(null); setDone(null);
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const rows = [];
        const monthSheets = wb.SheetNames.filter(n =>
          ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC",
           "JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"].includes(n.toUpperCase().slice(0,3))
        );
        const sheets = monthSheets.length > 0 ? monthSheets : wb.SheetNames.slice(0, 12);

        sheets.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (data.length < 2) return;

          // Find header row
          let headerIdx = 0;
          for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i].map(c => String(c).toLowerCase());
            if (row.some(c => c.includes("date") || c.includes("data") || c.includes("revenue") || c.includes("fatura"))) {
              headerIdx = i; break;
            }
          }

          const headers = data[headerIdx].map(h => String(h).toLowerCase().trim());
          const getCol = (keywords) => {
            for (const kw of keywords) {
              const idx = headers.findIndex(h => h.includes(kw));
              if (idx !== -1) return idx;
            }
            return -1;
          };

          const dateCol    = getCol(["date","data"]);
          const revCol     = getCol(["revenue","receita","fatura","vendas"]);
          const cogCol     = getCol(["cost of goods","cog","custo"]);
          const adsFbCol   = getCol(["adspend fb","ads fb","facebook","meta","adspend"]);
          const ads2Col    = getCol(["adspend 2","ads2","ads 2"]);
          const ads3Col    = getCol(["adspend 3","ads3","ads 3"]);
          const refundsCol = getCol(["refund","devoluc","return"]);

          for (let i = headerIdx + 1; i < data.length; i++) {
            const row = data[i];
            const rawDate = row[dateCol];
            if (!rawDate) continue;

            const date = typeof rawDate === "number"
              ? excelDateToISO(rawDate)
              : String(rawDate).slice(0, 10);

            if (!date || date < "2020-01-01") continue;

            const revenue = parseFloat(row[revCol]) || 0;
            const cog     = parseFloat(row[cogCol]) || 0;
            const ads_fb  = parseFloat(row[adsFbCol]) || 0;
            const ads2    = parseFloat(row[ads2Col]) || 0;
            const ads3    = parseFloat(row[ads3Col]) || 0;
            const refunds = parseFloat(row[refundsCol]) || 0;

            if (revenue === 0 && ads_fb === 0 && cog === 0) continue;

            rows.push({ user_id: userId, date, revenue, cog, ads_fb, ads2, ads3, refunds });
          }
        });

        if (rows.length === 0) {
          setError("Não encontrei dados válidos no ficheiro. Verifica se o Excel tem o formato correcto.");
          return;
        }

        // Deduplicate by date
        const unique = Object.values(rows.reduce((acc, r) => { acc[r.date] = r; return acc; }, {}));
        unique.sort((a, b) => a.date.localeCompare(b.date));
        setPreview(unique);
      } catch (err) {
        setError("Erro ao ler o ficheiro: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function handleImport() {
    if (!preview?.length) return;
    setImporting(true);
    // Delete existing entries for these dates first (upsert by date)
    const dates = preview.map(r => r.date);
    await supabase.from("Storepnl").delete().eq("user_id", userId).in("date", dates);
    // Insert all
    const { error: err } = await supabase.from("Storepnl").insert(preview);
    if (err) { setError("Erro ao importar: " + err.message); }
    else { setDone(preview.length); onImportDone(); }
    setImporting(false);
  }

  return (
    <div>
      {!preview && !done && (
        <div
          onClick={() => fileRef.current.click()}
          style={{ border:`2px dashed ${T.border}`, borderRadius:14, padding:"28px 20px", textAlign:"center", cursor:"pointer", background:T.bg, transition:"border 0.15s" }}>
          <div style={{ fontSize:32, marginBottom:10 }}>📊</div>
          <div style={{ fontSize:15, fontWeight:700, color:T.text, marginBottom:4 }}>Importar Excel / P&L</div>
          <div style={{ fontSize:13, color:T.textMuted }}>Clica para seleccionar o ficheiro .xlsx</div>
          <div style={{ fontSize:11, color:T.textLight, marginTop:8 }}>Compatível com o formato P&L Sheet</div>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={parseFile} style={{ display:"none" }} />
        </div>
      )}

      {error && (
        <div style={{ background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:12, padding:"14px 16px", color:T.red, fontSize:13, marginTop:12 }}>
          {error}
          <button onClick={() => { setError(""); if (fileRef.current) fileRef.current.value = ""; }}
            style={{ display:"block", marginTop:8, background:"transparent", border:"none", color:T.red, fontSize:12, cursor:"pointer", textDecoration:"underline", padding:0 }}>
            Tentar novamente
          </button>
        </div>
      )}

      {preview && !done && (
        <div>
          <div style={{ background:T.greenBg, border:`1px solid ${T.greenBorder}`, borderRadius:12, padding:"14px 16px", marginBottom:12 }}>
            <div style={{ color:T.green, fontWeight:700, fontSize:14, marginBottom:2 }}>✓ Ficheiro lido com sucesso</div>
            <div style={{ color:T.green, fontSize:13 }}>Encontrei <strong>{preview.length} dias</strong> com dados para importar.</div>
          </div>

          {/* Preview table */}
          <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:12, overflow:"hidden", marginBottom:12 }}>
            <div style={{ padding:"12px 16px", borderBottom:`1px solid ${T.border}`, display:"flex", justifyContent:"space-between" }}>
              <span style={{ fontSize:11, fontWeight:700, color:T.textMuted, textTransform:"uppercase", letterSpacing:"0.06em" }}>Pré-visualização</span>
              <span style={{ fontSize:11, color:T.textMuted }}>primeiros 5 dias</span>
            </div>
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr style={{ background:T.bg }}>
                    {["Data","Faturação","Ads","COG","Devol."].map(h => (
                      <th key={h} style={{ padding:"8px 12px", textAlign:"left", color:T.textMuted, fontWeight:600, fontSize:11, whiteSpace:"nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {preview.slice(0, 5).map((r, i) => (
                    <tr key={i} style={{ borderTop:`1px solid ${T.border}` }}>
                      <td style={{ padding:"8px 12px", fontWeight:600 }}>{r.date}</td>
                      <td style={{ padding:"8px 12px", color:T.text }}>€{r.revenue.toFixed(2)}</td>
                      <td style={{ padding:"8px 12px", color:T.amber }}>€{r.ads_fb.toFixed(2)}</td>
                      <td style={{ padding:"8px 12px", color:T.purple }}>€{r.cog.toFixed(2)}</td>
                      <td style={{ padding:"8px 12px", color:T.red }}>€{r.refunds.toFixed(2)}</td>
                    </tr>
                  ))}
                  {preview.length > 5 && (
                    <tr style={{ borderTop:`1px solid ${T.border}` }}>
                      <td colSpan={5} style={{ padding:"8px 12px", color:T.textMuted, fontSize:12, textAlign:"center" }}>
                        + {preview.length - 5} dias adicionais
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div style={{ display:"flex", gap:10 }}>
            <button onClick={() => { setPreview(null); if(fileRef.current) fileRef.current.value=""; }}
              style={{ flex:1, background:T.bg, border:`1px solid ${T.border}`, borderRadius:12, padding:"13px", color:T.textMuted, fontSize:14, fontWeight:600, cursor:"pointer" }}>
              Cancelar
            </button>
            <button onClick={handleImport} disabled={importing}
              style={{ flex:2, background:T.text, border:"none", borderRadius:12, padding:"13px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>
              {importing ? "A importar..." : `Importar ${preview.length} dias →`}
            </button>
          </div>
        </div>
      )}

      {done && (
        <div style={{ background:T.greenBg, border:`1px solid ${T.greenBorder}`, borderRadius:12, padding:"20px", textAlign:"center" }}>
          <div style={{ fontSize:32, marginBottom:8 }}>🎉</div>
          <div style={{ color:T.green, fontWeight:800, fontSize:16, marginBottom:4 }}>{done} dias importados!</div>
          <div style={{ color:T.green, fontSize:13, marginBottom:14 }}>Os dados já aparecem na tua dashboard.</div>
          <button onClick={() => { setDone(null); setPreview(null); if(fileRef.current) fileRef.current.value=""; }}
            style={{ background:"transparent", border:`1px solid ${T.greenBorder}`, borderRadius:8, padding:"8px 16px", color:T.green, fontSize:13, fontWeight:600, cursor:"pointer" }}>
            Importar outro ficheiro
          </button>
        </div>
      )}
    </div>
  );
}

// ── MOBILE ADD FORM ───────────────────────────────────────────────────────────
function MobileAddForm({ form, setForm, editId, saving, preview, onSave, onBack, f, eur, pct, mono, T }) {
  const fields = [
    { key:"revenue", label:"Faturação", emoji:"💰" },
    { key:"cog",     label:"Custo Produto", emoji:"📦" },
    { key:"ads_fb",  label:"Meta Ads", emoji:"📣" },
    { key:"ads2",    label:"Adspend 2", emoji:"📢" },
    { key:"ads3",    label:"Adspend 3", emoji:"📡" },
    { key:"refunds", label:"Devoluções", emoji:"↩️" },
  ];

  const refs = fields.reduce((acc, f) => { acc[f.key] = { current: null }; return acc; }, {});
  const [activeField, setActiveField] = useState(null);

  function handleNext(currentKey) {
    const idx = fields.findIndex(f => f.key === currentKey);
    if (idx < fields.length - 1) {
      refs[fields[idx + 1].key].current?.focus();
    } else {
      refs[fields[idx].key].current?.blur();
    }
  }

  function handleChange(key, val) {
    // Allow digits, comma and dot only
    const clean = val.replace(/[^0-9.,]/g, "");
    setForm(p => ({ ...p, [key]: clean }));
  }

  return (
    <div style={{ minHeight:"100vh", background:T.bg, fontFamily:"'DM Sans',sans-serif", display:"flex", flexDirection:"column" }}>
      {/* Header */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"52px 20px 16px", display:"flex", alignItems:"center", gap:14, flexShrink:0 }}>
        <button onClick={onBack} style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", fontSize:16, flexShrink:0 }}>←</button>
        <div>
          <div style={{ fontSize:19, fontWeight:800, letterSpacing:"-0.02em" }}>{editId ? "Editar Dia" : "Adicionar Dia"}</div>
          <div style={{ fontSize:12, color:T.textMuted }}>Preenche os valores do dia</div>
        </div>
      </div>

      {/* Date picker */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"14px 20px", flexShrink:0 }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <span style={{ fontSize:18 }}>📅</span>
            <div>
              <div style={{ fontSize:11, color:T.textMuted, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase" }}>Data</div>
              <div style={{ fontSize:14, fontWeight:600, color:T.text }}>{new Date(form.date + "T00:00:00").toLocaleDateString("pt-PT", { weekday:"short", day:"numeric", month:"short" })}</div>
            </div>
          </div>
          <input type="date" value={form.date} onChange={f("date")}
            style={{ background:"transparent", border:"none", outline:"none", fontSize:13, color:T.textMuted, cursor:"pointer", fontFamily:"inherit" }} />
        </div>
      </div>

      {/* Fields */}
      <div style={{ flex:1, overflowY:"auto", padding:"8px 0 0" }}>
        {fields.map((field, idx) => {
          const val = form[field.key];
          const isActive = activeField === field.key;
          const hasValue = val && val !== "" && val !== "0" && val !== "0,00";
          return (
            <div key={field.key}
              style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"0", transition:"all 0.15s" }}>
              <div style={{ display:"flex", alignItems:"center", padding:"14px 20px", gap:14 }}>
                <span style={{ fontSize:22, flexShrink:0 }}>{field.emoji}</span>
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:11, color:T.textMuted, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:2 }}>{field.label}</div>
                  <div style={{ display:"flex", alignItems:"center", gap:4 }}>
                    <span style={{ color:hasValue?T.text:T.textLight, fontSize:15, fontWeight:hasValue?700:400 }}>€</span>
                    <input
                      ref={el => { refs[field.key].current = el; }}
                      type="text"
                      inputMode="decimal"
                      value={val}
                      onChange={e => handleChange(field.key, e.target.value)}
                      onFocus={() => setActiveField(field.key)}
                      onBlur={() => setActiveField(null)}
                      onKeyDown={e => { if (e.key === "Enter" || e.key === "Next") { e.preventDefault(); handleNext(field.key); } }}
                      enterKeyHint={idx < fields.length - 1 ? "next" : "done"}
                      placeholder="0,00"
                      style={{ flex:1, background:"transparent", border:"none", outline:"none", fontSize:18, fontWeight:700, color:hasValue?T.text:T.textLight, fontFamily:mono, padding:0, width:"100%" }}
                    />
                  </div>
                </div>
                {hasValue && (
                  <button onClick={() => setForm(p => ({ ...p, [field.key]: "" }))}
                    style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:"50%", width:24, height:24, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", fontSize:12, color:T.textMuted, flexShrink:0, padding:0 }}>✕</button>
                )}
              </div>
              {isActive && (
                <div style={{ height:2, background:T.text, margin:"0 20px", borderRadius:2 }} />
              )}
            </div>
          );
        })}
      </div>

      {/* Preview + Save — always visible above keyboard */}
      <div style={{ background:T.surface, borderTop:`1px solid ${T.border}`, padding:"14px 20px 32px", flexShrink:0 }}>
        {preview && (
          <div style={{ display:"flex", justifyContent:"space-around", marginBottom:14, background:T.bg, borderRadius:12, padding:"12px 8px" }}>
            {[
              ["Lucro", eur(preview.profit), preview.profit>=0?T.green:T.red],
              ["ROAS", preview.roas?preview.roas.toFixed(2)+"x":"—", T.text],
              ["Margem", pct(preview.margin), preview.margin>=0?T.green:T.red],
            ].map(([l,v,c]) => (
              <div key={l} style={{ textAlign:"center" }}>
                <div style={{ color:T.textMuted, fontSize:10, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:2 }}>{l}</div>
                <div style={{ color:c, fontSize:15, fontWeight:800, fontFamily:mono }}>{v}</div>
              </div>
            ))}
          </div>
        )}
        <button onClick={onSave} disabled={saving}
          style={{ width:"100%", background:T.text, border:"none", borderRadius:14, padding:"16px", color:"#fff", fontSize:16, fontWeight:800, cursor:"pointer", letterSpacing:"-0.01em" }}>
          {saving ? "A guardar..." : editId ? "Atualizar Dia" : "Guardar Dia"}
        </button>
      </div>
    </div>
  );
}


// ── LANDING PAGE ─────────────────────────────────────────────────────────────
function LandingPage({ onStart, onCheckout, checking }) {
  const isMobile = useIsMobile();

  const S = {
    bg: "#F8F7F5", surface: "#FFFFFF", border: "#EBEBEB",
    text: "#1A1A1A", muted: "#8A8A8A", light: "#C4C4C4",
    green: "#16A34A", greenBg: "#F0FDF4", greenBorder: "#86EFAC",
    mono: "'DM Mono','Courier New',monospace",
  };

  return (
    <div style={{ minHeight:"100vh", background:S.bg, fontFamily:"'DM Sans',sans-serif", color:S.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        html { scroll-behavior: smooth; }
        .lp-feature:hover { border-color: #D4D4D4 !important; }
        .lp-faq { cursor:pointer; }
        .lp-faq-answer { display:none; }
        .lp-faq.open .lp-faq-answer { display:block; }
      `}</style>

      {/* NAV */}
      <nav style={{ position:"fixed", top:0, left:0, right:0, zIndex:100, background:"rgba(248,247,245,0.92)", backdropFilter:"blur(12px)", borderBottom:`1px solid ${S.border}`, padding:"0 24px", height:60, display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          <div style={{ width:34, height:34, background:S.text, borderRadius:10, display:"flex", alignItems:"center", justifyContent:"center" }}>
            <div style={{ display:"flex", alignItems:"flex-end", gap:2, height:16 }}>
              {[6,9,12,16].map((h,i) => <div key={i} style={{ width:4, height:h, background:"#4ade80", borderRadius:"2px 2px 0 0", opacity:[0.3,0.5,0.75,1][i] }} />)}
            </div>
          </div>
          <span style={{ fontSize:16, fontWeight:800, letterSpacing:"-0.02em" }}>StorePNL</span>
        </div>
        <button onClick={() => onStart("login")} style={{ background:S.text, color:"#fff", border:"none", borderRadius:8, padding:"8px 18px", fontSize:13, fontWeight:700, cursor:"pointer", fontFamily:"inherit" }}>Entrar</button>
      </nav>

      {/* HERO */}
      <div style={{ minHeight:"100vh", paddingTop:60, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", textAlign:"center", padding:"80px 24px 60px" }}>
        <div style={{ display:"inline-flex", alignItems:"center", gap:6, background:S.surface, border:`1px solid ${S.border}`, borderRadius:20, padding:"6px 14px", fontSize:12, fontWeight:600, color:S.muted, marginBottom:28 }}>
          <div style={{ width:6, height:6, background:S.green, borderRadius:"50%" }} />
          Para lojas Shopify
        </div>

        <h1 style={{ fontSize:isMobile?"38px":"68px", fontWeight:800, letterSpacing:"-0.04em", lineHeight:1.05, marginBottom:20, maxWidth:760 }}>
          Para de adivinhar.<br/>Sabe o teu{" "}
          <span style={{ color:S.green }}>lucro real</span><br/>
          todos os dias.
        </h1>

        <p style={{ fontSize:isMobile?"16px":"19px", color:S.muted, maxWidth:500, marginBottom:36, lineHeight:1.6 }}>
          O StorePNL substitui o teu Excel de P&L por uma dashboard limpa, rápida e sempre no teu bolso.
        </p>

        <div style={{ display:"flex", gap:12, flexWrap:"wrap", justifyContent:"center", marginBottom:14 }}>
          <button onClick={onCheckout} disabled={checking} style={{ background:S.text, color:"#fff", border:"none", borderRadius:12, padding:"14px 28px", fontSize:16, fontWeight:700, cursor:checking?"not-allowed":"pointer", fontFamily:"inherit", opacity:checking?0.7:1 }}>
            {checking ? "A redirecionar..." : "Começar agora — €4/mês"}
          </button>
          <a href="#features" style={{ background:S.surface, color:S.text, border:`1px solid ${S.border}`, borderRadius:12, padding:"14px 28px", fontSize:16, fontWeight:600, cursor:"pointer", textDecoration:"none" }}>
            Ver como funciona
          </a>
        </div>
        <div style={{ fontSize:12, color:S.light }}>Cancela quando quiseres · Sem compromissos</div>

        {/* App mockup */}
        <div style={{ width:"100%", maxWidth:320, marginTop:52, background:S.surface, border:`1px solid ${S.border}`, borderRadius:24, overflow:"hidden", boxShadow:"0 24px 60px rgba(0,0,0,0.09)" }}>
          <div style={{ background:S.surface, borderBottom:`1px solid ${S.border}`, padding:"14px 18px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <div style={{ width:30, height:30, background:S.text, borderRadius:8, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontSize:11, fontWeight:800 }}>M</div>
              <div><div style={{ fontSize:12, fontWeight:700 }}>A minha loja</div><div style={{ fontSize:9, color:S.muted }}>Dashboard P&L</div></div>
            </div>
            <div style={{ background:S.text, color:"#fff", borderRadius:7, padding:"5px 10px", fontSize:10, fontWeight:700 }}>+ Dia</div>
          </div>
          <div style={{ padding:12 }}>
            <div style={{ background:S.greenBg, border:`1px solid ${S.greenBorder}`, borderRadius:12, padding:14, marginBottom:8 }}>
              <div style={{ fontSize:9, fontWeight:700, color:S.muted, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:3 }}>Hoje</div>
              <div style={{ fontFamily:S.mono, fontSize:28, fontWeight:800, color:S.green, letterSpacing:"-0.03em" }}>€127.43</div>
              <div style={{ fontSize:9, color:S.muted, marginTop:3 }}>38.2% margem · ROAS 2.8x</div>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:6, marginBottom:8 }}>
              {[["Faturação","€334.00","#2563EB","↑ €42"],[" Adspend","€89.50","#D97706","↑ €12"]].map(([l,v,c,d])=>(
                <div key={l} style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:8, padding:"9px 10px" }}>
                  <div style={{ fontSize:8, fontWeight:700, color:S.muted, textTransform:"uppercase", letterSpacing:"0.06em" }}>{l}</div>
                  <div style={{ fontFamily:S.mono, fontSize:14, fontWeight:800, color:c, marginTop:3 }}>{v}</div>
                  <div style={{ fontSize:8, color:S.green, marginTop:2 }}>{d} vs ontem</div>
                </div>
              ))}
            </div>
            <div style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:8, padding:"9px 10px" }}>
              <div style={{ fontSize:8, fontWeight:700, color:S.muted, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>30 dias</div>
              <div style={{ display:"flex", alignItems:"flex-end", gap:2, height:32 }}>
                {[40,30,55,45,70,80,60,35,90,75,85,100].map((h,i)=>(
                  <div key={i} style={{ flex:1, height:`${h}%`, background:h>50?"#86EFAC":"#FCA5A5", borderRadius:"2px 2px 0 0" }} />
                ))}
              </div>
            </div>
          </div>
          <div style={{ borderTop:`1px solid ${S.border}`, padding:"8px 0 14px", display:"flex", justifyContent:"space-around" }}>
            {[["▦","Início",true],["◈","Analytics",false],["⚙","Config",false]].map(([icon,label,active])=>(
              <div key={label} style={{ textAlign:"center" }}>
                <div style={{ fontSize:16, color:active?S.text:S.light }}>{icon}</div>
                <div style={{ fontSize:8, fontWeight:700, color:active?S.text:S.light, letterSpacing:"0.06em", textTransform:"uppercase" }}>{label}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* FEATURES */}
      <div id="features" style={{ background:S.surface, borderTop:`1px solid ${S.border}`, borderBottom:`1px solid ${S.border}`, padding:"72px 24px" }}>
        <div style={{ maxWidth:1000, margin:"0 auto" }}>
          <div style={{ textAlign:"center", marginBottom:48 }}>
            <div style={{ fontSize:11, fontWeight:700, color:S.muted, letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:10 }}>Como funciona</div>
            <h2 style={{ fontSize:isMobile?"28px":"40px", fontWeight:800, letterSpacing:"-0.03em", marginBottom:12 }}>Tudo o que precisas.<br/>Nada do que não usas.</h2>
            <p style={{ fontSize:16, color:S.muted, maxWidth:460, margin:"0 auto" }}>Criado por um lojista Shopify que estava farto de Excel.</p>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:`repeat(${isMobile?1:3}, 1fr)`, gap:14 }}>
            {[
              ["⚡","Registo em segundos","Faturação, adspend, custos. Em 30 segundos sabes se hoje foi verde ou vermelho."],
              ["📊","Lucro real","Vê o lucro líquido depois de COG, adspend, taxas e devoluções. Sem surpresas."],
              ["📈","Compara com ontem","Cada dia mostra a comparação automática com o anterior. Sabes se estás a melhorar."],
              ["📱","Mobile first","Adiciona ao ecrã inicial do iPhone como atalho. Abre em 2 segundos, sem browser, sem login todas as vezes."],
              ["📁","Importa o teu Excel","Tens um P&L Sheet? Faz upload directo e tudo aparece na dashboard em segundos."],
              ["🏪","A tua marca","Adiciona o logo e nome da tua loja. A dashboard fica personalizada só para ti."],
            ].map(([icon,title,desc])=>(
              <div key={title} className="lp-feature" style={{ background:S.bg, border:`1px solid ${S.border}`, borderRadius:16, padding:24, transition:"border 0.15s" }}>
                <div style={{ fontSize:26, marginBottom:12 }}>{icon}</div>
                <div style={{ fontSize:16, fontWeight:700, marginBottom:8 }}>{title}</div>
                <div style={{ fontSize:13, color:S.muted, lineHeight:1.6 }}>{desc}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* STORY */}
      <div style={{ padding:"72px 24px" }}>
        <div style={{ maxWidth:620, margin:"0 auto", background:S.surface, border:`1px solid ${S.border}`, borderRadius:22, padding:isMobile?"24px":"44px" }}>
          <p style={{ fontSize:isMobile?"18px":"22px", fontWeight:500, lineHeight:1.5, letterSpacing:"-0.01em", marginBottom:24 }}>
            "Estava a gerir a minha loja com um Excel enorme que demorava imenso a atualizar. Nunca sabia exactamente{" "}
            <span style={{ color:S.green, fontWeight:700 }}>quanto estava a ganhar</span>{" "}
            depois de todos os custos. Construí o StorePNL para mim — e percebi que toda a gente precisava disto."
          </p>
          <div style={{ display:"flex", alignItems:"center", gap:12 }}>
            <div style={{ width:42, height:42, background:S.text, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontSize:15, fontWeight:800, flexShrink:0 }}>L</div>
            <div>
              <div style={{ fontSize:14, fontWeight:700 }}>Luís — Fundador do StorePNL</div>
              <div style={{ fontSize:12, color:S.muted }}>Criador de conteúdo de e-commerce</div>
            </div>
          </div>
        </div>
      </div>

      {/* PRICING */}
      <div style={{ background:S.surface, borderTop:`1px solid ${S.border}`, borderBottom:`1px solid ${S.border}`, padding:"72px 24px" }}>
        <div style={{ maxWidth:400, margin:"0 auto", textAlign:"center" }}>
          <div style={{ fontSize:11, fontWeight:700, color:S.muted, letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:10 }}>Preço</div>
          <h2 style={{ fontSize:isMobile?"28px":"40px", fontWeight:800, letterSpacing:"-0.03em", marginBottom:12 }}>Simples como devia ser.</h2>
          <p style={{ fontSize:16, color:S.muted, marginBottom:40 }}>Um plano. Um preço. Sem surpresas.</p>
          <div style={{ background:S.text, borderRadius:20, padding:"32px 28px", color:"#fff", textAlign:"left" }}>
            <div style={{ fontSize:12, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", color:"rgba(255,255,255,0.5)", marginBottom:12 }}>StorePNL</div>
            <div style={{ fontFamily:S.mono, fontSize:52, fontWeight:800, letterSpacing:"-0.04em", marginBottom:4 }}>€4</div>
            <div style={{ fontSize:13, color:"rgba(255,255,255,0.5)", marginBottom:24 }}>por mês · cancela quando quiseres</div>
            <ul style={{ listStyle:"none", marginBottom:28 }}>
              {["Dashboard completa","Registo diário em segundos","Comparação vs ontem","Analytics por período","Import de Excel","Logo e nome da loja","Mobile + Desktop","Sincronização multi-dispositivo"].map(f=>(
                <li key={f} style={{ fontSize:14, padding:"7px 0", borderBottom:"1px solid rgba(255,255,255,0.1)", display:"flex", alignItems:"center", gap:8 }}>
                  <span style={{ color:"#4ade80", fontWeight:700 }}>✓</span> {f}
                </li>
              ))}
            </ul>
            <button onClick={onCheckout} disabled={checking} style={{ width:"100%", background:"#fff", border:"none", borderRadius:12, padding:14, fontSize:15, fontWeight:700, cursor:checking?"not-allowed":"pointer", fontFamily:"inherit", color:S.text, opacity:checking?0.7:1 }}>
              {checking ? "A redirecionar..." : "Começar agora →"}
            </button>
          </div>
        </div>
      </div>

      {/* FAQ */}
      <div style={{ padding:"72px 24px" }}>
        <div style={{ maxWidth:600, margin:"0 auto" }}>
          <div style={{ textAlign:"center", marginBottom:40 }}>
            <div style={{ fontSize:11, fontWeight:700, color:S.muted, letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:10 }}>FAQ</div>
            <h2 style={{ fontSize:isMobile?"28px":"36px", fontWeight:800, letterSpacing:"-0.03em" }}>Tens dúvidas?</h2>
          </div>
          {[
            ["Preciso de ligar a minha loja Shopify?", "Não. O StorePNL é manual — tu inseris os dados que queres. Não acessa à tua loja, não precisa de permissões. Simples e rápido."],
            ["Os meus dados são privados?", "Sim. Cada conta é completamente isolada. Os teus dados só são acessíveis por ti, protegidos por autenticação segura."],
            ["Posso cancelar quando quiser?", "Sim, sem qualquer penalização. Cancelas a qualquer momento."],
            ["Funciona para outras plataformas?", "Por agora focamos em Shopify, mas como é inserção manual podes usar para qualquer loja."],
            ["Tenho dados no Excel, consigo importar?", "Sim! Nas Definições há um botão de import directo de ficheiros .xlsx. Em segundos tens o histórico completo."],
          ].map(([q,a],i) => (
            <div key={i} className="lp-faq" style={{ background:S.surface, border:`1px solid ${S.border}`, borderRadius:14, padding:"18px 22px", marginBottom:8 }}
              onClick={e => e.currentTarget.classList.toggle("open")}>
              <div style={{ fontSize:15, fontWeight:600, display:"flex", justifyContent:"space-between", alignItems:"center", gap:12 }}>
                {q} <span style={{ color:S.muted, fontSize:12, flexShrink:0 }}>▼</span>
              </div>
              <div className="lp-faq-answer" style={{ fontSize:14, color:S.muted, marginTop:12, lineHeight:1.6 }}>{a}</div>
            </div>
          ))}
        </div>
      </div>

      {/* CTA BOTTOM */}
      <div style={{ background:S.text, borderRadius:24, padding:isMobile?"40px 20px":"64px 32px", textAlign:"center", margin:`0 ${isMobile?"12px":"24px"} 80px`, maxWidth:900, marginLeft:"auto", marginRight:"auto" }}>
        <h2 style={{ fontSize:isMobile?"28px":"44px", fontWeight:800, color:"#fff", letterSpacing:"-0.03em", marginBottom:12 }}>Começa hoje.<br/>€4/mês.</h2>
        <p style={{ color:"rgba(255,255,255,0.5)", fontSize:16, marginBottom:32 }}>Junta-te a lojistas que já sabem exactamente quanto ganham todos os dias.</p>
        <button onClick={onCheckout} disabled={checking} style={{ background:"#fff", color:S.text, border:"none", borderRadius:12, padding:"14px 32px", fontSize:16, fontWeight:700, cursor:checking?"not-allowed":"pointer", fontFamily:"inherit", opacity:checking?0.7:1 }}>
          {checking ? "A redirecionar..." : "Começar agora →"}
        </button>
        <p style={{ color:"rgba(255,255,255,0.3)", fontSize:12, marginTop:14 }}>€4/mês · Cancela quando quiseres · Sem compromissos</p>
      </div>

      {/* FOOTER */}
      <div style={{ borderTop:`1px solid ${S.border}`, padding:"24px", display:"flex", justifyContent:"space-between", alignItems:"center", flexWrap:"wrap", gap:12, maxWidth:1100, margin:"0 auto" }}>
        <div style={{ fontSize:13, color:S.muted }}>© 2026 StorePNL · Feito em Portugal 🇵🇹</div>
        <div style={{ display:"flex", gap:20 }}>
          {["Termos","Privacidade","Contacto"].map(l=>(
            <a key={l} href="#" style={{ fontSize:13, color:S.muted, textDecoration:"none" }}>{l}</a>
          ))}
        </div>
      </div>
    </div>
  );
}


// ── PAYWALL ───────────────────────────────────────────────────────────────────
function PaywallScreen({ onUpgrade, onLogout, email, checking, T }) {
  const isMobile = typeof window !== "undefined" && window.innerWidth < 768;
  return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", padding:20, fontFamily:"'DM Sans',sans-serif" }}>
      <div style={{ width:"100%", maxWidth:400, textAlign:"center" }}>
        <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", width:52, height:52, background:T.text, borderRadius:16, marginBottom:20 }}>
          <div style={{ display:"flex", alignItems:"flex-end", gap:2, height:20 }}>
            {[8,12,16,20].map((h,i) => <div key={i} style={{ width:4, height:h, background:"#4ade80", borderRadius:"2px 2px 0 0", opacity:[0.3,0.5,0.75,1][i] }} />)}
          </div>
        </div>
        <div style={{ fontSize:24, fontWeight:800, letterSpacing:"-0.03em", marginBottom:8 }}>Começa a usar o StorePNL</div>
        <div style={{ color:T.textMuted, fontSize:15, marginBottom:36, lineHeight:1.6 }}>Subscreve para ter acesso completo à tua dashboard P&L.</div>

        <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:20, padding:"28px 24px", marginBottom:14, textAlign:"left" }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"baseline", marginBottom:20 }}>
            <div style={{ fontSize:16, fontWeight:700 }}>StorePNL Pro</div>
            <div style={{ fontFamily:"'DM Mono',monospace", fontSize:28, fontWeight:800, letterSpacing:"-0.03em" }}>€4<span style={{ fontSize:14, fontWeight:400, color:T.textMuted }}>/mês</span></div>
          </div>
          <ul style={{ listStyle:"none", marginBottom:24 }}>
            {["Dashboard P&L diária","Comparação vs ontem","Analytics por período","Import de Excel (.xlsx)","Mobile + Desktop","Logo e nome da loja"].map(f => (
              <li key={f} style={{ fontSize:14, padding:"7px 0", borderBottom:`1px solid ${T.border}`, display:"flex", alignItems:"center", gap:8 }}>
                <span style={{ color:T.green, fontWeight:700 }}>✓</span> {f}
              </li>
            ))}
          </ul>
          <button onClick={onUpgrade} disabled={checking}
            style={{ width:"100%", background:T.text, border:"none", borderRadius:12, padding:"14px", color:"#fff", fontSize:15, fontWeight:700, cursor:checking?"not-allowed":"pointer", fontFamily:"inherit", opacity:checking?0.7:1 }}>
            {checking ? "A redirecionar..." : "Subscrever agora — €4/mês →"}
          </button>
        </div>
        <div style={{ fontSize:12, color:T.textMuted, marginBottom:16 }}>Cancela quando quiseres · Sem compromissos</div>
        <button onClick={onLogout} style={{ background:"transparent", border:"none", color:T.textMuted, fontSize:12, cursor:"pointer", textDecoration:"underline", fontFamily:"inherit" }}>
          Sair ({email})
        </button>
      </div>
    </div>
  );
}

// ── AUTH ──────────────────────────────────────────────────────────────────────
function AuthScreen({ initialModo = "login", paymentSuccess = false }) {
  const [modo, setModo] = useState(initialModo);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [erro, setErro] = useState("");
  const [sucesso, setSucesso] = useState("");
  const isMobile = useIsMobile();

  const inp = { width:"100%", background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:12, padding:"13px 16px", color:T.text, fontSize:15, outline:"none", fontFamily:"inherit", boxSizing:"border-box" };

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
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", fontFamily:"'DM Sans',sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap'); *{box-sizing:border-box;}`}</style>

      {/* Left panel — only desktop */}
      {!isMobile && (
        <div style={{ flex:1, background:T.text, display:"flex", flexDirection:"column", justifyContent:"space-between", padding:"48px 56px" }}>
          <div style={{ display:"flex", alignItems:"center", gap:10 }}>
            <div style={{ width:36, height:36, background:"#fff", borderRadius:10, display:"flex", alignItems:"center", justifyContent:"center" }}>
              <span style={{ color:T.text, fontSize:16, fontWeight:800 }}>S</span>
            </div>
            <span style={{ color:"#fff", fontSize:18, fontWeight:800, letterSpacing:"-0.02em" }}>StorePNL</span>
          </div>
          <div>
            <div style={{ color:"#fff", fontSize:36, fontWeight:800, letterSpacing:"-0.03em", lineHeight:1.2, marginBottom:16 }}>O teu P&L<br/>diário.<br/>Simples.</div>
            <div style={{ color:"rgba(255,255,255,0.5)", fontSize:15, lineHeight:1.6 }}>Regista a tua faturação, custos e adspend.<br/>Vê o teu lucro em tempo real.</div>
          </div>
          <div style={{ color:"rgba(255,255,255,0.3)", fontSize:12 }}>© 2026 StorePNL · Para lojas Shopify</div>
        </div>
      )}

      {/* Right panel — form */}
      <div style={{ width:isMobile?"100%":480, display:"flex", alignItems:"center", justifyContent:"center", padding:isMobile?"20px":"48px" }}>
        <div style={{ width:"100%", maxWidth:380 }}>
          {isMobile && (
            <div style={{ textAlign:"center", marginBottom:32 }}>
              <div style={{ display:"inline-flex", alignItems:"center", justifyContent:"center", width:48, height:48, background:T.text, borderRadius:14, marginBottom:12 }}>
                <span style={{ color:"#fff", fontSize:20, fontWeight:800 }}>S</span>
              </div>
              <div style={{ fontSize:22, fontWeight:800, color:T.text, letterSpacing:"-0.03em" }}>StorePNL</div>
              <div style={{ color:T.textMuted, fontSize:13, marginTop:4 }}>P&L diário para lojas Shopify</div>
            </div>
          )}

          {!isMobile && <div style={{ fontSize:26, fontWeight:800, color:T.text, letterSpacing:"-0.03em", marginBottom:8 }}>Bem-vindo de volta</div>}
          {!isMobile && <div style={{ color:T.textMuted, fontSize:14, marginBottom:32 }}>Entra na tua conta para continuar</div>}

          <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:20, padding:"28px 24px", boxShadow:"0 2px 20px rgba(0,0,0,0.06)" }}>
            {!paymentSuccess && initialModo !== "registo" && (
              <div style={{ display:"flex", background:T.bg, borderRadius:12, padding:4, marginBottom:24, gap:4 }}>
                {[["login","Entrar"]].map(([m,l]) => (
                  <button key={m} onClick={() => { setModo(m); setErro(""); setSucesso(""); }}
                    style={{ flex:1, background:modo===m?T.surface:"transparent", border:modo===m?`1px solid ${T.border}`:"1px solid transparent", borderRadius:9, padding:"9px", color:modo===m?T.text:T.textMuted, fontSize:13, fontWeight:600, cursor:"pointer", transition:"all 0.15s", boxShadow:modo===m?"0 1px 4px rgba(0,0,0,0.06)":"none" }}>
                    {l}
                  </button>
                ))}
              </div>
            )}
            {(paymentSuccess || initialModo === "registo") && (
              <div style={{ marginBottom:20 }}>
                <div style={{ fontSize:16, fontWeight:700, color:T.text, marginBottom:4 }}>Cria a tua conta</div>
                <div style={{ fontSize:13, color:T.textMuted }}>Define o teu email e password para entrar.</div>
              </div>
            )}

            <div style={{ marginBottom:12 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>Email</div>
              <input type="email" value={email} onChange={e=>setEmail(e.target.value)} placeholder="teu@email.com" style={inp} />
            </div>
            <div style={{ marginBottom:20 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>Password</div>
              <input type="password" value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleSubmit()} placeholder="••••••••" style={inp} />
            </div>

            {paymentSuccess && <div style={{ background:"#f0fdf4", border:"1px solid #86EFAC", borderRadius:10, padding:"10px 14px", color:"#16A34A", fontSize:13, marginBottom:14, fontWeight:600 }}>✓ Pagamento confirmado! Cria a tua conta para entrar.</div>}
          {erro && <div style={{ background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:10, padding:"10px 14px", color:T.red, fontSize:13, marginBottom:14 }}>{erro}</div>}
            {sucesso && <div style={{ background:T.greenBg, border:`1px solid ${T.greenBorder}`, borderRadius:10, padding:"10px 14px", color:T.green, fontSize:13, marginBottom:14 }}>{sucesso}</div>}

            <button onClick={handleSubmit} disabled={loading}
              style={{ width:"100%", background:T.text, border:"none", borderRadius:12, padding:"14px", color:"#fff", fontSize:15, fontWeight:700, cursor:"pointer", letterSpacing:"-0.01em" }}>
              {loading ? "A carregar..." : (paymentSuccess || initialModo === "registo") ? "Criar conta e entrar →" : "Entrar"}
            </button>
          </div>
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
      const { error: upErr } = await supabase.storage.from("Logos").upload(path, logoFile, { upsert: true });
      if (!upErr) {
        const { data } = supabase.storage.from("Logos").getPublicUrl(path);
        logo_url = data.publicUrl;
      }
    }
    await supabase.from("profiles").upsert({ id: userId, store_name: storeName, logo_url });
    setLoading(false);
    onComplete({ store_name: storeName, logo_url });
  }

  return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center", padding:20, fontFamily:"'DM Sans',sans-serif" }}>
      <div style={{ width:"100%", maxWidth:400 }}>
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ fontSize:32, marginBottom:10 }}>👋</div>
          <div style={{ fontSize:24, fontWeight:800, color:T.text, letterSpacing:"-0.03em" }}>Bem-vindo ao StorePNL</div>
          <div style={{ color:T.textMuted, fontSize:14, marginTop:6 }}>Personaliza a tua dashboard em 30 segundos</div>
        </div>
        <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:20, padding:"28px 24px", boxShadow:"0 2px 20px rgba(0,0,0,0.06)" }}>
          <div style={{ textAlign:"center", marginBottom:24 }}>
            <div onClick={() => fileRef.current.click()} style={{ width:80, height:80, borderRadius:20, background:logoPreview?"transparent":T.bg, border:`2px dashed ${T.border}`, display:"inline-flex", alignItems:"center", justifyContent:"center", cursor:"pointer", overflow:"hidden" }}>
              {logoPreview ? <img src={logoPreview} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover" }} /> : <span style={{ fontSize:28 }}>🏪</span>}
            </div>
            <div style={{ color:T.textMuted, fontSize:12, marginTop:8 }}>Clica para adicionar o teu logo</div>
            <input ref={fileRef} type="file" accept="image/*" onChange={handleLogo} style={{ display:"none" }} />
          </div>
          <div style={{ marginBottom:20 }}>
            <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>Nome da loja</div>
            <input type="text" value={storeName} onChange={e=>setStoreName(e.target.value)} placeholder="Nome da tua loja"
              style={{ width:"100%", background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:12, padding:"13px 16px", color:T.text, fontSize:15, outline:"none", fontFamily:"inherit", boxSizing:"border-box" }} />
          </div>
          <button onClick={handleSave} disabled={loading||!storeName.trim()}
            style={{ width:"100%", background:storeName.trim()?T.text:T.border, border:"none", borderRadius:12, padding:"14px", color:storeName.trim()?"#fff":T.textMuted, fontSize:15, fontWeight:700, cursor:storeName.trim()?"pointer":"not-allowed", transition:"all 0.2s" }}>
            {loading?"A guardar...":"Começar →"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── MAIN APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const [session, setSession] = useState(undefined);
  const [authModo, setAuthModo] = useState(null); // null = landing, 'login'/'registo' = auth form
  const [checkingOut, setCheckingOut] = useState(false);
  const [landingCheckout, setLandingCheckout] = useState(false);
  const [profile, setProfile] = useState(undefined);
  const [profileLoading, setProfileLoading] = useState(true);
  const [entries, setEntries] = useState([]);
  const [loadingData, setLoadingData] = useState(false);
  const [tab, setTab] = useState("home");
  const [selected, setSelected] = useState(null);
  const [editId, setEditId] = useState(null);
  const [settingsName, setSettingsName] = useState("");
  const [settingsTxRate, setSettingsTxRate] = useState("");
  const [txRate, setTxRate] = useState(0.05);
  const [settingsLogo, setSettingsLogo] = useState(null);
  const [settingsLogoPreview, setSettingsLogoPreview] = useState(null);
  const [settingsSaving, setSettingsSaving] = useState(false);
  const [settingsSaved, setSettingsSaved] = useState(false);
  const settingsFileRef = useRef();
  const [statFilter, setStatFilter] = useState("mes");
  const [statValue, setStatValue] = useState(() => ({ mes: new Date().getMonth(), quarter:"Q2", ano: new Date().getFullYear() }));
  const [form, setForm] = useState({ date:TODAY, revenue:"", cog:"", ads_fb:"", ads2:"", ads3:"", refunds:"" });
  const [saving, setSaving] = useState(false);
  const [deleteConfirm, setDeleteConfirm] = useState(false);
  const isMobile = useIsMobile();

  // Auth — persistent session
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session || null));
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, session) => setSession(session || null));
    return () => subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (session) { loadProfile(); loadEntries(); }
    else if (session === null) setProfileLoading(false);
  }, [session]);

  async function loadProfile() {
    setProfileLoading(true);
    const { data } = await supabase.from("profiles").select("*").eq("id", session.user.id).single();
    setProfile(data || false);
    if (data?.tx_rate) setTxRate(parseFloat(data.tx_rate));
    setProfileLoading(false);

    // Handle payment success redirect
    const params = new URLSearchParams(window.location.search);
    if (params.get("payment") === "success") {
      await supabase.from("profiles").upsert({ id: session.user.id, plan: "pro" });
      window.history.replaceState({}, "", "/");
    }
  }

  async function loadEntries() {
    setLoadingData(true);
    const { data } = await supabase.from("Storepnl").select("*").eq("user_id", session.user.id).order("date", { ascending: false });
    setEntries(data || []);
    setLoadingData(false);
  }

  async function handleSave() {
    setSaving(true);
    const payload = { user_id:session.user.id, date:form.date, revenue:n(form.revenue), cog:n(form.cog), ads_fb:n(form.ads_fb), ads2:n(form.ads2), ads3:n(form.ads3), refunds:n(form.refunds) };
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
    setForm({ date:r.date, revenue:r.revenue, cog:r.cog, ads_fb:r.ads_fb, ads2:r.ads2||"", ads3:r.ads3||"", refunds:r.refunds });
    setEditId(r.id); setTab("add");
  }
  function resetForm() { setForm({ date:TODAY, revenue:"", cog:"", ads_fb:"", ads2:"", ads3:"", refunds:"" }); setEditId(null); }

  async function handleUpgrade() {
    setCheckingOut(true);
    try {
      const res = await fetch('/api/create-checkout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ userId: session.user.id, email: session.user.email }),
      });
      const data = await res.json();
      if (data.url) window.location.href = data.url;
      else { alert('Erro ao iniciar pagamento. Tenta novamente.'); setCheckingOut(false); }
    } catch (err) {
      alert('Erro de ligação. Tenta novamente.');
      setCheckingOut(false);
    }
  }

  async function handleLandingCheckout() {
    setLandingCheckout(true);
    try {
      const res = await fetch('/api/create-checkout', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fromLanding: true }),
      });
      const data = await res.json();
      if (data.url) window.location.href = data.url;
      else { alert('Erro ao iniciar pagamento. Tenta novamente.'); setLandingCheckout(false); }
    } catch (err) {
      alert('Erro de ligação. Tenta novamente.');
      setLandingCheckout(false);
    }
  }

  async function handleSettingsSave() {
    setSettingsSaving(true);
    let logo_url = profile?.logo_url || null;
    if (settingsLogo) {
      const ext = settingsLogo.name.split(".").pop();
      const path = `${session.user.id}/logo.${ext}`;
      const { error: upErr } = await supabase.storage.from("Logos").upload(path, settingsLogo, { upsert: true });
      if (!upErr) {
        const { data } = supabase.storage.from("Logos").getPublicUrl(path);
        logo_url = data.publicUrl;
      }
    }
    const name = settingsName.trim() || profile?.store_name;
    const newTxRate = settingsTxRate !== "" ? parseFloat(settingsTxRate.replace(",",".")) / 100 : (profile?.tx_rate || 0.05);
    await supabase.from("profiles").upsert({ id: session.user.id, store_name: name, logo_url, tx_rate: newTxRate });
    setTxRate(newTxRate);
    setProfile({ ...profile, store_name: name, logo_url });
    setSettingsSaving(false);
    setSettingsSaved(true);
    setTimeout(() => setSettingsSaved(false), 2000);
  }

  function handleSettingsLogo(e) {
    const file = e.target.files[0];
    if (!file) return;
    setSettingsLogo(file);
    setSettingsLogoPreview(URL.createObjectURL(file));
  }
  const f = k => e => setForm(p => ({ ...p, [k]:e.target.value }));

  const all = useMemo(() => entries.map(d => calc(d, txRate)), [entries]);

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
    const t = statsData.reduce((a,r) => ({ revenue:a.revenue+r.revenue, profit:a.profit+r.profit, ads:a.ads+r.ads, cog:a.cog+r.cog, refunds:a.refunds+r.refunds }), {revenue:0,profit:0,ads:0,cog:0,refunds:0});
    t.roas = t.ads>0?t.revenue/t.ads:null;
    t.margin = t.revenue>0?t.profit/t.revenue:null;
    t.dias = statsData.length;
    return t;
  }, [statsData]);

  const homeChart = useMemo(() => [...all].sort((a,b)=>new Date(a.date)-new Date(b.date)).slice(-30).map(r=>({ date:fmtDate(r.date), profit:r.profit })), [all]);
  const preview = useMemo(() => (form.revenue||form.ads_fb)?calc(form, txRate):null, [form]);

  const thisMes = useMemo(() => {
    const mes = all.filter(r => new Date(r.date+"T00:00:00").getMonth()===new Date().getMonth());
    return { rev: mes.reduce((a,r)=>a+r.revenue,0), prof: mes.reduce((a,r)=>a+r.profit,0), ads: mes.reduce((a,r)=>a+r.ads,0), dias: mes.length };
  }, [all]);

  const todayStats = useMemo(() => {
    const todayStr = new Date().toISOString().slice(0,10);
    const yest = new Date(); yest.setDate(yest.getDate()-1);
    const yesterdayStr = yest.toISOString().slice(0,10);
    const todayRow = all.find(r => r.date === todayStr);
    const yesterdayRow = all.find(r => r.date === yesterdayStr);
    const lastEntry = all.length > 0 ? all[0] : null;
    return { today: todayRow || null, yesterday: yesterdayRow || null, lastEntry };
  }, [all]);

  function delta(today, yesterday, key) {
    if (!today || !yesterday) return null;
    const diff = today[key] - yesterday[key];
    const pctDiff = yesterday[key] !== 0 ? (diff / Math.abs(yesterday[key])) * 100 : null;
    return { diff, pctDiff, up: diff >= 0 };
  }

  if (session === undefined || profileLoading) return (
    <div style={{ minHeight:"100vh", background:T.bg, display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ width:32, height:32, border:`2px solid ${T.border}`, borderTop:`2px solid ${T.text}`, borderRadius:"50%", animation:"spin 0.7s linear infinite" }} />
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
  if (!session) {
    // After Stripe payment, redirect to register
    const params = new URLSearchParams(window.location.search);
    if (params.get("payment") === "success") {
      if (authModo === null) setAuthModo("registo");
    }
    if (authModo === null) return <LandingPage onStart={(m) => setAuthModo(m)} onCheckout={handleLandingCheckout} checking={landingCheckout} />;
    return <AuthScreen initialModo={authModo} paymentSuccess={new URLSearchParams(window.location.search).get("payment") === "success"} />;
  }
  if (profile === false) return <OnboardingScreen userId={session.user.id} onComplete={p => setProfile(p)} />;


  const storeName = profile?.store_name || "A minha loja";
  const logoUrl = profile?.logo_url;
  const mono = "'DM Mono','Courier New',monospace";
  const card = { background:T.surface, border:`1px solid ${T.border}`, borderRadius:16 };

  // ── SIDEBAR ITEMS ──
  const navItems = [
    { id:"home", label:"Dashboard", icon:"▦" },
    { id:"analytics", label:"Analytics", icon:"◈" },
    { id:"add", label:"Adicionar Dia", icon:"+" },
    { id:"definicoes", label:"Definições", icon:"⚙" },
  ];

  // ── DESKTOP LAYOUT ─────────────────────────────────────────────────────────
  if (!isMobile) {
    return (
      <div style={{ display:"flex", minHeight:"100vh", background:T.bg, fontFamily:"'DM Sans',sans-serif", color:T.text }}>
        <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap'); *{box-sizing:border-box;margin:0;padding:0;} ::-webkit-scrollbar{width:4px} ::-webkit-scrollbar-track{background:transparent} ::-webkit-scrollbar-thumb{background:${T.border};border-radius:4px}`}</style>

        {/* Sidebar */}
        <div style={{ width:240, background:T.sidebar, borderRight:`1px solid ${T.border}`, display:"flex", flexDirection:"column", position:"fixed", height:"100vh", zIndex:10 }}>
          {/* Logo */}
          <div style={{ padding:"28px 20px 24px", borderBottom:`1px solid ${T.border}` }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              {logoUrl
                ? <img src={logoUrl} alt="logo" style={{ width:36, height:36, borderRadius:10, objectFit:"cover", border:`1px solid ${T.border}` }} />
                : <div style={{ width:36, height:36, background:T.text, borderRadius:10, display:"flex", alignItems:"center", justifyContent:"center" }}><span style={{ color:"#fff", fontSize:14, fontWeight:800 }}>{storeName[0]}</span></div>
              }
              <div>
                <div style={{ fontSize:14, fontWeight:700, color:T.text, letterSpacing:"-0.01em" }}>{storeName}</div>
                <div style={{ fontSize:11, color:T.textMuted }}>StorePNL</div>
              </div>
            </div>
          </div>

          {/* Nav */}
          <nav style={{ padding:"16px 12px", flex:1 }}>
            {navItems.map(item => (
              <button key={item.id} onClick={() => { setTab(item.id); if(item.id!=="add") resetForm(); }}
                style={{ width:"100%", display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:10, border:"none", background:tab===item.id?T.accentLight:"transparent", color:tab===item.id?T.text:T.textMuted, fontSize:14, fontWeight:tab===item.id?600:400, cursor:"pointer", textAlign:"left", marginBottom:2, transition:"all 0.15s" }}>
                <span style={{ fontSize:16, width:20, textAlign:"center" }}>{item.icon}</span>
                {item.label}
              </button>
            ))}
          </nav>

          {/* User */}
          <div style={{ padding:"16px 20px", borderTop:`1px solid ${T.border}` }}>
            <div style={{ fontSize:12, color:T.textMuted, marginBottom:8, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{session.user.email}</div>
            <button onClick={() => supabase.auth.signOut()} style={{ width:"100%", background:"transparent", border:`1px solid ${T.border}`, borderRadius:8, padding:"8px", color:T.textMuted, fontSize:12, cursor:"pointer" }}>Sair da conta</button>
          </div>
        </div>

        {/* Main content */}
        <div style={{ marginLeft:240, flex:1, padding:"32px 40px", maxWidth:"calc(100vw - 240px)", overflowX:"hidden" }}>

          {/* HOME */}
          {tab==="home" && (
            <div>
              {/* Header */}
              <div style={{ marginBottom:24, display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                <div>
                  <div style={{ fontSize:26, fontWeight:800, letterSpacing:"-0.03em", marginBottom:4 }}>Dashboard</div>
                  <div style={{ color:T.textMuted, fontSize:14 }}>{new Date().toLocaleDateString("pt-PT", { weekday:"long", day:"numeric", month:"long" })}</div>
                </div>
                <button onClick={()=>{resetForm();setTab("add");}} style={{ background:T.text, border:"none", borderRadius:10, padding:"10px 20px", color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>+ Adicionar Dia</button>
              </div>

              {all.length===0 ? (
                <div style={{ ...card, padding:"80px", textAlign:"center" }}>
                  <div style={{ fontSize:48, marginBottom:12 }}>📊</div>
                  <div style={{ fontSize:18, fontWeight:700, marginBottom:6 }}>Sem dados ainda</div>
                  <div style={{ color:T.textMuted, fontSize:14, marginBottom:24 }}>Adiciona o teu primeiro dia para começar</div>
                  <button onClick={()=>{resetForm();setTab("add");}} style={{ background:T.text, border:"none", borderRadius:12, padding:"12px 28px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>+ Adicionar dia</button>
                </div>
              ) : (
                <>
                  {/* TODAY HERO */}
                  {(() => {
                    const t = todayStats.today;
                    const y = todayStats.yesterday;
                    const last = todayStats.lastEntry;
                    const display = t || last;
                    const isToday = !!t;
                    if (!display) return null;

                    const profDelta = delta(t, y, "profit");
                    const revDelta = delta(t, y, "revenue");
                    const adsDelta = delta(t, y, "ads");

                    return (
                      <div style={{ marginBottom:20 }}>
                        {/* Hero profit card */}
                        <div style={{ ...card, padding:"28px 32px", marginBottom:14, background: display.profit>=0 ? "#f0fdf4" : "#fef2f2", border:`1px solid ${display.profit>=0?T.greenBorder:T.redBorder}` }}>
                          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                            <div>
                              <div style={{ fontSize:12, fontWeight:600, color:T.textMuted, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:8 }}>
                                {isToday ? "Lucro de Hoje" : `Último dia registado — ${fmtDate(display.date)}`}
                              </div>
                              <div style={{ fontSize:48, fontWeight:800, color:display.profit>=0?T.green:T.red, fontFamily:mono, letterSpacing:"-0.04em", lineHeight:1 }}>
                                {eur(display.profit)}
                              </div>
                              <div style={{ color:T.textMuted, fontSize:13, marginTop:8 }}>{pct(display.margin)} margem · ROAS {display.roas?display.roas.toFixed(2)+"x":"—"}</div>
                            </div>
                            {profDelta && (
                              <div style={{ textAlign:"right" }}>
                                <div style={{ fontSize:11, color:T.textMuted, marginBottom:4 }}>vs ontem</div>
                                <div style={{ fontSize:20, fontWeight:800, color:profDelta.up?T.green:T.red, fontFamily:mono }}>
                                  {profDelta.up?"↑":"↓"} {eur(Math.abs(profDelta.diff))}
                                </div>
                                {profDelta.pctDiff !== null && (
                                  <div style={{ fontSize:12, color:profDelta.up?T.green:T.red }}>{Math.abs(profDelta.pctDiff).toFixed(1)}%</div>
                                )}
                              </div>
                            )}
                          </div>
                        </div>

                        {/* 3 metric cards */}
                        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12, marginBottom:14 }}>
                          {[
                            { label:"Faturação", value:eur(display.revenue), color:T.blue, d:revDelta },
                            { label:"Adspend", value:eur(display.ads), color:T.amber, d:adsDelta },
                            { label:"ROAS", value:display.roas?display.roas.toFixed(2)+"x":"—", color:display.roas>=2?T.green:display.roas>=1?T.amber:T.red, d:null },
                          ].map(({label,value,color,d}) => (
                            <div key={label} style={{ ...card, padding:"18px 20px" }}>
                              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:8 }}>{label}</div>
                              <div style={{ color, fontSize:22, fontWeight:800, fontFamily:mono, marginBottom:4 }}>{value}</div>
                              {d && (
                                <div style={{ fontSize:12, color:d.up?T.green:T.red, fontWeight:500 }}>
                                  {d.up?"↑":"↓"} {eur(Math.abs(d.diff))} vs ontem
                                </div>
                              )}
                            </div>
                          ))}
                        </div>

                        {/* COG + Tx + Refunds row */}
                        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:12 }}>
                          {[
                            ["Custo de Produto", eur(display.cog), T.purple],
                            ["Taxa (5%)", eur(display.tx), T.textMuted],
                            ["Devoluções", eur(display.refunds), T.red],
                          ].map(([l,v,c]) => (
                            <div key={l} style={{ ...card, padding:"14px 18px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                              <div style={{ color:T.textMuted, fontSize:12 }}>{l}</div>
                              <div style={{ color:c, fontSize:15, fontWeight:700, fontFamily:mono }}>{v}</div>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })()}

                  {/* Chart + recent */}
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 360px", gap:16, marginBottom:20 }}>
                    <div style={{ ...card, padding:"24px" }}>
                      <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:16 }}>Lucro / Prejuízo — 30 dias</div>
                      <ResponsiveContainer width="100%" height={180}>
                        <BarChart data={homeChart} barSize={homeChart.length>20?8:14} margin={{top:4,right:4,left:0,bottom:0}}>
                          <XAxis dataKey="date" tick={{fill:T.textLight,fontSize:10}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                          <YAxis tick={{fill:T.textLight,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={52} />
                          <Tooltip content={<ChartTip />} cursor={{fill:"rgba(0,0,0,0.02)"}} />
                          <ReferenceLine y={0} stroke={T.border} />
                          <Bar dataKey="profit" radius={[4,4,0,0]}>
                            {homeChart.map((e,i)=><Cell key={i} fill={e.profit>=0?"#86EFAC":"#FCA5A5"} />)}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                    <div style={{ ...card, padding:"24px" }}>
                      <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:14 }}>Dias Recentes</div>
                      <div style={{ overflowY:"auto", maxHeight:220 }}>
                        {all.slice(0,8).map(r => (
                          <div key={r.id} onClick={()=>{setSelected(r);setTab("detalhe");}}
                            style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"10px 0", borderBottom:`1px solid ${T.border}`, cursor:"pointer" }}>
                            <div>
                              <div style={{ fontSize:13, fontWeight:600, color:T.text }}>{fmtDate(r.date)}</div>
                              <div style={{ fontSize:11, color:T.textMuted }}>{eur(r.revenue)} faturação</div>
                            </div>
                            <div style={{ color:r.profit>=0?T.green:T.red, fontSize:14, fontWeight:700, fontFamily:mono }}>{eur(r.profit)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  {/* Full table */}
                  <div style={{ ...card, padding:"24px" }}>
                    <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:16 }}>Todos os Dias</div>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
                      <thead>
                        <tr>{["Data","Faturação","Adspend","COG","Devoluções","Taxa","Lucro","ROAS","Margem",""].map(h => (
                          <th key={h} style={{ textAlign:"left", color:T.textMuted, fontWeight:600, fontSize:11, letterSpacing:"0.05em", textTransform:"uppercase", paddingBottom:10, paddingRight:16, borderBottom:`1px solid ${T.border}` }}>{h}</th>
                        ))}</tr>
                      </thead>
                      <tbody>
                        {all.map(r => (
                          <tr key={r.id} style={{ borderBottom:`1px solid ${T.border}`, cursor:"pointer" }} onClick={()=>{setSelected(r);setTab("detalhe");}}>
                            <td style={{ padding:"12px 16px 12px 0", fontWeight:500 }}>{fmtDate(r.date)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.blue }}>{eur(r.revenue)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.amber }}>{eur(r.ads)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.purple }}>{eur(r.cog)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.red }}>{eur(r.refunds)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.textMuted }}>{eur(r.tx)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, fontWeight:700, color:r.profit>=0?T.green:T.red }}>{eur(r.profit)}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:T.textMuted }}>{r.roas?r.roas.toFixed(2)+"x":"—"}</td>
                            <td style={{ padding:"12px 16px 12px 0", fontFamily:mono, color:r.margin>=0?T.green:T.red }}>{pct(r.margin)}</td>
                            <td style={{ padding:"12px 0 12px 0" }}>
                              <button onClick={e=>{e.stopPropagation();openEdit(r);}} style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:6, padding:"4px 10px", fontSize:11, cursor:"pointer", color:T.textMuted }}>Editar</button>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          )}

          {/* ANALYTICS */}
          {tab==="analytics" && (
            <div>
              <div style={{ marginBottom:28 }}>
                <div style={{ fontSize:26, fontWeight:800, letterSpacing:"-0.03em", marginBottom:4 }}>Analytics</div>
                <div style={{ color:T.textMuted, fontSize:14 }}>Analisa o desempenho da tua loja por período</div>
              </div>

              {/* Filters */}
              <div style={{ display:"flex", gap:8, marginBottom:16, flexWrap:"wrap" }}>
                {[["mes","Mês"],["quarter","Trimestre"],["ano","Ano"],["tudo","Tudo"]].map(([v,l]) => (
                  <button key={v} onClick={()=>setStatFilter(v)} style={{ background:statFilter===v?T.text:T.surface, border:`1px solid ${statFilter===v?T.text:T.border}`, borderRadius:20, padding:"6px 16px", color:statFilter===v?"#fff":T.textMuted, fontSize:13, fontWeight:600, cursor:"pointer" }}>{l}</button>
                ))}
                {statFilter==="mes" && MESES.map((m,i) => (
                  <button key={m} onClick={()=>setStatValue(p=>({...p,mes:i}))} style={{ background:statValue.mes===i?T.accentLight:T.surface, border:`1px solid ${statValue.mes===i?T.borderStrong:T.border}`, borderRadius:20, padding:"6px 14px", color:statValue.mes===i?T.text:T.textMuted, fontSize:13, fontWeight:600, cursor:"pointer" }}>{m}</button>
                ))}
                {statFilter==="quarter" && ["Q1","Q2","Q3","Q4"].map(q => (
                  <button key={q} onClick={()=>setStatValue(p=>({...p,quarter:q}))} style={{ background:statValue.quarter===q?T.accentLight:T.surface, border:`1px solid ${statValue.quarter===q?T.borderStrong:T.border}`, borderRadius:20, padding:"6px 14px", color:statValue.quarter===q?T.text:T.textMuted, fontSize:13, fontWeight:600, cursor:"pointer" }}>{q}</button>
                ))}
                {statFilter==="ano" && [2024,2025,2026].map(y => (
                  <button key={y} onClick={()=>setStatValue(p=>({...p,ano:y}))} style={{ background:statValue.ano===y?T.accentLight:T.surface, border:`1px solid ${statValue.ano===y?T.borderStrong:T.border}`, borderRadius:20, padding:"6px 14px", color:statValue.ano===y?T.text:T.textMuted, fontSize:13, fontWeight:600, cursor:"pointer" }}>{y}</button>
                ))}
              </div>

              {statsData.length===0
                ? <div style={{ ...card, padding:"60px", textAlign:"center", color:T.textMuted }}>Sem dados para este período</div>
                : (
                  <>
                    <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:14, marginBottom:20 }}>
                      {[["Faturação",eur(totals.revenue),T.text],["Lucro Líquido",eur(totals.profit),totals.profit>=0?T.green:T.red],["Adspend",eur(totals.ads),T.amber],["ROAS",totals.roas?totals.roas.toFixed(2)+"x":"—",T.green],["Margem",pct(totals.margin),totals.margin>=0?T.green:T.red],["Dias",totals.dias,T.purple]].map(([l,v,c])=>(
                        <div key={l} style={{ ...card, padding:"18px 20px" }}>
                          <div style={{ color:T.textMuted, fontSize:12, marginBottom:6 }}>{l}</div>
                          <div style={{ color:c, fontSize:22, fontWeight:800, fontFamily:mono }}>{v}</div>
                        </div>
                      ))}
                    </div>

                    <div style={{ display:"grid", gridTemplateColumns:"2fr 1fr", gap:16 }}>
                      <div style={{ ...card, padding:"24px" }}>
                        <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:16 }}>Lucro / Prejuízo</div>
                        <ResponsiveContainer width="100%" height={200}>
                          <BarChart data={statsData.map(r=>({date:fmtDate(r.date),profit:r.profit}))} barSize={statsData.length>20?6:12} margin={{top:4,right:4,left:0,bottom:0}}>
                            <XAxis dataKey="date" tick={{fill:T.textLight,fontSize:10}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                            <YAxis tick={{fill:T.textLight,fontSize:10}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={52} />
                            <Tooltip content={<ChartTip />} cursor={{fill:"rgba(0,0,0,0.02)"}} />
                            <ReferenceLine y={0} stroke={T.border} />
                            <Bar dataKey="profit" radius={[4,4,0,0]}>
                              {statsData.map((e,i)=><Cell key={i} fill={e.profit>=0?"#86EFAC":"#FCA5A5"} />)}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>

                      <div style={{ ...card, padding:"24px" }}>
                        <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:16 }}>Detalhe P&L</div>
                        {[["Faturação",totals.revenue,T.text],["− COG",totals.cog,T.purple],["− Adspend",totals.ads,T.amber],["− Devoluções",totals.refunds,T.red],["− Taxas",totals.revenue*0.05,T.textMuted],["= Lucro",totals.profit,totals.profit>=0?T.green:T.red]].map(([l,v,c],i)=>(
                          <div key={l} style={{ display:"flex", justifyContent:"space-between", padding:"9px 0", borderTop:i>0?`1px solid ${T.border}`:"none" }}>
                            <span style={{ color:i===5?T.text:T.textMuted, fontSize:13, fontWeight:i===5?700:400 }}>{l}</span>
                            <span style={{ color:c, fontSize:13, fontWeight:i===5?800:600, fontFamily:mono }}>{eur(v)}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </>
                )
              }
            </div>
          )}

          {/* DETALHE */}
          {tab==="detalhe" && selected && (() => {
            const r = calc(selected, txRate);
            return (
              <div>
                <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:28 }}>
                  <button onClick={()=>{setTab("home");setDeleteConfirm(false);}} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", fontSize:16 }}>←</button>
                  <div>
                    <div style={{ color:T.textMuted, fontSize:12 }}>Relatório do Dia</div>
                    <div style={{ fontSize:22, fontWeight:800, letterSpacing:"-0.02em" }}>{fmtDateFull(r.date)}</div>
                  </div>
                </div>

                <div style={{ display:"grid", gridTemplateColumns:"280px 1fr", gap:20 }}>
                  <div>
                    <div style={{ ...card, padding:"28px", textAlign:"center", marginBottom:12 }}>
                      <div style={{ color:T.textMuted, fontSize:12, marginBottom:8 }}>Lucro Líquido</div>
                      <div style={{ color:r.profit>=0?T.green:T.red, fontSize:40, fontWeight:800, fontFamily:mono, letterSpacing:"-0.04em" }}>{eur(r.profit)}</div>
                      <div style={{ display:"inline-block", background:r.margin>=0?T.greenBg:T.redBg, color:r.margin>=0?T.green:T.red, fontSize:12, fontWeight:600, borderRadius:20, padding:"4px 12px", marginTop:8 }}>{pct(r.margin)} margem</div>
                    </div>
                    <div style={{ display:"flex", gap:8 }}>
                      <button onClick={()=>openEdit(r)} style={{ flex:1, background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, padding:"11px", color:T.text, fontSize:13, fontWeight:600, cursor:"pointer" }}>Editar</button>
                      {!deleteConfirm
                        ? <button onClick={()=>setDeleteConfirm(true)} style={{ flex:1, background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:10, padding:"11px", color:T.red, fontSize:13, fontWeight:600, cursor:"pointer" }}>Apagar</button>
                        : <button onClick={()=>handleDelete(r.id)} style={{ flex:1, background:T.red, border:"none", borderRadius:10, padding:"11px", color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>Confirmar?</button>
                      }
                    </div>
                  </div>

                  <div style={{ ...card, padding:"24px" }}>
                    <div style={{ color:T.textMuted, fontSize:12, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:16 }}>Detalhe</div>
                    {[["Faturação",eur(r.revenue),T.text],["Meta Adspend",eur(r.ads),T.amber],["Custo de Produto",eur(r.cog),T.purple],["ROAS",r.roas?r.roas.toFixed(2)+"x":"—",T.green],["Devoluções",eur(r.refunds),T.red],["Taxa (5%)",eur(r.tx),T.textMuted]].map(([l,v,c])=>(
                      <div key={l} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"13px 0", borderBottom:`1px solid ${T.border}` }}>
                        <span style={{ color:T.textMuted, fontSize:14 }}>{l}</span>
                        <span style={{ color:c, fontSize:16, fontWeight:700, fontFamily:mono }}>{v}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            );
          })()}

          {/* ADD/EDIT */}
          {tab==="add" && (
            <div style={{ maxWidth:560 }}>
              <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:28 }}>
                <button onClick={()=>{setTab("home");resetForm();}} style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", fontSize:16 }}>←</button>
                <div style={{ fontSize:22, fontWeight:800, letterSpacing:"-0.02em" }}>{editId?"Editar Dia":"Adicionar Dia"}</div>
              </div>

              <div style={{ ...card, padding:"28px" }}>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
                  {[["date","Data","date",true],["revenue","Faturação (€)","decimal",false],["cog","Custo de Produto (€)","decimal",false],["ads_fb","Meta Adspend (€)","decimal",false],["ads2","Adspend 2 (€)","decimal",false],["ads3","Adspend 3 (€)","decimal",false],["refunds","Devoluções (€)","decimal",false]].map(([key,label,mode,full])=>(
                    <div key={key} style={{ gridColumn:full?"1/-1":"auto" }}>
                      <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>{label}</div>
                      <input type={mode==="date"?"date":"text"} inputMode={mode} value={form[key]} onChange={f(key)} placeholder="0,00"
                        style={{ width:"100%", background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:15, outline:"none", fontFamily:mono }} />
                    </div>
                  ))}
                </div>

                {preview && (
                  <div style={{ background:T.bg, borderRadius:12, padding:"16px 20px", marginTop:20 }}>
                    <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:12 }}>Pré-visualização</div>
                    <div style={{ display:"flex", gap:32 }}>
                      {[["Lucro",eur(preview.profit),preview.profit>=0?T.green:T.red],["ROAS",preview.roas?preview.roas.toFixed(2)+"x":"—",T.green],["Margem",pct(preview.margin),preview.margin>=0?T.green:T.red]].map(([l,v,c])=>(
                        <div key={l}><div style={{ color:T.textMuted, fontSize:11, marginBottom:4 }}>{l}</div><div style={{ color:c, fontSize:18, fontWeight:800, fontFamily:mono }}>{v}</div></div>
                      ))}
                    </div>
                  </div>
                )}

                <button onClick={handleSave} disabled={saving}
                  style={{ width:"100%", background:T.text, border:"none", borderRadius:12, padding:"14px", color:"#fff", fontSize:15, fontWeight:700, cursor:"pointer", marginTop:20, letterSpacing:"-0.01em" }}>
                  {saving?"A guardar...":editId?"Atualizar Dia":"Guardar Dia"}
                </button>
              </div>
            </div>
          )}

          {/* DEFINICOES */}
          {tab==="definicoes" && (
            <div style={{ maxWidth:520 }}>
              <div style={{ marginBottom:28 }}>
                <div style={{ fontSize:26, fontWeight:800, letterSpacing:"-0.03em", marginBottom:4 }}>Definições</div>
                <div style={{ color:T.textMuted, fontSize:14 }}>Personaliza a tua loja</div>
              </div>
              <div style={{ ...card, padding:"28px" }}>
                {/* Logo */}
                <div style={{ marginBottom:24 }}>
                  <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:12 }}>Logo da Loja</div>
                  <div style={{ display:"flex", alignItems:"center", gap:16 }}>
                    <div onClick={() => settingsFileRef.current.click()} style={{ width:72, height:72, borderRadius:16, background:settingsLogoPreview||logoUrl?"transparent":T.bg, border:`2px dashed ${T.border}`, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", overflow:"hidden", flexShrink:0 }}>
                      {settingsLogoPreview
                        ? <img src={settingsLogoPreview} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                        : logoUrl
                          ? <img src={logoUrl} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                          : <span style={{ fontSize:24 }}>🏪</span>
                      }
                    </div>
                    <div>
                      <button onClick={() => settingsFileRef.current.click()} style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:8, padding:"8px 16px", fontSize:13, fontWeight:600, cursor:"pointer", color:T.text, marginBottom:6, display:"block" }}>Alterar logo</button>
                      <div style={{ color:T.textMuted, fontSize:12 }}>PNG, JPG ou GIF. Max 2MB.</div>
                    </div>
                    <input ref={settingsFileRef} type="file" accept="image/*" onChange={handleSettingsLogo} style={{ display:"none" }} />
                  </div>
                </div>

                {/* Name */}
                <div style={{ marginBottom:20 }}>
                  <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>Nome da Loja</div>
                  <input type="text" value={settingsName} onChange={e=>setSettingsName(e.target.value)}
                    placeholder={profile?.store_name || "Nome da tua loja"}
                    style={{ width:"100%", background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:15, outline:"none", fontFamily:"inherit" }} />
                </div>

                {/* Tx Rate */}
                <div style={{ marginBottom:24 }}>
                  <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>Taxa de Transacção (%)</div>
                  <div style={{ display:"flex", alignItems:"center", gap:10 }}>
                    <input type="text" inputMode="decimal" value={settingsTxRate} onChange={e=>setSettingsTxRate(e.target.value)}
                      placeholder={`${(txRate*100).toFixed(1)}`}
                      style={{ width:120, background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:15, outline:"none", fontFamily:"'DM Mono',monospace" }} />
                    <span style={{ color:T.textMuted, fontSize:13 }}>% · actual: {(txRate*100).toFixed(1)}%</span>
                  </div>
                  <div style={{ color:T.textLight, fontSize:12, marginTop:6 }}>Taxa cobrada pelo Shopify por cada transacção. Por defeito 5%.</div>
                </div>

                <button onClick={handleSettingsSave} disabled={settingsSaving}
                  style={{ background:settingsSaved?"#16A34A":T.text, border:"none", borderRadius:12, padding:"13px 28px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer", transition:"background 0.2s" }}>
                  {settingsSaved?"✓ Guardado!":settingsSaving?"A guardar...":"Guardar alterações"}
                </button>
              </div>

              <div style={{ ...card, padding:"24px", marginTop:16 }}>
                <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:12 }}>Conta</div>
                <div style={{ color:T.textMuted, fontSize:13, marginBottom:14 }}>{session.user.email}</div>
                <button onClick={() => supabase.auth.signOut()} style={{ background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:10, padding:"10px 20px", color:T.red, fontSize:13, fontWeight:600, cursor:"pointer" }}>Sair da conta</button>
              </div>

              {/* Excel Import */}
              <div style={{ ...card, padding:"24px", marginTop:16 }}>
                <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:4 }}>Importar Dados</div>
                <div style={{ color:T.textMuted, fontSize:13, marginBottom:16 }}>Importa o teu ficheiro P&L Sheet (.xlsx) directamente para a dashboard.</div>
                <ExcelImport userId={session.user.id} onImportDone={loadEntries} T={T} />
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ── MOBILE LAYOUT ──────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight:"100vh", background:T.bg, fontFamily:"'DM Sans',sans-serif", color:T.text, maxWidth:430, margin:"0 auto" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500;600&display=swap'); *{box-sizing:border-box;-webkit-tap-highlight-color:transparent;margin:0;padding:0;} input{font-family:'DM Mono',monospace!important} ::-webkit-scrollbar{display:none} body{background:${T.bg}}`}</style>

      {tab==="home" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"52px 20px 16px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
            <div style={{ display:"flex", alignItems:"center", gap:10 }}>
              {logoUrl
                ? <img src={logoUrl} alt="logo" style={{ width:36, height:36, borderRadius:10, objectFit:"cover", border:`1px solid ${T.border}` }} />
                : <div style={{ width:36, height:36, background:T.text, borderRadius:10, display:"flex", alignItems:"center", justifyContent:"center" }}><span style={{ color:"#fff", fontSize:14, fontWeight:800 }}>{storeName[0]}</span></div>
              }
              <div>
                <div style={{ fontSize:15, fontWeight:700, color:T.text }}>{storeName}</div>
                <div style={{ fontSize:11, color:T.textMuted }}>Dashboard P&L</div>
              </div>
            </div>
            <button onClick={()=>{resetForm();setTab("add");}} style={{ background:T.text, border:"none", borderRadius:10, padding:"9px 16px", color:"#fff", fontSize:13, fontWeight:700, cursor:"pointer" }}>+ Dia</button>
          </div>

          {loadingData ? <div style={{ textAlign:"center", padding:"80px 20px", color:T.textMuted, fontSize:14 }}>A carregar...</div>
          : all.length===0 ? (
            <div style={{ textAlign:"center", padding:"80px 20px" }}>
              <div style={{ fontSize:48, marginBottom:12 }}>📊</div>
              <div style={{ fontSize:18, fontWeight:700, marginBottom:6 }}>Sem dados ainda</div>
              <div style={{ fontSize:14, color:T.textMuted, marginBottom:24 }}>Adiciona o teu primeiro dia</div>
              <button onClick={()=>{resetForm();setTab("add");}} style={{ background:T.text, border:"none", borderRadius:12, padding:"12px 24px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>+ Adicionar dia</button>
            </div>
          ) : (
            <div style={{ padding:"12px 16px" }}>
              {(() => {
                const t = todayStats.today;
                const y = todayStats.yesterday;
                const display = t || todayStats.lastEntry;
                const isToday = !!t;
                if (!display) return null;
                const profDelta = delta(t, y, "profit");
                const revDelta = delta(t, y, "revenue");
                const adsDelta = delta(t, y, "ads");
                return (
                  <>
                    {/* Hero */}
                    <div style={{ background:display.profit>=0?"#f0fdf4":"#fef2f2", border:`1px solid ${display.profit>=0?T.greenBorder:T.redBorder}`, borderRadius:16, padding:"20px", marginBottom:10 }}>
                      <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:6 }}>
                        {isToday ? "Hoje" : `Último dia — ${fmtDate(display.date)}`}
                      </div>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
                        <div style={{ color:display.profit>=0?T.green:T.red, fontSize:38, fontWeight:800, fontFamily:mono, letterSpacing:"-0.03em" }}>{eur(display.profit)}</div>
                        {profDelta && (
                          <div style={{ textAlign:"right" }}>
                            <div style={{ fontSize:11, color:T.textMuted }}>vs ontem</div>
                            <div style={{ fontSize:16, fontWeight:800, color:profDelta.up?T.green:T.red, fontFamily:mono }}>{profDelta.up?"↑":"↓"} {eur(Math.abs(profDelta.diff))}</div>
                          </div>
                        )}
                      </div>
                      <div style={{ color:T.textMuted, fontSize:12, marginTop:6 }}>{pct(display.margin)} margem · ROAS {display.roas?display.roas.toFixed(2)+"x":"—"}</div>
                    </div>

                    {/* Rev + Ads */}
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:8 }}>
                      {[
                        ["Faturação", eur(display.revenue), T.blue, revDelta],
                        ["Adspend", eur(display.ads), T.amber, adsDelta],
                      ].map(([l,v,c,d]) => (
                        <div key={l} style={{ ...card, padding:"14px 16px" }}>
                          <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, textTransform:"uppercase", letterSpacing:"0.05em", marginBottom:6 }}>{l}</div>
                          <div style={{ color:c, fontSize:20, fontWeight:800, fontFamily:mono }}>{v}</div>
                          {d && <div style={{ fontSize:11, color:d.up?T.green:T.red, marginTop:4 }}>{d.up?"↑":"↓"} {eur(Math.abs(d.diff))}</div>}
                        </div>
                      ))}
                    </div>

                    {/* COG + Refunds */}
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8, marginBottom:12 }}>
                      {[["COG", eur(display.cog), T.purple],["Devoluções", eur(display.refunds), T.red]].map(([l,v,c]) => (
                        <div key={l} style={{ ...card, padding:"12px 16px", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                          <span style={{ color:T.textMuted, fontSize:12 }}>{l}</span>
                          <span style={{ color:c, fontSize:14, fontWeight:700, fontFamily:mono }}>{v}</span>
                        </div>
                      ))}
                    </div>
                  </>
                );
              })()}

              {/* Chart */}
              <div style={{ ...card, padding:"16px 14px 10px", marginBottom:12 }}>
                <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:10 }}>Lucro — 30 dias</div>
                <ResponsiveContainer width="100%" height={120}>
                  <BarChart data={homeChart} barSize={homeChart.length>20?5:10} margin={{top:4,right:4,left:0,bottom:0}}>
                    <XAxis dataKey="date" tick={{fill:T.textLight,fontSize:9}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                    <YAxis tick={{fill:T.textLight,fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={40} />
                    <Tooltip content={<ChartTip />} cursor={{fill:"rgba(0,0,0,0.02)"}} />
                    <ReferenceLine y={0} stroke={T.border} />
                    <Bar dataKey="profit" radius={[3,3,0,0]}>{homeChart.map((e,i)=><Cell key={i} fill={e.profit>=0?"#86EFAC":"#FCA5A5"} />)}</Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>

              {/* Recent days */}
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:8, paddingLeft:2 }}>Dias Recentes</div>
              {all.slice(0,60).map(r=>(
                <div key={r.id} onClick={()=>{setSelected(r);setTab("detalhe");}} style={{ ...card, padding:"13px 16px", marginBottom:6, display:"flex", justifyContent:"space-between", alignItems:"center", cursor:"pointer" }}>
                  <div><div style={{ fontSize:14, fontWeight:600, marginBottom:2 }}>{fmtDate(r.date)}</div><div style={{ color:T.textMuted, fontSize:12 }}>{eur(r.revenue)} · {eur(r.ads)} ads</div></div>
                  <div style={{ textAlign:"right" }}><div style={{ color:r.profit>=0?T.green:T.red, fontSize:15, fontWeight:700, fontFamily:mono }}>{eur(r.profit)}</div><div style={{ color:T.textMuted, fontSize:11, marginTop:1 }}>{pct(r.margin)}</div></div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {tab==="detalhe" && selected && (() => {
        const r = calc(selected, txRate);
        return (
          <div style={{ paddingBottom:90 }}>
            <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"52px 20px 16px", display:"flex", alignItems:"center", gap:14 }}>
              <button onClick={()=>{setTab("home");setDeleteConfirm(false);}} style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:10, width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", fontSize:16 }}>←</button>
              <div><div style={{ color:T.textMuted, fontSize:12 }}>Relatório</div><div style={{ fontSize:19, fontWeight:800 }}>{fmtDate(r.date)}</div></div>
            </div>
            <div style={{ padding:"16px" }}>
              <div style={{ ...card, padding:"24px", textAlign:"center", marginBottom:10 }}>
                <div style={{ color:T.textMuted, fontSize:12, marginBottom:6 }}>Lucro Líquido</div>
                <div style={{ color:r.profit>=0?T.green:T.red, fontSize:40, fontWeight:800, fontFamily:mono }}>{eur(r.profit)}</div>
                <div style={{ display:"inline-block", background:r.margin>=0?T.greenBg:T.redBg, color:r.margin>=0?T.green:T.red, fontSize:12, fontWeight:600, borderRadius:20, padding:"4px 12px", marginTop:8 }}>{pct(r.margin)}</div>
              </div>
              {[["Faturação",eur(r.revenue),T.text],["Meta Adspend",eur(r.ads),T.amber],["Custo de Produto",eur(r.cog),T.purple],["ROAS",r.roas?r.roas.toFixed(2)+"x":"—",T.green],["Devoluções",eur(r.refunds),T.red],["Taxa (5%)",eur(r.tx),T.textMuted]].map(([l,v,c])=>(
                <div key={l} style={{ ...card, padding:"13px 18px", marginBottom:6, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                  <span style={{ color:T.textMuted, fontSize:13 }}>{l}</span>
                  <span style={{ color:c, fontSize:15, fontWeight:700, fontFamily:mono }}>{v}</span>
                </div>
              ))}
              <div style={{ display:"flex", gap:10, marginTop:14 }}>
                <button onClick={()=>openEdit(r)} style={{ flex:1, background:T.bg, border:`1px solid ${T.border}`, borderRadius:12, padding:"13px", color:T.text, fontSize:14, fontWeight:600, cursor:"pointer" }}>Editar</button>
                {!deleteConfirm
                  ? <button onClick={()=>setDeleteConfirm(true)} style={{ flex:1, background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:12, padding:"13px", color:T.red, fontSize:14, fontWeight:600, cursor:"pointer" }}>Apagar</button>
                  : <button onClick={()=>handleDelete(r.id)} style={{ flex:1, background:T.red, border:"none", borderRadius:12, padding:"13px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer" }}>Confirmar?</button>
                }
              </div>
            </div>
          </div>
        );
      })()}

      {tab==="analytics" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"52px 20px 16px" }}>
            <div style={{ fontSize:22, fontWeight:800, letterSpacing:"-0.02em" }}>Analytics</div>
          </div>
          <div style={{ padding:"16px" }}>
            <div style={{ display:"flex", gap:6, marginBottom:10, overflowX:"auto", paddingBottom:2 }}>
              {[["mes","Mês"],["quarter","Trimestre"],["ano","Ano"],["tudo","Tudo"]].map(([v,l])=>(
                <button key={v} onClick={()=>setStatFilter(v)} style={{ background:statFilter===v?T.text:T.surface, border:`1px solid ${statFilter===v?T.text:T.border}`, borderRadius:20, padding:"6px 14px", color:statFilter===v?"#fff":T.textMuted, fontSize:12, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap" }}>{l}</button>
              ))}
            </div>
            {statFilter==="mes" && <div style={{ display:"flex", gap:5, marginBottom:10, overflowX:"auto", paddingBottom:2 }}>{MESES.map((m,i)=>(<button key={m} onClick={()=>setStatValue(p=>({...p,mes:i}))} style={{ background:statValue.mes===i?T.accentLight:"transparent", border:`1px solid ${statValue.mes===i?T.borderStrong:T.border}`, borderRadius:20, padding:"4px 11px", color:statValue.mes===i?T.text:T.textMuted, fontSize:11, fontWeight:600, cursor:"pointer", whiteSpace:"nowrap" }}>{m}</button>))}</div>}
            {statFilter==="quarter" && <div style={{ display:"flex", gap:6, marginBottom:10 }}>{["Q1","Q2","Q3","Q4"].map(q=>(<button key={q} onClick={()=>setStatValue(p=>({...p,quarter:q}))} style={{ background:statValue.quarter===q?T.accentLight:"transparent", border:`1px solid ${statValue.quarter===q?T.borderStrong:T.border}`, borderRadius:20, padding:"5px 14px", color:statValue.quarter===q?T.text:T.textMuted, fontSize:12, fontWeight:600, cursor:"pointer" }}>{q}</button>))}</div>}
            {statFilter==="ano" && <div style={{ display:"flex", gap:6, marginBottom:10 }}>{[2024,2025,2026].map(y=>(<button key={y} onClick={()=>setStatValue(p=>({...p,ano:y}))} style={{ background:statValue.ano===y?T.accentLight:"transparent", border:`1px solid ${statValue.ano===y?T.borderStrong:T.border}`, borderRadius:20, padding:"5px 14px", color:statValue.ano===y?T.text:T.textMuted, fontSize:12, fontWeight:600, cursor:"pointer" }}>{y}</button>))}</div>}

            {statsData.length===0
              ? <div style={{ textAlign:"center", padding:"60px 20px", color:T.textMuted, fontSize:14 }}>Sem dados para este período</div>
              : (
                <>
                  <div style={{ ...card, padding:"20px", marginBottom:10 }}>
                    <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
                      {[["Faturação",eur(totals.revenue),T.text],["Lucro",eur(totals.profit),totals.profit>=0?T.green:T.red],["Adspend",eur(totals.ads),T.amber],["ROAS",totals.roas?totals.roas.toFixed(2)+"x":"—",T.green],["Margem",pct(totals.margin),totals.margin>=0?T.green:T.red],["Dias",totals.dias,T.purple]].map(([l,v,c])=>(
                        <div key={l}><div style={{ color:T.textMuted, fontSize:11, marginBottom:3 }}>{l}</div><div style={{ color:c, fontSize:18, fontWeight:800, fontFamily:mono }}>{v}</div></div>
                      ))}
                    </div>
                  </div>
                  <div style={{ ...card, padding:"18px 16px 10px", marginBottom:10 }}>
                    <ResponsiveContainer width="100%" height={130}>
                      <BarChart data={statsData.map(r=>({date:fmtDate(r.date),profit:r.profit}))} barSize={statsData.length>20?5:10} margin={{top:4,right:4,left:0,bottom:0}}>
                        <XAxis dataKey="date" tick={{fill:T.textLight,fontSize:9}} axisLine={false} tickLine={false} interval="preserveStartEnd" />
                        <YAxis tick={{fill:T.textLight,fontSize:9}} axisLine={false} tickLine={false} tickFormatter={v=>"€"+v} width={44} />
                        <Tooltip content={<ChartTip />} cursor={{fill:"rgba(0,0,0,0.02)"}} />
                        <ReferenceLine y={0} stroke={T.border} />
                        <Bar dataKey="profit" radius={[4,4,0,0]}>{statsData.map((e,i)=><Cell key={i} fill={e.profit>=0?"#86EFAC":"#FCA5A5"} />)}</Bar>
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                  <div style={{ ...card, padding:"18px 20px" }}>
                    {[["Faturação",totals.revenue,T.text],["− COG",totals.cog,T.purple],["− Adspend",totals.ads,T.amber],["− Devoluções",totals.refunds,T.red],["− Taxas",totals.revenue*0.05,T.textMuted],["= Lucro",totals.profit,totals.profit>=0?T.green:T.red]].map(([l,v,c],i)=>(
                      <div key={l} style={{ display:"flex", justifyContent:"space-between", padding:"9px 0", borderTop:i>0?`1px solid ${T.border}`:"none" }}>
                        <span style={{ color:i===5?T.text:T.textMuted, fontSize:13, fontWeight:i===5?700:400 }}>{l}</span>
                        <span style={{ color:c, fontSize:13, fontWeight:i===5?800:600, fontFamily:mono }}>{eur(v)}</span>
                      </div>
                    ))}
                  </div>
                </>
              )
            }
          </div>
        </div>
      )}

      {tab==="add" && (
        <MobileAddForm
          form={form} setForm={setForm} editId={editId}
          saving={saving} preview={preview}
          onSave={handleSave}
          onBack={() => { setTab("home"); resetForm(); }}
          f={f} eur={eur} pct={pct} mono={mono} T={T}
        />
      )}


      {tab==="definicoes" && (
        <div style={{ paddingBottom:90 }}>
          <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`, padding:"52px 20px 16px" }}>
            <div style={{ fontSize:22, fontWeight:800 }}>Definições</div>
          </div>
          <div style={{ padding:"16px" }}>
            <div style={{ ...card, padding:"22px", marginBottom:12 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:14 }}>Logo da Loja</div>
              <div style={{ display:"flex", alignItems:"center", gap:14, marginBottom:4 }}>
                <div onClick={() => settingsFileRef.current.click()} style={{ width:64, height:64, borderRadius:14, background:"transparent", border:`2px dashed ${T.border}`, display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", overflow:"hidden", flexShrink:0 }}>
                  {settingsLogoPreview
                    ? <img src={settingsLogoPreview} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                    : logoUrl
                      ? <img src={logoUrl} alt="logo" style={{ width:"100%", height:"100%", objectFit:"cover" }} />
                      : <span style={{ fontSize:22 }}>🏪</span>
                  }
                </div>
                <button onClick={() => settingsFileRef.current.click()} style={{ background:T.bg, border:`1px solid ${T.border}`, borderRadius:8, padding:"9px 16px", fontSize:13, fontWeight:600, cursor:"pointer", color:T.text }}>Alterar logo</button>
                <input ref={settingsFileRef} type="file" accept="image/*" onChange={handleSettingsLogo} style={{ display:"none" }} />
              </div>
            </div>

            <div style={{ ...card, padding:"22px", marginBottom:12 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:8 }}>Nome da Loja</div>
              <input type="text" value={settingsName} onChange={e=>setSettingsName(e.target.value)}
                placeholder={profile?.store_name || "Nome da tua loja"}
                style={{ width:"100%", background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:15, outline:"none", fontFamily:"inherit" }} />
            </div>

            <div style={{ ...card, padding:"22px", marginBottom:12 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:8 }}>Taxa de Transacção</div>
              <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:6 }}>
                <input type="text" inputMode="decimal" value={settingsTxRate} onChange={e=>setSettingsTxRate(e.target.value)}
                  placeholder={`${(txRate*100).toFixed(1)}`}
                  style={{ width:100, background:T.bg, border:`1.5px solid ${T.border}`, borderRadius:10, padding:"12px 14px", color:T.text, fontSize:16, outline:"none", fontFamily:"'DM Mono',monospace" }} />
                <span style={{ color:T.textMuted, fontSize:13 }}>% · actual: {(txRate*100).toFixed(1)}%</span>
              </div>
              <div style={{ color:T.textLight, fontSize:12 }}>Taxa do Shopify por transacção. Por defeito 5%.</div>
            </div>

            <button onClick={handleSettingsSave} disabled={settingsSaving}
              style={{ width:"100%", background:settingsSaved?"#16A34A":T.text, border:"none", borderRadius:12, padding:"14px", color:"#fff", fontSize:15, fontWeight:700, cursor:"pointer", marginBottom:10, transition:"background 0.2s" }}>
              {settingsSaved?"✓ Guardado!":settingsSaving?"A guardar...":"Guardar alterações"}
            </button>

            <div style={{ ...card, padding:"18px 20px", marginBottom:12 }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:10 }}>Plano</div>
              {profile?.plan === "pro" ? (
                <div style={{ background:T.greenBg, border:`1px solid ${T.greenBorder}`, borderRadius:10, padding:"12px 14px", marginBottom:12, display:"flex", alignItems:"center", gap:8 }}>
                  <span>✓</span>
                  <div><div style={{ color:T.green, fontWeight:700, fontSize:13 }}>StorePNL Pro</div><div style={{ color:T.green, fontSize:11 }}>€4/mês · activo</div></div>
                </div>
              ) : (
                <button onClick={handleUpgrade} style={{ width:"100%", background:T.text, border:"none", borderRadius:10, padding:"12px", color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer", marginBottom:12 }}>
                  Subscrever — €4/mês →
                </button>
              )}
              <div style={{ color:T.textMuted, fontSize:12, marginBottom:10 }}>{session.user.email}</div>
              <button onClick={() => supabase.auth.signOut()} style={{ width:"100%", background:T.redBg, border:`1px solid ${T.redBorder}`, borderRadius:10, padding:"11px", color:T.red, fontSize:13, fontWeight:600, cursor:"pointer" }}>Sair da conta</button>
            </div>

            <div style={{ ...card, padding:"20px" }}>
              <div style={{ color:T.textMuted, fontSize:11, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase", marginBottom:4 }}>Importar Dados</div>
              <div style={{ color:T.textMuted, fontSize:13, marginBottom:14 }}>Importa o teu ficheiro P&L Sheet (.xlsx).</div>
              <ExcelImport userId={session.user.id} onImportDone={loadEntries} T={T} />
            </div>
          </div>
        </div>
      )}

      {tab!=="add" && tab!=="detalhe" && (
        <div style={{ position:"fixed", bottom:0, left:"50%", transform:"translateX(-50%)", width:"100%", maxWidth:430, background:T.surface, borderTop:`1px solid ${T.border}`, padding:"10px 0 28px", display:"flex" }}>
          {[["home","▦","Início"],["analytics","◈","Analytics"],["definicoes","⚙","Config"]].map(([t,icon,label])=>(
            <button key={t} onClick={()=>setTab(t)} style={{ background:"none", border:"none", cursor:"pointer", display:"flex", flexDirection:"column", alignItems:"center", gap:3, flex:1 }}>
              <span style={{ fontSize:20, color:tab===t?T.text:T.textLight }}>{icon}</span>
              <span style={{ fontSize:10, fontWeight:700, color:tab===t?T.text:T.textLight, letterSpacing:"0.06em", textTransform:"uppercase" }}>{label}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
