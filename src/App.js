import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import { inject } from "@vercel/analytics";

inject();

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


      {/* MOCKUPS SECTION */}
      <div style={{ background:S.bg, padding:"72px 24px", borderTop:`1px solid ${S.border}` }}>
        <div style={{ maxWidth:1000, margin:"0 auto" }}>
          <div style={{ textAlign:"center", marginBottom:48 }}>
            <div style={{ fontSize:11, fontWeight:700, color:S.muted, letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:10 }}>A dashboard</div>
            <h2 style={{ fontSize:isMobile?"28px":"40px", fontWeight:800, letterSpacing:"-0.03em", marginBottom:12 }}>Exactamente assim,<br/>no teu bolso e no PC.</h2>
            <p style={{ fontSize:16, color:S.muted, maxWidth:460, margin:"0 auto" }}>Mobile, desktop, sempre sincronizado.</p>
          </div>

          {/* Mobile + Desktop side by side */}
          <div style={{ display:"flex", flexDirection:"column", gap:32, alignItems:"center", justifyContent:"center", marginBottom:48 }}>

            {/* MOBILE MOCKUP */}
            <div style={{ flexShrink:0 }}>
              <div style={{ fontSize:11, fontWeight:700, color:S.muted, letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:12, textAlign:"center" }}>📱 iPhone</div>
              <div style={{ width:240, background:"#F8F7F5", borderRadius:30, border:"6px solid #1A1A1A", boxShadow:"0 24px 48px rgba(0,0,0,0.18)", overflow:"hidden" }}>
                <div style={{ width:70, height:18, background:"#1A1A1A", borderRadius:"0 0 10px 10px", margin:"0 auto" }} />
                <div style={{ padding:"8px 10px 0", background:"#F8F7F5" }}>
                  {/* Header */}
                  <div style={{ background:"#fff", borderBottom:"1px solid #EBEBEB", padding:"8px 12px", display:"flex", justifyContent:"space-between", alignItems:"center", borderRadius:"10px 10px 0 0" }}>
                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                      <div style={{ width:24, height:24, background:"#1A1A1A", borderRadius:6, display:"flex", alignItems:"center", justifyContent:"center" }}>
                        <div style={{ display:"flex", alignItems:"flex-end", gap:1.5, height:11 }}>
                          {[3,5,8,11].map((h,i)=><div key={i} style={{ width:2.5, height:h, background:"#4ade80", borderRadius:"1px 1px 0 0", opacity:[0.3,0.5,0.75,1][i] }} />)}
                        </div>
                      </div>
                      <div>
                        <div style={{ fontSize:9, fontWeight:700 }}>A minha loja</div>
                        <div style={{ fontSize:7, color:"#8A8A8A" }}>Dashboard P&L</div>
                      </div>
                    </div>
                    <div style={{ background:"#1A1A1A", color:"#fff", borderRadius:5, padding:"3px 7px", fontSize:8, fontWeight:700 }}>+ Dia</div>
                  </div>
                  {/* Hero */}
                  <div style={{ background:"#F0FDF4", border:"1px solid #86EFAC", borderRadius:10, padding:10, margin:"6px 0" }}>
                    <div style={{ fontSize:7, fontWeight:700, color:"#8A8A8A", letterSpacing:"0.08em", textTransform:"uppercase", marginBottom:2 }}>Hoje</div>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
                      <div style={{ fontFamily:"'DM Mono',monospace", fontSize:22, fontWeight:800, color:"#16A34A", letterSpacing:"-0.03em" }}>€127.43</div>
                      <div style={{ textAlign:"right" }}>
                        <div style={{ fontSize:7, color:"#8A8A8A" }}>vs ontem</div>
                        <div style={{ fontFamily:"'DM Mono',monospace", fontSize:10, fontWeight:800, color:"#16A34A" }}>↑ €42.10</div>
                      </div>
                    </div>
                    <div style={{ fontSize:7, color:"#8A8A8A", marginTop:3 }}>38.2% margem · ROAS 2.8x</div>
                  </div>
                  {/* Metrics */}
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:5, marginBottom:6 }}>
                    {[["Faturação","€334.00","#2563EB","↑ €54","#16A34A"],["Adspend","€89.50","#D97706","↑ €12","#DC2626"]].map(([l,v,c,d,dc])=>(
                      <div key={l} style={{ background:"#fff", border:"1px solid #EBEBEB", borderRadius:8, padding:"8px 9px" }}>
                        <div style={{ fontSize:7, fontWeight:700, color:"#8A8A8A", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:3 }}>{l}</div>
                        <div style={{ fontFamily:"'DM Mono',monospace", fontSize:13, fontWeight:800, color:c }}>{v}</div>
                        <div style={{ fontSize:7, color:dc, marginTop:2, fontWeight:600 }}>{d} vs ontem</div>
                      </div>
                    ))}
                  </div>
                  {/* Chart */}
                  <div style={{ background:"#fff", border:"1px solid #EBEBEB", borderRadius:8, padding:"8px 10px", marginBottom:6 }}>
                    <div style={{ fontSize:7, fontWeight:700, color:"#8A8A8A", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:6 }}>Lucro — 30 dias</div>
                    <div style={{ display:"flex", alignItems:"flex-end", gap:2, height:30 }}>
                      {[35,28,55,42,68,78,30,58,88,72,95,100].map((h,i)=>(
                        <div key={i} style={{ flex:1, height:`${h}%`, background:h>50?"#86EFAC":"#FCA5A5", borderRadius:"1px 1px 0 0" }} />
                      ))}
                    </div>
                  </div>
                  {/* Days */}
                  {[["18 Abr","€334 · €89 ads","€127.43","#16A34A"],["17 Abr","€280 · €77 ads","€85.33","#16A34A"],["16 Abr","€95 · €88 ads","-€18.50","#DC2626"]].map(([d,s,p,c])=>(
                    <div key={d} style={{ background:"#fff", border:"1px solid #EBEBEB", borderRadius:8, padding:"8px 10px", marginBottom:4, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div><div style={{ fontSize:10, fontWeight:600 }}>{d}</div><div style={{ fontSize:7, color:"#8A8A8A" }}>{s}</div></div>
                      <div style={{ fontFamily:"'DM Mono',monospace", fontSize:11, fontWeight:800, color:c }}>{p}</div>
                    </div>
                  ))}
                </div>
                {/* Nav */}
                <div style={{ background:"#fff", borderTop:"1px solid #EBEBEB", padding:"6px 0 12px", display:"flex", justifyContent:"space-around", marginTop:6 }}>
                  {[["▦","Início",true],["◈","Analytics",false],["⚙","Config",false]].map(([icon,label,active])=>(
                    <div key={label} style={{ textAlign:"center" }}>
                      <div style={{ fontSize:13, color:active?"#1A1A1A":"#C4C4C4" }}>{icon}</div>
                      <div style={{ fontSize:6, fontWeight:700, color:active?"#1A1A1A":"#C4C4C4", letterSpacing:"0.06em", textTransform:"uppercase" }}>{label}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
