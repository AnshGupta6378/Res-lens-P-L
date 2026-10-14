import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react';
import {
  AreaChart, Area, BarChart, Bar, LineChart, Line, ComposedChart,
  XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid,
  ReferenceLine, PieChart, Pie, Cell, Legend
} from 'recharts';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import SAMPLE_CSV from '../samples/sample-portfolio.csv?raw';
import { Upload, ArrowLeft, TrendingUp, TrendingDown, Building2, MapPin, Calendar, Search, X, BarChart3, Calculator, RefreshCw, Info, ArrowRight } from 'lucide-react';

// ─────────────────────────────────────────────────────────────────────────────
// HOOKS
// ─────────────────────────────────────────────────────────────────────────────
function useIsMobile(breakpoint = 768) {
  const [isMobile, setIsMobile] = useState(
    typeof window !== 'undefined' ? window.innerWidth < breakpoint : false
  );
  useEffect(() => {
    const handler = () => setIsMobile(window.innerWidth < breakpoint);
    window.addEventListener('resize', handler);
    return () => window.removeEventListener('resize', handler);
  }, [breakpoint]);
  return isMobile;
}


// ─────────────────────────────────────────────────────────────────────────────
// THEME
// ─────────────────────────────────────────────────────────────────────────────
const FONTS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,300;9..144,400;9..144,500;9..144,600;9..144,700&family=Geist:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
`;

const C = {
  bg:        '#F7F5F0',   // warm off-white
  surface:   '#FFFFFF',
  card:      '#FFFFFF',
  cardHover: '#FAF8F3',
  border:    '#E8E4DC',   // hairline warm grey
  borderLt:  '#D6D1C5',
  text:      '#1A1A1A',   // near-black
  textDim:   '#6B6660',   // mid warm grey
  textMute:  '#A39E94',   // light warm grey
  accent:    '#C2410C',   // burnt orange
  accentDk:  '#9A3412',
  success:   '#15803D',
  successDk: '#166534',
  danger:    '#B91C1C',
  info:      '#1D4ED8',
  violet:    '#6D28D9',
  rose:      '#BE123C',
};

const fmtINR = (n) => {
  if (n == null || isNaN(n)) return '-';
  const abs = Math.abs(n);
  if (abs >= 1e7) return `₹${(n/1e7).toFixed(2)}Cr`;
  if (abs >= 1e5) return `₹${(n/1e5).toFixed(2)}L`;
  if (abs >= 1e3) return `₹${(n/1e3).toFixed(1)}K`;
  return `₹${Math.round(n).toLocaleString('en-IN')}`;
};
const fmtINRFull = (n) => n == null || isNaN(n) ? '-' : `₹${Math.round(n).toLocaleString('en-IN')}`;
const fmtNum = (n) => n == null || isNaN(n) ? '-' : Math.round(n).toLocaleString('en-IN');
const fmtPct = (n) => n == null || isNaN(n) ? '-' : `${(n*100).toFixed(1)}%`;
const fmtMul = (n) => n == null || isNaN(n) ? '-' : `${n.toFixed(2)}x`;

// ─────────────────────────────────────────────────────────────────────────────
// DATA PARSING
// ─────────────────────────────────────────────────────────────────────────────
const normalizeMonth = (v) => {
  if (!v) return null;
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 7);
  // Excel serial date
  if (/^\d+(\.\d+)?$/.test(s) && parseFloat(s) > 30000) {
    const date = new Date(Date.UTC(1899, 11, 30) + parseFloat(s) * 86400000);
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  return s;
};

const num = (v) => {
  if (v == null || v === '') return 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
};

const parseRows = (rawRows) => {
  return rawRows.map(r => ({
    month:    normalizeMonth(r.month),
    res:      String(r.restaurant || '').trim(),
    outletId:      String(r.outlet_id || '').trim(),     // unique outlet ID
    area:  String(r.area || '').trim(),
    city:     String(r.city || '').trim(),
    cuisine:  String(r.cuisine || '').trim().toLowerCase(),
    orders:   num(r.orders),
    cv:       num(r.net_sales),
    gmv:      num(r.gmv),
    pc:       num(r.packaging),
    platDisc:      num(r.platform_discount),
    restDisc:      num(r.restaurant_discount),
    billed:   num(r.ad_spend),
    adsCv:    num(r.ad_sales),
    menuOpens: num(r.menu_opens),
    m2o:      num(r.menu_to_order),
  })).filter(r => r.month && r.res);
};

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────
export default function App() {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState('');
  const [view, setView] = useState('portfolio'); // 'portfolio' | 'restaurant' | 'simulator'
  const [selRes, setSelRes] = useState(null);
  const [selArea, setSelArea] = useState('All');
  const [search, setSearch] = useState('');
  const isMobile = useIsMobile();

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setFileName(file.name);
    const isCSV = file.name.toLowerCase().endsWith('.csv');
    if (isCSV) {
      Papa.parse(file, {
        header: true,
        dynamicTyping: false,
        skipEmptyLines: true,
        complete: (results) => setRows(parseRows(results.data))
      });
    } else {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      // Look for Raw Data sheet, or first sheet
      let sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('raw')) || wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const ref = sheet['!ref'];
      let rawJson;
      // Some exports carry two title rows above the header, so try row 3 first
      try {
        rawJson = XLSX.utils.sheet_to_json(sheet, { range: 2, defval: '' });
        if (!rawJson.length || !rawJson[0].month) {
          // Fall back to row 1
          rawJson = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        }
      } catch {
        rawJson = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      }
      setRows(parseRows(rawJson));
    }
  }, []);

  const loadSample = useCallback(() => {
    const parsed = Papa.parse(SAMPLE_CSV, { header: true, skipEmptyLines: true });
    setFileName('sample-portfolio.csv');
    setRows(parseRows(parsed.data));
  }, []);

  if (!rows && view !== 'simulator') return <UploadScreen onFile={handleFile} onSample={loadSample} onGoToSimulator={() => setView('simulator')} isMobile={isMobile} />;
  if (rows && rows.length === 0) {
    return <UploadScreen onFile={handleFile} onSample={loadSample} onGoToSimulator={() => setView('simulator')} isMobile={isMobile} error="No valid rows found. Expected columns: month, restaurant, area, gmv, net_sales, orders, ad_spend, platform_discount, restaurant_discount, packaging" />;
  }

  return (
    <div style={{ minHeight: '100vh', background: C.bg, color: C.text, fontFamily: "'Geist', sans-serif" }}>
      <style>{FONTS}</style>
      <style>{`
        * { box-sizing: border-box; }
        body { margin: 0; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-track { background: ${C.bg}; }
        ::-webkit-scrollbar-thumb { background: ${C.borderLt}; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: ${C.textMute}; }
        .num { font-family: 'JetBrains Mono', monospace; font-variant-numeric: tabular-nums; }
        .display { font-family: 'Fraunces', serif; font-feature-settings: "ss01"; }
        button { font-family: 'Geist', sans-serif; }
        input, select { font-family: 'Geist', sans-serif; }
        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
        input[type="number"] { -moz-appearance: textfield; }
      `}</style>

      <Header
        fileName={fileName}
        view={view}
        onBack={() => { setView('portfolio'); setSelRes(null); }}
        onReset={() => { setRows(null); setFileName(''); setView('portfolio'); setSelRes(null); }}
        onChangeView={setView}
        hasData={!!rows}
        rowCount={rows?.length || 0}
        isMobile={isMobile}
      />

      {view === 'portfolio' && rows && (
        <PortfolioView
          rows={rows}
          onSelectRestaurant={(r) => { setSelRes(r); setView('restaurant'); }}
          selArea={selArea}
          setSelArea={setSelArea}
          search={search}
          setSearch={setSearch}
          isMobile={isMobile}
        />
      )}

      {view === 'restaurant' && selRes && rows && (
        <RestaurantView rows={rows} resName={selRes} isMobile={isMobile} />
      )}

      {view === 'simulator' && (
        <SimulatorView isMobile={isMobile} />
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// UPLOAD SCREEN
// ─────────────────────────────────────────────────────────────────────────────
function UploadScreen({ onFile, onSample, error, onGoToSimulator, isMobile }) {
  const [dragging, setDragging] = useState(false);
  const preview = computeSim(SIM_PRESETS.base);
  const inputRef = useRef();

  return (
    <div style={{
      minHeight: '100vh',
      background: `radial-gradient(ellipse at top, #FFFFFF 0%, ${C.bg} 60%)`,
      color: C.text,
      fontFamily: "'Geist', sans-serif",
      padding: isMobile ? '40px 20px' : '64px 40px',
    }}>
      <style>{FONTS}</style>
      <div style={{ maxWidth: 1120, width: '100%', margin: '0 auto' }}>

        {/* ── Header pill ─────────────────────────────────────────────── */}
        <div style={{
          display: 'inline-flex',
          alignItems: 'center',
          gap: 10,
          padding: '6px 12px',
          border: `1px solid ${C.border}`,
          borderRadius: 100,
          background: C.surface,
          fontSize: 11,
          letterSpacing: '0.12em',
          textTransform: 'uppercase',
          color: C.textDim,
          marginBottom: 24,
        }}>
          <div style={{ width: 6, height: 6, borderRadius: '50%', background: C.accent }} />
          Res Lens · P&L Suite
        </div>

        {/* ── Hero ────────────────────────────────────────────────────── */}
        <h1 className="display" style={{
          fontFamily: "'Fraunces', serif",
          fontSize: isMobile ? 36 : 54,
          fontWeight: 400,
          letterSpacing: '-0.02em',
          margin: '0 0 16px',
          lineHeight: 1.05,
          color: C.text,
        }}>
          See the real <em style={{ color: C.accent, fontStyle: 'italic' }}>economics</em><br/>
          of your online business.
        </h1>
        <p style={{
          color: C.textDim,
          fontSize: isMobile ? 14 : 16,
          lineHeight: 1.6,
          margin: '0 0 40px',
          maxWidth: 560,
        }}>
          Two ways to begin: upload monthly actuals to see the full P&L, or open Price Lens to model a what-if.
        </p>

        {/* ── Two-card grid ───────────────────────────────────────────── */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr',
          gap: isMobile ? 16 : 20,
        }}>

          {/* ── Card 1: Upload (Res Lens) ──────────────────────────── */}
          <label
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setDragging(false);
              if (e.dataTransfer.files?.[0]) onFile(e.dataTransfer.files[0]);
            }}
            style={{
              display: 'block',
              background: C.surface,
              border: `1px solid ${dragging ? C.accent : C.border}`,
              borderRadius: 16,
              padding: isMobile ? 24 : 32,
              cursor: 'pointer',
              transition: 'all 200ms ease',
              boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
              position: 'relative',
              overflow: 'hidden',
              minHeight: isMobile ? 'auto' : 380,
              display: 'flex',
              flexDirection: 'column',
            }}
          >
            <div style={{
              position: 'absolute', top: 0, left: 0, width: 40, height: 3,
              background: C.accent,
            }} />
            <div style={{
              fontSize: 10,
              letterSpacing: '0.18em',
              textTransform: 'uppercase',
              color: C.accent,
              marginBottom: 6,
              fontWeight: 600,
            }}>
              Res Lens
            </div>
            <div className="display" style={{
              fontFamily: "'Fraunces', serif",
              fontSize: isMobile ? 22 : 26,
              fontWeight: 500,
              letterSpacing: '-0.01em',
              margin: '0 0 6px',
              color: C.text,
            }}>
              Upload your data
            </div>
            <div style={{ fontSize: 13, color: C.textDim, marginBottom: 20, lineHeight: 1.5 }}>
              Drop a monthly restaurant-level export. Full P&L per restaurant, per month, all in your browser.
            </div>

            <input
              ref={inputRef}
              type="file"
              accept=".csv,.xlsx,.xls"
              onChange={(e) => onFile(e.target.files?.[0])}
              style={{ display: 'none' }}
            />

            <div style={{
              flex: 1,
              border: `1.5px dashed ${dragging ? C.accent : C.borderLt}`,
              borderRadius: 12,
              padding: isMobile ? 24 : 32,
              background: dragging ? 'rgba(194,65,12,0.04)' : C.bg,
              transition: 'all 200ms ease',
              textAlign: 'center',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              minHeight: 160,
            }}>
              <div style={{
                width: 44, height: 44, borderRadius: 12,
                background: dragging ? C.accent : C.surface,
                color: dragging ? '#FFFFFF' : C.accent,
                border: `1px solid ${dragging ? C.accent : C.border}`,
                display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                marginBottom: 12,
                transition: 'all 200ms ease',
              }}>
                <Upload size={20} strokeWidth={1.8} />
              </div>
              <div style={{ fontSize: 14, fontWeight: 500, marginBottom: 4 }}>
                Drop here, or click to browse
              </div>
              <div style={{ fontSize: 12, color: C.textMute }}>
                <span className="num">.csv</span> · <span className="num">.xlsx</span> · <span className="num">.xls</span>
              </div>
            </div>

            <div style={{
              marginTop: 14,
              fontSize: 11,
              color: C.textMute,
              lineHeight: 1.6,
            }}>
              Needs: <span className="num" style={{ color: C.textDim }}>month, restaurant, area, gmv, net_sales, orders, ad_spend, platform_discount, restaurant_discount, packaging</span>
            </div>
            <button
              type="button"
              onClick={(e) => { e.preventDefault(); e.stopPropagation(); onSample(); }}
              style={{
                marginTop: 14, alignSelf: 'flex-start',
                padding: '9px 14px', background: C.text, color: C.surface,
                border: 'none', borderRadius: 8, fontSize: 13, fontWeight: 500, cursor: 'pointer',
                display: 'inline-flex', alignItems: 'center', gap: 8,
              }}
            >
              <BarChart3 size={14} /> Try it with sample data
            </button>
          </label>

          {/* ── Card 2: Price Lens ──────────────────────────────────── */}
          <button
            onClick={onGoToSimulator}
            style={{
              display: 'flex',
              flexDirection: 'column',
              justifyContent: 'flex-start',
              alignItems: 'stretch',
              textAlign: 'left',
              background: C.surface,
              border: `1px solid ${C.border}`,
              borderRadius: 16,
              padding: isMobile ? 24 : 32,
              cursor: 'pointer',
              transition: 'all 200ms ease',
              boxShadow: '0 1px 2px rgba(0,0,0,0.04)',
              position: 'relative',
              overflow: 'hidden',
              minHeight: isMobile ? 'auto' : 380,
              fontFamily: 'inherit',
              color: C.text,
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.borderColor = C.accent;
              e.currentTarget.style.boxShadow = '0 4px 12px rgba(0,0,0,0.06)';
              e.currentTarget.style.transform = 'translateY(-2px)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.borderColor = C.border;
              e.currentTarget.style.boxShadow = '0 1px 2px rgba(0,0,0,0.04)';
              e.currentTarget.style.transform = 'translateY(0)';
            }}
          >
            <div style={{
              position: 'absolute', top: 0, left: 0, width: 40, height: 3,
              background: C.success,
            }} />
            <div style={{
              fontSize: 10,
              letterSpacing: '0.18em',
              textTransform: 'uppercase',
              color: C.success,
              marginBottom: 6,
              fontWeight: 600,
            }}>
              Price Lens · New
            </div>
            <div className="display" style={{
              fontFamily: "'Fraunces', serif",
              fontSize: isMobile ? 22 : 26,
              fontWeight: 500,
              letterSpacing: '-0.01em',
              margin: '0 0 6px',
              color: C.text,
            }}>
              Simulate a what-if
            </div>
            <div style={{ fontSize: 13, color: C.textDim, marginBottom: 20, lineHeight: 1.5 }}>
              No data needed. Model pricing, discounts and ads, and see where every rupee goes.
            </div>

            {/* Mini preview tile */}
            <div style={{
              background: C.bg,
              border: `1px solid ${C.border}`,
              borderRadius: 12,
              padding: 18,
              marginBottom: 16,
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: 12 }}>
                <div>
                  <div style={{ fontSize: 10, letterSpacing: '0.14em', textTransform: 'uppercase', color: C.textMute, marginBottom: 4 }}>
                    Sample payout
                  </div>
                  <div className="num" style={{ fontSize: 28, fontWeight: 500, color: C.success, lineHeight: 1 }}>
                    {fmtINR(preview.payout)}
                  </div>
                </div>
                <div style={{
                  padding: '4px 10px',
                  background: `${C.success}15`,
                  color: C.success,
                  borderRadius: 100,
                  fontSize: 11,
                  fontWeight: 600,
                }}>
                  <span className="num">{(preview.payoutPctOffline * 100).toFixed(0)}%</span> of offline
                </div>
              </div>
              <div style={{ display: 'flex', gap: 8 }}>
                {[
                  { label: 'Net sales', val: fmtINR(preview.cv) },
                  { label: 'Platform',  val: fmtINR(preview.totalPlatform) },
                  { label: 'Ads',       val: fmtINR(preview.totalMarketing) },
                ].map((s, i) => (
                  <div key={i} style={{
                    flex: 1,
                    padding: '8px 10px',
                    background: C.surface,
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                  }}>
                    <div style={{ fontSize: 9, letterSpacing: '0.1em', textTransform: 'uppercase', color: C.textMute, marginBottom: 2 }}>
                      {s.label}
                    </div>
                    <div className="num" style={{ fontSize: 12, fontWeight: 500, color: C.text }}>
                      {s.val}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div style={{
              display: 'inline-flex', alignSelf: 'flex-start',
              alignItems: 'center',
              gap: 8,
              padding: '10px 16px',
              background: C.text,
              color: C.surface,
              borderRadius: 8,
              fontSize: 13,
              fontWeight: 500,
            }}>
              <Calculator size={14} />
              Launch Price Lens
              <ArrowRight size={14} />
            </div>
          </button>
        </div>

        {error && (
          <div style={{
            marginTop: 20,
            padding: 14,
            border: `1px solid ${C.danger}40`,
            background: `${C.danger}10`,
            borderRadius: 10,
            fontSize: 13,
            color: C.danger,
          }}>
            {error}
          </div>
        )}

        {/* ── Footer ──────────────────────────────────────────────────── */}
        <div style={{
          marginTop: 60,
          paddingTop: 24,
          borderTop: `1px solid ${C.border}`,
          fontSize: 12,
          color: C.textMute,
          display: 'flex',
          justifyContent: 'space-between',
          flexWrap: 'wrap',
          gap: 12,
        }}>
          <div>Built for people who manage restaurants on delivery apps.</div>
          <div>Everything runs in your browser. No data leaves your device.</div>
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// HEADER
// ─────────────────────────────────────────────────────────────────────────────
function Header({ fileName, view, onBack, onReset, onChangeView, hasData, rowCount, isMobile }) {
  const isDashboardActive = view === 'portfolio' || view === 'restaurant';
  return (
    <div style={{
      borderBottom: `1px solid ${C.border}`,
      background: `${C.bg}E6`,
      backdropFilter: 'blur(12px)',
      WebkitBackdropFilter: 'blur(12px)',
      position: 'sticky',
      top: 0,
      zIndex: 50,
    }}>
      <div style={{
        maxWidth: 1400,
        margin: '0 auto',
        padding: isMobile ? '12px 16px' : '14px 32px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: 12,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: isMobile ? 10 : 16, flex: '0 1 auto', minWidth: 0 }}>
          {view === 'restaurant' && (
            <button
              onClick={onBack}
              style={{
                background: C.card, border: `1px solid ${C.border}`, color: C.text,
                padding: '7px 10px', borderRadius: 8, cursor: 'pointer',
                display: 'flex', alignItems: 'center', gap: 6, fontSize: 13,
                flexShrink: 0,
              }}
            >
              <ArrowLeft size={14} /> {!isMobile && 'Back'}
            </button>
          )}
          <div className="display" style={{
            fontFamily: "'Fraunces', serif",
            fontSize: isMobile ? 18 : 22, fontWeight: 500, letterSpacing: '-0.01em',
            whiteSpace: 'nowrap',
          }}>
            Res Lens
            <span style={{ color: C.accent, marginLeft: 4 }}>.</span>
          </div>

          {/* Tabs */}
          <div style={{
            display: 'flex',
            gap: 2,
            padding: 3,
            background: C.bg,
            border: `1px solid ${C.border}`,
            borderRadius: 9,
            marginLeft: isMobile ? 4 : 8,
          }}>
            <button
              onClick={() => onChangeView(hasData ? 'portfolio' : 'portfolio')}
              style={{
                padding: isMobile ? '6px 10px' : '7px 14px',
                background: isDashboardActive ? C.surface : 'transparent',
                border: 'none',
                borderRadius: 7,
                color: isDashboardActive ? C.text : C.textDim,
                cursor: 'pointer',
                fontSize: isMobile ? 12 : 13,
                fontWeight: 500,
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                boxShadow: isDashboardActive ? '0 1px 2px rgba(0,0,0,0.05)' : 'none',
                transition: 'all 120ms ease',
              }}
            >
              <BarChart3 size={13} strokeWidth={2} />
              {!isMobile && 'Dashboard'}
            </button>
            <button
              onClick={() => onChangeView('simulator')}
              style={{
                padding: isMobile ? '6px 10px' : '7px 14px',
                background: view === 'simulator' ? C.surface : 'transparent',
                border: 'none',
                borderRadius: 7,
                color: view === 'simulator' ? C.text : C.textDim,
                cursor: 'pointer',
                fontSize: isMobile ? 12 : 13,
                fontWeight: 500,
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                boxShadow: view === 'simulator' ? '0 1px 2px rgba(0,0,0,0.05)' : 'none',
                transition: 'all 120ms ease',
              }}
            >
              <Calculator size={13} strokeWidth={2} />
              {!isMobile && 'Price Lens'}
            </button>
          </div>
        </div>

        {!isMobile && hasData && view !== 'simulator' && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexShrink: 0 }}>
            <div style={{ fontSize: 12, color: C.textDim }}>
              <span className="num" style={{ color: C.text }}>{fmtNum(rowCount)}</span> rows
              {fileName && <span> · {fileName}</span>}
            </div>
            <button
              onClick={onReset}
              style={{
                background: 'transparent', border: `1px solid ${C.border}`, color: C.textDim,
                padding: '6px 10px', borderRadius: 6, cursor: 'pointer', fontSize: 12,
              }}
            >
              Replace
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PORTFOLIO VIEW
// ─────────────────────────────────────────────────────────────────────────────
function PortfolioView({ rows, onSelectRestaurant, selArea, setSelArea, search, setSearch, isMobile }) {
  // Compute filters
  const areas = useMemo(() => {
    const set = new Set(rows.map(r => r.area).filter(Boolean));
    return ['All', ...Array.from(set).sort()];
  }, [rows]);

  const filtered = useMemo(() => {
    return rows.filter(r => {
      if (selArea !== 'All' && r.area !== selArea) return false;
      return true;
    });
  }, [rows, selArea]);

  // KPIs
  const kpis = useMemo(() => {
    const totals = filtered.reduce((acc, r) => {
      acc.gmv    += r.gmv;
      acc.cv     += r.cv;
      acc.orders += r.orders;
      acc.billed += r.billed;
      acc.platDisc    += r.platDisc;
      acc.restDisc    += r.restDisc;
      acc.adsCv  += r.adsCv;
      return acc;
    }, { gmv: 0, cv: 0, orders: 0, billed: 0, platDisc: 0, restDisc: 0, adsCv: 0 });
    const uniqueRes = new Set(filtered.map(r => r.res)).size;
    return { ...totals, uniqueRes, aov: totals.orders ? totals.gmv / totals.orders : 0,
             roas: totals.billed ? totals.cv / totals.billed : 0 };
  }, [filtered]);

  // Monthly trend
  const monthlyTrend = useMemo(() => {
    const byMonth = {};
    filtered.forEach(r => {
      if (!byMonth[r.month]) byMonth[r.month] = { month: r.month, gmv: 0, orders: 0, billed: 0, cv: 0 };
      byMonth[r.month].gmv    += r.gmv;
      byMonth[r.month].orders += r.orders;
      byMonth[r.month].billed += r.billed;
      byMonth[r.month].cv     += r.cv;
    });
    return Object.values(byMonth).sort((a, b) => a.month.localeCompare(b.month))
      .map(m => ({ ...m, label: monthLabel(m.month) }));
  }, [filtered]);

  // Top restaurants, grouped by brand, tracking distinct outlets by outlet ID
  const topRes = useMemo(() => {
    const byRes = {};
    filtered.forEach(r => {
      if (!byRes[r.res]) byRes[r.res] = {
        res: r.res, cuisine: r.cuisine,
        gmv: 0, orders: 0, billed: 0, cv: 0,
        months: new Set(),
        outlets: {},          // keyed by outlet ID
        areas: new Set(),
      };
      const b = byRes[r.res];
      b.gmv    += r.gmv;
      b.orders += r.orders;
      b.billed += r.billed;
      b.cv     += r.cv;
      b.months.add(r.month);
      if (r.area) b.areas.add(r.area);
      // track per-outlet
      const key = r.outletId || r.area || 'unknown';
      if (!b.outlets[key]) b.outlets[key] = { outletId: r.outletId, area: r.area, gmv: 0 };
      b.outlets[key].gmv += r.gmv;
    });
    let list = Object.values(byRes).map(b => {
      const outletArr = Object.values(b.outlets).sort((a, c) => c.gmv - a.gmv);
      const areaArr = Array.from(b.areas);
      return {
        res: b.res,
        cuisine: b.cuisine,
        gmv: b.gmv, orders: b.orders, billed: b.billed, cv: b.cv,
        months: b.months.size,
        aov: b.orders ? b.gmv / b.orders : 0,
        roas: b.billed ? b.cv / b.billed : 0,
        outletCount: outletArr.length,
        outlets: outletArr,
        // Location label: single area name, or "N locations" for chains
        locationLabel: outletArr.length > 1
          ? `${outletArr.length} locations`
          : (areaArr[0] || '-'),
        isChain: outletArr.length > 1,
      };
    });
    if (search) {
      const q = search.toLowerCase();
      list = list.filter(r =>
        r.res.toLowerCase().includes(q) ||
        r.outlets.some(o => (o.area || '').toLowerCase().includes(q))
      );
    }
    return list.sort((a, b) => b.gmv - a.gmv);
  }, [filtered, search]);

  return (
    <div style={{ maxWidth: 1400, margin: '0 auto', padding: isMobile ? '20px 16px' : '32px' }}>
      {/* Title + filters */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', marginBottom: 28, gap: 24, flexWrap: 'wrap' }}>
        <div>
          <div style={{ fontSize: 11, letterSpacing: '0.18em', textTransform: 'uppercase', color: C.accent, marginBottom: 6 }}>
            Portfolio Overview
          </div>
          <h1 className="display" style={{
            fontFamily: "'Fraunces', serif",
            fontSize: isMobile ? 28 : 40, fontWeight: 400, letterSpacing: '-0.02em',
            margin: 0, color: C.text, lineHeight: 1.15,
          }}>
            <span className="num" style={{ fontFamily: 'JetBrains Mono', fontSize: isMobile ? 26 : 36 }}>{kpis.uniqueRes}</span> restaurants <em style={{ color: C.textDim, fontStyle: 'italic' }}>across</em> {selArea === 'All' ? Object.keys(rows.reduce((s, r) => ({ ...s, [r.area]: 1 }), {})).length : 1} {selArea === 'All' ? 'areas' : 'area'}
          </h1>
        </div>

        <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
          <FilterSelect
            label="Location"
            value={selArea}
            options={areas}
            onChange={setSelArea}
            icon={<MapPin size={13} />}
          />
        </div>
      </div>

      {/* KPI Grid */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2, 1fr)' : 'repeat(4, 1fr)', gap: isMobile ? 10 : 14, marginBottom: isMobile ? 22 : 32 }}>
        <KPICard label="Gross GMV"     value={fmtINR(kpis.gmv)}    sub={`${fmtNum(kpis.orders)} orders`} accent={C.accent} />
        <KPICard label="Net Sales"     value={fmtINR(kpis.cv)}     sub={`AOV ${fmtINR(kpis.aov)}`}      accent={C.info} />
        <KPICard label="Ad Spend"      value={fmtINR(kpis.billed)} sub={`RoAS ${fmtMul(kpis.roas)}`}    accent={C.violet} />
        <KPICard label="Discount Mix"  value={fmtINR(kpis.platDisc + kpis.restDisc)} sub={`Platform ${fmtPct(kpis.platDisc / (kpis.platDisc + kpis.restDisc || 1))} / You ${fmtPct(kpis.restDisc / (kpis.platDisc + kpis.restDisc || 1))}`} accent={C.rose} />
      </div>

      {/* Charts row */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '2fr 1fr', gap: 14, marginBottom: isMobile ? 22 : 32 }}>
        <Card title="Monthly GMV & Ad Spend" subtitle={`${monthlyTrend.length} months of data`}>
          <ResponsiveContainer width="100%" height={300}>
            <ComposedChart data={monthlyTrend} margin={{ top: 20, right: 12, left: -10, bottom: 0 }}>
              <defs>
                <linearGradient id="gmvGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%"   stopColor={C.accent} stopOpacity={0.18} />
                  <stop offset="100%" stopColor={C.accent} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis yAxisId="L" stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={fmtINR} />
              <YAxis yAxisId="R" orientation="right" stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={fmtINR} />
              <Tooltip content={<CustomTooltip />} />
              <Area yAxisId="L" type="monotone" dataKey="gmv"    stroke={C.accent} strokeWidth={2} fill="url(#gmvGrad)" name="GMV" />
              <Bar  yAxisId="R" dataKey="billed" fill={C.violet} fillOpacity={0.85} name="Ad Spend" radius={[3,3,0,0]} barSize={14} />
            </ComposedChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Discount Funding Split" subtitle="Who paid for the discounts">
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie
                data={[
                  { name: 'Platform-funded', value: kpis.platDisc, color: C.info },
                  { name: 'Restaurant-funded', value: kpis.restDisc, color: C.accent },
                ]}
                cx="50%" cy="48%" innerRadius={60} outerRadius={95}
                paddingAngle={2}
                dataKey="value"
              >
                <Cell fill={C.info} />
                <Cell fill={C.accent} />
              </Pie>
              <Tooltip content={<CustomTooltip />} />
              <Legend
                verticalAlign="bottom"
                iconType="circle"
                wrapperStyle={{ fontSize: 11, color: C.textDim }}
              />
            </PieChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Restaurants table */}
      <Card
        title="Restaurants"
        subtitle={`Click any row to drill into single-restaurant P&L`}
        action={
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, background: C.bg, padding: '6px 10px', borderRadius: 6, border: `1px solid ${C.border}` }}>
            <Search size={12} color={C.textDim} />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Search…"
              style={{
                background: 'transparent', border: 'none', outline: 'none',
                color: C.text, fontSize: 12, width: 140,
              }}
            />
          </div>
        }
      >
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ color: C.textDim, fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase' }}>
                <th style={th()}>Restaurant</th>
                <th style={th()}>Location</th>
                <th style={th()}>Cuisine</th>
                <th style={{ ...th(), textAlign: 'right' }}>GMV</th>
                <th style={{ ...th(), textAlign: 'right' }}>Orders</th>
                <th style={{ ...th(), textAlign: 'right' }}>AOV</th>
                <th style={{ ...th(), textAlign: 'right' }}>Ad Spend</th>
                <th style={{ ...th(), textAlign: 'right' }}>RoAS</th>
              </tr>
            </thead>
            <tbody>
              {topRes.slice(0, 30).map((r, i) => (
                <tr
                  key={r.res}
                  onClick={() => onSelectRestaurant(r.res)}
                  style={{
                    cursor: 'pointer',
                    borderTop: `1px solid ${C.border}`,
                    transition: 'background 100ms',
                  }}
                  onMouseEnter={(e) => e.currentTarget.style.background = C.cardHover}
                  onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                >
                  <td style={td()}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      <span className="num" style={{ color: C.textMute, fontSize: 11, width: 18 }}>{String(i+1).padStart(2,'0')}</span>
                      <span style={{ fontWeight: 500 }}>{r.res}</span>
                      {r.isChain && (
                        <span style={{
                          fontSize: 10,
                          fontWeight: 600,
                          letterSpacing: '0.04em',
                          color: C.violet,
                          background: `${C.violet}14`,
                          border: `1px solid ${C.violet}30`,
                          padding: '1px 7px',
                          borderRadius: 100,
                          whiteSpace: 'nowrap',
                        }}>
                          {r.outletCount} outlets
                        </span>
                      )}
                    </div>
                  </td>
                  <td style={td(C.textDim)}>
                    {r.isChain
                      ? <span style={{ color: C.violet, fontWeight: 500 }}>{r.locationLabel}</span>
                      : r.locationLabel}
                  </td>
                  <td style={td(C.textDim)}>
                    <span style={{ textTransform: 'capitalize' }}>{r.cuisine || '-'}</span>
                  </td>
                  <td className="num" style={{ ...td(), textAlign: 'right' }}>{fmtINR(r.gmv)}</td>
                  <td className="num" style={{ ...td(C.textDim), textAlign: 'right' }}>{fmtNum(r.orders)}</td>
                  <td className="num" style={{ ...td(C.textDim), textAlign: 'right' }}>{fmtINR(r.aov)}</td>
                  <td className="num" style={{ ...td(), textAlign: 'right', color: r.billed > 0 ? C.violet : C.textMute }}>{r.billed > 0 ? fmtINR(r.billed) : '-'}</td>
                  <td className="num" style={{ ...td(), textAlign: 'right', color: r.roas >= 5 ? C.success : r.roas > 0 ? C.accent : C.textMute }}>{r.billed > 0 ? fmtMul(r.roas) : '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {topRes.length > 30 && (
          <div style={{ padding: '12px 0 0', textAlign: 'center', color: C.textDim, fontSize: 12 }}>
            Showing top 30 of {topRes.length}. Use search to find more.
          </div>
        )}
      </Card>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// RESTAURANT VIEW (deep dive)
// ─────────────────────────────────────────────────────────────────────────────
function RestaurantView({ rows, resName, isMobile }) {
  const [assumptions, setAssumptions] = useState({
    hike:       0.20,
    commission: 0.20,
    pgRate:     0.02,
    gst:        0.18,
    foodCost:   0.30,
  });
  const [priceAdj, setPriceAdj] = useState(0.05);  // default +5% for the live wow factor
  const [selOutlet, setSelOutlet] = useState('__all__');  // '__all__' or an outlet ID

  // All rows for this brand (across every outlet)
  const brandRows = useMemo(() =>
    rows.filter(r => r.res === resName),
    [rows, resName]
  );

  // Distinct outlets within this brand: { outletId, area, gmv }
  const outlets = useMemo(() => {
    const byOutlet = {};
    brandRows.forEach(r => {
      const key = r.outletId || r.area || 'unknown';
      if (!byOutlet[key]) byOutlet[key] = { outletId: r.outletId, area: r.area, gmv: 0 };
      byOutlet[key].gmv += r.gmv;
    });
    return Object.values(byOutlet).sort((a, b) => b.gmv - a.gmv);
  }, [brandRows]);

  const isChain = outlets.length > 1;

  // Rows scoped to the current selection (all outlets, or one specific outlet)
  const resRows = useMemo(() => {
    const scoped = selOutlet === '__all__'
      ? brandRows
      : brandRows.filter(r => (r.outletId || r.area || 'unknown') === selOutlet);
    // Aggregate by month so chain view sums across outlets per month
    const byMonth = {};
    scoped.forEach(r => {
      if (!byMonth[r.month]) byMonth[r.month] = {
        month: r.month, res: r.res, outletId: r.outletId, area: r.area,
        city: r.city, cuisine: r.cuisine,
        orders: 0, cv: 0, gmv: 0, pc: 0, platDisc: 0, restDisc: 0,
        billed: 0, adsCv: 0, menuOpens: 0, m2o: 0, m2oWeight: 0,
      };
      const m = byMonth[r.month];
      m.orders += r.orders; m.cv += r.cv; m.gmv += r.gmv; m.pc += r.pc;
      m.platDisc += r.platDisc; m.restDisc += r.restDisc; m.billed += r.billed; m.adsCv += r.adsCv;
      m.menuOpens += r.menuOpens;
      // weight M2O by menu opens for a meaningful blended rate
      m.m2o += r.m2o * (r.menuOpens || 0);
      m.m2oWeight += (r.menuOpens || 0);
    });
    return Object.values(byMonth)
      .map(m => ({ ...m, m2o: m.m2oWeight ? m.m2o / m.m2oWeight : 0 }))
      .sort((a, b) => a.month.localeCompare(b.month));
  }, [brandRows, selOutlet]);

  // Reset outlet selection when switching restaurants
  useEffect(() => { setSelOutlet('__all__'); }, [resName]);

  const meta = resRows[0] || {};
  const selectedArea = selOutlet === '__all__'
    ? (isChain ? `${outlets.length} locations` : (outlets[0]?.area || meta.area))
    : (outlets.find(o => (o.outletId || o.area || 'unknown') === selOutlet)?.area || '');

  // Monthly P&L computation
  const CUSTOMER_GST = 0.05;  // GST collected from customer (passes through, not part of payout)
  const monthly = useMemo(() => {
    return resRows.map(r => {
      const commission   = r.cv * assumptions.commission;
      const customerGst  = r.cv * CUSTOMER_GST;
      const nov          = r.cv + customerGst;
      const pg           = nov * assumptions.pgRate;       // PG is on NOV, not CV
      const gstOnFees     = (commission + pg) * assumptions.gst;
      const adSpend      = r.billed * (1 + assumptions.gst);
      const totalDed     = commission + pg + gstOnFees + adSpend;
      const payout       = r.cv - totalDed;
      const realRev      = r.gmv / (1 + assumptions.hike);
      const foodCost     = realRev * assumptions.foodCost;
      const netProfit    = payout - foodCost;
      const dedPctCV     = r.cv ? totalDed / r.cv : 0;
      const earnPctCV    = r.cv ? payout / r.cv   : 0;
      const roas         = r.billed ? r.cv / r.billed : 0;
      const netMargin    = realRev ? netProfit / realRev : 0;
      // Discount fields
      const totalDisc    = r.platDisc + r.restDisc;
      const platDiscPctGMV    = r.gmv ? r.platDisc / r.gmv : 0;
      const restDiscPctGMV    = r.gmv ? r.restDisc / r.gmv : 0;
      const totalDiscPctGMV = r.gmv ? totalDisc / r.gmv : 0;
      const platDiscShare     = totalDisc ? r.platDisc / totalDisc : 0;
      const restDiscShare     = totalDisc ? r.restDisc / totalDisc : 0;
      return {
        ...r,
        commission, customerGst, nov, pg, gstOnFees, adSpend, totalDed, payout,
        realRev, foodCost, netProfit, dedPctCV, earnPctCV, roas, netMargin,
        totalDisc, platDiscPctGMV, restDiscPctGMV, totalDiscPctGMV, platDiscShare, restDiscShare,
        label: monthLabel(r.month),
      };
    });
  }, [resRows, assumptions]);

  const totals = useMemo(() => {
    return monthly.reduce((acc, m) => ({
      gmv: acc.gmv + m.gmv,
      cv: acc.cv + m.cv,
      orders: acc.orders + m.orders,
      billed: acc.billed + m.billed,
      adSpend: acc.adSpend + m.adSpend,
      payout: acc.payout + m.payout,
      netProfit: acc.netProfit + m.netProfit,
      realRev: acc.realRev + m.realRev,
      foodCost: acc.foodCost + m.foodCost,
      platDisc: acc.platDisc + m.platDisc,
      restDisc: acc.restDisc + m.restDisc,
      commission: acc.commission + m.commission,
      pg: acc.pg + m.pg,
      gstOnFees: acc.gstOnFees + m.gstOnFees,
      totalDed: acc.totalDed + m.totalDed,
      totalDisc: acc.totalDisc + m.totalDisc,
    }), { gmv:0, cv:0, orders:0, billed:0, adSpend:0, payout:0, netProfit:0, realRev:0, foodCost:0, platDisc:0, restDisc:0, commission:0, pg:0, gstOnFees:0, totalDed:0, totalDisc:0 });
  }, [monthly]);

  const totalsRoas         = totals.billed ? totals.cv / totals.billed : 0;
  const totalsNetMargin    = totals.realRev ? totals.netProfit / totals.realRev : 0;
  const totalsAOV          = totals.orders ? totals.gmv / totals.orders : 0;
  const totalsDedPct       = totals.cv ? totals.totalDed / totals.cv : 0;
  const totalsEarnPct      = totals.cv ? totals.payout / totals.cv : 0;
  const totalsPlatDiscPctGMV    = totals.gmv ? totals.platDisc / totals.gmv : 0;
  const totalsRestDiscPctGMV    = totals.gmv ? totals.restDisc / totals.gmv : 0;
  const totalsDiscPctGMV   = totals.gmv ? totals.totalDisc / totals.gmv : 0;
  const totalsPlatDiscShare     = totals.totalDisc ? totals.platDisc / totals.totalDisc : 0;
  const totalsRestDiscShare     = totals.totalDisc ? totals.restDisc / totals.totalDisc : 0;

  // ── RE-PRICING IMPACT ────────────────────────────────────────────────
  // Model: raise menu prices by priceAdj. Orders stay constant.
  //   GMV and both discounts scale (they're % of price). Packaging stays per order.
  //   Ad spend stays absolute. Food cost stays absolute (same items cooked).
  //   Recompute the full P&L waterfall on the new CV.
  const adjusted = useMemo(() => {
    const k = 1 + priceAdj;
    // Net sales = GMV - platform discount - restaurant discount + packaging.
    // Only the menu-price part scales with a price change; packaging does not.
    const packTot = totals.cv && totals.gmv ? Math.max(totals.cv - (totals.gmv - totals.platDisc - totals.restDisc), 0) : 0;
    const cvScalable   = totals.cv - packTot;
    const newCv        = cvScalable * k + packTot;
    const newGmv       = totals.gmv * k;
    const newPlatDisc       = totals.platDisc * k;
    const newRestDisc       = totals.restDisc * k;
    const newCustGst   = newCv * 0.05;
    const newNov       = newCv + newCustGst;
    const newCommission= newCv * assumptions.commission;
    const newPg        = newNov * assumptions.pgRate;
    const newGstOnZom  = (newCommission + newPg) * assumptions.gst;
    const newAdSpend   = totals.adSpend;     // unchanged
    const newTotalDed  = newCommission + newPg + newGstOnZom + newAdSpend;
    const newPayout    = newCv - newTotalDed;
    const newRealRev   = totals.realRev * k;
    const newFoodCost  = totals.foodCost;    // unchanged - same orders, same items cooked
    const newNetProfit = newPayout - newFoodCost;
    const newNetMargin = newRealRev ? newNetProfit / newRealRev : 0;
    return {
      gmv: newGmv, cv: newCv, platDisc: newPlatDisc, restDisc: newRestDisc,
      commission: newCommission, pg: newPg, payout: newPayout,
      adSpend: newAdSpend, totalDed: newTotalDed,
      realRev: newRealRev, foodCost: newFoodCost,
      netProfit: newNetProfit, netMargin: newNetMargin,
      dedPctCV: newCv ? newTotalDed / newCv : 0,
      earnPctCV: newCv ? newPayout / newCv : 0,
    };
  }, [priceAdj, totals, assumptions]);

  return (
    <div style={{ maxWidth: 1400, margin: '0 auto', padding: isMobile ? '20px 16px' : '32px' }}>
      {/* Restaurant header */}
      <div style={{ marginBottom: 28 }}>
        <div style={{ fontSize: 11, letterSpacing: '0.18em', textTransform: 'uppercase', color: C.accent, marginBottom: 6 }}>
          {isChain ? 'Chain Deep Dive' : 'Restaurant Deep Dive'}
        </div>
        <h1 className="display" style={{
          fontFamily: "'Fraunces', serif",
          fontSize: isMobile ? 32 : 44, fontWeight: 400, letterSpacing: '-0.02em',
          margin: '0 0 10px', color: C.text, lineHeight: 1.1,
        }}>
          {resName}
        </h1>
        <div style={{ display: 'flex', gap: 12, color: C.textDim, fontSize: 13, flexWrap: 'wrap', alignItems: 'center' }}>
          {selectedArea && (
            <Pill icon={<MapPin size={12} />}>
              {selOutlet === '__all__' && isChain
                ? <span style={{ color: C.violet, fontWeight: 500 }}>{selectedArea}</span>
                : selectedArea}
            </Pill>
          )}
          {meta.cuisine && <Pill icon={<Building2 size={12} />}><span style={{ textTransform: 'capitalize' }}>{meta.cuisine}</span></Pill>}
          {monthly.length > 0 && <Pill icon={<Calendar size={12} />}>{monthly[0].label} → {monthly[monthly.length-1].label}</Pill>}
        </div>

        {/* Outlet selector, only for chains */}
        {isChain && (
          <div style={{
            marginTop: 16,
            display: 'flex',
            alignItems: 'center',
            gap: 12,
            flexWrap: 'wrap',
            padding: isMobile ? 14 : '14px 18px',
            background: `${C.violet}08`,
            border: `1px solid ${C.violet}25`,
            borderRadius: 12,
          }}>
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              fontSize: 12, color: C.violet, fontWeight: 600,
              letterSpacing: '0.04em', textTransform: 'uppercase',
            }}>
              <Building2 size={14} />
              Viewing
            </div>
            <select
              value={selOutlet}
              onChange={(e) => setSelOutlet(e.target.value)}
              style={{
                flex: isMobile ? '1 1 100%' : '0 1 auto',
                minWidth: 220,
                background: C.surface,
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                padding: '10px 14px',
                color: C.text,
                fontSize: 13,
                fontWeight: 500,
                cursor: 'pointer',
                outline: 'none',
                fontFamily: 'inherit',
              }}
            >
              <option value="__all__">All outlets (brand) · {outlets.length} locations</option>
              {outlets.map(o => {
                const key = o.outletId || o.area || 'unknown';
                return (
                  <option key={key} value={key}>
                    {o.outletId ? `${o.outletId} · ${o.area || 'Unknown'}` : (o.area || 'Unknown')}
                  </option>
                );
              })}
            </select>
            <div style={{ fontSize: 12, color: C.textMute }}>
              {selOutlet === '__all__'
                ? 'Showing combined performance across all outlets'
                : 'Showing one outlet'}
            </div>
          </div>
        )}
      </div>

      {/* KPI row */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2, 1fr)' : 'repeat(5, 1fr)', gap: isMobile ? 10 : 14, marginBottom: 28 }}>
        <KPICard label="Total GMV"        value={fmtINR(totals.gmv)}       sub={`${fmtNum(totals.orders)} orders · AOV ${fmtINR(totalsAOV)}`} accent={C.accent} />
        <KPICard label="Payout"           value={fmtINR(totals.payout)}    sub={`${fmtPct(totalsEarnPct)} of net sales`} accent={C.success} />
        <KPICard label="Net Profit"       value={fmtINR(totals.netProfit)} sub={`${fmtPct(totalsNetMargin)} margin`} accent={totals.netProfit >= 0 ? C.success : C.danger} />
        <KPICard label="Ad Spend (w/GST)" value={fmtINR(totals.adSpend)}   sub={`RoAS ${fmtMul(totalsRoas)}`} accent={C.violet} />
        <KPICard label="Total Deductions" value={fmtPct(totalsDedPct)}     sub={`of net sales (${fmtINR(totals.totalDed)})`} accent={C.rose} />
      </div>

      {/* Charts grid */}
      <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(2, 1fr)', gap: 14, marginBottom: 28 }}>
        <Card title="GMV Trend">
          <ResponsiveContainer width="100%" height={220}>
            <AreaChart data={monthly} margin={{ top: 16, right: 12, left: -12, bottom: 0 }}>
              <defs>
                <linearGradient id="gmvA" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%"   stopColor={C.accent} stopOpacity={0.2} />
                  <stop offset="100%" stopColor={C.accent} stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={fmtINR} />
              <Tooltip content={<CustomTooltip />} />
              <Area type="monotone" dataKey="gmv" stroke={C.accent} strokeWidth={2} fill="url(#gmvA)" name="GMV" />
            </AreaChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Orders Trend">
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={monthly} margin={{ top: 16, right: 12, left: -12, bottom: 0 }}>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} />
              <Tooltip content={<CustomTooltip />} />
              <Bar dataKey="orders" fill={C.info} fillOpacity={0.9} name="Orders" radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card title="Ad Spend (billed, ex-GST)">
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={monthly} margin={{ top: 16, right: 12, left: -12, bottom: 0 }}>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={fmtINR} />
              <Tooltip content={<CustomTooltip />} />
              <Bar dataKey="billed" fill={C.violet} fillOpacity={0.9} name="Ad Spend (billed)" radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card title="RoAS Trend" subtitle="Net sales per ₹ of ad spend · 5x+ is healthy">
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={monthly} margin={{ top: 16, right: 12, left: -12, bottom: 0 }}>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={(v) => `${v.toFixed(0)}x`} />
              <Tooltip content={<CustomTooltip />} />
              <ReferenceLine y={5} stroke={C.success} strokeDasharray="3 3" strokeOpacity={0.5} />
              <Line type="monotone" dataKey="roas" stroke={C.success} strokeWidth={2.5} dot={{ fill: C.success, r: 3 }} name="RoAS" />
            </LineChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* ── RE-PRICING IMPACT (Price Lens · live) ─────────────────────────────── */}
      <div style={{ marginBottom: 14 }}>
        <div style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 14,
          padding: isMobile ? 18 : 24,
          boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
          position: 'relative',
          overflow: 'hidden',
        }}>
          {/* Accent strip */}
          <div style={{ position: 'absolute', top: 0, left: 0, width: 50, height: 3, background: C.success }} />

          <div style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'flex-start',
            marginBottom: 18,
            gap: 12,
            flexWrap: 'wrap',
          }}>
            <div>
              <div style={{ fontSize: 10, letterSpacing: '0.18em', textTransform: 'uppercase', color: C.success, marginBottom: 6, fontWeight: 600 }}>
                Price Lens · Re-pricing Impact
              </div>
              <div className="display" style={{
                fontFamily: "'Fraunces', serif",
                fontSize: isMobile ? 22 : 26,
                fontWeight: 500,
                letterSpacing: '-0.01em',
                color: C.text,
                marginBottom: 4,
              }}>
                What if you adjusted menu prices?
              </div>
              <div style={{ fontSize: 12, color: C.textMute, lineHeight: 1.5, maxWidth: 600 }}>
                Live simulation on the actuals above. Assumes orders stay constant, with ad spend and food cost unchanged.
              </div>
            </div>
          </div>

          {/* Adjustment buttons + custom input */}
          <div style={{
            display: 'flex',
            gap: 8,
            marginBottom: 22,
            flexWrap: 'wrap',
            alignItems: 'center',
          }}>
            {[-0.10, -0.05, 0, 0.05, 0.10, 0.15, 0.20].map(v => {
              const active = Math.abs(priceAdj - v) < 0.001;
              return (
                <button
                  key={v}
                  onClick={() => setPriceAdj(v)}
                  style={{
                    padding: '7px 12px',
                    background: active ? C.text : C.surface,
                    border: `1px solid ${active ? C.text : C.border}`,
                    borderRadius: 100,
                    color: active ? C.surface : C.text,
                    cursor: 'pointer',
                    fontSize: 12,
                    fontWeight: 500,
                    fontFamily: "'JetBrains Mono', monospace",
                    transition: 'all 120ms ease',
                  }}
                >
                  {v > 0 ? '+' : ''}{(v * 100).toFixed(0)}%
                </button>
              );
            })}
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              marginLeft: isMobile ? 0 : 6,
              padding: '4px 10px 4px 12px',
              background: C.surface,
              border: `1px solid ${C.border}`,
              borderRadius: 100,
            }}>
              <span style={{ fontSize: 11, color: C.textDim }}>Custom</span>
              <input
                type="number"
                step="0.5"
                value={(priceAdj * 100).toFixed(1)}
                onChange={(e) => {
                  const v = parseFloat(e.target.value);
                  setPriceAdj(isNaN(v) ? 0 : v / 100);
                }}
                style={{
                  width: 50,
                  background: 'transparent',
                  border: 'none',
                  outline: 'none',
                  color: C.text,
                  fontSize: 13,
                  fontFamily: "'JetBrains Mono', monospace",
                  fontVariantNumeric: 'tabular-nums',
                  textAlign: 'right',
                }}
              />
              <span style={{ fontSize: 12, color: C.textMute }}>%</span>
            </div>
          </div>

          {/* Comparison table */}
          <div style={{
            background: C.bg,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            overflow: 'hidden',
          }}>
            {/* Header row */}
            <div style={{
              display: 'grid',
              gridTemplateColumns: isMobile ? '1.2fr 1fr 1.1fr' : '1.4fr 1fr 1fr 1.2fr',
              padding: '12px 16px',
              background: C.surface,
              borderBottom: `1px solid ${C.border}`,
              fontSize: 10,
              letterSpacing: '0.14em',
              textTransform: 'uppercase',
              color: C.textDim,
              fontWeight: 600,
            }}>
              <div>Metric</div>
              {!isMobile && <div style={{ textAlign: 'right' }}>Current</div>}
              <div style={{ textAlign: 'right' }}>Adjusted</div>
              <div style={{ textAlign: 'right' }}>Δ Impact</div>
            </div>

            {/* Rows */}
            {[
              { label: 'Gross GMV',         curr: totals.gmv,        adj: adjusted.gmv,        fmt: fmtINR, accent: false },
              { label: 'Net Sales',         curr: totals.cv,         adj: adjusted.cv,         fmt: fmtINR, accent: false },
              { label: 'Total Discounts',   curr: totals.totalDisc,  adj: adjusted.platDisc + adjusted.restDisc, fmt: fmtINR, accent: false, neg: true },
              { label: 'Payout Receivable', curr: totals.payout,     adj: adjusted.payout,     fmt: fmtINR, accent: true,  bold: true },
              { label: 'Net Online Profit', curr: totals.netProfit,  adj: adjusted.netProfit,  fmt: fmtINR, accent: true,  bold: true },
            ].map((row, i) => {
              const delta = row.adj - row.curr;
              const pctChg = row.curr ? delta / row.curr : 0;
              const deltaColor = row.neg
                ? (delta <= 0 ? C.success : C.danger)   // for costs, lower is better
                : (delta >= 0 ? C.success : C.danger);  // for revenue/profit, higher is better
              return (
                <div key={i} style={{
                  display: 'grid',
                  gridTemplateColumns: isMobile ? '1.2fr 1fr 1.1fr' : '1.4fr 1fr 1fr 1.2fr',
                  padding: isMobile ? '11px 14px' : '12px 16px',
                  borderBottom: i < 4 ? `1px solid ${C.border}` : 'none',
                  fontSize: isMobile ? 12 : 13,
                  alignItems: 'center',
                  background: row.accent ? `${C.success}06` : 'transparent',
                }}>
                  <div style={{
                    color: row.accent ? C.success : C.text,
                    fontWeight: row.bold ? 600 : 400,
                  }}>
                    {row.label}
                  </div>
                  {!isMobile && (
                    <div className="num" style={{
                      textAlign: 'right',
                      color: C.textDim,
                      fontWeight: row.bold ? 500 : 400,
                    }}>
                      {row.fmt(row.curr)}
                    </div>
                  )}
                  <div className="num" style={{
                    textAlign: 'right',
                    color: row.accent ? C.success : C.text,
                    fontWeight: row.bold ? 600 : 500,
                  }}>
                    {row.fmt(row.adj)}
                  </div>
                  <div className="num" style={{
                    textAlign: 'right',
                    color: deltaColor,
                    fontWeight: 600,
                    display: 'flex',
                    justifyContent: 'flex-end',
                    alignItems: 'center',
                    gap: 4,
                  }}>
                    {delta >= 0 ? <TrendingUp size={12} /> : <TrendingDown size={12} />}
                    <span>{delta >= 0 ? '+' : '−'}{row.fmt(Math.abs(delta))}</span>
                    <span style={{
                      fontSize: 10,
                      color: `${deltaColor}CC`,
                      fontWeight: 500,
                      marginLeft: 2,
                    }}>
                      ({delta >= 0 ? '+' : '−'}{(Math.abs(pctChg) * 100).toFixed(1)}%)
                    </span>
                  </div>
                </div>
              );
            })}

            {/* Margin row (special: uses pp not %) */}
            <div style={{
              display: 'grid',
              gridTemplateColumns: isMobile ? '1.2fr 1fr 1.1fr' : '1.4fr 1fr 1fr 1.2fr',
              padding: isMobile ? '11px 14px' : '12px 16px',
              borderTop: `1px solid ${C.border}`,
              fontSize: isMobile ? 12 : 13,
              alignItems: 'center',
              background: `${C.success}06`,
            }}>
              <div style={{ color: C.success, fontWeight: 600 }}>Net Margin %</div>
              {!isMobile && (
                <div className="num" style={{ textAlign: 'right', color: C.textDim, fontWeight: 500 }}>
                  {fmtPct(totalsNetMargin)}
                </div>
              )}
              <div className="num" style={{ textAlign: 'right', color: C.success, fontWeight: 600 }}>
                {fmtPct(adjusted.netMargin)}
              </div>
              <div className="num" style={{
                textAlign: 'right',
                color: adjusted.netMargin >= totalsNetMargin ? C.success : C.danger,
                fontWeight: 600,
                display: 'flex',
                justifyContent: 'flex-end',
                alignItems: 'center',
                gap: 4,
              }}>
                {adjusted.netMargin >= totalsNetMargin ? <TrendingUp size={12} /> : <TrendingDown size={12} />}
                <span>{adjusted.netMargin >= totalsNetMargin ? '+' : '−'}{(Math.abs(adjusted.netMargin - totalsNetMargin) * 100).toFixed(2)} pp</span>
              </div>
            </div>
          </div>

          {/* Insight tile */}
          {priceAdj !== 0 && (() => {
            const profitDelta = adjusted.netProfit - totals.netProfit;
            const orientation = priceAdj > 0 ? 'lift' : 'cut';
            const profitDirection = profitDelta >= 0 ? 'add' : 'cost';
            const pctText = `${priceAdj > 0 ? '+' : '−'}${(Math.abs(priceAdj) * 100).toFixed(1)}%`;
            return (
              <div style={{
                marginTop: 16,
                padding: '14px 18px',
                background: profitDelta >= 0 ? `${C.success}10` : `${C.danger}10`,
                border: `1px solid ${profitDelta >= 0 ? `${C.success}30` : `${C.danger}30`}`,
                borderRadius: 10,
                display: 'flex',
                alignItems: 'flex-start',
                gap: 12,
              }}>
                <div style={{
                  flexShrink: 0,
                  width: 28, height: 28,
                  borderRadius: 8,
                  background: profitDelta >= 0 ? `${C.success}25` : `${C.danger}25`,
                  color: profitDelta >= 0 ? C.success : C.danger,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                }}>
                  {profitDelta >= 0 ? <TrendingUp size={15} /> : <TrendingDown size={15} />}
                </div>
                <div style={{ fontSize: 13, color: C.text, lineHeight: 1.55 }}>
                  A <strong className="num">{pctText}</strong> price {orientation} would {profitDirection}{' '}
                  <strong className="num" style={{ color: profitDelta >= 0 ? C.success : C.danger }}>
                    {fmtINRFull(Math.abs(profitDelta))}
                  </strong>{' '}
                  to your net profit across the period
                  {totals.netProfit > 0 && profitDelta > 0 && (
                    <>, a <strong className="num">+{((profitDelta / Math.abs(totals.netProfit)) * 100).toFixed(1)}%</strong> swing</>
                  )}
                  . Food cost and ad spend stay the same, so almost all the extra revenue flows to the bottom line.
                </div>
              </div>
            );
          })()}
        </div>
      </div>

      {/* Discounts breakdown */}
      <div style={{ marginBottom: 14 }}>
        <Card title="Discounts Breakdown" subtitle="Who paid for each discount: the platform, or the restaurant itself">
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr 1fr', gap: 14, marginBottom: 16 }}>
            <div style={{ padding: 16, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <div style={{ width: 8, height: 8, borderRadius: 2, background: C.info }} />
                <div style={{ fontSize: 10, letterSpacing: '0.14em', textTransform: 'uppercase', color: C.textDim }}>Platform-funded</div>
              </div>
              <div className="num" style={{ fontSize: 22, fontWeight: 500, color: C.text, marginBottom: 4 }}>
                {fmtINR(totals.platDisc)}
              </div>
              <div style={{ fontSize: 12, color: C.textMute }}>
                <span className="num">{fmtPct(totalsPlatDiscPctGMV)}</span> of GMV · <span className="num">{fmtPct(totalsPlatDiscShare)}</span> of total discount
              </div>
            </div>

            <div style={{ padding: 16, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <div style={{ width: 8, height: 8, borderRadius: 2, background: C.accent }} />
                <div style={{ fontSize: 10, letterSpacing: '0.14em', textTransform: 'uppercase', color: C.textDim }}>Restaurant-funded</div>
              </div>
              <div className="num" style={{ fontSize: 22, fontWeight: 500, color: C.text, marginBottom: 4 }}>
                {fmtINR(totals.restDisc)}
              </div>
              <div style={{ fontSize: 12, color: C.textMute }}>
                <span className="num">{fmtPct(totalsRestDiscPctGMV)}</span> of GMV · <span className="num">{fmtPct(totalsRestDiscShare)}</span> of total discount
              </div>
            </div>

            <div style={{ padding: 16, background: C.bg, border: `1px solid ${C.border}`, borderRadius: 10 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                <div style={{ width: 8, height: 8, borderRadius: 2, background: C.rose }} />
                <div style={{ fontSize: 10, letterSpacing: '0.14em', textTransform: 'uppercase', color: C.textDim }}>Total Discount Cost</div>
              </div>
              <div className="num" style={{ fontSize: 22, fontWeight: 500, color: C.text, marginBottom: 4 }}>
                {fmtINR(totals.totalDisc)}
              </div>
              <div style={{ fontSize: 12, color: C.textMute }}>
                <span className="num">{fmtPct(totalsDiscPctGMV)}</span> of GMV · ratio <span className="num">{totals.restDisc ? (totals.platDisc / totals.restDisc).toFixed(2) : '-'}x</span> platform : you
              </div>
            </div>
          </div>

          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={monthly} margin={{ top: 10, right: 12, left: -12, bottom: 0 }}>
              <CartesianGrid stroke={C.border} strokeDasharray="2 4" vertical={false} />
              <XAxis dataKey="label" stroke={C.textMute} fontSize={11} tickLine={false} axisLine={{ stroke: C.border }} />
              <YAxis stroke={C.textMute} fontSize={10} tickLine={false} axisLine={false} tickFormatter={fmtINR} />
              <Tooltip content={<CustomTooltip />} />
              <Legend verticalAlign="top" iconType="circle" wrapperStyle={{ fontSize: 11, color: C.textDim, paddingBottom: 8 }} />
              <Bar dataKey="platDisc" stackId="d" fill={C.info}   name="Platform-funded" radius={[0,0,0,0]} />
              <Bar dataKey="restDisc" stackId="d" fill={C.accent} name="Restaurant-funded"   radius={[4,4,0,0]} />
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>

      {/* Assumptions panel */}
      <Card title="Assumptions" subtitle="Edit these to recalculate the P&L">
        <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2, 1fr)' : 'repeat(5, 1fr)', gap: isMobile ? 10 : 14 }}>
          <NumInput label="Menu Hike %"     value={assumptions.hike}       onChange={(v) => setAssumptions(a => ({...a, hike: v}))}       pct />
          <NumInput label="Commission %"    value={assumptions.commission} onChange={(v) => setAssumptions(a => ({...a, commission: v}))} pct />
          <NumInput label="PG Rate"         value={assumptions.pgRate}     onChange={(v) => setAssumptions(a => ({...a, pgRate: v}))}     pct />
          <NumInput label="GST"             value={assumptions.gst}        onChange={(v) => setAssumptions(a => ({...a, gst: v}))}        pct />
          <NumInput label="Food Cost %"     value={assumptions.foodCost}   onChange={(v) => setAssumptions(a => ({...a, foodCost: v}))}   pct />
        </div>
      </Card>

      {/* P&L Table */}
      <div style={{ marginTop: 14 }}>
        <Card title="Monthly P&L" subtitle="Numbers update live with assumptions">
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, minWidth: 900 }}>
              <thead>
                <tr style={{ color: C.textDim, fontSize: 10, letterSpacing: '0.1em', textTransform: 'uppercase' }}>
                  <th style={{ ...th(), textAlign: 'left' }}>Metric</th>
                  {monthly.map(m => <th key={m.month} style={{ ...th(), textAlign: 'right' }}>{m.label}</th>)}
                  <th style={{ ...th(), textAlign: 'right', color: C.accent }}>Total</th>
                </tr>
              </thead>
              <tbody>
                <PLRow label="Gross GMV"       data={monthly} field="gmv"        total={totals.gmv} fmt={fmtINR} />
                <PLRow label="Orders"          data={monthly} field="orders"     total={totals.orders} fmt={fmtNum} dim />
                <PLRow label="Net Sales"       data={monthly} field="cv"         total={totals.cv} fmt={fmtINR} dim />
                <PLSep label="Discounts  (on GMV)" />
                <PLRow label="Platform-funded" data={monthly} field="platDisc" total={totals.platDisc} fmt={fmtINR}
                       pctField="platDiscPctGMV" totalPct={totalsPlatDiscPctGMV} />
                <PLRow label="Restaurant-funded" data={monthly} field="restDisc" total={totals.restDisc} fmt={fmtINR}
                       pctField="restDiscPctGMV" totalPct={totalsRestDiscPctGMV} />
                <PLRow label="Total Discount" data={monthly} field="totalDisc" total={totals.totalDisc} fmt={fmtINR}
                       pctField="totalDiscPctGMV" totalPct={totalsDiscPctGMV} bold />
                <PLSep label="Platform Deductions" />
                <PLRow label="(−) Commission"  data={monthly} field="commission" total={totals.commission} fmt={fmtINR} negative />
                <PLRow label="(−) PG Charges"  data={monthly} field="pg"         total={totals.pg} fmt={fmtINR} negative />
                <PLRow label="(−) GST on Comm+PG" data={monthly} field="gstOnFees" total={totals.gstOnFees} fmt={fmtINR} negative />
                <PLRow label="(−) Ad Spend (w/GST)" data={monthly} field="adSpend" total={totals.adSpend} fmt={fmtINR} negative />
                <PLRow label="Total Deductions" data={monthly} field="totalDed"  total={totals.totalDed} fmt={fmtINR} negative bold
                       pctField="dedPctCV" totalPct={totalsDedPct} />
                <PLRow label="Payout Receivable" data={monthly} field="payout"   total={totals.payout} fmt={fmtINR} bold success
                       pctField="earnPctCV" totalPct={totalsEarnPct} />
                <PLSep label="Restaurant Economics" />
                <PLRow label="Real Revenue (ex hike)" data={monthly} field="realRev" total={totals.realRev} fmt={fmtINR} dim />
                <PLRow label="(−) Food Cost"   data={monthly} field="foodCost"   total={totals.foodCost} fmt={fmtINR} negative />
                <PLRow label="Net Online Profit" data={monthly} field="netProfit" total={totals.netProfit} fmt={fmtINR} bold success
                       pctField="netMargin" totalPct={totalsNetMargin} />
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SIMULATOR VIEW: payout calculator
// ─────────────────────────────────────────────────────────────────────────────
const SIM_PRESETS = {
  base: {
    label: 'Base Case',
    desc: 'Typical delivery restaurant',
    direction: 'on-to-off',
    onlineVal: 200000, offlineVal: 0,
    hike: 0.30,
    promo: 15000,   promoType: 'rupees',
    salt: 0,        saltType: 'rupees',
    packaging: 1500,
    commission: 0.20,
    pgRate: 0.02,
    gstCollected: 0.05,
    gstOnComm: 0.18,
    gstOnAds: 0.18,
    adSpend: 8,     adSpendType: 'percent',
  },
  noAds: {
    label: 'No Ads',
    desc: 'Same setup, zero ad spend',
    direction: 'on-to-off',
    onlineVal: 200000, offlineVal: 0,
    hike: 0.30,
    promo: 15000,   promoType: 'rupees',
    salt: 0,        saltType: 'rupees',
    packaging: 1500,
    commission: 0.20,
    pgRate: 0.02,
    gstCollected: 0.05,
    gstOnComm: 0.18,
    gstOnAds: 0.18,
    adSpend: 0,     adSpendType: 'percent',
  },
  highHike: {
    label: 'Aggressive Hike',
    desc: '40% hike, see what it does',
    direction: 'on-to-off',
    onlineVal: 200000, offlineVal: 0,
    hike: 0.40,
    promo: 15000,   promoType: 'rupees',
    salt: 0,        saltType: 'rupees',
    packaging: 1500,
    commission: 0.20,
    pgRate: 0.02,
    gstCollected: 0.05,
    gstOnComm: 0.18,
    gstOnAds: 0.18,
    adSpend: 8,     adSpendType: 'percent',
  },
  heavyDisc: {
    label: 'Heavy Discount',
    desc: '25% promo discount',
    direction: 'on-to-off',
    onlineVal: 200000, offlineVal: 0,
    hike: 0.30,
    promo: 25,      promoType: 'percent',
    salt: 0,        saltType: 'rupees',
    packaging: 1500,
    commission: 0.20,
    pgRate: 0.02,
    gstCollected: 0.05,
    gstOnComm: 0.18,
    gstOnAds: 0.18,
    adSpend: 8,     adSpendType: 'percent',
  },
};

function computeSim(inp) {
  const { direction, onlineVal, offlineVal, hike } = inp;
  let onlinePrice, offlinePrice;
  if (direction === 'on-to-off') {
    onlinePrice  = onlineVal;
    offlinePrice = hike > -1 ? onlineVal / (1 + hike) : 0;
  } else {
    offlinePrice = offlineVal;
    onlinePrice  = offlineVal * (1 + hike);
  }
  const priceIncrease = onlinePrice - offlinePrice;

  const promoAmt = inp.promoType === 'percent' ? onlinePrice * inp.promo / 100 : inp.promo;
  const saltAmt  = inp.saltType  === 'percent' ? onlinePrice * inp.salt  / 100 : inp.salt;
  const totalDisc = promoAmt + saltAmt;

  const cv  = onlinePrice - promoAmt - saltAmt + inp.packaging;
  const customerGst = cv * inp.gstCollected;
  const nov = cv + customerGst;

  const commission = cv * inp.commission;
  const pg         = nov * inp.pgRate;
  const gstOnComm  = commission * inp.gstOnComm;
  const gstOnPg    = pg * inp.gstOnComm;

  const adAmt    = inp.adSpendType === 'percent' ? cv * inp.adSpend / 100 : inp.adSpend;
  const gstOnAds = adAmt * inp.gstOnAds;

  const totalPlatform  = commission + pg + gstOnComm + gstOnPg;
  const totalMarketing = adAmt + gstOnAds;
  const totalTaxes     = customerGst + gstOnComm + gstOnPg + gstOnAds;
  const totalCosts     = totalPlatform + totalMarketing;

  const payout = cv - commission - pg - gstOnComm - gstOnPg - adAmt - gstOnAds;

  const payoutPctOffline = offlinePrice ? payout / offlinePrice : 0;
  const payoutPctOnline  = onlinePrice  ? payout / onlinePrice  : 0;
  const effCommPct       = onlinePrice  ? (commission + pg) / onlinePrice : 0;
  const effCostPct       = onlinePrice  ? (onlinePrice - payout) / onlinePrice : 0;

  return {
    onlinePrice, offlinePrice, priceIncrease,
    promoAmt, saltAmt, totalDisc,
    cv, customerGst, nov,
    commission, pg, gstOnComm, gstOnPg,
    adAmt, gstOnAds,
    totalPlatform, totalMarketing, totalTaxes, totalCosts,
    payout, payoutPctOffline, payoutPctOnline, effCommPct, effCostPct,
  };
}

function SimulatorView({ isMobile }) {
  const [activePreset, setActivePreset] = useState('base');
  const [inp, setInp] = useState(SIM_PRESETS.base);

  const set = (key, val) => {
    setInp(prev => ({ ...prev, [key]: val }));
    setActivePreset(null);
  };
  const loadPreset = (key) => {
    setInp(SIM_PRESETS[key]);
    setActivePreset(key);
  };
  const reset = () => loadPreset('base');

  // ── CALCULATIONS ─────────────────────────────────────────────────
  const calc = useMemo(() => computeSim(inp), [inp]);

  // Bar fill colour based on payout health
  const payoutHealth = calc.payoutPctOffline >= 0.85 ? C.success
                     : calc.payoutPctOffline >= 0.70 ? C.accent
                     : C.danger;

  return (
    <div style={{ maxWidth: 1400, margin: '0 auto', padding: isMobile ? '20px 16px' : '32px' }}>
      {/* ── HERO ───────────────────────────────────────────────────── */}
      <div style={{ marginBottom: isMobile ? 20 : 28 }}>
        <div style={{ fontSize: 11, letterSpacing: '0.18em', textTransform: 'uppercase', color: C.success, marginBottom: 6, fontWeight: 600 }}>
          Price Lens
        </div>
        <h1 className="display" style={{
          fontFamily: "'Fraunces', serif",
          fontSize: isMobile ? 28 : 40, fontWeight: 400, letterSpacing: '-0.02em',
          margin: '0 0 8px', color: C.text, lineHeight: 1.1,
        }}>
          Where every <em style={{ color: C.accent, fontStyle: 'italic' }}>rupee</em> of an order goes.
        </h1>
        <p style={{ color: C.textDim, fontSize: isMobile ? 13 : 14, lineHeight: 1.55, margin: 0, maxWidth: 720 }}>
          Tweak any input (pricing, discount, commission, ads) and watch the full payout waterfall update instantly. Built for what-if conversations with restaurant owners.
        </p>
      </div>

      {/* ── PRESETS ────────────────────────────────────────────────── */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 20, overflowX: 'auto', paddingBottom: 4 }}>
        {Object.entries(SIM_PRESETS).map(([key, p]) => (
          <button
            key={key}
            onClick={() => loadPreset(key)}
            style={{
              flexShrink: 0,
              padding: '8px 14px',
              background: activePreset === key ? C.text : C.card,
              color: activePreset === key ? C.surface : C.text,
              border: `1px solid ${activePreset === key ? C.text : C.border}`,
              borderRadius: 100,
              cursor: 'pointer',
              fontSize: 12,
              fontWeight: 500,
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              transition: 'all 120ms ease',
            }}
          >
            {p.label}
            {!isMobile && (
              <span style={{ fontSize: 10, color: activePreset === key ? C.surface + 'AA' : C.textMute, fontWeight: 400 }}>
                {p.desc}
              </span>
            )}
          </button>
        ))}
        <button
          onClick={reset}
          style={{
            flexShrink: 0,
            padding: '8px 12px',
            background: 'transparent',
            border: `1px solid ${C.border}`,
            borderRadius: 100,
            color: C.textDim,
            cursor: 'pointer',
            fontSize: 12,
            display: 'flex',
            alignItems: 'center',
            gap: 6,
          }}
        >
          <RefreshCw size={11} /> Reset
        </button>
      </div>

      {/* ── HERO METRIC CARD ───────────────────────────────────────── */}
      <div style={{
        background: C.card,
        border: `1px solid ${C.border}`,
        borderRadius: 14,
        padding: isMobile ? 20 : 28,
        marginBottom: 14,
        boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
        position: 'relative',
        overflow: 'hidden',
      }}>
        <div style={{
          display: 'grid',
          gridTemplateColumns: isMobile ? '1fr' : '1.4fr 1fr',
          gap: isMobile ? 20 : 32,
          alignItems: 'center',
        }}>
          {/* Big number */}
          <div>
            <div style={{ fontSize: 10, letterSpacing: '0.16em', textTransform: 'uppercase', color: C.textDim, marginBottom: 8 }}>
              Final Payout Received
            </div>
            <div className="num display" style={{
              fontFamily: "'Fraunces', serif",
              fontSize: isMobile ? 44 : 64,
              fontWeight: 500,
              color: C.success,
              lineHeight: 1,
              letterSpacing: '-0.02em',
              marginBottom: 10,
            }}>
              {fmtINRFull(calc.payout)}
            </div>

            {/* Absolute delta vs offline */}
            {(() => {
              const delta = calc.payout - calc.offlinePrice;
              const positive = delta >= 0;
              const deltaColor = positive ? C.success : C.danger;
              return (
                <div style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 6,
                  padding: '5px 12px',
                  background: `${deltaColor}12`,
                  border: `1px solid ${deltaColor}30`,
                  borderRadius: 100,
                  fontSize: 12,
                  fontWeight: 500,
                  color: deltaColor,
                  marginBottom: 14,
                }}>
                  {positive ? <TrendingUp size={13} /> : <TrendingDown size={13} />}
                  <span className="num">
                    {positive ? '+' : '−'}{fmtINRFull(Math.abs(delta))}
                  </span>
                  <span style={{ color: `${deltaColor}AA`, fontWeight: 400 }}>vs offline price</span>
                </div>
              );
            })()}

            {/* Comparison bar */}
            <div style={{ marginTop: 6 }}>
              <div style={{
                display: 'flex',
                justifyContent: 'space-between',
                fontSize: 11,
                color: C.textDim,
                marginBottom: 6,
              }}>
                <span>Offline equivalent</span>
                <span className="num" style={{ color: C.text }}>{fmtINRFull(calc.offlinePrice)}</span>
              </div>
              <div style={{
                position: 'relative',
                height: 12,
                background: `${C.border}`,
                borderRadius: 100,
                overflow: 'hidden',
              }}>
                <div style={{
                  position: 'absolute',
                  left: 0, top: 0, bottom: 0,
                  width: `${Math.min(Math.max(calc.payoutPctOffline * 100, 0), 100)}%`,
                  background: `linear-gradient(90deg, ${payoutHealth} 0%, ${payoutHealth}EE 100%)`,
                  borderRadius: 100,
                  transition: 'width 300ms ease, background 200ms ease',
                }} />
              </div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 6, fontSize: 11 }}>
                <span style={{ color: C.textMute }}>0%</span>
                <span className="num" style={{ color: payoutHealth, fontWeight: 600, fontSize: 13 }}>
                  {fmtPct(calc.payoutPctOffline)} of offline price
                </span>
                <span style={{ color: C.textMute }}>100%</span>
              </div>
            </div>
          </div>

          {/* Sub-KPI grid */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: isMobile ? 'repeat(2, 1fr)' : 'repeat(2, 1fr)',
            gap: 10,
          }}>
            <MiniStat label="Net Sales" value={fmtINR(calc.cv)} />
            <MiniStat label="Net Order Value" value={fmtINR(calc.nov)} />
            <MiniStat label="Platform Takes" value={fmtINR(calc.totalPlatform)} hue={C.danger} />
            <MiniStat label="Effective Cost" value={fmtPct(calc.effCostPct)} hue={C.violet} />
          </div>
        </div>
      </div>

      {/* ── INPUTS + BREAKDOWN ─────────────────────────────────────── */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: isMobile ? '1fr' : '1.4fr 1fr',
        gap: 14,
      }}>
        {/* LEFT: INPUTS */}
        <div style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 14,
          padding: isMobile ? 16 : 24,
          boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
        }}>
          <SectionLabel>Menu Pricing</SectionLabel>
          {/* Direction toggle */}
          <div style={{
            display: 'inline-flex',
            background: C.bg,
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            padding: 3,
            marginBottom: 14,
          }}>
            <button
              onClick={() => set('direction', 'off-to-on')}
              style={{
                padding: '6px 12px',
                background: inp.direction === 'off-to-on' ? C.surface : 'transparent',
                border: 'none',
                borderRadius: 6,
                color: inp.direction === 'off-to-on' ? C.text : C.textDim,
                cursor: 'pointer',
                fontSize: 12,
                fontWeight: 500,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              <ArrowRight size={11} /> Offline → Online
            </button>
            <button
              onClick={() => set('direction', 'on-to-off')}
              style={{
                padding: '6px 12px',
                background: inp.direction === 'on-to-off' ? C.surface : 'transparent',
                border: 'none',
                borderRadius: 6,
                color: inp.direction === 'on-to-off' ? C.text : C.textDim,
                cursor: 'pointer',
                fontSize: 12,
                fontWeight: 500,
                display: 'flex', alignItems: 'center', gap: 6,
              }}
            >
              <ArrowLeft size={11} /> Online → Offline
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(3, 1fr)', gap: 12, marginBottom: 22 }}>
            <SimInput
              label="Online Menu Price"
              value={inp.direction === 'on-to-off' ? inp.onlineVal : calc.onlinePrice}
              onChange={(v) => set('onlineVal', v)}
              prefix="₹"
              auto={inp.direction !== 'on-to-off'}
            />
            <SimInput
              label="Price Hike %"
              value={inp.hike * 100}
              onChange={(v) => set('hike', v / 100)}
              suffix="%"
              step="1"
            />
            <SimInput
              label="Offline Menu Price"
              value={inp.direction === 'off-to-on' ? inp.offlineVal : calc.offlinePrice}
              onChange={(v) => set('offlineVal', v)}
              prefix="₹"
              auto={inp.direction !== 'off-to-on'}
            />
          </div>

          <SectionLabel>Discounts</SectionLabel>
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(2, 1fr)', gap: 12, marginBottom: 22 }}>
            <SimInputDual
              label="Promo Discount"
              value={inp.promo}
              type={inp.promoType}
              onChangeValue={(v) => set('promo', v)}
              onChangeType={(t) => set('promoType', t)}
            />
            <SimInputDual
              label="Extra Discount"
              value={inp.salt}
              type={inp.saltType}
              onChangeValue={(v) => set('salt', v)}
              onChangeType={(t) => set('saltType', t)}
            />
          </div>

          <SectionLabel>Packaging & Commission</SectionLabel>
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr 1fr' : 'repeat(2, 1fr)', gap: 12, marginBottom: 22 }}>
            <SimInput
              label="Packaging Charge"
              value={inp.packaging}
              onChange={(v) => set('packaging', v)}
              prefix="₹"
            />
            <SimInput
              label="Base Commission %"
              value={inp.commission * 100}
              onChange={(v) => set('commission', v / 100)}
              suffix="%"
              step="0.1"
            />
          </div>

          <SectionLabel>GST & Charges</SectionLabel>
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? 'repeat(2, 1fr)' : 'repeat(4, 1fr)', gap: 12, marginBottom: 22 }}>
            <SimInput label="PG Charges %"     value={inp.pgRate * 100}       onChange={(v) => set('pgRate', v / 100)}       suffix="%" step="0.01" />
            <SimInput label="GST Collected %"  value={inp.gstCollected * 100} onChange={(v) => set('gstCollected', v / 100)} suffix="%" step="0.1" />
            <SimInput label="GST Comm/PG %"    value={inp.gstOnComm * 100}    onChange={(v) => set('gstOnComm', v / 100)}    suffix="%" step="0.1" />
            <SimInput label="GST on Ads %"     value={inp.gstOnAds * 100}     onChange={(v) => set('gstOnAds', v / 100)}     suffix="%" step="0.1" />
          </div>

          <SectionLabel>Advertising</SectionLabel>
          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 12 }}>
            <SimInputDual
              label="Ad Spend"
              value={inp.adSpend}
              type={inp.adSpendType}
              onChangeValue={(v) => set('adSpend', v)}
              onChangeType={(t) => set('adSpendType', t)}
              pctHint="% of net sales"
            />
          </div>
        </div>

        {/* RIGHT: BREAKDOWN */}
        <div style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 14,
          padding: isMobile ? 16 : 24,
          boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
          height: 'fit-content',
          position: isMobile ? 'static' : 'sticky',
          top: 88,
        }}>
          <div style={{ marginBottom: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 4 }}>Calculation Breakdown</div>
            <div style={{ fontSize: 11, color: C.textMute }}>How online menu price becomes final payout</div>
          </div>

          <BreakdownRow op=" " label="Offline Menu Price"   value={calc.offlinePrice} muted />
          <BreakdownRow op="+" label="Price Hike"           value={calc.priceIncrease} positive />
          <BreakdownRow op="=" label="Online Menu Price"    value={calc.onlinePrice} highlight />
          <BreakdownRow op="−" label="Promo Discount"       value={calc.promoAmt} negative />
          <BreakdownRow op="−" label="Extra Discount"        value={calc.saltAmt} negative />
          <BreakdownRow op="+" label="Packaging Charge"     value={inp.packaging} positive />
          <BreakdownRow op="=" label="Net Sales (commissionable)" value={calc.cv} highlight />
          <BreakdownRow op="+" label={`Customer GST (${fmtPct(inp.gstCollected)})`} value={calc.customerGst} muted />
          <BreakdownRow op="=" label="Net Order Value"      value={calc.nov} highlight />
          <BreakdownRow op="−" label="Commission"           value={calc.commission} negative />
          <BreakdownRow op="−" label="PG Charges"           value={calc.pg} negative />
          <BreakdownRow op="−" label="GST on Commission"    value={calc.gstOnComm} negative />
          <BreakdownRow op="−" label="GST on PG"            value={calc.gstOnPg} negative />
          <BreakdownRow op="−" label="Advertisement"        value={calc.adAmt} negative />
          <BreakdownRow op="−" label="GST on Ads"           value={calc.gstOnAds} negative />
          <BreakdownRow op="=" label="Final Payout"         value={calc.payout} payout />
        </div>
      </div>

      {/* ── ECONOMICS CARD (collapsible-feel section) ──────────────── */}
      <div style={{ marginTop: 14 }}>
        <div style={{
          background: C.card,
          border: `1px solid ${C.border}`,
          borderRadius: 14,
          padding: isMobile ? 16 : 24,
          boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
        }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 18, gap: 12, flexWrap: 'wrap' }}>
            <div>
              <div style={{ fontSize: 13, fontWeight: 600, color: C.text, marginBottom: 4 }}>Detailed Economics</div>
              <div style={{ fontSize: 11, color: C.textMute }}>The full picture: every cost, every metric</div>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr 1fr' : 'repeat(4, 1fr)', gap: 12, marginBottom: 22 }}>
            <SimStat label="You Receive"     value={fmtINR(calc.payout)}       sub="Final payout"          hue={C.success} />
            <SimStat label="Platform Takes"  value={fmtINR(calc.totalPlatform)} sub={fmtPct(calc.effCommPct)} hue={C.danger} />
            <SimStat label="Discounts Given" value={fmtINR(calc.totalDisc)}    sub="Promo + extra"          hue={C.accent} />
            <SimStat label="Ads Spent"       value={fmtINR(calc.totalMarketing)} sub="Incl. GST"           hue={C.violet} />
          </div>

          <div style={{
            background: C.bg,
            border: `1px solid ${C.border}`,
            borderRadius: 10,
            padding: isMobile ? 14 : 18,
            display: 'grid',
            gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr',
            gap: isMobile ? 12 : 20,
          }}>
            <EconRow label="Offline Selling Price"  value={fmtINRFull(calc.offlinePrice)} />
            <EconRow label="Total Marketing Cost"   value={fmtINRFull(calc.totalMarketing)} sub="Ads + GST on Ads" hue={C.danger} />
            <EconRow label="Online Selling Price"   value={fmtINRFull(calc.onlinePrice)} />
            <EconRow label="Effective Commission %" value={fmtPct(calc.effCommPct)} sub="(Commission + PG) ÷ Online Price" hue={C.accent} />
            <EconRow label="Price Increase"         value={`+${fmtINRFull(calc.priceIncrease)}`} sub={`+${(inp.hike * 100).toFixed(1)}% over offline`} hue={C.success} />
            <EconRow label="Effective Cost %"       value={fmtPct(calc.effCostPct)} sub="All costs ÷ Online Price" hue={C.violet} />
            <EconRow label="Total Discounts"        value={fmtINRFull(calc.totalDisc)} sub="Promo + extra" hue={C.danger} />
            <EconRow label="Final Payout"           value={fmtINRFull(calc.payout)} bold hue={C.success} />
            <EconRow label="Total Taxes"            value={fmtINRFull(calc.totalTaxes)} sub="GST on order + GST on costs" hue={C.danger} />
            <EconRow label="Payout % of Offline"    value={fmtPct(calc.payoutPctOffline)} sub="Payout ÷ Offline Price" hue={C.success} bold />
            <EconRow label="Total Platform Costs"   value={fmtINRFull(calc.totalPlatform)} sub="Commission + PG + GST" hue={C.danger} />
            <EconRow label="Payout % of Online"     value={fmtPct(calc.payoutPctOnline)} sub="Payout ÷ Online Price" hue={C.accent} />
          </div>
        </div>
      </div>

      {/* Footer */}
      <div style={{
        marginTop: 28,
        padding: '14px 0',
        borderTop: `1px solid ${C.border}`,
        fontSize: 11,
        color: C.textMute,
        textAlign: 'center',
        lineHeight: 1.7,
      }}>
        Price Lens · Numbers are indicative. Check the real rates in your own contract.
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SIMULATOR SUB-COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────
function SectionLabel({ children }) {
  return (
    <div style={{
      fontSize: 10,
      letterSpacing: '0.18em',
      textTransform: 'uppercase',
      color: C.textDim,
      marginBottom: 10,
      paddingBottom: 6,
      borderBottom: `1px solid ${C.border}`,
      display: 'flex',
      alignItems: 'center',
      gap: 6,
    }}>
      <span style={{ width: 4, height: 4, background: C.accent, borderRadius: 1 }} />
      {children}
    </div>
  );
}

function SimInput({ label, value, onChange, prefix, suffix, step = '1', auto }) {
  // Display value: round nicely. Auto fields just show the computed value.
  const displayVal = auto
    ? (Math.round(value * 100) / 100)
    : value;

  return (
    <div>
      <div style={{ fontSize: 10, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.textDim, marginBottom: 6, display: 'flex', justifyContent: 'space-between' }}>
        <span>{label}</span>
        {auto && <span style={{ color: C.textMute, fontWeight: 500 }}>AUTO</span>}
      </div>
      <div style={{ position: 'relative' }}>
        {prefix && (
          <span style={{
            position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)',
            color: C.textDim, fontSize: 13, fontFamily: 'JetBrains Mono',
            pointerEvents: 'none',
          }}>{prefix}</span>
        )}
        <input
          type="number"
          step={step}
          value={typeof displayVal === 'number' && !isNaN(displayVal) ? displayVal : ''}
          readOnly={auto}
          onChange={(e) => {
            const v = parseFloat(e.target.value);
            onChange(isNaN(v) ? 0 : v);
          }}
          style={{
            width: '100%',
            background: auto ? C.bg : C.surface,
            border: `1px solid ${C.border}`,
            borderRadius: 7,
            padding: `9px ${suffix ? 26 : 10}px 9px ${prefix ? 22 : 10}px`,
            color: auto ? C.textDim : C.text,
            fontSize: 14,
            fontFamily: 'JetBrains Mono',
            fontVariantNumeric: 'tabular-nums',
            outline: 'none',
          }}
        />
        {suffix && (
          <span style={{
            position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
            color: C.textMute, fontSize: 12, fontFamily: 'JetBrains Mono',
            pointerEvents: 'none',
          }}>{suffix}</span>
        )}
      </div>
    </div>
  );
}

function SimInputDual({ label, value, type, onChangeValue, onChangeType, pctHint }) {
  return (
    <div>
      <div style={{ fontSize: 10, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.textDim, marginBottom: 6, display: 'flex', justifyContent: 'space-between' }}>
        <span>{label}</span>
        {pctHint && type === 'percent' && <span style={{ color: C.textMute }}>{pctHint}</span>}
      </div>
      <div style={{ display: 'flex', gap: 6 }}>
        <div style={{
          display: 'inline-flex',
          background: C.bg,
          border: `1px solid ${C.border}`,
          borderRadius: 7,
          padding: 2,
          flexShrink: 0,
        }}>
          <button
            type="button"
            onClick={() => onChangeType('percent')}
            style={{
              padding: '6px 10px',
              background: type === 'percent' ? C.surface : 'transparent',
              border: 'none',
              borderRadius: 5,
              color: type === 'percent' ? C.text : C.textDim,
              cursor: 'pointer',
              fontSize: 11,
              fontWeight: 600,
              fontFamily: 'JetBrains Mono',
            }}
          >%</button>
          <button
            type="button"
            onClick={() => onChangeType('rupees')}
            style={{
              padding: '6px 10px',
              background: type === 'rupees' ? C.surface : 'transparent',
              border: 'none',
              borderRadius: 5,
              color: type === 'rupees' ? C.text : C.textDim,
              cursor: 'pointer',
              fontSize: 11,
              fontWeight: 600,
              fontFamily: 'JetBrains Mono',
            }}
          >₹</button>
        </div>
        <div style={{ position: 'relative', flex: 1 }}>
          <input
            type="number"
            step={type === 'percent' ? '0.5' : '100'}
            value={value}
            onChange={(e) => {
              const v = parseFloat(e.target.value);
              onChangeValue(isNaN(v) ? 0 : v);
            }}
            style={{
              width: '100%',
              background: C.surface,
              border: `1px solid ${C.border}`,
              borderRadius: 7,
              padding: '9px 10px',
              color: C.text,
              fontSize: 14,
              fontFamily: 'JetBrains Mono',
              fontVariantNumeric: 'tabular-nums',
              outline: 'none',
            }}
          />
        </div>
      </div>
    </div>
  );
}

function BreakdownRow({ op, label, value, highlight, negative, positive, muted, payout }) {
  const color = payout ? C.success
              : highlight ? C.text
              : negative ? C.danger
              : positive ? C.success
              : muted ? C.textDim
              : C.text;
  const bg = payout ? `${C.success}10`
            : highlight ? C.bg
            : 'transparent';
  const fontWeight = payout || highlight ? 600 : 400;

  return (
    <div style={{
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      padding: '8px 10px',
      borderRadius: 6,
      background: bg,
      marginBottom: 2,
      borderLeft: payout ? `3px solid ${C.success}` : highlight ? `3px solid ${C.accent}` : '3px solid transparent',
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, fontSize: 13, color }}>
        <span style={{
          width: 14,
          textAlign: 'center',
          color: C.textMute,
          fontFamily: 'JetBrains Mono',
          fontSize: 12,
        }}>{op}</span>
        <span style={{ fontWeight }}>{label}</span>
      </div>
      <span className="num" style={{ fontSize: 13, color, fontWeight, letterSpacing: '-0.01em' }}>
        {negative && value !== 0 ? '−' : ''}{fmtINRFull(Math.abs(value))}
      </span>
    </div>
  );
}

function MiniStat({ label, value, hue }) {
  return (
    <div style={{
      padding: 12,
      background: C.bg,
      border: `1px solid ${C.border}`,
      borderRadius: 10,
    }}>
      <div style={{ fontSize: 10, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.textDim, marginBottom: 4 }}>
        {label}
      </div>
      <div className="num" style={{ fontSize: 16, fontWeight: 600, color: hue || C.text }}>
        {value}
      </div>
    </div>
  );
}

function SimStat({ label, value, sub, hue }) {
  return (
    <div style={{
      padding: 14,
      background: C.bg,
      border: `1px solid ${C.border}`,
      borderRadius: 10,
      position: 'relative',
      overflow: 'hidden',
    }}>
      <div style={{
        position: 'absolute', top: 0, left: 0, width: 24, height: 2,
        background: hue || C.accent,
      }} />
      <div style={{ fontSize: 10, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.textDim, marginBottom: 4 }}>
        {label}
      </div>
      <div className="num" style={{ fontSize: 18, fontWeight: 600, color: hue || C.text, marginBottom: 2 }}>
        {value}
      </div>
      {sub && <div style={{ fontSize: 11, color: C.textMute }}>{sub}</div>}
    </div>
  );
}

function EconRow({ label, value, sub, hue, bold }) {
  return (
    <div style={{
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'flex-start',
      padding: '6px 0',
      borderBottom: `1px dashed ${C.border}`,
      gap: 12,
    }}>
      <div>
        <div style={{ fontSize: 12.5, color: C.text, fontWeight: bold ? 600 : 400 }}>{label}</div>
        {sub && <div style={{ fontSize: 10.5, color: C.textMute, marginTop: 1 }}>{sub}</div>}
      </div>
      <div className="num" style={{
        fontSize: bold ? 14.5 : 13,
        color: hue || C.text,
        fontWeight: bold ? 700 : 500,
        textAlign: 'right',
        whiteSpace: 'nowrap',
      }}>
        {value}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SHARED COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────
function KPICard({ label, value, sub, accent }) {
  return (
    <div style={{
      background: C.card,
      border: `1px solid ${C.border}`,
      borderRadius: 12,
      padding: '18px 18px 16px',
      position: 'relative',
      overflow: 'hidden',
      boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
    }}>
      <div style={{
        position: 'absolute', top: 0, left: 0, width: 30, height: 2,
        background: accent || C.accent,
      }} />
      <div style={{ fontSize: 10, letterSpacing: '0.14em', textTransform: 'uppercase', color: C.textDim, marginBottom: 8 }}>
        {label}
      </div>
      <div className="num" style={{ fontSize: 24, fontWeight: 500, color: C.text, marginBottom: 4 }}>
        {value}
      </div>
      <div style={{ fontSize: 11, color: C.textMute }}>
        {sub}
      </div>
    </div>
  );
}

function Card({ title, subtitle, children, action }) {
  return (
    <div style={{
      background: C.card,
      border: `1px solid ${C.border}`,
      borderRadius: 12,
      padding: 20,
      boxShadow: '0 1px 2px rgba(0,0,0,0.03)',
    }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 14 }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{title}</div>
          {subtitle && <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>{subtitle}</div>}
        </div>
        {action}
      </div>
      {children}
    </div>
  );
}

function FilterSelect({ label, value, options, onChange, icon }) {
  return (
    <div style={{
      display: 'flex', alignItems: 'center', gap: 8,
      background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
      padding: '8px 12px',
    }}>
      {icon && <span style={{ color: C.textDim }}>{icon}</span>}
      <span style={{ fontSize: 11, color: C.textDim, letterSpacing: '0.05em' }}>{label}</span>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        style={{
          background: 'transparent', border: 'none', outline: 'none',
          color: C.text, fontSize: 13, cursor: 'pointer',
          appearance: 'none', paddingRight: 16,
        }}
      >
        {options.map(o => <option key={o} value={o} style={{ background: C.card }}>{o}</option>)}
      </select>
    </div>
  );
}

function Pill({ icon, children }) {
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 6,
      padding: '4px 10px', border: `1px solid ${C.border}`, borderRadius: 100,
      background: C.card, fontSize: 12, color: C.textDim,
    }}>
      {icon}
      {children}
    </span>
  );
}

function NumInput({ label, value, onChange, pct }) {
  const displayVal = pct ? (value * 100).toFixed(1) : value;
  return (
    <div>
      <div style={{ fontSize: 10, letterSpacing: '0.12em', textTransform: 'uppercase', color: C.textDim, marginBottom: 6 }}>
        {label}
      </div>
      <div style={{ position: 'relative' }}>
        <input
          type="number"
          step="0.1"
          value={displayVal}
          onChange={(e) => {
            const v = parseFloat(e.target.value);
            onChange(pct ? (isNaN(v) ? 0 : v / 100) : (isNaN(v) ? 0 : v));
          }}
          style={{
            width: '100%',
            background: C.surface,
            border: `1px solid ${C.border}`,
            borderRadius: 6,
            padding: '8px 24px 8px 10px',
            color: C.text,
            fontSize: 14,
            fontFamily: 'JetBrains Mono',
            fontVariantNumeric: 'tabular-nums',
            outline: 'none',
          }}
        />
        {pct && (
          <span style={{
            position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
            color: C.textMute, fontSize: 12, fontFamily: 'JetBrains Mono',
          }}>%</span>
        )}
      </div>
    </div>
  );
}

function CustomTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{
      background: C.surface,
      border: `1px solid ${C.border}`,
      borderRadius: 8,
      padding: '8px 12px',
      fontSize: 12,
      boxShadow: '0 4px 12px rgba(0,0,0,0.08)',
    }}>
      <div style={{ color: C.textDim, marginBottom: 4, fontSize: 11 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ display: 'flex', justifyContent: 'space-between', gap: 16, alignItems: 'center' }}>
          <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 8, height: 8, background: p.color, borderRadius: 2 }} />
            <span style={{ color: C.textDim }}>{p.name}</span>
          </span>
          <span className="num" style={{ color: C.text, fontWeight: 500 }}>
            {typeof p.value === 'number' && p.value > 1000 ? fmtINR(p.value) :
             typeof p.value === 'number' && p.value < 100 ? p.value.toFixed(2) :
             fmtNum(p.value)}
          </span>
        </div>
      ))}
    </div>
  );
}

function PLRow({ label, data, field, total, fmt, negative, bold, accent, success, dim, pctField, totalPct }) {
  const color = success ? C.success : accent ? C.accent : negative ? C.danger : dim ? C.textDim : C.text;
  const totalColor = success ? C.success : accent ? C.accent : color;
  const renderCell = (val, pct) => {
    const main = negative ? `(${fmt(val)})` : fmt(val);
    if (pctField == null || pct == null) return main;
    return (
      <span>
        {main}
        <span style={{ color: C.textMute, fontSize: 11, marginLeft: 6, fontWeight: 400 }}>
          {fmtPct(pct)}
        </span>
      </span>
    );
  };
  return (
    <tr style={{ borderTop: `1px solid ${C.border}` }}>
      <td style={{ ...td(), color, fontWeight: bold ? 600 : 400 }}>{label}</td>
      {data.map(d => (
        <td key={d.month} className="num" style={{ ...td(), textAlign: 'right', color, fontWeight: bold ? 600 : 400 }}>
          {renderCell(d[field], pctField ? d[pctField] : null)}
        </td>
      ))}
      <td className="num" style={{ ...td(), textAlign: 'right', color: totalColor, fontWeight: 600 }}>
        {pctField != null && totalPct != null
          ? <span>{negative ? `(${fmt(total)})` : fmt(total)}<span style={{ color: C.textMute, fontSize: 11, marginLeft: 6, fontWeight: 400 }}>{fmtPct(totalPct)}</span></span>
          : (negative ? `(${fmt(total)})` : fmt(total))}
      </td>
    </tr>
  );
}

function PLSep({ label }) {
  return (
    <tr>
      <td colSpan={20} style={{
        padding: '14px 12px 4px',
        fontSize: 10,
        letterSpacing: '0.18em',
        textTransform: 'uppercase',
        color: C.textMute,
        borderTop: `1px solid ${C.border}`,
      }}>
        {label}
      </td>
    </tr>
  );
}

const th = () => ({ padding: '10px 12px', textAlign: 'left', fontWeight: 500, borderBottom: `1px solid ${C.border}` });
const td = (color) => ({ padding: '10px 12px', color: color || C.text, fontSize: 13 });

function monthLabel(yyyymm) {
  if (!yyyymm) return '';
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const [y, m] = yyyymm.split('-');
  return `${months[parseInt(m)-1]} ${y.slice(2)}`;
}