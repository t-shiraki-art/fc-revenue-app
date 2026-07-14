import { useState, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────
// デザイントークン — カーセブン ブランドデザイン
// 白ベース × カーセブンレッド × Inter + Noto Sans JP
// ─────────────────────────────────────────────
const C = {
  // 背景・サーフェス（白ベース）
  bg:          "#F5F6F8",
  surface:     "#FFFFFF",
  surfaceHigh: "#F0F2F5",
  surfaceMid:  "#F8F9FB",
  // ボーダー
  border:      "#E2E6EC",
  borderLight: "#EDF0F4",
  // テキスト
  text:        "#1A1D23",
  textSub:     "#4B5563",
  textMuted:   "#9CA3AF",
  // ブランドカラー（カーセブンレッド）
  red:         "#E31F25",
  redLight:    "#FF4D52",
  redDim:      "#FDEAEA",
  redBorder:   "#FBBFBF",
  // 機能色
  green:       "#059669",
  greenLight:  "#D1FAE5",
  blue:        "#2563EB",
  blueLight:   "#DBEAFE",
  amber:       "#D97706",
  amberLight:  "#FEF3C7",
  purple:      "#7C3AED",
  purpleLight: "#EDE9FE",
  teal:        "#0891B2",
  tealLight:   "#CFFAFE",
  orange:      "#EA580C",
  orangeLight: "#FFEDD5",
  // 期区切り（ブランドレッド）
  periodLine:  "#E31F25",
  periodBg:    "rgba(227,31,37,0.04)",
};

// PKG定義（実データに合わせる）
const PKG_DEFS = {
  toku:    { name:"創業時特別PKG", color:C.amber   },
  norm:    { name:"通常PKG",       color:C.blue    },
  area:    { name:"エリア限定PKG", color:C.green   },
  okinawa: { name:"沖縄PKG",       color:C.purple  },
  kyushu:  { name:"九州PKG",       color:C.orange  },
  direct:  { name:"直営",          color:C.teal    },
  other:   { name:"その他",        color:C.textMuted},
};
const PKG_MAP = {
  '創業時特別PKG':'toku','通常PKG':'norm','エリア限定PKG':'area',
  '沖縄PKG':'okinawa','九州PKG':'kyushu','直営':'direct',
};

// 収入7項目
const ITEMS = [
  { id:"royalty",    label:"純粋ロイヤリティ", color:C.amber,      adOnly:false },
  { id:"sv",         label:"SV費用",           color:C.purple,     adOnly:false },
  { id:"renewal",    label:"更新料",            color:C.orange,     adOnly:false },
  { id:"membership", label:"加盟金",            color:C.green,      adOnly:false },
  { id:"cs",         label:"CS向上",            color:C.teal,       adOnly:false },
  { id:"system",     label:"システム費用",       color:C.blue,       adOnly:false },
  { id:"ad",         label:"広告費用",           color:C.textMuted,  adOnly:true  },
];
const ITEM_LABEL_MAP = {
  '純粋ロイヤリティ':'royalty','SV費用':'sv','更新料':'renewal',
  '加盟金':'membership','CS向上':'cs','システム費用':'system','広告費用':'ad',
};

// ─────────────────────────────────────────────
// XLSXパーサー（「20260713更新_売上明細」シート）
// ─────────────────────────────────────────────
function parseXLSX(buffer, sheetName) {
  const wb = XLSX.read(buffer, { type:"array", cellDates:true });

  // シート名が指定されていなければ自動選択（「売上明細」を含むシート）
  const targetSheet = sheetName ||
    wb.SheetNames.find(n => n.includes("売上明細") && !n.includes("ひな")) ||
    wb.SheetNames[0];

  const ws = wb.Sheets[targetSheet];
  if (!ws) return { error:`シート「${targetSheet}」が見つかりません`, sheets: wb.SheetNames };

  const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
  if (raw.length < 9) return { error:"データが少なすぎます" };

  // 行8（index=7）がヘッダー
  const headerRow = raw[7];

  // 月列を検出（DateオブジェクトまたはYYYYMM数値）
  const monthCols = [];
  headerRow.forEach((v, i) => {
    if (v instanceof Date) {
      monthCols.push({ col:i, ym:`${v.getFullYear()}/${String(v.getMonth()+1).padStart(2,"0")}` });
    } else if (typeof v === 'number' && v > 200000 && v < 210000) {
      const y = Math.floor(v / 100);
      const m = v % 100;
      monthCols.push({ col:i, ym:`${y}/${String(m).padStart(2,"0")}` });
    }
  });

  const allMonths = monthCols.map(mc => mc.ym);

  // データ行（行9〜）を集約
  const storeMap = {};
  for (let ri = 8; ri < raw.length; ri++) {
    const row = raw[ri];
    if (!row || !row[1]) continue;
    const name = String(row[1]).trim();
    const itemJa = row[14] ? String(row[14]).trim() : null;
    if (!itemJa) continue;
    const itemId = ITEM_LABEL_MAP[itemJa];
    if (!itemId) continue;

    if (!storeMap[name]) {
      const retire = row[0];
      let retireYM = null;
      if (retire instanceof Date) {
        retireYM = `${retire.getFullYear()}/${String(retire.getMonth()+1).padStart(2,"0")}`;
      } else if (retire && typeof retire === 'string' && retire !== '撤退候補') {
        const d = new Date(retire);
        if (!isNaN(d)) retireYM = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}`;
      }

      const ce = row[5];
      let contractEnd = null;
      if (ce instanceof Date) {
        contractEnd = `${ce.getFullYear()}/${String(ce.getMonth()+1).padStart(2,"0")}`;
      } else if (ce && typeof ce === 'string') {
        contractEnd = ce.substring(0, 7).replace('-', '/');
      }

      const pkgRaw = row[8] ? String(row[8]) : '';
      storeMap[name] = {
        id:            String(row[7] || name),
        storeId:       String(row[7] || ''),        // 店舗ID
        memberId:      String(row[3] || ''),         // 加盟店ID
        company:       row[2] ? String(row[2]) : '', // 会社名
        name,
        pref:          row[13] ? String(row[13]) : '',
        pkgKey:        PKG_MAP[pkgRaw] || 'other',
        pkgName:       pkgRaw,
        contractYears: row[9] ? String(row[9]) : '',
        renewalPeriod: row[12] ? String(row[12]) : '', // 更新期
        openDate:      row[6] instanceof Date
                         ? `${row[6].getFullYear()}/${String(row[6].getMonth()+1).padStart(2,"0")}`
                         : String(row[6]||''),
        lastRenewal:   row[4] ? (row[4] instanceof Date
                         ? `${row[4].getFullYear()}/${String(row[4].getMonth()+1).padStart(2,"0")}`
                         : String(row[4])) : '',      // 直近契約更新日
        contractEnd,
        retireYM,
        retireRaw:     row[0] ? String(row[0]) : '',  // 撤退日原文
        monthly:       {},
      };
    }

    // 月次金額を格納
    for (const mc of monthCols) {
      const val = row[mc.col];
      if (val !== null && val !== '' && !isNaN(Number(val)) && Number(val) !== 0) {
        if (!storeMap[name].monthly[mc.ym]) storeMap[name].monthly[mc.ym] = {};
        storeMap[name].monthly[mc.ym][itemId] = Number(val);
      }
    }
  }

  const stores = Object.values(storeMap);
  return { stores, months: allMonths, sheetName: targetSheet, sheets: wb.SheetNames };
}

// 期番号を取得（10月始まり）
function getPeriodNum(ym) {
  const y = parseInt(ym.substring(0,4));
  const m = parseInt(ym.substring(5,7));
  return m >= 10 ? y : y - 1;
}

// 合計
const sumItems = r => r ? ITEMS.reduce((s,it) => s + (r[it.id]||0), 0) : 0;

// 金額変更履歴を適用してその月の収入行を返す
function applyPriceChanges(baseRow, storeId, ym, priceChanges) {
  if (!baseRow || !priceChanges || priceChanges.length === 0) return baseRow;
  // storeIdに対して fromYM <= ym の変更を適用（複数項目・複数タイミング対応）
  const applicable = priceChanges
    .filter(c => c.storeId === storeId && ym >= c.fromYM)
    .sort((a, b) => a.fromYM < b.fromYM ? -1 : 1);
  if (applicable.length === 0) return baseRow;
  const row = { ...baseRow };
  // 同じ項目で複数変更がある場合、最新（fromYM最大）を使う
  const latest = {};
  applicable.forEach(c => {
    if (!latest[c.itemId] || c.fromYM >= latest[c.itemId].fromYM) {
      latest[c.itemId] = c;
    }
  });
  Object.values(latest).forEach(c => { row[c.itemId] = c.newValue; });
  return row;
}
const fmtK = n => !n ? "" : Math.round(n/1000).toLocaleString()+"千";
const fmtM = n => !n ? "¥0" : "¥"+Math.round(n/10000).toLocaleString()+"万";
const fmtFull = n => n == null ? "—" : "¥"+Math.round(n).toLocaleString();

// ─────────────────────────────────────────────
// スタイル — 白ベース クリーンデザイン
// ─────────────────────────────────────────────
const baseCell = {
  height:36, lineHeight:"36px",
  borderBottom:`1px solid ${C.borderLight}`,
  whiteSpace:"nowrap", fontSize:13, padding:"0 14px",
};
const numCell = (isFirst, retired, event) => ({
  ...baseCell, textAlign:"right",
  borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}`,
  background: retired ? C.redDim : event ? C.amberLight : "transparent",
  color: retired ? C.red : event ? C.amber : C.text,
  minWidth: 96,
});
const stickyL  = (w, bg, bold) => ({
  ...baseCell, position:"sticky", left:0, zIndex:4,
  width:w, minWidth:w, maxWidth:w,
  background:bg||C.surface, fontWeight:bold?"600":"400",
  borderRight:`1px solid ${C.border}`,
  overflow:"hidden", textOverflow:"ellipsis",
});
const stickyL2 = (bg) => ({
  ...baseCell, position:"sticky", left:200, zIndex:4,
  width:120, minWidth:120, maxWidth:120,
  background:bg||C.surface,
  borderRight:`1px solid ${C.border}`,
  fontSize:12, color:C.textSub,
});
const inputSt = {
  padding:"5px 10px",
  background:C.surface,
  border:`1px solid ${C.border}`,
  borderRadius:6, color:C.text, fontSize:13, outline:"none",
  boxShadow:"inset 0 1px 2px rgba(0,0,0,0.04)",
};
const selSt = { ...inputSt };
const btnSt = (c=C.red) => ({
  padding:"5px 14px", borderRadius:6,
  border:`1px solid ${c}`,
  background: c===C.red ? C.red : `${c}10`,
  color: c===C.red ? "#fff" : c,
  fontSize:13, fontWeight:600, cursor:"pointer",
  transition:"opacity .15s",
});
const pillSt = a => ({
  padding:"4px 12px", borderRadius:20, cursor:"pointer",
  border:`1px solid ${a ? C.red : C.border}`,
  background: a ? C.red : "transparent",
  color: a ? "#fff" : C.textSub,
  fontWeight: a ? 600 : 400, fontSize:12,
});

// ─────────────────────────────────────────────
// APP ROOT
// ─────────────────────────────────────────────
export default function App({
  initialData         = null,
  initialPriceChanges = [],
  onXLSXLoaded        = null,
  onStoreUpdate       = null,
  onPriceChangeAdd    = null,
  onPriceChangeDelete = null,
}) {
  const [tab, setTab]             = useState(initialData ? "sim" : "upload");
  const [appData, setAppData]     = useState(initialData);
  const [showItems, setShowItems] = useState(Object.fromEntries(ITEMS.map(it=>[it.id,true])));
  // 金額変更履歴: [{ id, storeId, fromYM, itemId, newValue }]
  const [priceChanges, setPriceChanges] = useState(initialPriceChanges);

  const onLoaded = useCallback(async (data) => {
    setAppData(data);
    setTab("sim");
    if (onXLSXLoaded) await onXLSXLoaded(data);
  }, [onXLSXLoaded]);

  const { periods, periodStarts } = useMemo(() => {
    if (!appData) return { periods:[], periodStarts:{} };
    const ps = {};
    appData.months.forEach(ym => {
      const pn = getPeriodNum(ym);
      if (!ps[pn]) ps[pn] = ym;
    });
    return { periods: Object.keys(ps).map(Number).sort(), periodStarts: ps };
  }, [appData]);

  const shared = { appData, periods, periodStarts, showItems, setShowItems, priceChanges };

  return (
    <div style={{ fontFamily:"'Inter','Noto Sans JP',sans-serif", background:C.bg, minHeight:"100vh", color:C.text }}>
      {/* ヘッダー */}
      <div style={{
        position:"sticky", top:0, zIndex:100,
        background:C.surface,
        borderBottom:`1px solid ${C.border}`,
        boxShadow:"0 1px 3px rgba(0,0,0,0.06)",
        height:56, display:"flex", alignItems:"center", padding:"0 20px",
      }}>
        {/* ロゴ */}
        <div style={{ display:"flex", alignItems:"center", gap:10, marginRight:32 }}>
          <div style={{
            width:36, height:36, borderRadius:8,
            background:C.red,
            display:"flex", alignItems:"center", justifyContent:"center",
            fontWeight:900, fontSize:15, color:"#fff", letterSpacing:"-0.5px",
            boxShadow:`0 2px 8px ${C.red}40`,
          }}>C7</div>
          <div>
            <div style={{ fontSize:13, fontWeight:700, color:C.text, letterSpacing:"0.01em" }}>
              カーセブン FC収入管理
            </div>
            <div style={{ fontSize:10, color:C.textMuted }}>
              {appData ? `${appData.stores.length}店舗 | ${appData.sheetName}` : "XLSXをアップロードして開始"}
            </div>
          </div>
        </div>
        {/* ナビ */}
        {[
          { id:"upload", label:"データ読込" },
          { id:"sim",    label:"売上明細",    disabled:!appData },
          { id:"summary",label:"期別サマリー", disabled:!appData },
          { id:"master", label:"店舗マスタ",  disabled:!appData },
        ].map(({ id, label, disabled }) => (
          <button key={id} onClick={() => !disabled && setTab(id)} style={{
            height:56, padding:"0 18px",
            background:"transparent", border:"none",
            cursor: disabled ? "default" : "pointer",
            fontSize:13, fontWeight: tab===id ? 700 : 400,
            color: disabled ? C.textMuted : tab===id ? C.red : C.textSub,
            borderBottom: tab===id ? `2px solid ${C.red}` : "2px solid transparent",
            opacity: disabled ? 0.35 : 1,
            letterSpacing:"0.01em",
          }}>{label}</button>
        ))}
        {/* 右端情報 */}
        {appData && (
          <div style={{ marginLeft:"auto", display:"flex", gap:16, alignItems:"center", fontSize:11, color:C.textMuted }}>
            <span>{appData.stores.length}店舗</span>
            <div style={{ display:"flex", alignItems:"center", gap:4 }}>
              <div style={{ width:10, height:2, background:C.red, borderRadius:1 }} />
              <span style={{ color:C.red, fontSize:10 }}>期区切り</span>
            </div>
          </div>
        )}
      </div>

      <div style={{ padding:"16px 12px" }}>
        {tab === "upload"  && <UploadTab onLoaded={onLoaded} />}
        {tab === "sim"     && appData && <SimTab     {...shared} />}
        {tab === "summary" && appData && <SummaryTab {...shared} />}
        {tab === "master"  && appData && <MasterTab  appData={appData} priceChanges={priceChanges} setPriceChanges={setPriceChanges} />}
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════
// アップロードタブ
// ══════════════════════════════════════════════
function UploadTab({ onLoaded }) {
  const [loading, setLoading] = useState(false);
  const [error,   setError]   = useState(null);
  const [sheets,  setSheets]  = useState(null);
  const [buffer,  setBuffer]  = useState(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const fileRef = useRef();

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true); setError(null); setSheets(null);
    try {
      const buf = await file.arrayBuffer();
      // まずシート一覧だけ取得
      const wb = XLSX.read(buf, { type:"array", cellDates:true });
      setSheets(wb.SheetNames);
      setBuffer(buf);
      // 売上明細シートを自動選択
      const auto = wb.SheetNames.find(n => n.includes("売上明細") && !n.includes("ひな")) || wb.SheetNames[0];
      setSelectedSheet(auto);
    } catch(err) {
      setError("ファイルの読み込みに失敗しました: " + err.message);
    }
    setLoading(false);
    e.target.value = "";
  };

  const handleLoad = () => {
    if (!buffer || !selectedSheet) return;
    setLoading(true); setError(null);
    try {
      const result = parseXLSX(buffer, selectedSheet);
      if (result.error) { setError(result.error); setLoading(false); return; }
      onLoaded(result);
    } catch(err) {
      setError("データ解析エラー: " + err.message);
    }
    setLoading(false);
  };

  return (
    <div style={{ maxWidth:560, margin:"48px auto" }}>
      {/* メインカード */}
      <div style={{
        background:C.surface, borderRadius:12,
        border:`1px solid ${C.border}`,
        boxShadow:"0 2px 12px rgba(0,0,0,0.06)",
        overflow:"hidden",
      }}>
        {/* 上部レッドバー */}
        <div style={{ height:4, background:C.red }} />
        <div style={{ padding:"32px 32px 28px", textAlign:"center" }}>
          <div style={{
            width:56, height:56, borderRadius:12,
            background:C.redDim, border:`1px solid ${C.redBorder}`,
            display:"flex", alignItems:"center", justifyContent:"center",
            fontSize:24, margin:"0 auto 16px",
          }}>📊</div>
          <div style={{ fontSize:18, fontWeight:700, color:C.text, marginBottom:8 }}>
            売上明細をアップロード
          </div>
          <div style={{ fontSize:13, color:C.textMuted, lineHeight:1.7, marginBottom:24 }}>
            売上明細シートを含むXLSXファイルをアップロードしてください。<br />
            シート選択後に「読み込む」ボタンでデータを反映します。
          </div>

          <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display:"none" }} onChange={handleFile} />
          <button onClick={() => fileRef.current.click()} disabled={loading} style={{
            padding:"10px 28px", borderRadius:8,
            background: loading ? C.border : C.red,
            color:"#fff", border:"none",
            fontSize:14, fontWeight:600, cursor: loading ? "default" : "pointer",
            boxShadow:`0 2px 8px ${C.red}30`,
          }}>
            {loading ? "読み込み中..." : "ファイルを選択"}
          </button>

          {error && (
            <div style={{
              marginTop:16, padding:"10px 14px",
              background:C.redDim, border:`1px solid ${C.redBorder}`,
              borderRadius:8, fontSize:12, color:C.red, textAlign:"left",
            }}>
              {error}
            </div>
          )}

          {sheets && (
            <div style={{ marginTop:20, textAlign:"left" }}>
              <div style={{ fontSize:12, color:C.textSub, marginBottom:8, fontWeight:600 }}>
                シートを選択（{sheets.length}件）
              </div>
              <div style={{
                maxHeight:200, overflowY:"auto",
                background:C.bg, borderRadius:8,
                border:`1px solid ${C.border}`,
              }}>
                {sheets.map(s => (
                  <div key={s} onClick={() => setSelectedSheet(s)} style={{
                    padding:"9px 14px", cursor:"pointer", fontSize:12,
                    background: selectedSheet===s ? C.redDim : "transparent",
                    color: selectedSheet===s ? C.red : C.text,
                    borderBottom:`1px solid ${C.borderLight}`,
                    fontWeight: selectedSheet===s ? 600 : 400,
                    display:"flex", alignItems:"center", justifyContent:"space-between",
                  }}>
                    <span>{s}</span>
                    <div style={{ display:"flex", gap:6 }}>
                      {selectedSheet===s && <span style={{ fontSize:11, color:C.red }}>✓</span>}
                      {s.includes("売上明細") && !s.includes("ひな") && (
                        <span style={{
                          fontSize:10, padding:"1px 6px", borderRadius:4,
                          background:C.greenLight, color:C.green, fontWeight:600,
                        }}>推奨</span>
                      )}
                    </div>
                  </div>
                ))}
              </div>
              <button onClick={handleLoad} disabled={!selectedSheet || loading} style={{
                marginTop:12, width:"100%", padding:"10px",
                borderRadius:8, border:"none",
                background: (!selectedSheet || loading) ? C.border : C.red,
                color:"#fff", fontSize:13, fontWeight:600,
                cursor: (!selectedSheet || loading) ? "default" : "pointer",
              }}>
                {loading ? "解析中..." : `「${selectedSheet}」を読み込む`}
              </button>
            </div>
          )}
        </div>  {/* padding div 閉じ */}
      </div>  {/* メインカード 閉じ */}

      {/* フォーマット説明 */}
      <div style={{
        marginTop:12, padding:"14px 16px",
        background:C.surface, border:`1px solid ${C.border}`,
        borderRadius:10, fontSize:12, color:C.textMuted, lineHeight:1.8,
      }}>
        <div style={{ fontWeight:600, color:C.textSub, marginBottom:6 }}>対応フォーマット</div>
        <div>・8行目：ヘッダー行（撤退日、店舗名、契約プラン、都道府県、品目、月次列…）</div>
        <div>・9行目以降：店舗×品目のデータ行</div>
        <div>・品目：純粋ロイヤリティ / SV費用 / 更新料 / 加盟金 / CS向上 / システム費用 / 広告費用</div>
        <div>・月次列：日付形式（2025/04 等）で自動検出</div>
      </div>
    </div>
  );
}
// ══════════════════════════════════════════════
function SimTab({ appData, periods, periodStarts, showItems, setShowItems, priceChanges }) {
  const { stores, months } = appData;

  const [expanded,   setExpanded]   = useState(new Set());
  const [filterPref, setFilterPref] = useState("all");
  const [filterPkg,  setFilterPkg]  = useState("all");
  const [search,     setSearch]     = useState("");
  const [filterPeriod, setFilterPeriod] = useState(null);

  // 表示月を絞り込み
  const visibleMonths = useMemo(() => {
    if (!filterPeriod) return months;
    return months.filter(ym => getPeriodNum(ym) === filterPeriod);
  }, [months, filterPeriod]);

  const prefs = useMemo(() => ["all", ...new Set(stores.map(s => s.pref).filter(Boolean))], [stores]);
  const pkgs  = useMemo(() => ["all", ...new Set(stores.map(s => s.pkgKey))], [stores]);

  const filtered = useMemo(() => stores.filter(s => {
    if (filterPref !== "all" && s.pref !== filterPref) return false;
    if (filterPkg !== "all" && s.pkgKey !== filterPkg) return false;
    if (search && !s.name.includes(search)) return false;
    return true;
  }), [stores, filterPref, filterPkg, search]);

  // 全店合計（金額変更反映済み）
  const grandTotals = useMemo(() => {
    const t = {};
    visibleMonths.forEach(ym => {
      t[ym] = { _total:0 };
      ITEMS.forEach(it => { t[ym][it.id] = 0; });
      stores.forEach(s => {
        const baseRow = s.monthly[ym];
        if (!baseRow) return;
        const row = applyPriceChanges(baseRow, s.id, ym, priceChanges);
        ITEMS.forEach(it => { t[ym][it.id] += row[it.id]||0; });
        t[ym]._total += sumItems(row);
      });
    });
    return t;
  }, [stores, visibleMonths, priceChanges]);

  const toggle = id => { const n = new Set(expanded); n.has(id) ? n.delete(id) : n.add(id); setExpanded(n); };
  const visItems = ITEMS.filter(it => showItems[it.id]);

  return (
    <div>
      {/* コントロールバー */}
      <div style={{
        display:"flex", gap:8, marginBottom:8, alignItems:"center", flexWrap:"wrap",
        padding:"10px 12px", background:C.surface, borderRadius:8,
        border:`1px solid ${C.border}`, boxShadow:"0 1px 3px rgba(0,0,0,0.04)",
      }}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="店舗名で検索..."
          style={{ ...inputSt, width:150 }} />
        <select value={filterPref} onChange={e=>setFilterPref(e.target.value)} style={selSt}>
          <option value="all">全都道府県</option>
          {prefs.slice(1).map(p => <option key={p} value={p}>{p}</option>)}
        </select>
        <select value={filterPkg} onChange={e=>setFilterPkg(e.target.value)} style={selSt}>
          <option value="all">全PKG</option>
          {pkgs.slice(1).map(k => <option key={k} value={k}>{PKG_DEFS[k]?.name||k}</option>)}
        </select>
        <div style={{ display:"flex", gap:4, alignItems:"center" }}>
          <span style={{ fontSize:11, color:C.textMuted }}>期：</span>
          <button onClick={() => setFilterPeriod(null)} style={pillSt(!filterPeriod)}>全期</button>
          {periods.map(p => (
            <button key={p} onClick={() => setFilterPeriod(p===filterPeriod?null:p)} style={pillSt(filterPeriod===p)}>
              {p}年度
            </button>
          ))}
        </div>
        <div style={{ display:"flex", gap:4 }}>
          <button onClick={() => setExpanded(new Set(stores.map(s=>s.id)))} style={{ ...btnSt(C.textSub), background:"transparent", color:C.textSub, border:`1px solid ${C.border}`, fontSize:11 }}>全展開</button>
          <button onClick={() => setExpanded(new Set())} style={{ ...btnSt(C.textSub), background:"transparent", color:C.textSub, border:`1px solid ${C.border}`, fontSize:11 }}>全折畳</button>
        </div>
        <div style={{ marginLeft:"auto", display:"flex", gap:4, flexWrap:"wrap" }}>
          {ITEMS.map(it => (
            <button key={it.id} onClick={() => setShowItems({ ...showItems, [it.id]:!showItems[it.id] })}
              style={{
                padding:"3px 10px", borderRadius:16, fontSize:11, cursor:"pointer",
                background: showItems[it.id] ? `${it.color}15` : "transparent",
                color: showItems[it.id] ? it.color : C.textMuted,
                border: `1px solid ${showItems[it.id] ? it.color : C.border}`,
                fontWeight: showItems[it.id] ? 600 : 400,
              }}>{it.label}</button>
          ))}
        </div>
      </div>

      {/* テーブル */}
      <div style={{
        overflowX:"auto", overflowY:"auto",
        maxHeight:"calc(100vh - 168px)",
        border:`1px solid ${C.border}`, borderRadius:8,
        background:C.surface,
        boxShadow:"0 1px 4px rgba(0,0,0,0.05)",
      }}>
        <table style={{ borderCollapse:"collapse", whiteSpace:"nowrap", minWidth:"100%" }}>
          <thead>
            {/* 期ラベル行 */}
            <tr style={{ background:C.surfaceHigh }}>
              <th style={{ ...stickyL(200,C.surfaceHigh,true), position:"sticky", left:0, top:0, zIndex:12, height:22, lineHeight:"22px", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}` }}>店舗名</th>
              <th style={{ ...stickyL2(C.surfaceHigh), position:"sticky", left:200, top:0, zIndex:12, height:22, lineHeight:"22px", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}` }}>項目</th>
              {visibleMonths.map(ym => {
                const isFirst = ym === periodStarts[getPeriodNum(ym)];
                const pn = getPeriodNum(ym);
                return (
                  <th key={ym} style={{ height:22, lineHeight:"22px", textAlign:"center", fontSize:10,
                    background: isFirst ? C.periodBg : C.surfaceHigh,
                    borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}`,
                    borderBottom:`1px solid ${C.border}`,
                    color: isFirst ? C.red : C.textMuted,
                    fontWeight: isFirst ? 700 : 400,
                    minWidth:72, padding:"0 4px",
                    position:"sticky", top:0, zIndex:8,
                  }}>
                    {isFirst ? `${pn}年度` : ""}
                  </th>
                );
              })}
            </tr>
            {/* 年月行 */}
            <tr style={{ background:C.surface }}>
              <th style={{ ...stickyL(200,C.surface,false), position:"sticky", left:0, top:22, zIndex:12, height:22, lineHeight:"22px", borderBottom:`1px solid ${C.border}` }}></th>
              <th style={{ ...stickyL2(C.surface), position:"sticky", left:200, top:22, zIndex:12, height:22, lineHeight:"22px", borderBottom:`1px solid ${C.border}` }}></th>
              {visibleMonths.map(ym => {
                const isFirst = ym === periodStarts[getPeriodNum(ym)];
                const mo = ym.substring(5);
                return (
                  <th key={ym} style={{ height:22, lineHeight:"22px", textAlign:"center", fontSize:9,
                    color: isFirst ? C.red : C.textMuted, fontWeight: isFirst ? 600 : 400,
                    background: isFirst ? C.periodBg : C.surface,
                    borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}`,
                    borderBottom:`1px solid ${C.border}`,
                    position:"sticky", top:22, zIndex:8,
                  }}>{mo}月</th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {/* 全店合計 */}
            <GrandTotalRow grandTotals={grandTotals} visibleMonths={visibleMonths} periodStarts={periodStarts} visItems={visItems} filteredCount={filtered.length} totalCount={stores.length} />
            {/* 店舗行 */}
            {filtered.map(store => (
              <StoreRow key={store.id} store={store} visibleMonths={visibleMonths} periodStarts={periodStarts} visItems={visItems} expanded={expanded.has(store.id)} onToggle={() => toggle(store.id)} priceChanges={priceChanges} />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function GrandTotalRow({ grandTotals, visibleMonths, periodStarts, visItems, filteredCount, totalCount }) {
  const [open, setOpen] = useState(true);
  return (
    <>
      <tr style={{ background:C.redDim, cursor:"pointer" }} onClick={() => setOpen(!open)}>
        <td style={{ ...stickyL(200,C.redDim,true), color:C.red, fontSize:12 }}>
          <span style={{ marginRight:6, fontSize:10, color:C.red }}>{open?"▼":"▶"}</span>
          全店合計
          {filteredCount < totalCount && <span style={{ fontSize:11, color:C.textMuted, marginLeft:5 }}>({filteredCount}店表示)</span>}
        </td>
        <td style={{ ...stickyL2(C.redDim), color:C.red, fontWeight:700 }}>合計</td>
        {visibleMonths.map(ym => {
          const isFirst = ym === periodStarts[getPeriodNum(ym)];
          const val = grandTotals[ym]?._total || 0;
          return (
            <td key={ym} style={{ ...numCell(isFirst,false,false), background:isFirst?C.periodBg:C.redDim, color:C.red, fontWeight:700 }}>
              {fmtK(val)}
            </td>
          );
        })}
      </tr>
      {open && visItems.map(item => (
        <tr key={item.id} style={{ background:"#FFF8F8" }}>
          <td style={{ ...stickyL(200,"#FFF8F8"), paddingLeft:24 }}></td>
          <td style={{ ...stickyL2("#FFF8F8") }}>
            <span style={{ display:"inline-block", width:6, height:6, borderRadius:3, background:item.color, marginRight:6, verticalAlign:"middle" }} />
            <span style={{ fontSize:11, color:item.color, fontWeight:500 }}>{item.label}</span>
          </td>
          {visibleMonths.map(ym => {
            const isFirst = ym === periodStarts[getPeriodNum(ym)];
            const val = grandTotals[ym]?.[item.id] || 0;
            return <td key={ym} style={{ ...numCell(isFirst,false,false), fontSize:11, color:C.textSub, background:"#FFF8F8" }}>{val>0?fmtK(val):""}</td>;
          })}
        </tr>
      ))}
    </>
  );
}

function StoreRow({ store, visibleMonths, periodStarts, visItems, expanded, onToggle, priceChanges }) {
  const pkgDef = PKG_DEFS[store.pkgKey] || PKG_DEFS.other;
  const isRetired = ym => store.retireYM && ym >= store.retireYM;
  const isContractEnd = ym => store.contractEnd && ym.substring(0,7) === store.contractEnd.substring(0,7);
  // 金額変更が発生する最初の月
  const storeChanges = (priceChanges||[]).filter(c => c.storeId === store.id);
  const hasChangeAt = ym => storeChanges.some(c => c.fromYM === ym);

  return (
    <>
      <tr style={{ cursor:"pointer" }} onClick={onToggle}
        onMouseEnter={e=>e.currentTarget.style.background=C.surfaceHigh}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <td style={{ ...stickyL(200,C.surface,true), fontSize:11 }}>
          <span style={{ marginRight:5, fontSize:9, color:C.textMuted }}>{expanded?"▼":"▶"}</span>
          {store.name}
          {store.retireYM && <span style={{ fontSize:8, marginLeft:4, padding:"1px 4px", borderRadius:3, background:`${C.red}20`, color:C.red }}>撤退</span>}
          {storeChanges.length > 0 && <span style={{ fontSize:8, marginLeft:4, padding:"1px 4px", borderRadius:3, background:`${C.purple}20`, color:C.purple }}>変更{storeChanges.length}件</span>}
        </td>
        <td style={{ ...stickyL2(C.surface), fontSize:10 }}>
          <span style={{ padding:"1px 5px", borderRadius:3, fontSize:9, background:`${pkgDef.color}20`, color:pkgDef.color, fontWeight:600 }}>{pkgDef.name}</span>
        </td>
        {visibleMonths.map(ym => {
          const isFirst = ym === periodStarts[getPeriodNum(ym)];
          const retired = isRetired(ym);
          const contractEnd = isContractEnd(ym);
          const changeStart = hasChangeAt(ym);
          const baseRow = store.monthly[ym];
          const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
          const total = row ? sumItems(row) : 0;
          return (
            <td key={ym} style={{ ...numCell(isFirst, retired, contractEnd || changeStart), fontWeight:700 }}>
              {retired
                ? <span style={{ fontSize:9, opacity:0.5 }}>撤退</span>
                : total > 0
                  ? <>
                      {contractEnd && <span style={{ fontSize:7, color:C.amber, marginRight:1 }}>🔔</span>}
                      {changeStart && <span style={{ fontSize:7, color:C.purple, marginRight:1 }}>↕</span>}
                      {fmtK(total)}
                    </>
                  : <span style={{ color:C.textMuted, fontSize:9 }}>—</span>
              }
            </td>
          );
        })}
      </tr>
      {/* 展開：項目別 */}
      {expanded && (
        <>
          {visItems.filter(it => !it.adOnly).map(item => (
            <tr key={item.id} style={{ background:`${C.surfaceHigh}50` }}>
              <td style={{ ...stickyL(200,`${C.surfaceHigh}50`), paddingLeft:24 }}></td>
              <td style={{ ...stickyL2(`${C.surfaceHigh}50`) }}>
                <span style={{ display:"inline-block", width:4, height:4, borderRadius:2, background:item.color, marginRight:5, verticalAlign:"middle" }} />
                <span style={{ fontSize:10, color:item.color }}>{item.label}</span>
              </td>
              {visibleMonths.map(ym => {
                const isFirst = ym === periodStarts[getPeriodNum(ym)];
                const retired = isRetired(ym);
                const baseRow = store.monthly[ym];
                const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
                const val = row?.[item.id] || 0;
                // この項目・この月に変更があるか
                const changed = storeChanges.some(c => c.itemId === item.id && ym >= c.fromYM);
                return (
                  <td key={ym} style={{ ...numCell(isFirst, retired, false), fontSize:10, color: changed ? C.purple : C.textSub }}>
                    {val > 0 ? fmtK(val) : ""}
                    {changed && hasChangeAt(ym) && <span style={{ fontSize:7, color:C.purple, marginLeft:1 }}>↕</span>}
                  </td>
                );
              })}
            </tr>
          ))}
          {/* 広告費：合計のみ */}
          {visItems.find(it => it.adOnly) && (() => {
            const adIt = ITEMS.find(it => it.adOnly);
            return (
              <tr key="ad" style={{ background:`${C.surfaceHigh}50` }}>
                <td style={{ ...stickyL(200,`${C.surfaceHigh}50`), paddingLeft:24 }}></td>
                <td style={{ ...stickyL2(`${C.surfaceHigh}50`) }}>
                  <span style={{ display:"inline-block", width:4, height:4, borderRadius:2, background:adIt.color, marginRight:5, verticalAlign:"middle" }} />
                  <span style={{ fontSize:10, color:adIt.color }}>{adIt.label}</span>
                  <span style={{ fontSize:8, color:C.textMuted, marginLeft:3 }}>(合計)</span>
                </td>
                {visibleMonths.map(ym => {
                  const isFirst = ym === periodStarts[getPeriodNum(ym)];
                  const row = store.monthly[ym];
                  const val = row?.ad || 0;
                  return <td key={ym} style={{ ...numCell(isFirst,false,false), fontSize:10, color:C.textMuted }}>{val > 0 ? fmtK(val) : ""}</td>;
                })}
              </tr>
            );
          })()}
          {/* 契約情報 */}
          <tr style={{ background:`${C.surfaceHigh}30` }}>
            <td style={{ ...stickyL(200,`${C.surfaceHigh}30`), paddingLeft:24, fontSize:9, color:C.textMuted }}>
              契約満了: {store.contractEnd||"—"} | 都道府県: {store.pref||"—"} | 契約: {store.contractYears||"—"} | 開店: {store.openDate||"—"}
            </td>
            <td style={{ ...stickyL2(`${C.surfaceHigh}30`) }}></td>
            {visibleMonths.map(ym => <td key={ym} style={{ borderLeft:`1px solid ${C.borderLight}`, borderBottom:`1px solid ${C.borderLight}20` }}></td>)}
          </tr>
        </>
      )}
    </>
  );
}

// ══════════════════════════════════════════════
// 期別サマリータブ
// ══════════════════════════════════════════════
function SummaryTab({ appData, periods, periodStarts, priceChanges }) {
  const { stores, months } = appData;

  const periodData = useMemo(() => periods.map(pn => {
    const pMonths = months.filter(ym => getPeriodNum(ym) === pn);
    const storeData = stores.map(store => {
      const tot = {};
      ITEMS.forEach(it => { tot[it.id] = 0; });
      tot._total = 0;
      let activeMonths = 0;
      pMonths.forEach(ym => {
        if (store.retireYM && ym >= store.retireYM) return;
        const baseRow = store.monthly[ym];
        if (!baseRow) return;
        const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
        activeMonths++;
        ITEMS.forEach(it => { tot[it.id] += row[it.id]||0; });
        tot._total += sumItems(row);
      });
      return { store, tot, activeMonths };
    });
    const ptot = {};
    ITEMS.forEach(it => { ptot[it.id] = storeData.reduce((s,d) => s+d.tot[it.id], 0); });
    ptot._total = storeData.reduce((s,d) => s+d.tot._total, 0);
    const retiredCount = stores.filter(s => s.retireYM && getPeriodNum(s.retireYM) === pn).length;
    const activeCount  = storeData.filter(d => d.activeMonths > 0).length;
    return { pn, pMonths, storeData, ptot, retiredCount, activeCount };
  }), [stores, months, periods, priceChanges]);

  const periodColors = [C.red, C.blue, C.green, C.purple, C.orange, C.teal];
  const maxPeriodTotal = Math.max(...periodData.map(p => p.ptot._total), 1);

  return (
    <div>
      {/* KPIカード */}
      <div style={{ display:"grid", gridTemplateColumns:`repeat(${Math.min(periodData.length,4)},1fr)`, gap:12, marginBottom:16 }}>
        {periodData.slice(0,4).map(({ pn, ptot, retiredCount, activeCount, pMonths }, pi) => (
          <div key={pn} style={{
            background:C.surface, border:`1px solid ${C.border}`,
            borderTop:`3px solid ${periodColors[pi%6]}`,
            borderRadius:10, padding:16,
            boxShadow:"0 1px 4px rgba(0,0,0,0.04)",
          }}>
            <div style={{ fontSize:11, color:C.textMuted, marginBottom:4 }}>{pn}年度（{pMonths[0]}〜{pMonths[pMonths.length-1]}）</div>
            <div style={{ fontSize:24, fontWeight:800, color:C.text, letterSpacing:"-0.5px" }}>{fmtM(ptot._total)}</div>
            <div style={{ display:"flex", gap:10, marginTop:6, marginBottom:10, fontSize:11 }}>
              <span style={{ color:C.green, fontWeight:500 }}>稼働 {activeCount}店</span>
              {retiredCount > 0 && <span style={{ color:C.red, fontWeight:500 }}>撤退 {retiredCount}店</span>}
            </div>
            {ITEMS.filter(it=>!it.adOnly).map(it => (
              <div key={it.id} style={{ display:"flex", justifyContent:"space-between", padding:"3px 0", borderBottom:`1px solid ${C.borderLight}` }}>
                <span style={{ fontSize:11, color:C.textSub, display:"flex", alignItems:"center", gap:5 }}>
                  <span style={{ display:"inline-block", width:6, height:6, borderRadius:3, background:it.color }} />{it.label}
                </span>
                <span style={{ fontSize:11, fontWeight:600, color:C.text }}>{fmtM(ptot[it.id])}</span>
              </div>
            ))}
            <div style={{ display:"flex", justifyContent:"space-between", padding:"3px 0", fontSize:11, color:C.textMuted }}>
              <span>広告費</span><span>{fmtM(ptot.ad)}</span>
            </div>
          </div>
        ))}
      </div>

      {/* 年度別棒グラフ */}
      <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, padding:16, marginBottom:16, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
        <div style={{ fontSize:13, fontWeight:700, marginBottom:16, color:C.text }}>年度別 収入推移</div>
        <div style={{ display:"flex", gap:10, alignItems:"flex-end", height:90 }}>
          {periodData.map(({ pn, ptot }, pi) => {
            const h = Math.round((ptot._total / maxPeriodTotal) * 80);
            return (
              <div key={pn} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", gap:6 }}>
                <div style={{ fontSize:11, color:C.textSub, fontWeight:500 }}>{fmtM(ptot._total)}</div>
                <div style={{
                  width:"100%", height:h,
                  background:periodColors[pi%6],
                  borderRadius:"4px 4px 0 0",
                  opacity:0.85,
                }} />
                <div style={{ fontSize:11, color:C.textSub, fontWeight:600 }}>{pn}年度</div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 店舗×期マトリクス */}
      <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, overflowX:"auto", boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
        <div style={{ padding:"12px 16px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.text }}>店舗別 年度別収入</div>
        <table style={{ borderCollapse:"collapse", width:"100%", fontSize:12 }}>
          <thead>
            <tr style={{ background:C.surfaceHigh }}>
              <th style={{ padding:"8px 14px", textAlign:"left", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}`, position:"sticky", left:0, background:C.surfaceHigh, minWidth:180 }}>店舗名</th>
              <th style={{ padding:"8px 14px", textAlign:"left", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}`, minWidth:80 }}>PKG</th>
              {periodData.map(({pn}) => (
                <th key={pn} style={{ padding:"8px 14px", textAlign:"right", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}`, minWidth:90 }}>{pn}年度</th>
              ))}
              <th style={{ padding:"8px 14px", textAlign:"right", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}` }}>累計</th>
            </tr>
          </thead>
          <tbody>
            {stores.map((store, i) => {
              const pTotals = periodData.map(p => p.storeData.find(d => d.store.id===store.id)?.tot._total||0);
              const grand = pTotals.reduce((s,v)=>s+v,0);
              const pkgDef = PKG_DEFS[store.pkgKey]||PKG_DEFS.other;
              return (
                <tr key={store.id} style={{ background:i%2===0?"transparent":`${C.surfaceHigh}40`, borderBottom:`1px solid ${C.borderLight}` }}>
                  <td style={{ padding:"5px 12px", position:"sticky", left:0, background:i%2===0?C.surface:`${C.surfaceHigh}80`, fontWeight:600, minWidth:180, whiteSpace:"nowrap" }}>
                    {store.name}
                    {store.retireYM && <span style={{ fontSize:8, marginLeft:4, padding:"1px 4px", borderRadius:3, background:`${C.red}15`, color:C.red }}>撤退</span>}
                  </td>
                  <td style={{ padding:"5px 12px" }}>
                    <span style={{ padding:"1px 5px", borderRadius:3, fontSize:9, background:`${pkgDef.color}20`, color:pkgDef.color, fontWeight:600 }}>{pkgDef.name}</span>
                  </td>
                  {pTotals.map((t, pi) => (
                    <td key={pi} style={{ padding:"5px 12px", textAlign:"right", color:t>0?C.text:C.textMuted }}>{t>0?fmtM(t):"—"}</td>
                  ))}
                  <td style={{ padding:"5px 12px", textAlign:"right", fontWeight:800, color:C.red }}>{fmtM(grand)}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{ background:C.redDim, borderTop:`2px solid ${C.redBorder}` }}>
              <td style={{ padding:"9px 14px", fontWeight:700, color:C.red, position:"sticky", left:0, background:C.redDim }}>合計</td>
              <td style={{ padding:"9px 14px" }}></td>
              {periodData.map(({ pn, ptot }) => (
                <td key={pn} style={{ padding:"9px 14px", textAlign:"right", fontWeight:700, color:C.red }}>{fmtM(ptot._total)}</td>
              ))}
              <td style={{ padding:"9px 14px", textAlign:"right", fontWeight:800, color:C.red }}>
                {fmtM(periodData.reduce((s,p)=>s+p.ptot._total,0))}
              </td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════
// 店舗マスタタブ（編集・新規追加対応）
// ══════════════════════════════════════════════
const EMPTY_STORE = () => ({
  id: `new_${Date.now()}`,
  storeId: "", memberId: "", company: "", name: "",
  pref: "", pkgKey: "norm", pkgName: "通常PKG",
  contractYears: "4年", renewalPeriod: "", openDate: "",
  lastRenewal: "", contractEnd: "", retireYM: "", retireRaw: "",
  monthly: {},
});

function MasterTab({ appData, priceChanges, setPriceChanges }) {
  const [localStores, setLocalStores] = useState(() => appData.stores);
  const [search,       setSearch]       = useState("");
  const [filterPref,   setFilterPref]   = useState("all");
  const [filterPkg,    setFilterPkg]    = useState("all");
  const [filterStatus, setFilterStatus] = useState("all");
  const [sortKey,      setSortKey]      = useState("name");
  const [sortDir,      setSortDir]      = useState(1);
  const [editId,       setEditId]       = useState(null);
  const [editForm,     setEditForm]     = useState({});
  const [showAddForm,  setShowAddForm]  = useState(false);
  const [addForm,      setAddForm]      = useState(EMPTY_STORE());
  const [saveMsg,      setSaveMsg]      = useState("");

  const prefs = useMemo(() => ["all", ...new Set(localStores.map(s=>s.pref).filter(Boolean))], [localStores]);

  const filtered = useMemo(() => {
    let arr = localStores.filter(s => {
      if (filterPref !== "all" && s.pref !== filterPref) return false;
      if (filterPkg  !== "all" && s.pkgKey !== filterPkg) return false;
      if (filterStatus === "active"  &&  s.retireYM) return false;
      if (filterStatus === "retired" && !s.retireYM) return false;
      if (search && !s.name.includes(search) && !s.company?.includes(search) && !s.memberId?.includes(search)) return false;
      return true;
    });
    return [...arr].sort((a,b) => {
      const va=a[sortKey]||"", vb=b[sortKey]||"";
      return va<vb?-sortDir:va>vb?sortDir:0;
    });
  }, [localStores, filterPref, filterPkg, filterStatus, search, sortKey, sortDir]);

  const toggleSort = key => { if(sortKey===key)setSortDir(d=>-d); else{setSortKey(key);setSortDir(1);} };

  // 編集開始
  const startEdit = s => { setEditId(s.id); setEditForm({...s}); setShowAddForm(false); };
  const cancelEdit = () => { setEditId(null); setEditForm({}); };

  // 保存
  const saveEdit = () => {
    const updated = { ...editForm,
      pkgName: PKG_DEFS[editForm.pkgKey]?.name || editForm.pkgKey,
    };
    setLocalStores(localStores.map(s => s.id===editId ? updated : s));
    setEditId(null); setEditForm({});
    flash("✓ 保存しました");
  };

  // 新規追加
  const addStore = () => {
    if (!addForm.name) { flash("店舗名を入力してください", true); return; }
    const newStore = { ...addForm,
      id: `new_${Date.now()}`,
      pkgName: PKG_DEFS[addForm.pkgKey]?.name || addForm.pkgKey,
      monthly: {},
    };
    setLocalStores([...localStores, newStore]);
    setAddForm(EMPTY_STORE());
    setShowAddForm(false);
    flash("✓ 店舗を追加しました");
  };

  const [retireModal, setRetireModal] = useState(null); // { storeId, ym }

  // 撤退登録・取消
  const toggleRetire = (s) => {
    if (s.retireYM) {
      setLocalStores(localStores.map(st => st.id===s.id ? {...st, retireYM:"", retireRaw:""} : st));
      flash("✓ 撤退を取り消しました");
    } else {
      setRetireModal({ storeId: s.id, name: s.name, ym: "" });
    }
  };

  const confirmRetire = () => {
    if (!retireModal || !/^\d{4}\/\d{2}$/.test(retireModal.ym)) {
      flash("⚠️ 年月の形式が正しくありません（例: 2026/04）", true); return;
    }
    setLocalStores(localStores.map(st => st.id===retireModal.storeId
      ? {...st, retireYM:retireModal.ym, retireRaw:retireModal.ym} : st));
    setRetireModal(null);
    flash("✓ 撤退日を登録しました");
  };

  const flash = (msg, err=false) => {
    setSaveMsg(msg);
    setTimeout(() => setSaveMsg(""), 2500);
  };

  // CSV エクスポート
  const exportCSV = () => {
    const headers = ["店舗名","会社名","加盟店ID","店舗ID","都道府県","PKG","契約期間","更新期","オープン日","直近契約更新日","契約満了日","撤退日","ステータス"];
    const rows = filtered.map(s => [
      s.name, s.company||"", s.memberId||"", s.storeId||"",
      s.pref||"", s.pkgName||"", s.contractYears||"",
      s.renewalPeriod||"", s.openDate||"", s.lastRenewal||"",
      s.contractEnd||"", s.retireRaw||"",
      s.retireYM ? "撤退" : "稼働中",
    ]);
    const bom = "\uFEFF";
    const csv = bom + [headers,...rows].map(r=>r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",")).join("\n");
    const blob = new Blob([csv],{type:"text/csv;charset=utf-8;"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download="store_master.csv"; a.click();
    URL.revokeObjectURL(url);
  };

  const pkgCount = useMemo(()=>{ const t={}; localStores.forEach(s=>{t[s.pkgKey]=(t[s.pkgKey]||0)+1;}); return t; },[localStores]);
  const prefCount = useMemo(()=>{ const t={}; localStores.forEach(s=>{if(s.pref)t[s.pref]=(t[s.pref]||0)+1;}); return Object.entries(t).sort((a,b)=>b[1]-a[1]); },[localStores]);
  const activeCount  = localStores.filter(s=>!s.retireYM).length;
  const retiredCount = localStores.filter(s=> s.retireYM).length;

  const SortTh = ({k,label,right}) => (
    <th onClick={()=>toggleSort(k)} style={{padding:"7px 10px",textAlign:right?"right":"left",fontSize:10,color:sortKey===k?C.red:C.textMuted,borderBottom:`1px solid ${C.border}`,cursor:"pointer",whiteSpace:"nowrap",userSelect:"none"}}>
      {label}{sortKey===k?(sortDir===1?" ▲":" ▼"):""}
    </th>
  );

  // インライン編集フィールド
  const EF = ({k,placeholder,type="text",w=120}) => (
    <input type={type} value={editForm[k]||""} placeholder={placeholder||k}
      onChange={e=>setEditForm({...editForm,[k]:e.target.value})}
      style={{...inputSt,width:w,fontSize:10}} />
  );
  const AF = ({k,placeholder,type="text",w=120}) => (
    <input type={type} value={addForm[k]||""} placeholder={placeholder||k}
      onChange={e=>setAddForm({...addForm,[k]:e.target.value})}
      style={{...inputSt,width:w,fontSize:10}} />
  );

  return (
    <div>
      {/* 撤退登録モーダル */}
      {retireModal && (
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:24,width:320}}>
            <div style={{fontSize:13,fontWeight:700,marginBottom:4}}>撤退日を登録</div>
            <div style={{fontSize:11,color:C.textMuted,marginBottom:16}}>{retireModal.name}</div>
            <div style={{fontSize:11,color:C.textMuted,marginBottom:6}}>撤退年月（例: 2026/04）</div>
            <input value={retireModal.ym} onChange={e=>setRetireModal({...retireModal,ym:e.target.value})}
              placeholder="2026/04" style={{...inputSt,width:"100%",marginBottom:16}}
              onKeyDown={e=>e.key==="Enter"&&confirmRetire()} />
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button onClick={()=>setRetireModal(null)} style={{...btnSt(C.textMuted),fontSize:11}}>キャンセル</button>
              <button onClick={confirmRetire} style={{...btnSt(C.red),fontSize:11}}>撤退登録</button>
            </div>
          </div>
        </div>
      )}
      {/* KPI */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:14}}>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.amber}`,borderRadius:8,padding:"12px 14px"}}>
          <div style={{fontSize:10,color:C.textMuted,marginBottom:4}}>総店舗数</div>
          <div style={{fontSize:26,fontWeight:800,color:C.red}}>{localStores.length}<span style={{fontSize:13,color:C.textMuted}}>店</span></div>
          <div style={{fontSize:10,color:C.textMuted,marginTop:2}}>XLSXから{appData.stores.length}店 読込 + 手動{localStores.length-appData.stores.length}店追加</div>
        </div>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.green}`,borderRadius:8,padding:"12px 14px"}}>
          <div style={{fontSize:10,color:C.textMuted,marginBottom:4}}>稼働中</div>
          <div style={{fontSize:26,fontWeight:800,color:C.green}}>{activeCount}<span style={{fontSize:13,color:C.textMuted}}>店</span></div>
        </div>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.red}`,borderRadius:8,padding:"12px 14px"}}>
          <div style={{fontSize:10,color:C.textMuted,marginBottom:4}}>撤退済み</div>
          <div style={{fontSize:26,fontWeight:800,color:retiredCount>0?C.red:C.textMuted}}>{retiredCount}<span style={{fontSize:13,color:C.textMuted}}>店</span></div>
        </div>
        <div style={{background:C.surface,border:`1px solid ${C.border}`,borderTop:`3px solid ${C.blue}`,borderRadius:8,padding:"12px 14px"}}>
          <div style={{fontSize:10,color:C.textMuted,marginBottom:4}}>PKG内訳</div>
          {Object.entries(pkgCount).map(([k,v])=>{
            const def=PKG_DEFS[k]||PKG_DEFS.other;
            return <div key={k} style={{display:"flex",justifyContent:"space-between",fontSize:10,padding:"1px 0"}}><span style={{color:def.color}}>{def.name}</span><span style={{fontWeight:700}}>{v}店</span></div>;
          })}
        </div>
      </div>

      {/* フィルター + ボタン */}
      <div style={{display:"flex",gap:8,marginBottom:10,alignItems:"center",flexWrap:"wrap",padding:"8px 10px",background:C.surface,borderRadius:7,border:`1px solid ${C.border}`}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="店舗名・会社名・IDで検索..."
          style={{...inputSt,width:200}} />
        <select value={filterPref} onChange={e=>setFilterPref(e.target.value)} style={selSt}>
          <option value="all">全都道府県</option>
          {prefs.slice(1).map(p=><option key={p} value={p}>{p}</option>)}
        </select>
        <select value={filterPkg} onChange={e=>setFilterPkg(e.target.value)} style={selSt}>
          <option value="all">全PKG</option>
          {Object.entries(PKG_DEFS).map(([k,v])=><option key={k} value={k}>{v.name}</option>)}
        </select>
        <div style={{display:"flex",gap:4}}>
          {[{v:"all",l:"全店舗"},{v:"active",l:"稼働中"},{v:"retired",l:"撤退"}].map(({v,l})=>(
            <button key={v} onClick={()=>setFilterStatus(v)} style={pillSt(filterStatus===v)}>{l}</button>
          ))}
        </div>
        <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
          {saveMsg && <span style={{fontSize:11,color:C.green,fontWeight:600}}>{saveMsg}</span>}
          <span style={{fontSize:10,color:C.textMuted}}>{filtered.length}件</span>
          <button onClick={()=>{setShowAddForm(!showAddForm);setEditId(null);}} style={{...btnSt(C.green),fontSize:10}}>
            ＋ 新規店舗追加
          </button>
          <button onClick={exportCSV} style={{...btnSt(C.blue),fontSize:10}}>↓ CSVエクスポート</button>
        </div>
      </div>

      {/* 新規追加フォーム */}
      {showAddForm && (
        <div style={{marginBottom:12,padding:14,background:`rgba(53,201,138,0.06)`,border:`1px solid ${C.green}30`,borderRadius:8}}>
          <div style={{fontSize:12,fontWeight:700,color:C.green,marginBottom:10}}>＋ 新規店舗オープン登録</div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
            {/* 行1 */}
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>店舗名 <span style={{color:C.red}}>*</span></div><AF k="name" placeholder="カーセブン○○店" w={160}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>会社名</div><AF k="company" placeholder="株式会社○○" w={150}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>加盟店ID</div><AF k="memberId" placeholder="B1030XXX" w={100}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>店舗ID</div><AF k="storeId" placeholder="V1030XXX" w={100}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>都道府県</div><AF k="pref" placeholder="東京都" w={80}/></div>
            <div>
              <div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>PKG</div>
              <select value={addForm.pkgKey} onChange={e=>setAddForm({...addForm,pkgKey:e.target.value,pkgName:PKG_DEFS[e.target.value]?.name||""})} style={{...selSt,fontSize:10}}>
                {Object.entries(PKG_DEFS).map(([k,v])=><option key={k} value={k}>{v.name}</option>)}
              </select>
            </div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>契約期間</div>
              <select value={addForm.contractYears} onChange={e=>setAddForm({...addForm,contractYears:e.target.value})} style={{...selSt,fontSize:10}}>
                {["4年","5年","3年","その他"].map(y=><option key={y} value={y}>{y}</option>)}
              </select>
            </div>
          </div>
          <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end",marginTop:10}}>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>オープン日</div><AF k="openDate" placeholder="2026/04" w={90}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>直近契約更新日</div><AF k="lastRenewal" placeholder="2026/04/01" w={100}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>契約満了日</div><AF k="contractEnd" placeholder="2030/03" w={90}/></div>
            <div><div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>更新期</div><AF k="renewalPeriod" placeholder="31期" w={70}/></div>
            <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"flex-end"}}>
              <button onClick={addStore} style={{...btnSt(C.green),fontSize:11,padding:"6px 20px"}}>追加</button>
              <button onClick={()=>setShowAddForm(false)} style={{...btnSt(C.textMuted),fontSize:10}}>キャンセル</button>
            </div>
          </div>
        </div>
      )}

      {/* テーブル */}
      <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,overflowX:"auto",maxHeight:"calc(100vh - 320px)",overflowY:"auto"}}>
        <table style={{borderCollapse:"collapse",width:"100%",fontSize:11}}>
          <thead style={{position:"sticky",top:0,zIndex:5}}>
            <tr style={{background:C.surfaceHigh}}>
              <SortTh k="name"          label="店舗名" />
              <SortTh k="company"       label="会社名" />
              <SortTh k="memberId"      label="加盟店ID" />
              <SortTh k="pref"          label="都道府県" />
              <SortTh k="pkgName"       label="PKG" />
              <SortTh k="contractYears" label="契約期間" right />
              <SortTh k="openDate"      label="オープン日" right />
              <SortTh k="lastRenewal"   label="直近更新日" right />
              <SortTh k="contractEnd"   label="契約満了日" right />
              <th style={{padding:"7px 10px",fontSize:10,color:C.textMuted,borderBottom:`1px solid ${C.border}`}}>ステータス</th>
              <th style={{padding:"7px 10px",fontSize:10,color:C.textMuted,borderBottom:`1px solid ${C.border}`,minWidth:120}}>操作</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((s, i) => {
              const pkgDef = PKG_DEFS[s.pkgKey]||PKG_DEFS.other;
              const isRetired = !!s.retireYM;
              const nearExpiry = (()=>{
                if(!s.contractEnd||isRetired) return false;
                const [ey,em]=s.contractEnd.split("/").map(Number);
                const diffM=(new Date(ey,em-1,1)-new Date())/(1000*60*60*24*30);
                return diffM>0&&diffM<=12;
              })();

              // 編集行
              if (editId === s.id) return (
                <>
                <tr key={s.id} style={{background:`rgba(212,168,67,0.06)`}}>
                  <td style={{padding:"4px 8px"}}><EF k="name" w={140}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="company" w={120}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="memberId" w={90}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="pref" w={70}/></td>
                  <td style={{padding:"4px 6px"}}>
                    <select value={editForm.pkgKey||"norm"} onChange={e=>setEditForm({...editForm,pkgKey:e.target.value})} style={{...selSt,fontSize:10}}>
                      {Object.entries(PKG_DEFS).map(([k,v])=><option key={k} value={k}>{v.name}</option>)}
                    </select>
                  </td>
                  <td style={{padding:"4px 6px"}}>
                    <select value={editForm.contractYears||"4年"} onChange={e=>setEditForm({...editForm,contractYears:e.target.value})} style={{...selSt,fontSize:10,width:70}}>
                      {["4年","5年","3年","その他"].map(y=><option key={y} value={y}>{y}</option>)}
                    </select>
                  </td>
                  <td style={{padding:"4px 6px"}}><EF k="openDate" placeholder="2026/04" w={85}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="lastRenewal" placeholder="2026/04/01" w={95}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="contractEnd" placeholder="2030/03" w={85}/></td>
                  <td style={{padding:"4px 6px"}}><EF k="retireYM" placeholder="撤退年月(任意)" w={95}/></td>
                  <td style={{padding:"4px 8px",display:"flex",gap:4}}>
                    <button onClick={saveEdit} style={{...btnSt(C.amber),fontSize:10,padding:"3px 10px"}}>保存</button>
                    <button onClick={cancelEdit} style={{...btnSt(C.textMuted),fontSize:10,padding:"3px 8px"}}>取消</button>
                  </td>
                </tr>
                <PriceChangePanel key={`pc_${s.id}`} store={s} priceChanges={priceChanges} setPriceChanges={setPriceChanges} months={appData.months} />
                </>
              );

              // 表示行
              return (
                <tr key={s.id} style={{background:i%2===0?"transparent":`${C.surfaceHigh}40`,borderBottom:`1px solid ${C.borderLight}`,opacity:isRetired?0.65:1}}>
                  <td style={{padding:"5px 10px",fontWeight:600,whiteSpace:"nowrap"}}>
                    {s.name}
                    {!s.storeId && <span style={{fontSize:8,marginLeft:4,padding:"1px 4px",borderRadius:3,background:`${C.green}20`,color:C.green}}>新規</span>}
                    {(priceChanges||[]).filter(c=>c.storeId===s.id).length > 0 && (
                      <span style={{fontSize:8,marginLeft:4,padding:"1px 4px",borderRadius:3,background:`${C.purple}20`,color:C.purple}}>
                        変更{(priceChanges||[]).filter(c=>c.storeId===s.id).length}件
                      </span>
                    )}
                  </td>
                  <td style={{padding:"5px 10px",color:C.textSub,fontSize:10,whiteSpace:"nowrap"}}>{s.company||"—"}</td>
                  <td style={{padding:"5px 10px",color:C.textMuted,fontSize:10}}>{s.memberId||"—"}</td>
                  <td style={{padding:"5px 10px",fontSize:10}}>{s.pref||"—"}</td>
                  <td style={{padding:"5px 10px"}}>
                    <span style={{padding:"1px 6px",borderRadius:3,fontSize:9,fontWeight:600,background:`${pkgDef.color}20`,color:pkgDef.color}}>{pkgDef.name}</span>
                  </td>
                  <td style={{padding:"5px 10px",textAlign:"right",color:C.textSub,fontSize:10}}>{s.contractYears||"—"}</td>
                  <td style={{padding:"5px 10px",textAlign:"right",color:C.textSub,fontSize:10}}>{s.openDate||"—"}</td>
                  <td style={{padding:"5px 10px",textAlign:"right",color:C.textSub,fontSize:10}}>{s.lastRenewal||"—"}</td>
                  <td style={{padding:"5px 10px",textAlign:"right",fontSize:10}}>
                    <span style={{color:nearExpiry?C.amber:isRetired?C.textMuted:C.text,fontWeight:nearExpiry?700:400}}>
                      {s.contractEnd||"—"}
                      {nearExpiry&&<span style={{fontSize:8,marginLeft:3,color:C.amber}}>🔔</span>}
                    </span>
                  </td>
                  <td style={{padding:"5px 10px"}}>
                    {isRetired
                      ? <span style={{padding:"2px 7px",borderRadius:4,fontSize:9,fontWeight:700,background:`${C.red}15`,color:C.red}}>撤退 {s.retireYM}</span>
                      : <span style={{padding:"2px 7px",borderRadius:4,fontSize:9,fontWeight:700,background:`${C.green}15`,color:C.green}}>稼働中</span>
                    }
                  </td>
                  <td style={{padding:"5px 8px"}}>
                    <div style={{display:"flex",gap:4}}>
                      <button onClick={()=>startEdit(s)} style={{...btnSt(C.blue),fontSize:9,padding:"2px 8px"}}>編集</button>
                      <button onClick={()=>toggleRetire(s)} style={{...btnSt(isRetired?C.green:C.red),fontSize:9,padding:"2px 8px"}}>
                        {isRetired?"撤退解除":"撤退登録"}
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {filtered.length===0&&(
          <div style={{padding:32,textAlign:"center",color:C.textMuted,fontSize:12}}>条件に一致する店舗がありません</div>
        )}
      </div>

      {/* 都道府県別 */}
      <div style={{marginTop:14,background:C.surface,border:`1px solid ${C.border}`,borderRadius:8,padding:14}}>
        <div style={{fontSize:12,fontWeight:700,marginBottom:10}}>都道府県別 店舗数</div>
        <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
          {prefCount.map(([pref,count])=>(
            <div key={pref} onClick={()=>setFilterPref(filterPref===pref?"all":pref)} style={{
              padding:"4px 10px",borderRadius:6,cursor:"pointer",fontSize:11,
              background:filterPref===pref?`${C.amber}20`:C.surfaceHigh,
              border:`1px solid ${filterPref===pref?C.amber:C.border}`,
              color:filterPref===pref?C.red:C.textSub,
            }}>
              {pref} <span style={{fontWeight:700}}>{count}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ══════════════════════════════════════════════
// 金額変更履歴パネル（MasterTabの編集行直下に表示）
// ══════════════════════════════════════════════
function PriceChangePanel({ store, priceChanges, setPriceChanges, months }) {
  const storeChanges = (priceChanges||[]).filter(c => c.storeId === store.id)
    .sort((a,b) => a.fromYM < b.fromYM ? -1 : 1);

  const [newChange, setNewChange] = useState({ fromYM:"", itemId:"royalty", newValue:"" });

  const addChange = () => {
    if (!newChange.fromYM || !newChange.newValue) return;
    if (!/^\d{4}\/\d{2}$/.test(newChange.fromYM)) return;
    const entry = {
      id: `pc_${Date.now()}`,
      storeId: store.id,
      fromYM: newChange.fromYM,
      itemId: newChange.itemId,
      newValue: Number(newChange.newValue),
    };
    setPriceChanges([...(priceChanges||[]), entry]);
    setNewChange({ fromYM:"", itemId:"royalty", newValue:"" });
  };

  const removeChange = id => setPriceChanges((priceChanges||[]).filter(c => c.id !== id));

  const itemLabel = id => ITEMS.find(it=>it.id===id)?.label || id;

  return (
    <tr style={{background:`rgba(155,126,212,0.05)`}}>
      <td colSpan={11} style={{padding:"10px 16px",borderBottom:`1px solid ${C.border}`}}>
        <div style={{fontSize:11,fontWeight:700,color:C.purple,marginBottom:8}}>
          ↕ 金額変更履歴 — {store.name}
          <span style={{fontSize:10,fontWeight:400,color:C.textMuted,marginLeft:8}}>
            指定月以降の金額をシミュレーションに反映します
          </span>
        </div>

        {/* 既存の変更履歴 */}
        {storeChanges.length > 0 && (
          <div style={{marginBottom:10}}>
            <table style={{borderCollapse:"collapse",fontSize:11,width:"auto"}}>
              <thead>
                <tr style={{background:C.surfaceHigh}}>
                  {["変更開始月","収入項目","変更後金額",""].map((h,i)=>(
                    <th key={i} style={{padding:"4px 10px",textAlign:i===2?"right":"left",fontSize:10,color:C.textMuted,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap"}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {storeChanges.map(c => (
                  <tr key={c.id} style={{borderBottom:`1px solid ${C.borderLight}`}}>
                    <td style={{padding:"4px 10px",color:C.red,fontWeight:600}}>{c.fromYM}〜</td>
                    <td style={{padding:"4px 10px"}}>
                      <span style={{
                        padding:"1px 6px",borderRadius:3,fontSize:10,
                        background:`${ITEMS.find(it=>it.id===c.itemId)?.color||C.textMuted}20`,
                        color:ITEMS.find(it=>it.id===c.itemId)?.color||C.textMuted,
                        fontWeight:600,
                      }}>{itemLabel(c.itemId)}</span>
                    </td>
                    <td style={{padding:"4px 10px",textAlign:"right",fontWeight:700,color:C.text}}>
                      {fmtFull(c.newValue)}
                    </td>
                    <td style={{padding:"4px 8px"}}>
                      <button onClick={()=>removeChange(c.id)}
                        style={{padding:"2px 8px",background:`${C.red}15`,border:`1px solid ${C.red}40`,borderRadius:3,color:C.red,fontSize:10,cursor:"pointer"}}>
                        削除
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* 新規追加フォーム */}
        <div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>変更開始年月</div>
            <input value={newChange.fromYM} onChange={e=>setNewChange({...newChange,fromYM:e.target.value})}
              placeholder="2026/04" style={{...inputSt,width:85}} />
          </div>
          <div>
            <div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>収入項目</div>
            <select value={newChange.itemId} onChange={e=>setNewChange({...newChange,itemId:e.target.value})} style={{...selSt}}>
              {ITEMS.map(it=><option key={it.id} value={it.id}>{it.label}</option>)}
            </select>
          </div>
          <div>
            <div style={{fontSize:9,color:C.textMuted,marginBottom:2}}>変更後金額（円/月）</div>
            <input type="number" value={newChange.newValue} onChange={e=>setNewChange({...newChange,newValue:e.target.value})}
              placeholder="例: 120000" style={{...inputSt,width:110,textAlign:"right"}} />
          </div>
          <button onClick={addChange} style={{...btnSt(C.purple),fontSize:10,padding:"5px 14px"}}>
            ＋ 変更を追加
          </button>
        </div>
        {storeChanges.length === 0 && (
          <div style={{fontSize:10,color:C.textMuted,marginTop:6}}>
            変更履歴がありません。上のフォームから追加してください。
          </div>
        )}
      </td>
    </tr>
  );
}
