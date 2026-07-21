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

// 上期（10〜3月）か下期（4〜9月）か
function getHalf(ym) {
  const m = parseInt(ym.substring(5,7));
  return m >= 10 || m <= 3 ? 'first' : 'second'; // first=上期, second=下期
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
  initialActuals      = {},
  onXLSXLoaded        = null,
  onStoreUpdate       = null,
  onStoreDelete       = null,
  onPriceChangeAdd    = null,
  onPriceChangeDelete = null,
  onActualsSave       = null,
}) {
  const [tab, setTab]             = useState(initialData ? "sim" : "upload");
  const [appData, setAppData]     = useState(initialData);
  const [showItems, setShowItems] = useState(Object.fromEntries(ITEMS.map(it=>[it.id,true])));
  // 金額変更履歴
  const [priceChanges, setPriceChanges] = useState(initialPriceChanges);
  // 実績データ: { storeId: { ym: { royalty,sv,... } } }
  const [actuals, setActuals] = useState(initialActuals);

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

  const shared = { appData, periods, periodStarts, showItems, setShowItems, priceChanges, actuals, setActuals };

  return (
    <div style={{ fontFamily:"'Inter','Noto Sans JP',sans-serif", background:C.bg, minHeight:"100vh", color:C.text, overflowX:"hidden" }}>
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
          { id:"upload",  label:"データ読込" },
          { id:"sim",     label:"売上明細",    disabled:!appData },
          { id:"summary", label:"期別サマリー", disabled:!appData },
          { id:"budget",  label:"予実管理",    disabled:!appData },
          { id:"master",  label:"店舗マスタ",  disabled:!appData },
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

      <div style={{ padding:"12px 8px" }}>
        {tab === "upload"  && <UploadTab onLoaded={onLoaded} />}
        {tab === "sim"     && appData && <SimTab     {...shared} />}
        {tab === "summary" && appData && <SummaryTab {...shared} />}
        {tab === "budget"  && appData && <BudgetTab  {...shared} onActualsSave={onActualsSave} />}
        {tab === "master"  && appData && <MasterTab  appData={appData} priceChanges={priceChanges} setPriceChanges={setPriceChanges} onStoreDelete={onStoreDelete} />}
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

  const [expanded,    setExpanded]    = useState(new Set());
  const [filterPref,  setFilterPref]  = useState("all");
  const [filterPkg,   setFilterPkg]   = useState("all");
  const [search,      setSearch]      = useState("");
  const [filterPeriod, setFilterPeriod] = useState(null);
  const [filterHalf,  setFilterHalf]  = useState(null); // null=全期, 'first'=上期, 'second'=下期

  // 表示月を絞り込み
  const visibleMonths = useMemo(() => {
    let ms = months;
    if (filterPeriod) ms = ms.filter(ym => getPeriodNum(ym) === filterPeriod);
    if (filterHalf)   ms = ms.filter(ym => getHalf(ym) === filterHalf);
    return ms;
  }, [months, filterPeriod, filterHalf]);

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

  // 列幅リサイズ
  const [colWidths, setColWidths] = useState({ name:200, item:120, num:96 });
  const resizing = useRef(null);

  const onResizeStart = (col, e) => {
    e.preventDefault();
    resizing.current = { col, startX: e.clientX, startW: colWidths[col] };
    const onMove = (ev) => {
      const dx = ev.clientX - resizing.current.startX;
      setColWidths(w => ({ ...w, [resizing.current.col]: Math.max(60, resizing.current.startW + dx) }));
    };
    const onUp = () => {
      resizing.current = null;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  };

  const ResizeHandle = ({ col }) => (
    <span onMouseDown={e => onResizeStart(col, e)} style={{
      position:'absolute', right:0, top:0, bottom:0, width:6,
      cursor:'col-resize', userSelect:'none',
      display:'flex', alignItems:'center', justifyContent:'center',
    }}>
      <span style={{ width:2, height:14, borderRadius:1, background:C.border }} />
    </span>
  );

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
        <div style={{ display:"flex", gap:4, alignItems:"center" }}>
          <span style={{ fontSize:11, color:C.textMuted }}>半期：</span>
          <button onClick={() => setFilterHalf(null)} style={pillSt(!filterHalf)}>全期</button>
          <button onClick={() => setFilterHalf(filterHalf==='first'?null:'first')} style={pillSt(filterHalf==='first')}>上期（10〜3月）</button>
          <button onClick={() => setFilterHalf(filterHalf==='second'?null:'second')} style={pillSt(filterHalf==='second')}>下期（4〜9月）</button>
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
        overflowX:"scroll", overflowY:"auto",
        maxHeight:"calc(100vh - 168px)",
        border:`1px solid ${C.border}`, borderRadius:8,
        background:C.surface,
        boxShadow:"0 1px 4px rgba(0,0,0,0.05)",
        WebkitOverflowScrolling:"touch",
      }}>
        <table style={{ borderCollapse:"collapse", whiteSpace:"nowrap", tableLayout:"fixed" }}>
          <thead>
            {/* 期ラベル行 */}
            <tr style={{ background:C.surfaceHigh }}>
              <th style={{ ...stickyL(colWidths.name,C.surfaceHigh,true), position:"sticky", left:0, top:0, zIndex:12, height:22, lineHeight:"22px", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}`, overflow:'visible' }}>
                店舗名
                <ResizeHandle col="name" />
              </th>
              <th style={{ ...stickyL2(C.surfaceHigh), width:colWidths.item, minWidth:colWidths.item, maxWidth:colWidths.item, position:"sticky", left:colWidths.name, top:0, zIndex:12, height:22, lineHeight:"22px", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}`, overflow:'visible' }}>
                項目
                <ResizeHandle col="item" />
              </th>
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
                    minWidth:colWidths.num, padding:"0 4px",
                    position:"sticky", top:0, zIndex:8,
                  }}>
                    {isFirst ? `${pn}年度` : ""}
                  </th>
                );
              })}
            </tr>
            {/* 年月行 */}
            <tr style={{ background:C.surface }}>
              <th style={{ ...stickyL(colWidths.name,C.surface,false), position:"sticky", left:0, top:22, zIndex:12, height:22, lineHeight:"22px", borderBottom:`1px solid ${C.border}` }}></th>
              <th style={{ ...stickyL2(C.surface), width:colWidths.item, minWidth:colWidths.item, maxWidth:colWidths.item, position:"sticky", left:colWidths.name, top:22, zIndex:12, height:22, lineHeight:"22px", borderBottom:`1px solid ${C.border}` }}></th>
              {visibleMonths.map(ym => {
                const isFirst = ym === periodStarts[getPeriodNum(ym)];
                const mo = ym.substring(5);
                return (
                  <th key={ym} style={{ height:22, lineHeight:"22px", textAlign:"center", fontSize:9,
                    color: isFirst ? C.red : C.textMuted, fontWeight: isFirst ? 600 : 400,
                    background: isFirst ? C.periodBg : C.surface,
                    borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}`,
                    borderBottom:`1px solid ${C.border}`,
                    minWidth:colWidths.num,
                    position:"sticky", top:22, zIndex:8,
                  }}>{mo}月</th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {/* 全店合計 */}
            <GrandTotalRow grandTotals={grandTotals} visibleMonths={visibleMonths} periodStarts={periodStarts} visItems={visItems} filteredCount={filtered.length} totalCount={stores.length} colWidths={colWidths} />
            {/* 店舗行 */}
            {filtered.map(store => (
              <StoreRow key={store.id} store={store} visibleMonths={visibleMonths} periodStarts={periodStarts} visItems={visItems} expanded={expanded.has(store.id)} onToggle={() => toggle(store.id)} priceChanges={priceChanges} colWidths={colWidths} />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function GrandTotalRow({ grandTotals, visibleMonths, periodStarts, visItems, filteredCount, totalCount, colWidths }) {
  const cw = colWidths || { name:200, item:120, num:96 };
  const [open, setOpen] = useState(true);
  return (
    <>
      <tr style={{ background:C.redDim, cursor:"pointer" }} onClick={() => setOpen(!open)}>
        <td style={{ ...stickyL(cw.name,C.redDim,true), color:C.red, fontSize:12 }}>
          <span style={{ marginRight:6, fontSize:10, color:C.red }}>{open?"▼":"▶"}</span>
          全店合計
          {filteredCount < totalCount && <span style={{ fontSize:11, color:C.textMuted, marginLeft:5 }}>({filteredCount}店表示)</span>}
        </td>
        <td style={{ ...stickyL2(C.redDim), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name, color:C.red, fontWeight:700 }}>合計</td>
        {visibleMonths.map(ym => {
          const isFirst = ym === periodStarts[getPeriodNum(ym)];
          const val = grandTotals[ym]?._total || 0;
          return (
            <td key={ym} style={{ ...numCell(isFirst,false,false), minWidth:cw.num, background:isFirst?C.periodBg:C.redDim, color:C.red, fontWeight:700 }}>
              {fmtK(val)}
            </td>
          );
        })}
      </tr>
      {open && visItems.map(item => (
        <tr key={item.id} style={{ background:"#FFF8F8" }}>
          <td style={{ ...stickyL(cw.name,"#FFF8F8"), paddingLeft:24 }}></td>
          <td style={{ ...stickyL2("#FFF8F8"), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name }}>
            <span style={{ display:"inline-block", width:6, height:6, borderRadius:3, background:item.color, marginRight:6, verticalAlign:"middle" }} />
            <span style={{ fontSize:11, color:item.color, fontWeight:500 }}>{item.label}</span>
          </td>
          {visibleMonths.map(ym => {
            const isFirst = ym === periodStarts[getPeriodNum(ym)];
            const val = grandTotals[ym]?.[item.id] || 0;
            return <td key={ym} style={{ ...numCell(isFirst,false,false), minWidth:cw.num, fontSize:11, color:C.textSub, background:"#FFF8F8" }}>{val>0?fmtK(val):""}</td>;
          })}
        </tr>
      ))}
    </>
  );
}

function StoreRow({ store, visibleMonths, periodStarts, visItems, expanded, onToggle, priceChanges, colWidths }) {
  const cw = colWidths || { name:200, item:120, num:96 };
  const pkgDef = PKG_DEFS[store.pkgKey] || PKG_DEFS.other;
  const isRetired = ym => store.retireYM && ym >= store.retireYM;
  const isContractEnd = ym => store.contractEnd && ym.substring(0,7) === store.contractEnd.substring(0,7);
  const storeChanges = (priceChanges||[]).filter(c => c.storeId === store.id);
  const hasChangeAt = ym => storeChanges.some(c => c.fromYM === ym);

  return (
    <>
      <tr style={{ cursor:"pointer" }} onClick={onToggle}
        onMouseEnter={e=>e.currentTarget.style.background=C.surfaceHigh}
        onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
        <td style={{ ...stickyL(cw.name,C.surface,true), fontSize:11 }}>
          <span style={{ marginRight:5, fontSize:9, color:C.textMuted }}>{expanded?"▼":"▶"}</span>
          {store.name}
          {store.retireYM && <span style={{ fontSize:8, marginLeft:4, padding:"1px 4px", borderRadius:3, background:`${C.red}20`, color:C.red }}>撤退</span>}
          {storeChanges.length > 0 && <span style={{ fontSize:8, marginLeft:4, padding:"1px 4px", borderRadius:3, background:`${C.purple}20`, color:C.purple }}>変更{storeChanges.length}件</span>}
        </td>
        <td style={{ ...stickyL2(C.surface), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name, fontSize:10 }}>
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
            <td key={ym} style={{ ...numCell(isFirst, retired, contractEnd || changeStart), minWidth:cw.num, fontWeight:700 }}>
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
      {expanded && (
        <>
          {visItems.filter(it => !it.adOnly).map(item => (
            <tr key={item.id} style={{ background:`${C.surfaceHigh}50` }}>
              <td style={{ ...stickyL(cw.name,`${C.surfaceHigh}50`), paddingLeft:24 }}></td>
              <td style={{ ...stickyL2(`${C.surfaceHigh}50`), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name }}>
                <span style={{ display:"inline-block", width:4, height:4, borderRadius:2, background:item.color, marginRight:5, verticalAlign:"middle" }} />
                <span style={{ fontSize:10, color:item.color }}>{item.label}</span>
              </td>
              {visibleMonths.map(ym => {
                const isFirst = ym === periodStarts[getPeriodNum(ym)];
                const retired = isRetired(ym);
                const baseRow = store.monthly[ym];
                const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
                const val = row?.[item.id] || 0;
                const changed = storeChanges.some(c => c.itemId === item.id && ym >= c.fromYM);
                return (
                  <td key={ym} style={{ ...numCell(isFirst, retired, false), minWidth:cw.num, fontSize:10, color: changed ? C.purple : C.textSub }}>
                    {val > 0 ? fmtK(val) : ""}
                    {changed && hasChangeAt(ym) && <span style={{ fontSize:7, color:C.purple, marginLeft:1 }}>↕</span>}
                  </td>
                );
              })}
            </tr>
          ))}
          {visItems.find(it => it.adOnly) && (() => {
            const adIt = ITEMS.find(it => it.adOnly);
            return (
              <tr key="ad" style={{ background:`${C.surfaceHigh}50` }}>
                <td style={{ ...stickyL(cw.name,`${C.surfaceHigh}50`), paddingLeft:24 }}></td>
                <td style={{ ...stickyL2(`${C.surfaceHigh}50`), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name }}>
                  <span style={{ display:"inline-block", width:4, height:4, borderRadius:2, background:adIt.color, marginRight:5, verticalAlign:"middle" }} />
                  <span style={{ fontSize:10, color:adIt.color }}>{adIt.label}</span>
                  <span style={{ fontSize:8, color:C.textMuted, marginLeft:3 }}>(合計)</span>
                </td>
                {visibleMonths.map(ym => {
                  const isFirst = ym === periodStarts[getPeriodNum(ym)];
                  const row = store.monthly[ym];
                  const val = row?.ad || 0;
                  return <td key={ym} style={{ ...numCell(isFirst,false,false), minWidth:cw.num, fontSize:10, color:C.textMuted }}>{val > 0 ? fmtK(val) : ""}</td>;
                })}
              </tr>
            );
          })()}
          <tr style={{ background:`${C.surfaceHigh}30` }}>
            <td style={{ ...stickyL(cw.name,`${C.surfaceHigh}30`), paddingLeft:24, fontSize:9, color:C.textMuted }}>
              契約満了: {store.contractEnd||"—"} | 都道府県: {store.pref||"—"} | 契約: {store.contractYears||"—"} | 開店: {store.openDate||"—"}
            </td>
            <td style={{ ...stickyL2(`${C.surfaceHigh}30`), width:cw.item, minWidth:cw.item, maxWidth:cw.item, left:cw.name }}></td>
            {visibleMonths.map(ym => <td key={ym} style={{ minWidth:cw.num, borderLeft:`1px solid ${C.borderLight}`, borderBottom:`1px solid ${C.borderLight}20` }}></td>)}
          </tr>
        </>
      )}
    </>
  );
}

// ══════════════════════════════════════════════
// 期別サマリータブ
// ══════════════════════════════════════════════
function SummaryTab({ appData, periods, periodStarts, priceChanges, showItems, setShowItems }) {
  const { stores, months } = appData;

  // 選択中の項目だけで合計（showItems変化で再計算）
  const sumSelected = useCallback((tot) => {
    if (!tot) return 0;
    return ITEMS.reduce((s, it) => showItems[it.id] ? s + (tot[it.id]||0) : s, 0);
  }, [showItems]);

  // 前年比ラベル
  const yoyLabel = (curr, prev) => {
    if (!prev || prev === 0) return null;
    const pct = ((curr - prev) / prev * 100);
    const sign = pct >= 0 ? '+' : '';
    const color = pct >= 0 ? C.green : C.red;
    return { text: `${sign}${pct.toFixed(1)}%`, color };
  };

  const periodData = useMemo(() => periods.map(pn => {
    const pMonths = months.filter(ym => getPeriodNum(ym) === pn);
    const firstMonths  = pMonths.filter(ym => getHalf(ym) === 'first');
    const secondMonths = pMonths.filter(ym => getHalf(ym) === 'second');
    const calcTotal = (targetMonths) => {
      const t = {};
      ITEMS.forEach(it => { t[it.id] = 0; });
      t._total = 0;
      stores.forEach(store => {
        targetMonths.forEach(ym => {
          if (store.retireYM && ym >= store.retireYM) return;
          const baseRow = store.monthly[ym];
          if (!baseRow) return;
          const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
          ITEMS.forEach(it => { t[it.id] += row[it.id]||0; });
          t._total += sumItems(row);
        });
      });
      return t;
    };
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
    const firstTot  = calcTotal(firstMonths);
    const secondTot = calcTotal(secondMonths);
    return { pn, pMonths, storeData, ptot, retiredCount, activeCount, firstTot, secondTot };
  }), [stores, months, periods, priceChanges]);

  const periodColors = [C.red, C.blue, C.green, C.purple, C.orange, C.teal];
  const maxPeriodTotal = Math.max(...periodData.map(p => sumSelected(p.ptot)), 1);

  return (
    <div>
      {/* 項目フィルター */}
      <div style={{ display:"flex", gap:6, marginBottom:12, flexWrap:"wrap", alignItems:"center" }}>
        <span style={{ fontSize:12, color:C.textMuted, fontWeight:600 }}>表示項目：</span>
        <button onClick={() => setShowItems(Object.fromEntries(ITEMS.map(it=>[it.id,true])))}
          style={{ ...btnSt(C.textSub), background:"transparent", border:`1px solid ${C.border}`, color:C.textSub, fontSize:11 }}>全選択</button>
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

      {/* KPIカード */}
      <div style={{ display:"grid", gridTemplateColumns:`repeat(${Math.min(periodData.length,4)},1fr)`, gap:12, marginBottom:16 }}>
        {periodData.map(({ pn, ptot, retiredCount, activeCount, pMonths, firstTot, secondTot }, pi) => {
          const prevPtot = pi > 0 ? periodData[pi-1].ptot : null;
          const currTotal = sumSelected(ptot);
          const prevTotal = prevPtot ? sumSelected(prevPtot) : null;
          const yoy = yoyLabel(currTotal, prevTotal);
          return (
            <div key={pn} style={{
              background:C.surface, border:`1px solid ${C.border}`,
              borderTop:`3px solid ${periodColors[pi%6]}`,
              borderRadius:10, padding:16,
              boxShadow:"0 1px 4px rgba(0,0,0,0.04)",
            }}>
              <div style={{ fontSize:11, color:C.textMuted, marginBottom:4 }}>
                {pn}年度（{pMonths[0]}〜{pMonths[pMonths.length-1]}）
              </div>
              <div style={{ display:"flex", alignItems:"flex-end", gap:8, marginBottom:6 }}>
                <div style={{ fontSize:24, fontWeight:800, color:C.text, letterSpacing:"-0.5px" }}>{fmtM(currTotal)}</div>
                {yoy && <div style={{ fontSize:13, fontWeight:700, color:yoy.color, marginBottom:3 }}>{yoy.text}</div>}
              </div>
              <div style={{ display:"flex", gap:10, marginBottom:10, fontSize:11 }}>
                <span style={{ color:C.green, fontWeight:500 }}>稼働 {activeCount}店</span>
                {retiredCount > 0 && <span style={{ color:C.red, fontWeight:500 }}>撤退 {retiredCount}店</span>}
              </div>
              {/* 上期・下期 */}
              <div style={{ display:"flex", gap:6, marginBottom:10 }}>
                {[
                  { label:"▲ 上期（10〜3月）", tot:firstTot, prevTot:pi>0?periodData[pi-1].firstTot:null },
                  { label:"▽ 下期（4〜9月）",  tot:secondTot, prevTot:pi>0?periodData[pi-1].secondTot:null },
                ].map(({ label, tot, prevTot }) => {
                  const v = sumSelected(tot);
                  const hy = prevTot ? yoyLabel(v, sumSelected(prevTot)) : null;
                  return (
                    <div key={label} style={{ flex:1, background:C.surfaceHigh, borderRadius:6, padding:"6px 8px" }}>
                      <div style={{ fontSize:10, color:C.textMuted, marginBottom:2 }}>{label}</div>
                      <div style={{ fontSize:13, fontWeight:700, color:C.text }}>{fmtM(v)}</div>
                      {hy && <div style={{ fontSize:10, fontWeight:600, color:hy.color }}>{hy.text}</div>}
                    </div>
                  );
                })}
              </div>
              {/* 項目別 */}
              {ITEMS.filter(it => showItems[it.id] && !it.adOnly).map(it => {
                const v = ptot[it.id] || 0;
                const iy = prevPtot ? yoyLabel(v, prevPtot[it.id]) : null;
                return (
                  <div key={it.id} style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"3px 0", borderBottom:`1px solid ${C.borderLight}` }}>
                    <span style={{ fontSize:11, color:C.textSub, display:"flex", alignItems:"center", gap:5 }}>
                      <span style={{ display:"inline-block", width:6, height:6, borderRadius:3, background:it.color }} />{it.label}
                    </span>
                    <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                      {iy && <span style={{ fontSize:10, color:iy.color, fontWeight:600 }}>{iy.text}</span>}
                      <span style={{ fontSize:11, fontWeight:600, color:C.text }}>{fmtM(v)}</span>
                    </div>
                  </div>
                );
              })}
              {showItems['ad'] && (
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"3px 0", fontSize:11, color:C.textMuted }}>
                  <span>広告費</span>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                    {prevPtot && yoyLabel(ptot.ad, prevPtot.ad) && (
                      <span style={{ fontSize:10, color:yoyLabel(ptot.ad, prevPtot.ad).color, fontWeight:600 }}>
                        {yoyLabel(ptot.ad, prevPtot.ad).text}
                      </span>
                    )}
                    <span>{fmtM(ptot.ad)}</span>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* 棒グラフ */}
      <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, padding:16, marginBottom:16, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
        <div style={{ fontSize:13, fontWeight:700, marginBottom:16, color:C.text }}>年度別 収入推移</div>
        <div style={{ display:"flex", gap:10, alignItems:"flex-end", height:120 }}>
          {periodData.map(({ pn, ptot }, pi) => {
            const v = sumSelected(ptot);
            const pv = pi > 0 ? sumSelected(periodData[pi-1].ptot) : null;
            const yoy = pv ? yoyLabel(v, pv) : null;
            const h = Math.round((v / maxPeriodTotal) * 80);
            return (
              <div key={pn} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", gap:4 }}>
                <div style={{ fontSize:10, color:C.textSub, fontWeight:500 }}>{fmtM(v)}</div>
                {yoy && <div style={{ fontSize:10, fontWeight:700, color:yoy.color }}>{yoy.text}</div>}
                <div style={{ width:"100%", height:h, background:periodColors[pi%6], borderRadius:"4px 4px 0 0", opacity:0.85 }} />
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
              {periodData.map(({pn}, pi) => (
                <th key={pn} colSpan={pi>0?2:1} style={{ padding:"8px 14px", textAlign:"right", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}`, minWidth:pi>0?160:90, borderLeft:`1px solid ${C.borderLight}` }}>{pn}年度</th>
              ))}
              <th style={{ padding:"8px 14px", textAlign:"right", fontSize:11, color:C.textMuted, fontWeight:500, borderBottom:`1px solid ${C.border}` }}>累計</th>
            </tr>
            <tr style={{ background:C.surfaceHigh }}>
              <th style={{ padding:"3px 14px", borderBottom:`1px solid ${C.border}`, position:"sticky", left:0, background:C.surfaceHigh }}></th>
              <th style={{ padding:"3px 14px", borderBottom:`1px solid ${C.border}` }}></th>
              {periodData.map(({pn}, pi) => (
                pi === 0
                  ? <th key={pn} style={{ padding:"3px 10px", textAlign:"right", fontSize:10, color:C.textMuted, fontWeight:400, borderBottom:`1px solid ${C.border}`, borderLeft:`1px solid ${C.borderLight}` }}>金額</th>
                  : <>
                      <th key={`${pn}-v`} style={{ padding:"3px 10px", textAlign:"right", fontSize:10, color:C.textMuted, fontWeight:400, borderBottom:`1px solid ${C.border}`, borderLeft:`1px solid ${C.borderLight}` }}>金額</th>
                      <th key={`${pn}-y`} style={{ padding:"3px 10px", textAlign:"right", fontSize:10, color:C.textMuted, fontWeight:400, borderBottom:`1px solid ${C.border}` }}>前年比</th>
                    </>
              ))}
              <th style={{ padding:"3px 14px", borderBottom:`1px solid ${C.border}` }}></th>
            </tr>
          </thead>
          <tbody>
            {stores.map((store, i) => {
              const pTotals = periodData.map(p => sumSelected(p.storeData.find(d => d.store.id===store.id)?.tot||{}));
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
                  {pTotals.map((t, pi) => {
                    const prev = pi > 0 ? pTotals[pi-1] : null;
                    const yoy = prev !== null ? yoyLabel(t, prev) : null;
                    return pi === 0 ? (
                      <td key={pi} style={{ padding:"5px 10px", textAlign:"right", color:t>0?C.text:C.textMuted, borderLeft:`1px solid ${C.borderLight}` }}>{t>0?fmtM(t):"—"}</td>
                    ) : (
                      <>
                        <td key={`${pi}-v`} style={{ padding:"5px 10px", textAlign:"right", color:t>0?C.text:C.textMuted, borderLeft:`1px solid ${C.borderLight}` }}>{t>0?fmtM(t):"—"}</td>
                        <td key={`${pi}-y`} style={{ padding:"5px 10px", textAlign:"right", fontSize:11 }}>
                          {yoy ? <span style={{ color:yoy.color, fontWeight:600 }}>{yoy.text}</span> : ""}
                        </td>
                      </>
                    );
                  })}
                  <td style={{ padding:"5px 12px", textAlign:"right", fontWeight:800, color:C.red }}>{fmtM(grand)}</td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr style={{ background:C.redDim, borderTop:`2px solid ${C.redBorder}` }}>
              <td style={{ padding:"9px 14px", fontWeight:700, color:C.red, position:"sticky", left:0, background:C.redDim }}>合計</td>
              <td></td>
              {periodData.map(({ pn, ptot }, pi) => {
                const v = sumSelected(ptot);
                const prev = pi > 0 ? sumSelected(periodData[pi-1].ptot) : null;
                const yoy = prev !== null ? yoyLabel(v, prev) : null;
                return pi === 0 ? (
                  <td key={pn} style={{ padding:"9px 10px", textAlign:"right", fontWeight:700, color:C.red, borderLeft:`1px solid ${C.redBorder}` }}>{fmtM(v)}</td>
                ) : (
                  <>
                    <td key={`${pn}-v`} style={{ padding:"9px 10px", textAlign:"right", fontWeight:700, color:C.red, borderLeft:`1px solid ${C.redBorder}` }}>{fmtM(v)}</td>
                    <td key={`${pn}-y`} style={{ padding:"9px 10px", textAlign:"right" }}>
                      {yoy ? <span style={{ fontSize:12, color:yoy.color, fontWeight:700 }}>{yoy.text}</span> : ""}
                    </td>
                  </>
                );
              })}
              <td style={{ padding:"9px 14px", textAlign:"right", fontWeight:800, color:C.red }}>
                {fmtM(periodData.reduce((s,p)=>s+sumSelected(p.ptot),0))}
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

function MasterTab({ appData, priceChanges, setPriceChanges, onStoreDelete }) {
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
  const [deleteModal,  setDeleteModal]  = useState(null); // { id, name }

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

  // 削除確認
  const confirmDelete = () => {
    if (!deleteModal) return;
    setLocalStores(localStores.filter(s => s.id !== deleteModal.id));
    setPriceChanges((priceChanges||[]).filter(c => c.storeId !== deleteModal.id));
    if (onStoreDelete) onStoreDelete(deleteModal.id);
    setDeleteModal(null);
    flash("✓ 店舗を削除しました");
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
  const composingRef = useRef(false);
  const EF = ({k,placeholder,type="text",w=120}) => (
    <input key={`ef_${editId}_${k}`} type={type} defaultValue={editForm[k]||""} placeholder={placeholder||k}
      onCompositionStart={() => { composingRef.current = true; }}
      onCompositionEnd={e => { composingRef.current = false; setEditForm(f=>({...f,[k]:e.target.value})); }}
      onChange={e => { if (!composingRef.current) setEditForm(f=>({...f,[k]:e.target.value})); }}
      style={{...inputSt,width:w,fontSize:10}} />
  );
  const AF = ({k,placeholder,type="text",w=120}) => (
    <input key={`af_${k}`} type={type} defaultValue={addForm[k]||""} placeholder={placeholder||k}
      onCompositionStart={() => { composingRef.current = true; }}
      onCompositionEnd={e => { composingRef.current = false; setAddForm(f=>({...f,[k]:e.target.value})); }}
      onChange={e => { if (!composingRef.current) setAddForm(f=>({...f,[k]:e.target.value})); }}
      style={{...inputSt,width:w,fontSize:10}} />
  );

  return (
    <div>
      {/* 削除確認モーダル */}
      {deleteModal && (
        <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.6)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{background:C.surface,border:`1px solid ${C.border}`,borderRadius:10,padding:28,width:360,boxShadow:"0 8px 32px rgba(0,0,0,0.2)"}}>
            <div style={{fontSize:15,fontWeight:700,color:C.text,marginBottom:8}}>店舗を削除しますか？</div>
            <div style={{fontSize:13,color:C.textSub,marginBottom:6}}>
              <span style={{fontWeight:600,color:C.red}}>{deleteModal.name}</span>
            </div>
            <div style={{fontSize:12,color:C.textMuted,marginBottom:20,padding:"10px 12px",background:C.redDim,borderRadius:6,border:`1px solid ${C.redBorder}`}}>
              ⚠️ この操作は取り消せません。店舗の金額変更履歴もすべて削除されます。
            </div>
            <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
              <button onClick={()=>setDeleteModal(null)} style={{...btnSt(C.textSub),background:"transparent",border:`1px solid ${C.border}`,color:C.textSub,fontSize:12}}>キャンセル</button>
              <button onClick={confirmDelete} style={{...btnSt(C.red),fontSize:12}}>削除する</button>
            </div>
          </div>
        </div>
      )}
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
                {/* 現在の月次金額確認・修正パネル */}
                <CurrentAmountPanel key={`ca_${s.id}`} store={s} months={appData.months} onUpdate={(updatedStore) => {
                  setLocalStores(localStores.map(st => st.id===s.id ? updatedStore : st));
                  flash("✓ 金額を更新しました");
                }} />
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
                      <button onClick={()=>setDeleteModal({id:s.id, name:s.name})}
                        style={{padding:"2px 8px",borderRadius:4,border:`1px solid ${C.border}`,background:"transparent",color:C.textMuted,fontSize:9,cursor:"pointer"}}>
                        削除
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

// ══════════════════════════════════════════════
// 現在の月次金額確認・修正パネル
// ══════════════════════════════════════════════
function CurrentAmountPanel({ store, months, onUpdate }) {
  // データがある月だけ絞り込み
  const dataMonths = useMemo(() =>
    months.filter(ym => store.monthly[ym] && Object.values(store.monthly[ym]).some(v => v > 0)),
    [store, months]
  );

  const latestYm = dataMonths.length > 0 ? dataMonths[dataMonths.length - 1] : months[months.length - 1];
  const [selectedYm, setSelectedYm] = useState(latestYm || '');
  const [editAmounts, setEditAmounts] = useState({});
  const [saved, setSaved] = useState(false);

  const currentRow = store.monthly[selectedYm] || {};

  const handleSave = () => {
    const updatedMonthly = { ...store.monthly };
    const newRow = { ...currentRow };
    ITEMS.forEach(it => {
      if (editAmounts[it.id] !== undefined && editAmounts[it.id] !== '') {
        newRow[it.id] = Number(editAmounts[it.id]) || 0;
      }
    });
    updatedMonthly[selectedYm] = newRow;
    onUpdate({ ...store, monthly: updatedMonthly });
    setEditAmounts({});
    setSaved(true);
    setTimeout(() => setSaved(false), 2500);
  };

  const hasEdits = Object.keys(editAmounts).length > 0;

  return (
    <tr style={{ background:'#FFFDE7' }}>
      <td colSpan={11} style={{ padding:'14px 18px', borderBottom:`2px solid ${C.amber}30` }}>
        {/* ヘッダー */}
        <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:12 }}>
          <div style={{ display:'flex', alignItems:'center', gap:10 }}>
            <span style={{ fontSize:12, fontWeight:700, color:C.amber }}>📋 現在の月次金額 — {store.name}</span>
            {saved && <span style={{ fontSize:11, color:C.green, fontWeight:600, background:'#D1FAE5', padding:'2px 8px', borderRadius:4 }}>✓ 保存しました</span>}
          </div>
          {/* 年月セレクター */}
          <div style={{ display:'flex', alignItems:'center', gap:6 }}>
            <span style={{ fontSize:11, color:C.textMuted }}>確認月：</span>
            <select value={selectedYm} onChange={e => { setSelectedYm(e.target.value); setEditAmounts({}); }}
              style={{ ...selSt, fontSize:11, padding:'3px 8px' }}>
              {[...months].reverse().map(ym => (
                <option key={ym} value={ym}>{ym}{store.monthly[ym] ? '' : ' （データなし）'}</option>
              ))}
            </select>
          </div>
        </div>

        {/* 金額グリッド */}
        <div style={{ display:'grid', gridTemplateColumns:'repeat(4, 1fr)', gap:8, marginBottom:12 }}>
          {ITEMS.map(it => {
            const current = currentRow[it.id] || 0;
            const isEditing = editAmounts[it.id] !== undefined;
            const inputVal = isEditing ? editAmounts[it.id] : '';
            return (
              <div key={it.id} style={{
                padding:'10px 12px', borderRadius:8,
                background: isEditing ? `${it.color}10` : C.surface,
                border: `1.5px solid ${isEditing ? it.color : C.border}`,
                transition:'all .15s',
              }}>
                <div style={{ display:'flex', alignItems:'center', gap:5, marginBottom:6 }}>
                  <span style={{ display:'inline-block', width:7, height:7, borderRadius:3.5, background:it.color }} />
                  <span style={{ fontSize:11, color:it.color, fontWeight:600 }}>{it.label}</span>
                </div>
                {/* 現在金額 */}
                <div style={{ fontSize:14, fontWeight:700, color:current > 0 ? C.text : C.textMuted, marginBottom:6 }}>
                  {current > 0 ? `¥${current.toLocaleString()}` : '—'}
                </div>
                {/* 修正入力 */}
                <div>
                  <div style={{ fontSize:9, color:C.textMuted, marginBottom:2 }}>修正後の金額（円/月）</div>
                  <input
                    type="number"
                    value={inputVal}
                    placeholder={current > 0 ? String(current) : '例: 120000'}
                    onChange={e => setEditAmounts({ ...editAmounts, [it.id]: e.target.value })}
                    style={{ ...inputSt, width:'100%', fontSize:11, textAlign:'right', padding:'4px 6px' }}
                  />
                </div>
              </div>
            );
          })}
        </div>

        {/* 保存ボタン */}
        <div style={{ display:'flex', gap:8, alignItems:'center' }}>
          <button onClick={handleSave} disabled={!hasEdits} style={{
            ...btnSt(C.amber), fontSize:11,
            opacity: hasEdits ? 1 : 0.4,
            cursor: hasEdits ? 'pointer' : 'default',
          }}>
            {selectedYm} の金額を保存
          </button>
          {hasEdits && (
            <button onClick={() => setEditAmounts({})}
              style={{ ...btnSt(C.textMuted), background:'transparent', border:`1px solid ${C.border}`, color:C.textSub, fontSize:11 }}>
              リセット
            </button>
          )}
          <span style={{ fontSize:10, color:C.textMuted }}>
            修正した金額は売上明細・期別サマリーに即時反映されます
          </span>
        </div>
      </td>
    </tr>
  );
}

// ══════════════════════════════════════════════
// 予実管理タブ
// ══════════════════════════════════════════════
function BudgetTab({ appData, periods, periodStarts, priceChanges, showItems, setShowItems, actuals, setActuals, onActualsSave }) {
  const { stores, months } = appData;
  const fileRef = useRef(null);
  const [loading, setLoading] = useState(false);
  const [filterPeriod, setFilterPeriod] = useState(null);
  const [filterHalf, setFilterHalf] = useState(null);
  const [viewMode, setViewMode] = useState('summary'); // summary | detail

  // 表示月
  const visibleMonths = useMemo(() => {
    let ms = months;
    if (filterPeriod) ms = ms.filter(ym => getPeriodNum(ym) === filterPeriod);
    if (filterHalf)   ms = ms.filter(ym => getHalf(ym) === filterHalf);
    return ms;
  }, [months, filterPeriod, filterHalf]);

  // 実績CSVのアップロード処理
  const handleActualFile = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);
    try {
      const buf = await file.arrayBuffer();

      // 文字コード検出（Shift-JIS対応）
      const decoder = new TextDecoder('shift-jis');
      const text = decoder.decode(buf);
      const lines = text.split('\n').filter(l => l.trim());

      // ヘッダー行をスキップ
      const dataLines = lines.slice(1);

      // 店舗名→store IDのマップ作成
      const storeNameMap = {};
      const storeCodeMap = {};
      stores.forEach(s => {
        // 店舗名の正規化（スペース・記号を除去して照合）
        const norm = s.name.replace(/\s/g,'').replace(/[（(）)　]/g,'');
        storeNameMap[norm] = s.id;
        if (s.storeId) storeCodeMap[s.storeId] = s.id;
      });

      const newActuals = { ...actuals };
      let matched = 0, unmatched = 0;
      const unmatchedNames = [];

      dataLines.forEach(line => {
        // CSVパース（カンマ区切り・クォート対応）
        const cols = line.match(/(".*?"|[^,]+)(?=,|$)/g)
          ?.map(c => c.replace(/^"|"$/g,'').replace(/,/g,'').trim()) || [];
        if (cols.length < 9) return;

        const ymRaw   = cols[0]?.trim(); // 202606
        const storeRaw = cols[2]?.trim(); // T1000217　カーセブン川崎宮崎台店
        const royalty  = parseInt(cols[3]?.replace(/,/g,'') || '0') || 0;
        const sv       = parseInt(cols[4]?.replace(/,/g,'') || '0') || 0;
        const ad       = parseInt(cols[5]?.replace(/,/g,'') || '0') || 0;
        const system   = parseInt(cols[6]?.replace(/,/g,'') || '0') || 0;
        const cs       = parseInt(cols[7]?.replace(/,/g,'') || '0') || 0;
        const juryo    = parseInt(cols[8]?.replace(/,/g,'') || '0') || 0;

        if (!ymRaw || ymRaw.length < 6) return;

        // 年月フォーマット変換 202606 → 2026/06
        const ym = `${ymRaw.substring(0,4)}/${ymRaw.substring(4,6)}`;

        // 店舗コード抽出（先頭の英数字部分）
        const codeMatch = storeRaw?.match(/^([A-Z]\d+)/);
        const storeCode = codeMatch?.[1] || '';

        // 店舗名抽出（コードと全角スペースの後）
        const storeName = storeRaw?.replace(/^[A-Z]\d+[\s　]+/, '').trim() || '';
        const normName  = storeName.replace(/\s/g,'').replace(/[（(）)　]/g,'');

        // マッチング：コード優先→名前照合
        let storeId = storeCodeMap[storeCode] || storeNameMap[normName];

        // 部分一致フォールバック
        if (!storeId && normName) {
          const key = Object.keys(storeNameMap).find(k =>
            k.includes(normName.substring(0, 8)) || normName.includes(k.substring(0, 8))
          );
          if (key) storeId = storeNameMap[key];
        }

        if (!storeId) {
          unmatched++;
          unmatchedNames.push(storeName);
          return;
        }

        matched++;
        if (!newActuals[storeId]) newActuals[storeId] = {};

        // 従量額を分配（60% → 純粋ロイヤリティ、40% → 広告費用）
        const juryoRoyalty = Math.round(juryo * 0.6);
        const juryoAd      = juryo - juryoRoyalty;

        newActuals[storeId][ym] = {
          royalty:    royalty + juryoRoyalty,
          sv:         sv,
          renewal:    0,
          membership: 0,
          cs:         cs,
          system:     system,
          ad:         ad + juryoAd,
        };
      });

      setActuals(newActuals);
      if (onActualsSave) onActualsSave(newActuals);

      const msg = [
        `✅ 実績データを取り込みました`,
        `・マッチ: ${matched}店舗`,
        unmatched > 0 ? `・未マッチ: ${unmatched}店舗（${unmatchedNames.slice(0,3).join('、')}${unmatchedNames.length>3?'…':''}）` : '',
        `・従量額は純粋ロイヤリティ60%・広告費用40%で配分済み`,
      ].filter(Boolean).join('\n');
      alert(msg);

    } catch(err) {
      console.error(err);
      alert('読み込みに失敗しました: ' + err.message);
    }
    setLoading(false);
    if (fileRef.current) fileRef.current.value = '';
  };

  // showItemsで選択した項目だけ集計
  const sumFiltered = useCallback((row) => {
    if (!row) return 0;
    return ITEMS.reduce((s, it) => showItems[it.id] ? s + (row[it.id]||0) : s, 0);
  }, [showItems]);

  // 予算集計（priceChanges適用済み・showItemsフィルター）
  const budgetByYm = useMemo(() => {
    const result = {};
    visibleMonths.forEach(ym => {
      result[ym] = { _total:0 };
      ITEMS.forEach(it => { result[ym][it.id] = 0; });
      stores.forEach(store => {
        if (store.retireYM && ym >= store.retireYM) return;
        const baseRow = store.monthly[ym];
        if (!baseRow) return;
        const row = applyPriceChanges(baseRow, store.id, ym, priceChanges);
        ITEMS.forEach(it => { result[ym][it.id] += row[it.id]||0; });
        result[ym]._total += sumFiltered(row);
      });
    });
    return result;
  }, [stores, visibleMonths, priceChanges, sumFiltered]);

  // 実績集計（showItemsフィルター）
  const actualByYm = useMemo(() => {
    const result = {};
    visibleMonths.forEach(ym => {
      result[ym] = { _total:0 };
      ITEMS.forEach(it => { result[ym][it.id] = 0; });
      Object.values(actuals).forEach(storeActuals => {
        const row = storeActuals[ym];
        if (!row) return;
        ITEMS.forEach(it => { result[ym][it.id] += row[it.id]||0; });
        result[ym]._total += sumFiltered(row);
      });
    });
    return result;
  }, [actuals, visibleMonths, sumFiltered]);

  // 期別集計
  const periodSummary = useMemo(() => periods.map(pn => {
    const pMonths = months.filter(ym => getPeriodNum(ym) === pn);
    const budget = { _total:0 }; ITEMS.forEach(it => { budget[it.id] = 0; });
    const actual = { _total:0 }; ITEMS.forEach(it => { actual[it.id] = 0; });
    pMonths.forEach(ym => {
      const b = budgetByYm[ym] || {};
      const a = actualByYm[ym] || {};
      ITEMS.forEach(it => { budget[it.id] += b[it.id]||0; actual[it.id] += a[it.id]||0; });
      budget._total += b._total||0;
      actual._total += a._total||0;
    });
    const hasActual = actual._total > 0;
    const rate = budget._total > 0 ? (actual._total / budget._total * 100) : null;
    const diff = actual._total - budget._total;
    return { pn, pMonths, budget, actual, rate, diff, hasActual };
  }), [periods, months, budgetByYm, actualByYm]);

  const hasAnyActual = Object.keys(actuals).length > 0;
  const periodColors = [C.red, C.blue, C.green, C.purple, C.orange, C.teal];

  return (
    <div>
      {/* コントロールバー */}
      <div style={{ display:"flex", gap:8, marginBottom:8, alignItems:"center", flexWrap:"wrap", padding:"10px 12px", background:C.surface, borderRadius:8, border:`1px solid ${C.border}`, boxShadow:"0 1px 3px rgba(0,0,0,0.04)" }}>
        <div style={{ fontSize:13, fontWeight:700, color:C.text }}>予実管理</div>
        <div style={{ display:"flex", gap:4, alignItems:"center" }}>
          <span style={{ fontSize:11, color:C.textMuted }}>期：</span>
          <button onClick={() => setFilterPeriod(null)} style={pillSt(!filterPeriod)}>全期</button>
          {periods.map(p => (
            <button key={p} onClick={() => setFilterPeriod(p===filterPeriod?null:p)} style={pillSt(filterPeriod===p)}>{p}年度</button>
          ))}
        </div>
        <div style={{ display:"flex", gap:4, alignItems:"center" }}>
          <button onClick={() => setFilterHalf(null)} style={pillSt(!filterHalf)}>全期</button>
          <button onClick={() => setFilterHalf(filterHalf==='first'?null:'first')} style={pillSt(filterHalf==='first')}>上期</button>
          <button onClick={() => setFilterHalf(filterHalf==='second'?null:'second')} style={pillSt(filterHalf==='second')}>下期</button>
        </div>
        <div style={{ marginLeft:"auto", display:"flex", gap:8, alignItems:"center" }}>
          <button onClick={() => setViewMode(viewMode==='summary'?'detail':'summary')}
            style={{ ...btnSt(C.blue), background:"transparent", border:`1px solid ${C.border}`, color:C.textSub, fontSize:11 }}>
            {viewMode==='summary' ? '月次詳細' : 'サマリー'}
          </button>
          <input ref={fileRef} type="file" accept=".csv" style={{ display:"none" }} onChange={handleActualFile} />
          <button onClick={() => fileRef.current?.click()} style={{ ...btnSt(C.green), fontSize:11 }}>
            {loading ? '読込中...' : '実績CSVをアップロード'}
          </button>
        </div>
      </div>

      {/* 項目フィルター */}
      <div style={{ display:"flex", gap:6, marginBottom:12, flexWrap:"wrap", alignItems:"center", padding:"8px 12px", background:C.surface, borderRadius:8, border:`1px solid ${C.border}` }}>
        <span style={{ fontSize:11, color:C.textMuted, fontWeight:600 }}>表示項目：</span>
        <button onClick={() => setShowItems(Object.fromEntries(ITEMS.map(it=>[it.id,true])))}
          style={{ ...btnSt(C.textSub), background:"transparent", border:`1px solid ${C.border}`, color:C.textSub, fontSize:11 }}>全選択</button>
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

      {!hasAnyActual && (
        <div style={{ padding:"24px", background:C.surface, border:`1px dashed ${C.border}`, borderRadius:10, textAlign:"center", marginBottom:16 }}>
          <div style={{ fontSize:32, marginBottom:12 }}>📊</div>
          <div style={{ fontSize:14, fontWeight:700, color:C.text, marginBottom:8 }}>実績データがまだありません</div>
          <div style={{ fontSize:12, color:C.textMuted, marginBottom:16 }}>
            毎月締め後に実績XLSXをアップロードすることで予実比較ができます
          </div>
          <button onClick={() => fileRef.current?.click()} style={{ ...btnSt(C.green), fontSize:12 }}>
            実績XLSXをアップロード
          </button>
        </div>
      )}

      {/* 期別KPIカード */}
      <div style={{ display:"grid", gridTemplateColumns:`repeat(${Math.min(periodSummary.length,4)},1fr)`, gap:12, marginBottom:16 }}>
        {periodSummary.filter(p => !filterPeriod || p.pn === filterPeriod).map(({ pn, budget, actual, rate, diff, hasActual }, pi) => (
          <div key={pn} style={{ background:C.surface, border:`1px solid ${C.border}`, borderTop:`3px solid ${periodColors[pi%6]}`, borderRadius:10, padding:16, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
            <div style={{ fontSize:11, color:C.textMuted, marginBottom:4 }}>{pn}年度</div>

            {/* 予算 */}
            <div style={{ marginBottom:8 }}>
              <div style={{ fontSize:10, color:C.textMuted }}>予算</div>
              <div style={{ fontSize:20, fontWeight:800, color:C.text }}>{fmtM(budget._total)}</div>
            </div>

            {/* 実績 */}
            <div style={{ marginBottom:10 }}>
              <div style={{ fontSize:10, color:C.textMuted }}>実績</div>
              <div style={{ fontSize:20, fontWeight:800, color: hasActual ? (diff >= 0 ? C.green : C.red) : C.textMuted }}>
                {hasActual ? fmtM(actual._total) : '—'}
              </div>
            </div>

            {/* 達成率 */}
            {hasActual && rate !== null && (
              <>
                <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4, fontSize:11 }}>
                  <span style={{ color:C.textMuted }}>達成率</span>
                  <span style={{ fontWeight:700, color: rate >= 100 ? C.green : rate >= 90 ? C.amber : C.red }}>
                    {rate.toFixed(1)}%
                  </span>
                </div>
                {/* 達成率バー */}
                <div style={{ height:8, background:C.surfaceHigh, borderRadius:4, overflow:"hidden", marginBottom:8 }}>
                  <div style={{
                    height:"100%", borderRadius:4,
                    width:`${Math.min(rate,100)}%`,
                    background: rate >= 100 ? C.green : rate >= 90 ? C.amber : C.red,
                    transition:"width .5s",
                  }} />
                </div>
                {/* 差異 */}
                <div style={{ display:"flex", justifyContent:"space-between", fontSize:11 }}>
                  <span style={{ color:C.textMuted }}>差異</span>
                  <span style={{ fontWeight:600, color: diff >= 0 ? C.green : C.red }}>
                    {diff >= 0 ? '+' : ''}{fmtM(diff)}
                  </span>
                </div>
              </>
            )}

            {/* 項目別内訳 */}
            <div style={{ marginTop:10, paddingTop:10, borderTop:`1px solid ${C.borderLight}` }}>
              {ITEMS.filter(it => !it.adOnly).map(it => {
                const b = budget[it.id] || 0;
                const a = actual[it.id] || 0;
                const r = b > 0 ? (a / b * 100) : null;
                return (
                  <div key={it.id} style={{ display:"flex", justifyContent:"space-between", padding:"2px 0", fontSize:11 }}>
                    <span style={{ color:it.color, display:"flex", alignItems:"center", gap:4 }}>
                      <span style={{ display:"inline-block", width:5, height:5, borderRadius:2.5, background:it.color }} />
                      {it.label}
                    </span>
                    <span style={{ display:"flex", gap:8, alignItems:"center" }}>
                      {hasActual && r !== null && (
                        <span style={{ fontSize:10, color: r >= 100 ? C.green : r >= 90 ? C.amber : C.red, fontWeight:600 }}>
                          {r.toFixed(0)}%
                        </span>
                      )}
                      <span style={{ color:C.textSub }}>{fmtM(b)}</span>
                    </span>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>

      {/* 月次推移テーブル */}
      {viewMode === 'detail' && (
        <div style={{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:10, overflowX:"auto", boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
          <div style={{ padding:"12px 16px", borderBottom:`1px solid ${C.border}`, fontSize:13, fontWeight:700, color:C.text }}>月次 予実比較</div>
          <table style={{ borderCollapse:"collapse", fontSize:12, whiteSpace:"nowrap" }}>
            <thead>
              <tr style={{ background:C.surfaceHigh }}>
                <th style={{ padding:"8px 14px", textAlign:"left", fontSize:11, color:C.textMuted, borderBottom:`1px solid ${C.border}`, position:"sticky", left:0, background:C.surfaceHigh, minWidth:80 }}>項目</th>
                {visibleMonths.map(ym => {
                  const isFirst = ym === periodStarts[getPeriodNum(ym)];
                  return (
                    <th key={ym} colSpan={2} style={{
                      padding:"8px 10px", textAlign:"center", fontSize:11,
                      color: isFirst ? C.red : C.textMuted,
                      borderBottom:`1px solid ${C.border}`,
                      borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}`,
                      minWidth:140,
                    }}>{ym}</th>
                  );
                })}
              </tr>
              <tr style={{ background:C.surfaceHigh }}>
                <th style={{ padding:"4px 14px", borderBottom:`1px solid ${C.border}`, position:"sticky", left:0, background:C.surfaceHigh }}></th>
                {visibleMonths.map(ym => (
                  <>
                    <th key={`${ym}-b`} style={{ padding:"4px 8px", textAlign:"right", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}`, borderLeft:`1px solid ${C.borderLight}` }}>予算</th>
                    <th key={`${ym}-a`} style={{ padding:"4px 8px", textAlign:"right", fontSize:10, color:C.textMuted, borderBottom:`1px solid ${C.border}` }}>実績</th>
                  </>
                ))}
              </tr>
            </thead>
            <tbody>
              {/* 合計行 */}
              <tr style={{ background:C.redDim, fontWeight:700 }}>
                <td style={{ padding:"6px 14px", color:C.red, position:"sticky", left:0, background:C.redDim }}>合計</td>
                {visibleMonths.map(ym => {
                  const b = budgetByYm[ym]?._total || 0;
                  const a = actualByYm[ym]?._total || 0;
                  const isFirst = ym === periodStarts[getPeriodNum(ym)];
                  return (
                    <>
                      <td key={`${ym}-b`} style={{ padding:"6px 8px", textAlign:"right", color:C.red, borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}` }}>{fmtK(b)}</td>
                      <td key={`${ym}-a`} style={{ padding:"6px 8px", textAlign:"right", color: a > 0 ? (a >= b ? C.green : C.red) : C.textMuted }}>
                        {a > 0 ? fmtK(a) : '—'}
                      </td>
                    </>
                  );
                })}
              </tr>
              {/* 項目別行 */}
              {ITEMS.map((it, idx) => (
                <tr key={it.id} style={{ background:idx%2===0?"transparent":`${C.surfaceHigh}40`, borderBottom:`1px solid ${C.borderLight}` }}>
                  <td style={{ padding:"5px 14px", position:"sticky", left:0, background:idx%2===0?C.surface:`${C.surfaceHigh}80` }}>
                    <span style={{ display:"inline-flex", alignItems:"center", gap:5 }}>
                      <span style={{ display:"inline-block", width:6, height:6, borderRadius:3, background:it.color }} />
                      <span style={{ color:it.color, fontSize:11 }}>{it.label}</span>
                    </span>
                  </td>
                  {visibleMonths.map(ym => {
                    const b = budgetByYm[ym]?.[it.id] || 0;
                    const a = actualByYm[ym]?.[it.id] || 0;
                    const isFirst = ym === periodStarts[getPeriodNum(ym)];
                    return (
                      <>
                        <td key={`${ym}-b`} style={{ padding:"5px 8px", textAlign:"right", color:C.textSub, borderLeft: isFirst ? `2px solid ${C.periodLine}` : `1px solid ${C.borderLight}` }}>{b > 0 ? fmtK(b) : '—'}</td>
                        <td key={`${ym}-a`} style={{ padding:"5px 8px", textAlign:"right", color: a > 0 ? (a >= b ? C.green : C.red) : C.textMuted }}>
                          {a > 0 ? fmtK(a) : '—'}
                        </td>
                      </>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}
