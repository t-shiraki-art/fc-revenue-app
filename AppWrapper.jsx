/**
 * AppWrapper.jsx
 * 認証管理 + Supabaseデータロード + App.jsxへのデータ注入
 */
import { useState, useEffect, useCallback } from 'react'
import App from './App.jsx'
import AuthScreen from './AuthScreen.jsx'
import { auth, storesDB, monthlyDB, priceChangesDB, rowToStore, rowToMonthly, rowToPriceChange, monthlyToRows } from './db.js'

const C = { bg:'#F5F6F8', red:'#E31F25', text:'#1A1D23', textMuted:'#9CA3AF', surface:'#FFFFFF', border:'#E2E6EC' }

export default function AppWrapper() {
  const [user,         setUser]         = useState(null)
  const [authChecked,  setAuthChecked]  = useState(false)
  const [appData,      setAppData]      = useState(null)   // { stores, months, sheetName }
  const [priceChanges, setPriceChanges] = useState([])
  const [loading,      setLoading]      = useState(false)
  const [loadMsg,      setLoadMsg]      = useState('')
  const [saveStatus,   setSaveStatus]   = useState('')     // '', 'saving', 'saved', 'error'

  // ── 認証状態の監視 ─────────────────────────────────────────
  useEffect(() => {
    auth.getUser().then(({ data }) => {
      setUser(data.user || null)
      setAuthChecked(true)
    })
    const { data: { subscription } } = auth.onAuthStateChange((event, session) => {
      setUser(session?.user || null)
      if (event === 'SIGNED_OUT') { setAppData(null); setPriceChanges([]) }
    })
    return () => subscription.unsubscribe()
  }, [])

  // ── ログイン後にDBからデータロード ────────────────────────
  useEffect(() => {
    if (user) loadFromDB()
  }, [user])

  const loadFromDB = async () => {
    setLoading(true)
    setLoadMsg('店舗マスタを読み込み中...')
    try {
      const [storeRows, monthlyRows, pcRows] = await Promise.all([
        storesDB.getAll(),
        monthlyDB.getAll(),
        priceChangesDB.getAll(),
      ])

      if (storeRows.length === 0) {
        // DBが空 → アップロード画面を表示
        setAppData(null)
        setLoading(false)
        return
      }

      setLoadMsg('売上明細を組み立て中...')
      const monthlyMap = rowToMonthly(monthlyRows)
      const stores = storeRows.map(row => {
        const s = rowToStore(row)
        s.monthly = monthlyMap[s.id] || {}
        return s
      })

      // 全月リストを算出
      const monthSet = new Set()
      monthlyRows.forEach(r => monthSet.add(r.ym))
      const months = [...monthSet].sort()

      setAppData({ stores, months, sheetName: 'Supabaseから読み込み済み' })
      setPriceChanges(pcRows.map(rowToPriceChange))
    } catch (e) {
      console.error('DB読み込みエラー:', e)
      setAppData(null)
    }
    setLoading(false)
    setLoadMsg('')
  }

  // ── XLSXアップロード後にDBへ保存 ──────────────────────────
  const handleXLSXLoaded = useCallback(async (data) => {
    setAppData(data)
    setSaveStatus('saving')
    try {
      // 店舗マスタを保存
      await storesDB.upsertMany(data.stores)

      // 売上明細を保存（全店舗分）
      const allMonthlyRows = data.stores.flatMap(s => monthlyToRows(s))
      if (allMonthlyRows.length > 0) {
        await monthlyDB.upsertMany(allMonthlyRows)
      }

      setSaveStatus('saved')
      setTimeout(() => setSaveStatus(''), 3000)
    } catch (e) {
      console.error('DB保存エラー:', e)
      setSaveStatus('error')
    }
  }, [])

  // ── 店舗マスタ更新をDBに反映 ──────────────────────────────
  const handleStoreUpdate = useCallback(async (updatedStore) => {
    try {
      await storesDB.upsert(updatedStore)
    } catch (e) {
      console.error('店舗更新エラー:', e)
    }
  }, [])

  // ── 金額変更履歴をDBに反映 ────────────────────────────────
  const handlePriceChangeAdd = useCallback(async (change) => {
    try {
      const row = await priceChangesDB.insert(change)
      const newChange = rowToPriceChange(row)
      setPriceChanges(prev => [...prev, newChange])
      return newChange
    } catch (e) {
      console.error('金額変更追加エラー:', e)
      throw e
    }
  }, [])

  const handlePriceChangeDelete = useCallback(async (id) => {
    try {
      await priceChangesDB.delete(id)
      setPriceChanges(prev => prev.filter(c => c.id !== id))
    } catch (e) {
      console.error('金額変更削除エラー:', e)
    }
  }, [])

  // ── 未認証チェック中 ──────────────────────────────────────
  if (!authChecked) {
    return (
      <div style={{ fontFamily:"'Inter','Noto Sans JP',sans-serif", background:C.bg, minHeight:'100vh', display:'flex', alignItems:'center', justifyContent:'center' }}>
        <div style={{ textAlign:'center', color:C.textMuted }}>
          <div style={{ width:40, height:40, borderRadius:8, background:C.red, display:'inline-flex', alignItems:'center', justifyContent:'center', fontWeight:900, fontSize:16, color:'#fff', marginBottom:12 }}>C7</div>
          <div style={{ fontSize:13 }}>読み込み中...</div>
        </div>
      </div>
    )
  }

  // ── 未ログイン ─────────────────────────────────────────────
  if (!user) return <AuthScreen onLogin={setUser} />

  // ── DBロード中 ─────────────────────────────────────────────
  if (loading) {
    return (
      <div style={{ fontFamily:"'Inter','Noto Sans JP',sans-serif", background:C.bg, minHeight:'100vh', display:'flex', alignItems:'center', justifyContent:'center' }}>
        <div style={{ textAlign:'center' }}>
          <div style={{ width:48, height:48, borderRadius:10, background:C.red, display:'inline-flex', alignItems:'center', justifyContent:'center', fontWeight:900, fontSize:18, color:'#fff', marginBottom:16, boxShadow:`0 4px 14px ${C.red}40` }}>C7</div>
          <div style={{ fontSize:14, color:C.text, fontWeight:600, marginBottom:8 }}>データを読み込み中...</div>
          <div style={{ fontSize:12, color:C.textMuted }}>{loadMsg}</div>
        </div>
      </div>
    )
  }

  // ── メインアプリ ───────────────────────────────────────────
  return (
    <div style={{ position:'relative' }}>
      {/* DB保存ステータスバナー */}
      {saveStatus && (
        <div style={{
          position:'fixed', bottom:20, right:20, zIndex:9999,
          padding:'10px 18px', borderRadius:8, fontSize:12, fontWeight:600,
          background: saveStatus==='saved'?'#D1FAE5':saveStatus==='error'?'#FDEAEA':'#DBEAFE',
          color: saveStatus==='saved'?'#059669':saveStatus==='error'?'#E31F25':'#2563EB',
          border: `1px solid ${saveStatus==='saved'?'#6EE7B7':saveStatus==='error'?'#FBBFBF':'#93C5FD'}`,
          boxShadow:'0 2px 8px rgba(0,0,0,0.1)',
        }}>
          {saveStatus==='saving' && '💾 Supabaseに保存中...'}
          {saveStatus==='saved'  && '✓ Supabaseに保存しました'}
          {saveStatus==='error'  && '⚠️ 保存に失敗しました'}
        </div>
      )}

      {/* ログアウトボタン */}
      <button onClick={() => auth.signOut()} style={{
        position:'fixed', top:12, right:12, zIndex:200,
        padding:'4px 12px', borderRadius:6, border:'1px solid #E2E6EC',
        background:'#fff', color:'#9CA3AF', fontSize:11, cursor:'pointer',
      }}>
        {user.email} ／ ログアウト
      </button>

      <App
        initialData={appData}
        initialPriceChanges={priceChanges}
        onXLSXLoaded={handleXLSXLoaded}
        onStoreUpdate={handleStoreUpdate}
        onPriceChangeAdd={handlePriceChangeAdd}
        onPriceChangeDelete={handlePriceChangeDelete}
      />
    </div>
  )
}
