import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)

// ── 認証 ─────────────────────────────────────────────────────
export const auth = {
  signInWithGoogle: () => supabase.auth.signInWithOAuth({
    provider: 'google',
    options: {
      redirectTo: window.location.origin,
      queryParams: { access_type: 'offline', prompt: 'consent' },
    },
  }),
  signOut:           () => supabase.auth.signOut(),
  getUser:           () => supabase.auth.getUser(),
  onAuthStateChange: (cb) => supabase.auth.onAuthStateChange(cb),
}

// ── 店舗マスタ ────────────────────────────────────────────────
export const storesDB = {
  getAll: async () => {
    const { data, error } = await supabase.from('stores').select('*').order('name')
    if (error) throw error
    return data
  },
  upsert: async (store) => {
    const { data, error } = await supabase.from('stores').upsert(storeToRow(store), { onConflict:'id' }).select().single()
    if (error) throw error
    return data
  },
  upsertMany: async (stores) => {
    const { error } = await supabase.from('stores').upsert(stores.map(storeToRow), { onConflict:'id' })
    if (error) throw error
  },
  delete: async (id) => {
    const { error } = await supabase.from('stores').delete().eq('id', id)
    if (error) throw error
  },
}

// ── 売上明細（ページネーションで全件取得）─────────────────────
export const monthlyDB = {
  getAll: async () => {
    let allData = []
    let from = 0
    const pageSize = 1000
    while (true) {
      const { data, error } = await supabase
        .from('store_monthly')
        .select('*')
        .order('ym')
        .range(from, from + pageSize - 1)
      if (error) throw error
      allData = allData.concat(data)
      if (data.length < pageSize) break
      from += pageSize
    }
    return allData
  },
  upsertMany: async (rows) => {
    for (let i = 0; i < rows.length; i += 1000) {
      const { error } = await supabase
        .from('store_monthly')
        .upsert(rows.slice(i, i + 1000), { onConflict:'store_id,ym' })
      if (error) throw error
    }
  },
}

// ── 金額変更履歴 ──────────────────────────────────────────────
export const priceChangesDB = {
  getAll: async () => {
    const { data, error } = await supabase.from('price_changes').select('*').order('from_ym')
    if (error) throw error
    return data
  },
  insert: async (change) => {
    const { data, error } = await supabase.from('price_changes').insert(priceChangeToRow(change)).select().single()
    if (error) throw error
    return data
  },
  delete: async (id) => {
    const { error } = await supabase.from('price_changes').delete().eq('id', id)
    if (error) throw error
  },
}

// ── 変換ヘルパー ──────────────────────────────────────────────
function storeToRow(s) {
  return {
    id: s.id, store_id: s.storeId||null, member_id: s.memberId||null,
    company: s.company||null, name: s.name, pref: s.pref||null,
    pkg_key: s.pkgKey||null, pkg_name: s.pkgName||null,
    contract_years: s.contractYears||null, renewal_period: s.renewalPeriod||null,
    open_date: s.openDate||null, last_renewal: s.lastRenewal||null,
    contract_end: s.contractEnd||null, retire_ym: s.retireYM||null, retire_raw: s.retireRaw||null,
  }
}
export function rowToStore(row) {
  return {
    id: row.id, storeId: row.store_id, memberId: row.member_id,
    company: row.company||'', name: row.name, pref: row.pref||'',
    pkgKey: row.pkg_key||'other', pkgName: row.pkg_name||'',
    contractYears: row.contract_years||'', renewalPeriod: row.renewal_period||'',
    openDate: row.open_date||'', lastRenewal: row.last_renewal||'',
    contractEnd: row.contract_end||'', retireYM: row.retire_ym||'', retireRaw: row.retire_raw||'',
    monthly: {},
  }
}
export function rowToMonthly(rows) {
  const map = {}
  for (const row of rows) {
    if (!map[row.store_id]) map[row.store_id] = {}
    map[row.store_id][row.ym] = {
      royalty: Number(row.royalty||0), sv: Number(row.sv||0),
      renewal: Number(row.renewal||0), membership: Number(row.membership||0),
      cs: Number(row.cs||0), system: Number(row.system||0), ad: Number(row.ad||0),
    }
  }
  return map
}
export function monthlyToRows(store) {
  return Object.entries(store.monthly||{}).flatMap(([ym,items]) => {
    if (!items||!Object.keys(items).length) return []
    return [{ store_id:store.id, ym, royalty:items.royalty||0, sv:items.sv||0,
      renewal:items.renewal||0, membership:items.membership||0,
      cs:items.cs||0, system:items.system||0, ad:items.ad||0 }]
  })
}
function priceChangeToRow(c) {
  return { store_id:c.storeId, from_ym:c.fromYM, item_id:c.itemId, new_value:c.newValue, note:c.note||null }
}
export function rowToPriceChange(row) {
  return { id:String(row.id), storeId:row.store_id, fromYM:row.from_ym,
    itemId:row.item_id, newValue:Number(row.new_value), note:row.note||'' }
}
