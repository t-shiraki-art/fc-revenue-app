import { useState } from 'react'
import { auth } from './db'

const C = {
  bg: '#F5F6F8', surface: '#FFFFFF', border: '#E2E6EC',
  text: '#1A1D23', textMuted: '#9CA3AF', textSub: '#4B5563',
  red: '#E31F25', redDim: '#FDEAEA', redBorder: '#FBBFBF',
  green: '#059669',
}

export default function AuthScreen({ onLogin }) {
  const [loading, setLoading] = useState(false)
  const [error,   setError]   = useState(null)

  const handleGoogle = async () => {
    setLoading(true); setError(null)
    const { error } = await auth.signInWithGoogle()
    if (error) {
      setError('Googleログインに失敗しました: ' + error.message)
      setLoading(false)
    }
    // 成功時はSupabaseがリダイレクトするので何もしない
  }

  return (
    <div style={{
      fontFamily:"'Inter','Noto Sans JP',sans-serif",
      background:C.bg, minHeight:'100vh',
      display:'flex', alignItems:'center', justifyContent:'center',
    }}>
      <div style={{ width:380 }}>
        {/* ロゴ */}
        <div style={{ textAlign:'center', marginBottom:32 }}>
          <div style={{
            width:56, height:56, borderRadius:14,
            background:C.red, display:'inline-flex',
            alignItems:'center', justifyContent:'center',
            fontWeight:900, fontSize:22, color:'#fff',
            marginBottom:14,
            boxShadow:`0 4px 18px ${C.red}45`,
          }}>C7</div>
          <div style={{ fontSize:21, fontWeight:800, color:C.text, letterSpacing:'-0.3px' }}>
            カーセブン FC収入管理
          </div>
          <div style={{ fontSize:13, color:C.textMuted, marginTop:5 }}>
            FC本部 売上管理システム
          </div>
        </div>

        {/* カード */}
        <div style={{
          background:C.surface, borderRadius:14,
          border:`1px solid ${C.border}`,
          overflow:'hidden',
          boxShadow:'0 2px 20px rgba(0,0,0,0.08)',
        }}>
          <div style={{ height:4, background:C.red }} />
          <div style={{ padding:'32px 28px' }}>
            <div style={{ fontSize:15, fontWeight:700, color:C.text, marginBottom:6 }}>
              ログイン
            </div>
            <div style={{ fontSize:12, color:C.textMuted, marginBottom:24, lineHeight:1.6 }}>
              Googleアカウントでログインしてください。<br />
              FC事業部のGoogleアカウントを使用してください。
            </div>

            {error && (
              <div style={{
                marginBottom:16, padding:'10px 14px',
                background:C.redDim, border:`1px solid ${C.redBorder}`,
                borderRadius:8, fontSize:12, color:C.red,
              }}>{error}</div>
            )}

            {/* Googleログインボタン */}
            <button onClick={handleGoogle} disabled={loading} style={{
              width:'100%', padding:'12px 16px',
              borderRadius:9, border:`1px solid ${C.border}`,
              background: loading ? C.bg : '#fff',
              color:C.text, fontSize:14, fontWeight:600,
              cursor: loading ? 'default' : 'pointer',
              display:'flex', alignItems:'center', justifyContent:'center', gap:10,
              boxShadow:'0 1px 4px rgba(0,0,0,0.08)',
              transition:'box-shadow .15s',
            }}>
              {/* Google SVGアイコン */}
              {!loading && (
                <svg width="18" height="18" viewBox="0 0 48 48">
                  <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/>
                  <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/>
                  <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/>
                  <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/>
                </svg>
              )}
              {loading ? 'ログイン中...' : 'Googleでログイン'}
            </button>

            <div style={{ marginTop:20, textAlign:'center', fontSize:11, color:C.textMuted }}>
              ログインすることで社内利用規約に同意したものとみなします
            </div>
          </div>
        </div>

        <div style={{ textAlign:'center', marginTop:16, fontSize:11, color:C.textMuted }}>
          © カーセブン FC事業部
        </div>
      </div>
    </div>
  )
}
