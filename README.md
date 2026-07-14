# カーセブン FC収入管理システム

## 構成
- **フロントエンド**: React + Vite（GitHub Pages でホスティング）
- **データベース**: Supabase（PostgreSQL）
- **認証**: Supabase Auth（Googleログイン）

## GitHub Pages セットアップ

### 1. このリポジトリを GitHub に作成（Private推奨）

### 2. Supabase の環境変数を GitHub Secrets に登録
GitHub リポジトリ → Settings → Secrets and variables → Actions → New repository secret

| Secret名 | 値 |
|----------|---|
| `VITE_SUPABASE_URL` | `https://tyomtjykwsqyfgrdjkvm.supabase.co` |
| `VITE_SUPABASE_ANON_KEY` | （.env.exampleの値） |

### 3. GitHub Pages を有効化
リポジトリ → Settings → Pages → Source: **GitHub Actions**

### 4. main ブランチにpushすると自動デプロイ
公開URL: `https://<あなたのGitHubユーザー名>.github.io/fc-revenue-app/`

## Supabase設定
- Dashboard: https://supabase.com/dashboard/project/tyomtjykwsqyfgrdjkvm
- Google認証の設定: Authentication → Providers → Google
- Redirect URL: `https://<あなたのユーザー名>.github.io/fc-revenue-app/`

## ローカル開発
```bash
cp .env.example .env
# .env に Supabase の値を入力
npm install
npm run dev
```
