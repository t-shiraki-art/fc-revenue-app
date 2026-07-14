import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // GitHub Pages用: リポジトリ名に合わせる
  base: '/fc-revenue-app/',
  build: {
    outDir: 'dist',
    chunkSizeWarningLimit: 2000,
  }
})
