import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: './',           // works for GitHub Pages subfolder deploy
  build: {
    outDir: 'dist',
    chunkSizeWarningLimit: 4000,
  },
})
