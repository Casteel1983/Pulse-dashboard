import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    outDir: 'dist',
    chunkSizeWarningLimit: 4000,
  },
  define: {
    __VITE_ANTHROPIC_KEY__: JSON.stringify(process.env.VITE_ANTHROPIC_KEY || ''),
  },
})
