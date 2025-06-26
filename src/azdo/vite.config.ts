import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  base: './',
  plugins: [react()],
  build: {
    rollupOptions: {
      input: {
        // index: './web/index.html',
        mergequeue: './web/mergequeue/index.html'
      },
    }
  }
})
