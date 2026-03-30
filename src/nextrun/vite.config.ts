import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  base: './',
  plugins: [react()],
  css: {
    lightningcss: {
      errorRecovery: true // for legacy Azure DevOps CSS
    }
  },
  build: {
    rolldownOptions: {
      input: {
        web: './web/index.html'
      },
    }
  }
})
