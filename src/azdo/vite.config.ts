import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import babel from 'vite-plugin-babel'
import commonjs from 'vite-plugin-commonjs'

// https://vite.dev/config/
export default defineConfig({
  base: './',
  plugins: [
    babel({
      filter: /.*/,
      include: [
        /azure-devops-extension-api/, 
        /azure-devops-extension-sdk/
      ],
      babelConfig: {
        plugins: ['transform-amd-to-commonjs'],
      },
    }),
    commonjs({
      filter: (id) => 
        id.includes('azure-devops-extension-api') ||
        id.includes('azure-devops-extension-sdk'),
    }),
    react()
  ],
  css: {
    lightningcss: {
      errorRecovery: true // for legacy Azure DevOps CSS
    }
  },
  build: {
    rolldownOptions: {
      input: {
        // index: './web/index.html',
        currentruns: './web/currentruns/index.html',
        mergequeue: './web/mergequeue/index.html'
      },
    }
  }
})
