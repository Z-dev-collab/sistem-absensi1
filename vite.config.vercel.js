import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// Vercel build: standard Vite SPA. No Laravel plugin, so index.html is emitted
// into dist/ and can be served directly.
export default defineConfig({
    base: '/',
    publicDir: false,
    plugins: [react(), tailwindcss()],
    build: {
        outDir: 'dist',
        emptyOutDir: true,
    },
})
