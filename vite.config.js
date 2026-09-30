import { defineConfig } from 'vite'
import laravel from 'laravel-vite-plugin'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// GitHub Pages serves the app under /<repo>/, so assets need that base there.
// Laravel/Render (root domain) keep the default "/".
const base = process.env.GH_PAGES ? '/sistem-absensi1/' : '/'

export default defineConfig({
    base,
    plugins: [
        laravel({
            input: ['resources/js/main.tsx'],
            refresh: true,
        }),
        react(),
        tailwindcss(),
    ],
    server: {
        watch: {
            ignored: ['**/storage/framework/views/**'],
        },
    },
})

