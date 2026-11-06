import adapter from '@sveltejs/adapter-static';
import { sveltekit } from '@sveltejs/kit/vite';
import tailwindcss from '@tailwindcss/vite';
import { defineConfig } from 'vite';

// Base del sitio: '' en local; en GitHub Pages '/<repo>' (lo define el CI con BASE_PATH).
const base = (process.env.BASE_PATH ?? '') as '' | `/${string}`;
// Versión = short sha del commit (GITHUB_SHA en Actions); 'dev' en local.
const version = process.env.GITHUB_SHA?.slice(0, 7) || 'dev';
process.env.VITE_BUILT_AT ??= new Date().toISOString();

export default defineConfig({
	plugins: [
		tailwindcss(),
		sveltekit({
			compilerOptions: {
				runes: ({ filename }) => (filename.split(/[/\\]/).includes('node_modules') ? undefined : true)
			},
			// SPA: 404.html sirve cualquier ruta en GitHub Pages (el CI lo copia a index.html).
			adapter: adapter({ pages: 'build', assets: 'build', fallback: '404.html', precompress: false, strict: false }),
			paths: { base, relative: false },
			// Se registra a mano y solo en builds de producción (client/pwa.ts).
			serviceWorker: { register: false },
			version: { name: version, pollInterval: version === 'dev' ? 0 : 5 * 60_000 }
		})
	],
	worker: { format: 'es' },
	server: { port: Number(process.env.PORT ?? 5173), strictPort: true },
	build: { chunkSizeWarningLimit: 1500 }
});
