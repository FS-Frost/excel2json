/// <reference lib="webworker" />
// Service worker (PWA): la app abre y convierte sin conexión. Precachea todo el build (shell + chunks con hash) y el
// grafo del Worker de conversión (Kit solo lista su punto de entrada; el chunk de exceljs no). Origen: FS-Frost/mai-docs
// + precacheWorkerGraph de sensei.
import { immutable, assets } from '$app/manifest';
import { version } from '$app/env';
import { CACHE_PREFIX, KEPT_CACHES } from './lib/client/pwa-core.ts';

const sw = self as unknown as ServiceWorkerGlobalScope;

// Servido en `<base>/service-worker.js`: de ahí sale la base del sitio.
const BASE = sw.location.pathname.replace(/\/service-worker\.js$/, '');
const SHELL = `${BASE}/`;
const CACHE = `${CACHE_PREFIX}${version}`;

const url = (p: string) => `${BASE}/${p.replace(/^\//, '')}`;
const PRECACHE = [...immutable.map((f) => url(f.path)), ...assets.map((f) => url(f.path))];
const PRECACHED = new Set(PRECACHE);

/** Recorre las referencias desde los JS precacheados hacia `/_app/immutable/workers/…` y, dentro de los workers, sus
 *  imports relativos (`./chunks/x.js`), para que la conversión funcione offline desde la primera visita. */
async function precacheWorkerGraph(cache: Cache) {
	const seen = new Set<string>(PRECACHE);
	const isWorker = (p: string) => p.includes('/_app/immutable/workers/');
	const queue = PRECACHE.filter((p) => p.endsWith('.js'));
	while (queue.length) {
		const p = queue.shift()!;
		const res = await cache.match(p);
		if (!res) continue;
		const js = await res.text();
		const refs = [
			...[...js.matchAll(/[`"']([^`"'\s]*\/_app\/immutable\/workers\/[\w./-]+\.js)[`"']/g)].map((m) => m[1]),
			...(isWorker(p) ? [...js.matchAll(/[`"'](\.{1,2}\/[\w./-]+\.js)[`"']/g)].map((m) => new URL(m[1], new URL(p, sw.location.origin)).pathname) : [])
		];
		for (const r of refs) {
			const path = new URL(r, sw.location.origin).pathname;
			if (seen.has(path)) continue;
			seen.add(path);
			try {
				await cache.add(path);
				queue.push(path);
			} catch {
				/* referencia que no es un archivo real */
			}
		}
	}
}

sw.addEventListener('install', (event) => {
	event.waitUntil(
		(async () => {
			const cache = await caches.open(CACHE);
			await cache.addAll([...PRECACHE, SHELL]);
			await precacheWorkerGraph(cache);
			// Sin skipWaiting: una página vieja puede seguir cargando chunks de la caché vieja. La página lo pide
			// (client/pwa.ts → applyUpdate) justo antes de recargar.
		})()
	);
});

sw.addEventListener('message', (event) => {
	if (event.data === 'SKIP_WAITING') void sw.skipWaiting();
});

sw.addEventListener('activate', (event) => {
	event.waitUntil(
		(async () => {
			for (const key of await caches.keys()) if (key.startsWith(CACHE_PREFIX) && key !== CACHE && !KEPT_CACHES.includes(key)) await caches.delete(key);
			await sw.clients.claim();
		})()
	);
});

sw.addEventListener('fetch', (event) => {
	const req = event.request;
	if (req.method !== 'GET') return;
	const u = new URL(req.url);
	if (u.origin !== sw.location.origin) return;

	event.respondWith(
		(async () => {
			const cache = await caches.open(CACHE);
			// Archivos con hash y estáticos: caché primero.
			if (PRECACHED.has(u.pathname) || u.pathname.startsWith(`${BASE}/_app/immutable/`)) {
				const hit = await cache.match(u.pathname);
				if (hit) return hit;
			}
			// Navegación (SPA): red primero; sin conexión → el shell.
			if (req.mode === 'navigate') {
				try {
					return await fetch(req);
				} catch {
					return (await cache.match(SHELL)) ?? Response.error();
				}
			}
			// Resto: red, rellenando la caché; sin conexión → lo cacheado.
			try {
				const res = await fetch(req);
				if (res.ok && res.type === 'basic') void cache.put(req, res.clone());
				return res;
			} catch {
				return (await cache.match(req)) ?? Response.error();
			}
		})()
	);
});
