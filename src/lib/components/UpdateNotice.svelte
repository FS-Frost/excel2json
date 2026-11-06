<script lang="ts">
	// Versión nueva (short sha), con toasts (origen: FS-Frost/mai-docs):
	// - al volver a la pestaña sin archivos abiertos → se actualiza sola;
	// - con archivos abiertos o detectada con la pestaña visible → toast «Actualizar» (nunca recarga en medio del trabajo);
	// - tras actualizar, «excel2json actualizado (sha)»; si la recarga no la trajo, error con «Forzar».
	import { onMount } from 'svelte';
	import { updated } from '$app/state';
	import { version } from '$app/env';
	import { applyUpdate, failedUpdate, forceReload } from '../client/pwa.ts';
	import { noteVersion } from '../client/version.ts';
	import { toast } from '../client/toast.svelte.ts';
	import { session } from '../client/session.svelte.ts';
	import { flushSave } from '../client/persist.svelte.ts';
	import { ago, t } from '../i18n/index.svelte.ts';

	const builtAt = import.meta.env.VITE_BUILT_AT as string | undefined;

	function offer() {
		toast.info(t().update.available, { duration: null, key: 'version', actions: [{ label: t().update.action, run: update, keep: true }] });
	}

	async function update() {
		toast.info(t().update.updating, { duration: null, key: 'version', persistent: true, detail: t().update.updatingDetail });
		await flushSave();
		const r = await applyUpdate();
		if (r === 'publishing')
			toast.warn(t().update.publishing, {
				detail: t().update.publishingDetail,
				duration: null,
				key: 'version',
				actions: [{ label: t().common.retry, run: update, keep: true }]
			});
		else if (r === 'offline') toast.warn(t().update.offline, { key: 'version', actions: [{ label: t().common.retry, run: update, keep: true }] });
	}

	onMount(() => {
		const missed = version === 'dev' ? null : failedUpdate(version);
		if (missed) toast.error(t().update.failed(missed), { detail: t().update.failedDetail, key: 'version', actions: [{ label: t().update.force, run: forceReload, keep: true }] });
		else if (noteVersion(version) !== null)
			toast.success(t().update.done(version), { detail: builtAt ? t().update.publishedAgo(ago(Date.parse(builtAt))) : undefined, duration: 8000, key: 'version' });
		const onVisible = async () => {
			if (document.visibilityState !== 'visible' || version === 'dev') return;
			try {
				if (updated.current || (await updated.check())) {
					if (session.files.length) offer();
					else await applyUpdate();
				}
			} catch {
				/* sin conexión */
			}
		};
		document.addEventListener('visibilitychange', onVisible);
		return () => document.removeEventListener('visibilitychange', onVisible);
	});

	$effect(() => {
		if (updated.current) offer();
	});
</script>
