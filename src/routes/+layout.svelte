<script lang="ts">
	import '../app.css';
	import { onMount } from 'svelte';
	import type { LayoutProps } from './$types';
	import Toaster from '#lib/components/Toaster.svelte';
	import Dialogs from '#lib/components/Dialogs.svelte';
	import UpdateNotice from '#lib/components/UpdateNotice.svelte';
	import { initTheme } from '#lib/client/theme.svelte.ts';
	import { initI18n, t } from '#lib/i18n/index.svelte.ts';
	import { registerServiceWorker } from '#lib/client/pwa.ts';
	import { toast } from '#lib/client/toast.svelte.ts';

	let { children }: LayoutProps = $props();

	onMount(() => {
		initTheme();
		initI18n();
		void registerServiceWorker();
		// «Conexión recuperada» solo tras un aviso de «sin conexión» en esta sesión, nunca al cargar.
		let wasOffline = false;
		const offline = () => {
			wasOffline = true;
			toast.info(t().net.offline, { detail: t().net.offlineDetail, key: 'net' });
		};
		const online = () => {
			toast.dismissKey('net');
			if (wasOffline) toast.success(t().net.online, { key: 'net' });
			wasOffline = false;
		};
		addEventListener('offline', offline);
		addEventListener('online', online);
		return () => {
			removeEventListener('offline', offline);
			removeEventListener('online', online);
		};
	});
</script>

<svelte:head>
	<title>{t().app.name}</title>
	<meta name="description" content={t().app.tagline} />
</svelte:head>

{@render children()}

<Toaster />
<Dialogs />
<UpdateNotice />
