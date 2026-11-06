<script lang="ts">
	// Pantalla inicial: soltar/elegir archivos, pegar datos, probar un ejemplo y restaurar la sesión anterior.
	import { saved } from '../client/persist.svelte.ts';
	import { ago, t } from '../i18n/index.svelte.ts';
	import Icon from './Icon.svelte';
	import Logo from './Logo.svelte';

	interface Props {
		onchoose: () => void;
		onpaste: () => void;
		onsample: () => void;
		onrestore: () => void;
		ondiscard: () => void;
		busy: string | null;
	}

	let { onchoose, onpaste, onsample, onrestore, ondiscard, busy }: Props = $props();
</script>

<div class="mx-auto flex w-full max-w-3xl flex-1 flex-col gap-5 px-4 py-6 sm:py-10">
	{#if saved.meta}
		<div class="panel flex flex-wrap items-center gap-3 border-accent/40 px-4 py-3" role="region" aria-label={t().restore.title}>
			<Icon name="restore" class="text-accent" />
			<div class="min-w-0 flex-1">
				<div class="text-sm font-semibold">{t().restore.title}</div>
				<div class="text-xs text-muted">
					{t().restore.detail(saved.meta.files.length, ago(saved.meta.savedAt))}
					<span class="text-faint">{saved.meta.files.map((f) => f.name).join(', ')}</span>
				</div>
				<div class="mt-0.5 text-[11px] text-faint">{t().restore.replacedWarn}</div>
			</div>
			<div class="flex gap-2">
				<button class="btn" onclick={ondiscard}>{t().restore.discard}</button>
				<button class="btn btn-primary" onclick={onrestore}>{t().restore.restore}</button>
			</div>
		</div>
	{/if}

	<button
		class="group flex min-h-72 flex-1 flex-col items-center justify-center gap-4 rounded-2xl border-2 border-dashed border-line bg-surface px-6 py-10 text-center transition-colors hover:border-accent hover:bg-accent-soft/40 focus-visible:border-accent"
		onclick={onchoose}
		disabled={!!busy}
		aria-describedby="drop-sub"
	>
		{#if busy}
			<span class="size-10 animate-spin rounded-full border-4 border-line border-t-accent"></span>
			<span class="text-sm text-muted" aria-live="polite">{busy}</span>
		{:else}
			<Logo size={64} />
			<span class="text-xl font-semibold tracking-tight sm:text-2xl">{t().drop.title}</span>
			<span id="drop-sub" class="max-w-md text-sm text-muted">{t().drop.subtitle}</span>
			<span class="btn btn-primary pointer-events-none mt-1 px-4 py-2"><Icon name="upload" size={16} />{t().drop.choose}</span>
		{/if}
	</button>

	<div class="flex flex-wrap items-center justify-center gap-2">
		<button class="btn" onclick={onpaste} disabled={!!busy}><Icon name="clipboard" size={16} />{t().drop.paste}</button>
		<button class="btn" onclick={onsample} disabled={!!busy}><Icon name="sparkle" size={16} />{t().drop.sample}</button>
	</div>

	<ul class="grid gap-3 text-sm text-muted sm:grid-cols-3">
		<li class="flex gap-2"><Icon name="files" class="mt-0.5 text-accent" size={18} /><span>{t().export.modes.combined} · {t().export.modes.merged}</span></li>
		<li class="flex gap-2"><Icon name="table" class="mt-0.5 text-accent" size={18} /><span>{t().grid.find} · {t().grid.sortAsc} · {t().grid.filter}</span></li>
		<li class="flex gap-2"><Icon name="eyeOff" class="mt-0.5 text-accent" size={18} /><span>{t().app.privacy}</span></li>
	</ul>
</div>
