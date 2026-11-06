<script lang="ts">
	// Idioma: menú con «sistema», español e inglés.
	import ContextMenu, { type MenuAt } from './ContextMenu.svelte';
	import Icon from './Icon.svelte';
	import { i18n, setLangPref, t, type LangPref } from '../i18n/index.svelte.ts';

	let at = $state<MenuAt | null>(null);
	const OPTIONS: { pref: LangPref; label: () => string }[] = [
		{ pref: 'system', label: () => t().header.langSystem },
		{ pref: 'es', label: () => 'Español' },
		{ pref: 'en', label: () => 'English' }
	];

	function open(e: MouseEvent) {
		const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
		at = { x: r.left, y: r.bottom + 4 };
	}
</script>

<button class="tb gap-1 px-2 text-xs font-semibold uppercase" title={t().header.language} aria-label={t().header.language} aria-haspopup="menu" onclick={open}>
	<Icon name="globe" size={16} />
	{i18n.locale}
</button>
<ContextMenu {at} onclose={() => (at = null)} label={t().header.language}>
	{#snippet children(close)}
		{#each OPTIONS as o (o.pref)}
			<button
				class="menu-item"
				role="menuitemradio"
				aria-checked={i18n.pref === o.pref}
				onclick={() => {
					setLangPref(o.pref);
					close();
				}}
			>
				<span class="w-4">{#if i18n.pref === o.pref}<Icon name="check" size={14} />{/if}</span>
				{o.label()}
			</button>
		{/each}
	{/snippet}
</ContextMenu>
