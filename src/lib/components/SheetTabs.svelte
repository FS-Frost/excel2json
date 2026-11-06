<script lang="ts">
	// Pestañas de todas las tablas abiertas. Doble clic = renombrar la hoja (con deshacer).
	import { session } from '../client/session.svelte.ts';
	import { promptDialog } from '../client/dialogs.svelte.ts';
	import { baseName } from '../core/naming.ts';
	import { t } from '../i18n/index.svelte.ts';

	const tabs = $derived.by(() => {
		void session.rev;
		const many = session.files.length > 1;
		return session.files.flatMap((f) => f.tables.map((tb) => ({ id: tb.id, name: tb.name, file: many && f.tables.length > 0 ? baseName(f.name) : '' })));
	});

	let strip = $state<HTMLDivElement>();

	$effect(() => {
		const id = session.activeId;
		strip?.querySelector<HTMLElement>(`[data-tab="${id}"]`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
	});

	async function rename(id: string, current: string) {
		const name = await promptDialog(t().common.rename, current, {
			confirm: t().common.rename,
			validate: (v) => (!v.trim() ? t().grid.renameEmpty : null)
		});
		if (name && name.trim() !== current) session.exec(id, [{ t: 'name', name: name.trim() }]);
	}

	function keydown(e: KeyboardEvent) {
		if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
		const i = tabs.findIndex((x) => x.id === session.activeId);
		const next = tabs[(i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length];
		if (next) {
			e.preventDefault();
			session.setActive(next.id);
			strip?.querySelector<HTMLElement>(`[data-tab="${next.id}"]`)?.focus();
		}
	}
</script>

<div bind:this={strip} class="flex shrink-0 gap-0.5 overflow-x-auto border-b border-line bg-bg px-2 pt-1.5" role="tablist" tabindex="-1" onkeydown={keydown}>
	{#each tabs as tab (tab.id)}
		{@const on = session.activeId === tab.id}
		<button
			data-tab={tab.id}
			role="tab"
			aria-selected={on}
			tabindex={on ? 0 : -1}
			class="max-w-56 shrink-0 truncate rounded-t-md border border-b-0 px-3 py-1.5 text-sm {on ? 'border-line bg-surface font-medium text-fg' : 'border-transparent text-muted hover:bg-surface-2'}"
			title={tab.file ? `${tab.file} › ${tab.name}` : tab.name}
			onclick={() => session.setActive(tab.id)}
			ondblclick={() => rename(tab.id, tab.name)}
		>
			{#if tab.file}<span class="text-faint">{`${tab.file} › `}</span>{/if}{tab.name}
		</button>
	{/each}
</div>
