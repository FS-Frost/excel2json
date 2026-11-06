<script lang="ts">
	// Archivos abiertos con sus hojas: casilla = incluir en la exportación, clic = ver/editar. Opciones de importación.
	import { session } from '../client/session.svelte.ts';
	import { prefs, savePrefs } from '../client/prefs.svelte.ts';
	import { confirmDialog } from '../client/dialogs.svelte.ts';
	import { formatBytes, t, warningText } from '../i18n/index.svelte.ts';
	import type { SourceKind } from '../core/types.ts';
	import Icon, { type IconName } from './Icon.svelte';

	interface Props {
		onadd: () => void;
		onreapply: () => void;
		busy: boolean;
	}

	let { onadd, onreapply, busy }: Props = $props();

	const KIND_ICON: Record<SourceKind, IconName> = { xlsx: 'sheet', csv: 'doc', tsv: 'doc', json: 'braces' };
	const KIND_COLOR: Record<SourceKind, string> = { xlsx: 'text-sheet', csv: 'text-info', tsv: 'text-info', json: 'text-json' };

	// Resumen nuevo por revisión (las tablas se mutan en el lugar: la misma referencia no actualizaría los conteos).
	const files = $derived.by(() => {
		void session.rev;
		return session.files.map((f) => ({ ...f, tables: f.tables.map((tb) => ({ id: tb.id, name: tb.name, rows: tb.rows })) }));
	});
	const total = $derived(files.reduce((n, f) => n + f.tables.length, 0));
	const selectedCount = $derived(session.tables.filter((tb) => session.selected[tb.id]).length);
	const canReapply = $derived(files.some((f) => session.original(f.id)));

	async function remove(id: string, name: string) {
		const f = files.find((x) => x.id === id);
		if (f?.tables.some((tb) => session.edited.has(tb.id)) && !(await confirmDialog(t().files.removeConfirm(name), { confirm: t().files.remove, danger: true }))) return;
		session.removeFile(id);
	}

	async function clearAll() {
		if (await confirmDialog(t().files.clearConfirm, { confirm: t().files.clearAll, danger: true })) session.clear();
	}

	function setHeader(v: number) {
		prefs.import.headerRow = Math.max(0, Math.min(10_000, Math.floor(v) || 0));
		savePrefs();
	}
</script>

<section class="flex flex-col gap-3 p-3" aria-labelledby="files-title">
	<div class="flex items-center justify-between gap-2">
		<h2 id="files-title" class="panel-title">{t().files.heading}</h2>
		<div class="flex items-center gap-1">
			<button class="btn py-1 text-xs" onclick={onadd} disabled={busy}><Icon name="plus" size={14} />{t().files.addMore}</button>
			<button class="tb" title={t().files.clearAll} aria-label={t().files.clearAll} onclick={clearAll}><Icon name="trash" size={16} /></button>
		</div>
	</div>

	<div class="flex items-center justify-between text-xs text-muted">
		<span>{t().files.selected(selectedCount, total)}</span>
		<span class="flex gap-2">
			<button class="hover:text-accent" title={t().files.selectAll} onclick={() => session.selectAll(true)}>{t().common.all}</button>
			<button class="hover:text-accent" title={t().files.selectNone} onclick={() => session.selectAll(false)}>{t().common.none}</button>
		</span>
	</div>

	<ul class="flex flex-col gap-2">
		{#each files as f (f.id)}
			<li class="panel overflow-hidden">
				<div class="flex items-center gap-2 border-b border-line px-2.5 py-2">
					<Icon name={KIND_ICON[f.kind]} size={18} class={KIND_COLOR[f.kind]} />
					<div class="min-w-0 flex-1">
						<div class="truncate text-sm font-medium" title={f.name}>{f.name}</div>
						<div class="text-[11px] text-faint">
							{f.kind.toUpperCase()} · {formatBytes(f.size)} · {t().files.sheets(f.tables.length)}
							{#if session.restored.has(f.id)}<span title={t().restore.noOriginal}> · <Icon name="restore" size={11} class="inline" /></span>{/if}
						</div>
					</div>
					<button class="tb" title={t().files.remove} aria-label="{t().files.remove}: {f.name}" onclick={() => remove(f.id, f.name)}><Icon name="x" size={16} /></button>
				</div>
				<ul>
					{#each f.tables as tb (tb.id)}
						<li class="flex items-center gap-2 px-2.5 py-1 {session.activeId === tb.id ? 'bg-accent-soft' : 'hover:bg-surface-2'}">
							<input
								type="checkbox"
								class="check"
								checked={!!session.selected[tb.id]}
								aria-label={t().files.include(tb.name)}
								onchange={(e) => session.setSelected(tb.id, e.currentTarget.checked)}
							/>
							<button class="flex min-w-0 flex-1 items-center gap-2 text-left text-sm" onclick={() => session.setActive(tb.id)} aria-current={session.activeId === tb.id}>
								<span class="truncate {session.activeId === tb.id ? 'font-medium text-accent' : ''}">{tb.name}</span>
								{#if session.edited.has(tb.id)}<span class="size-1.5 shrink-0 rounded-full bg-warn" title={t().files.edited}></span>{/if}
								<span class="ml-auto shrink-0 text-[11px] text-faint tabular-nums">{t().files.rows(tb.rows.length)}</span>
							</button>
						</li>
					{/each}
				</ul>
				{#if f.warnings.length}
					<details class="border-t border-line px-2.5 py-1.5 text-xs">
						<summary class="cursor-pointer text-warn"><Icon name="info" size={13} class="inline" /> {t().files.warnings(f.warnings.length)}</summary>
						<ul class="mt-1 list-disc space-y-0.5 pl-4 text-muted">
							{#each f.warnings as w, i (i)}<li>{warningText(w)}</li>{/each}
						</ul>
					</details>
				{/if}
			</li>
		{/each}
	</ul>

	<details class="panel px-3 py-2 text-sm">
		<summary class="cursor-pointer panel-title">{t().importOpts.heading}</summary>
		<div class="mt-2 flex flex-col gap-2">
			<label class="flex items-center justify-between gap-2">
				<span>{t().importOpts.headerRow}</span>
				<input
					type="number"
					min="1"
					class="input w-20 py-1 text-right"
					value={prefs.import.headerRow || 1}
					disabled={prefs.import.headerRow === 0}
					onchange={(e) => setHeader(Number(e.currentTarget.value))}
				/>
			</label>
			<label class="flex items-center gap-2">
				<input type="checkbox" class="check" checked={prefs.import.headerRow === 0} onchange={(e) => setHeader(e.currentTarget.checked ? 0 : 1)} />
				{t().importOpts.noHeader}
			</label>
			<label class="flex items-center gap-2">
				<input type="checkbox" class="check" bind:checked={prefs.import.skipEmptyRows} onchange={savePrefs} />
				{t().importOpts.skipEmpty}
			</label>
			<label class="flex items-center gap-2">
				<input type="checkbox" class="check" bind:checked={prefs.import.nested} onchange={savePrefs} />
				{t().importOpts.nested}
			</label>
			<button class="btn mt-1 self-start" disabled={!canReapply || busy} onclick={onreapply}><Icon name="restore" size={14} />{t().importOpts.reapply}</button>
		</div>
	</details>
</section>
