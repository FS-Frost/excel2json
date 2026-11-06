<script lang="ts">
	import { onMount } from 'svelte';
	import { version } from '$app/env';
	import AppHeader from '#lib/components/AppHeader.svelte';
	import ExportPanel from '#lib/components/ExportPanel.svelte';
	import FilesPanel from '#lib/components/FilesPanel.svelte';
	import Icon from '#lib/components/Icon.svelte';
	import Landing from '#lib/components/Landing.svelte';
	import PasteDialog from '#lib/components/PasteDialog.svelte';
	import SheetTabs from '#lib/components/SheetTabs.svelte';
	import DataGrid, { type GridSel } from '#lib/components/grid/DataGrid.svelte';
	import FindBar from '#lib/components/grid/FindBar.svelte';
	import { ACCEPT, sourceKind } from '#lib/core/read.ts';
	import type { SourceFile } from '#lib/core/types.ts';
	import { isActive, visibleRows } from '#lib/edit/view.ts';
	import { cancelAll, converter, parseFile, sampleWorkbook } from '#lib/client/converter.svelte.ts';
	import { confirmDialog } from '#lib/client/dialogs.svelte.ts';
	import { discardSaved, initPersistence, persistReady, restoreSaved, saved, saveStatus } from '#lib/client/persist.svelte.ts';
	import { importOptions } from '#lib/client/prefs.svelte.ts';
	import { session } from '#lib/client/session.svelte.ts';
	import { toast } from '#lib/client/toast.svelte.ts';
	import { errorText, formatBytes, i18n, t, warningText } from '#lib/i18n/index.svelte.ts';

	const BIG = 50 * 1024 * 1024;

	let fileInput = $state<HTMLInputElement>();
	let grid = $state<ReturnType<typeof DataGrid>>();
	let pasteOpen = $state(false);
	let finding = $state(false);
	let replacing = $state(false);
	let highlight = $state<Set<string> | null>(null);
	let reading = $state<string | null>(null);
	let dragging = $state(0);
	let sel = $state<GridSel>({ r1: 0, c1: 0, r2: 0, c2: 0 });

	const hasFiles = $derived(session.files.length > 0);
	// Copia superficial por revisión: la tabla se muta en el lugar y la misma referencia no actualizaría la vista.
	const active = $derived.by(() => {
		void session.rev;
		const a = session.active;
		return a ? { ...a } : null;
	});
	const filters = $derived(session.activeId ? (session.views[session.activeId]?.filters ?? {}) : {});
	const filterCount = $derived(Object.values(filters).filter(isActive).length);
	const visibleCount = $derived.by(() => {
		void session.rev;
		if (!active) return 0;
		return filterCount ? visibleRows(active, filters).length : active.rows.length;
	});

	// ---------- Abrir archivos ----------

	async function addFiles(list: File[]) {
		if (!list.length) return;
		await persistReady;
		const added: { file: SourceFile; original?: File }[] = [];
		try {
			for (const f of list) {
				let bytes: ArrayBuffer;
				try {
					bytes = await f.arrayBuffer();
				} catch (e) {
					toast.error(t().drop.readFailed(f.name), { detail: String((e as Error)?.message ?? e) });
					continue;
				}
				if (!sourceKind(f.name, new Uint8Array(bytes, 0, Math.min(64, bytes.byteLength)))) {
					toast.error(t().drop.unsupported(f.name));
					continue;
				}
				if (f.size > BIG) toast.info(t().drop.bigFile(f.name, formatBytes(f.size)), { key: 'big' });
				reading = t().drop.reading(f.name);
				try {
					const file = await parseFile(f.name, bytes, importOptions(), (p) => {
						if (p.label) reading = t().drop.readingSheet(f.name, p.label);
					});
					added.push({ file, original: f });
					if (file.warnings.length)
						toast.warn(f.name, { detail: file.warnings.slice(0, 3).map(warningText).join(' · '), key: `warn-${f.name}` });
				} catch (e) {
					if ((e as { code?: string })?.code === 'cancelled') break;
					toast.error(t().drop.readFailed(f.name), { detail: errorText(e) });
				}
			}
		} finally {
			reading = null;
		}
		if (!added.length) return;
		// La sesión guardada se descarta recién cuando hay algo nuevo que la reemplace (un archivo inválido no la borra).
		if (saved.meta) await discardSaved();
		session.addFiles(added);
		session.setActive(added[0].file.tables[0]?.id ?? session.activeId);
		toast.success(t().drop.added(added.length), { key: 'added', duration: 2500 });
	}

	function choose() {
		fileInput?.click();
	}

	function onPicked(e: Event) {
		const input = e.currentTarget as HTMLInputElement;
		const files = [...(input.files ?? [])];
		// Limpiar recién al terminar: hacerlo antes puede invalidar los File en algunos navegadores.
		void addFiles(files).finally(() => (input.value = ''));
	}

	async function loadSample() {
		reading = t().common.loading;
		try {
			const buf = await sampleWorkbook(i18n.locale);
			await addFiles([new File([buf], t().sample.file)]);
		} catch (e) {
			toast.error(errorText(e));
		} finally {
			reading = null;
		}
	}

	async function reapply() {
		const edited = session.files.some((f) => session.original(f.id) && f.tables.some((tb) => session.edited.has(tb.id)));
		if (edited && !(await confirmDialog(t().importOpts.reapplyConfirm, { confirm: t().importOpts.reapply, danger: true }))) return;
		const opts = importOptions();
		for (const f of [...session.files]) {
			const original = session.original(f.id);
			if (!original) continue;
			reading = t().drop.reading(f.name);
			try {
				const next = await parseFile(original.name, await original.arrayBuffer(), opts);
				session.replaceFile(f.id, next);
			} catch (e) {
				toast.error(t().drop.readFailed(f.name), { detail: errorText(e) });
			}
		}
		reading = null;
		toast.success(t().importOpts.reapplied, { key: 'reapply' });
	}

	async function restore() {
		try {
			await restoreSaved();
			toast.success(t().restore.restored, { key: 'restore', duration: 2500 });
		} catch (e) {
			toast.error(errorText(e));
		}
	}

	// ---------- Arrastrar y soltar (en cualquier parte de la página) ----------

	const hasFilesIn = (e: DragEvent) => !!e.dataTransfer && [...e.dataTransfer.types].includes('Files');

	function dragenter(e: DragEvent) {
		if (!hasFilesIn(e)) return;
		e.preventDefault();
		dragging++;
	}
	function dragover(e: DragEvent) {
		if (!hasFilesIn(e)) return;
		e.preventDefault();
		if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
	}
	function dragleave(e: DragEvent) {
		if (!hasFilesIn(e)) return;
		dragging = Math.max(0, dragging - 1);
	}
	function drop(e: DragEvent) {
		if (!hasFilesIn(e)) return;
		e.preventDefault();
		dragging = 0;
		void addFiles([...(e.dataTransfer?.files ?? [])]);
	}

	// ---------- Atajos globales ----------

	function keydown(e: KeyboardEvent) {
		const mod = e.ctrlKey || e.metaKey;
		const target = e.target as HTMLElement | null;
		const typing = !!target?.closest('input, textarea, select, [contenteditable]');
		if (mod && e.key.toLowerCase() === 'o') {
			e.preventDefault();
			choose();
			return;
		}
		if (typing || !hasFiles) return;
		if (mod && e.key.toLowerCase() === 'z') {
			e.preventDefault();
			showTable(e.shiftKey ? session.redo() : session.undo());
		} else if (mod && e.key.toLowerCase() === 'y') {
			e.preventDefault();
			showTable(session.redo());
		} else if (mod && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'h')) {
			e.preventDefault();
			openFind(e.key.toLowerCase() === 'h');
		}
	}

	function showTable(id: string | null) {
		if (id && id !== session.activeId) session.setActive(id);
	}

	function undo() {
		showTable(session.undo());
		grid?.focusSink();
	}
	function redo() {
		showTable(session.redo());
		grid?.focusSink();
	}

	function openFind(replace: boolean) {
		finding = true;
		replacing = replace || replacing;
	}

	function closeFind() {
		finding = false;
		replacing = false;
		highlight = null;
		grid?.focusSink();
	}

	onMount(() => {
		void initPersistence().then(() => {
			const q = new URLSearchParams(location.search);
			if (q.has('demo') || q.has('ejemplo')) void loadSample();
		});
		// Abierto desde el sistema («Abrir con excel2json», manifest file_handlers, Chromium).
		window.launchQueue?.setConsumer(async ({ files }) => {
			const list: File[] = [];
			for (const h of files) {
				try {
					list.push(await h.getFile());
				} catch {
					/* sin permiso */
				}
			}
			void addFiles(list);
		});
		// Hay trabajo sin exportar → el navegador pregunta antes de cerrar (además del autoguardado).
		const beforeUnload = (e: BeforeUnloadEvent) => {
			if (saveStatus() !== 'saved' && session.edited.size) e.preventDefault();
		};
		addEventListener('beforeunload', beforeUnload);
		return () => removeEventListener('beforeunload', beforeUnload);
	});

	const statusText = $derived.by(() => {
		if (!hasFiles) return '';
		const s = saveStatus();
		return s === 'saved' ? t().autosave.saved : s === 'saving' ? t().autosave.saving : s === 'error' ? t().autosave.error : t().autosave.dirty;
	});

	$effect(() => {
		if (saved.lowSpace) toast.warn(t().autosave.quota, { key: 'quota' });
	});

	$effect(() => {
		if (saveStatus() === 'error') toast.error(t().autosave.error, { key: 'autosave' });
	});

	// Al cambiar de tabla, se cierra la búsqueda (sus resultados eran de la otra).
	$effect(() => {
		void session.activeId;
		highlight = null;
	});

	const selRows = $derived(sel.r2 - sel.r1 + 1);
	const selCols = $derived(sel.c2 - sel.c1 + 1);
</script>

<svelte:window ondragenter={dragenter} ondragover={dragover} ondragleave={dragleave} ondrop={drop} onkeydown={keydown} />

<input bind:this={fileInput} type="file" class="hidden" multiple accept={ACCEPT} onchange={onPicked} />

<div class="flex min-h-dvh flex-col lg:h-dvh">
	<AppHeader>
		{#snippet status()}
			{#if statusText}
				<span class="inline-flex items-center gap-1"><Icon name={saveStatus() === 'error' ? 'info' : 'database'} size={13} />{statusText}</span>
			{:else}
				{t().app.tagline}
			{/if}
		{/snippet}
	</AppHeader>

	{#if !hasFiles}
		<Landing busy={reading} onchoose={choose} onpaste={() => (pasteOpen = true)} onsample={loadSample} onrestore={restore} ondiscard={() => discardSaved()} />
		<footer class="flex flex-wrap items-center justify-center gap-x-4 gap-y-1 px-4 pb-5 text-xs text-faint">
			<span>{t().footer.version(version)}</span>
			<span class="inline-flex items-center gap-1"><Icon name="eyeOff" size={12} />{t().app.privacy}</span>
			<span>{t().footer.offline}</span>
			<a class="hover:text-accent" href="https://github.com/FS-Frost/excel2json" target="_blank" rel="noopener noreferrer">{t().footer.source}</a>
		</footer>
	{:else}
		<div class="flex flex-1 flex-col lg:min-h-0 lg:flex-row">
			<aside class="shrink-0 border-b border-line bg-bg lg:w-72 lg:overflow-y-auto lg:border-r lg:border-b-0" aria-label={t().files.heading}>
				<FilesPanel onadd={choose} onreapply={reapply} busy={!!reading} />
			</aside>

			<main class="flex h-[70dvh] min-w-0 flex-col lg:h-auto lg:min-h-0 lg:flex-1">
				<SheetTabs />
				{#if active}
					<div class="flex shrink-0 flex-wrap items-center gap-0.5 border-b border-line bg-surface px-2 py-1" role="toolbar" aria-label={active.name}>
						<button class="tb" title="{t().common.undo} (Ctrl+Z)" aria-label={t().common.undo} disabled={!session.canUndo} onclick={undo}><Icon name="undo" /></button>
						<button class="tb" title="{t().common.redo} (Ctrl+Y)" aria-label={t().common.redo} disabled={!session.canRedo} onclick={redo}><Icon name="redo" /></button>
						<span class="tb-sep"></span>
						<button class="tb" title={t().grid.addRow} aria-label={t().grid.addRow} onclick={() => grid?.addRow()}><Icon name="rowAdd" /></button>
						<button class="tb" title={t().grid.addCol} aria-label={t().grid.addCol} onclick={() => grid?.addCol()}><Icon name="colAdd" /></button>
						<button class="tb" title={t().grid.deleteRows(selRows)} aria-label={t().grid.deleteRows(selRows)} disabled={!active.rows.length} onclick={() => grid?.deleteRows()}><Icon name="rowDel" /></button>
						<button class="tb" title={t().grid.deleteCols(selCols)} aria-label={t().grid.deleteCols(selCols)} disabled={!active.headers.length} onclick={() => grid?.deleteCols()}><Icon name="colDel" /></button>
						<span class="tb-sep"></span>
						<button class="tb" title={t().grid.sortAsc} aria-label={t().grid.sortAsc} disabled={!active.rows.length} onclick={() => grid?.sortBy(sel.c1, 'asc')}><Icon name="sortAsc" /></button>
						<button class="tb" title={t().grid.sortDesc} aria-label={t().grid.sortDesc} disabled={!active.rows.length} onclick={() => grid?.sortBy(sel.c1, 'desc')}><Icon name="sortDesc" /></button>
						<button class="tb" title="{t().grid.find} (Ctrl+F)" aria-label={t().grid.find} aria-pressed={finding} onclick={() => (finding ? closeFind() : openFind(false))}><Icon name="search" /></button>
						{#if filterCount}
							<button class="tb gap-1 px-2 text-xs text-accent" title={t().grid.clearFilters} onclick={() => grid?.clearFilters()}>
								<Icon name="filter" size={15} />{t().grid.filtered(filterCount)}<Icon name="x" size={13} />
							</button>
						{/if}
						<details class="relative ml-auto">
							<summary class="tb cursor-pointer list-none" title={t().grid.shortcuts} aria-label={t().grid.shortcuts}><Icon name="keyboard" /></summary>
							<div class="menu absolute right-0 mt-1 w-80 p-3 text-xs">
								<table class="w-full">
									<tbody>
										{#each t().grid.shortcutList as [k, d] (k)}
											<tr><td class="py-0.5 pr-2 font-mono whitespace-nowrap text-muted">{k}</td><td class="py-0.5">{d}</td></tr>
										{/each}
									</tbody>
								</table>
							</div>
						</details>
					</div>
					{#if finding}
						<FindBar
							tableId={active.id}
							bind:replace={replacing}
							onreveal={(r, c) => grid?.reveal(r, c)}
							onmatches={(m) => (highlight = m)}
							onclose={closeFind}
						/>
					{/if}
					<div class="min-h-0 flex-1">
						{#key active.id}
							<DataGrid bind:this={grid} tableId={active.id} {highlight} onfind={openFind} onselect={(s) => (sel = s)} onundo={(r) => (r ? redo() : undo())} />
						{/key}
					</div>
					<div class="flex shrink-0 items-center gap-3 border-t border-line bg-surface px-3 py-1 text-xs text-muted tabular-nums">
						<span>{t().grid.rowsOf(filterCount ? visibleCount : active.rows.length, active.rows.length)} · {t().files.cols(active.headers.length)}</span>
						{#if selRows > 1 || selCols > 1}<span class="text-faint">{t().grid.sel(selRows, selCols)}</span>{/if}
						{#if converter.busy && converter.progress}
							<span class="ml-auto">{t().export.building(converter.progress.done, converter.progress.total)}</span>
						{/if}
					</div>
				{/if}
			</main>

			<aside class="shrink-0 border-t border-line bg-bg lg:w-80 lg:overflow-y-auto lg:border-t-0 lg:border-l" aria-label={t().export.heading}>
				<ExportPanel />
			</aside>
		</div>
	{/if}
</div>

{#if reading && hasFiles}
	<div class="fixed bottom-4 left-1/2 z-[60] flex -translate-x-1/2 items-center gap-3 rounded-lg border border-line bg-surface px-4 py-2 text-sm shadow-xl" role="status">
		<span class="size-4 animate-spin rounded-full border-2 border-line border-t-accent"></span>
		{reading}
		<button class="btn py-0.5 text-xs" onclick={cancelAll}>{t().export.cancel}</button>
	</div>
{/if}

{#if dragging}
	<div class="pointer-events-none fixed inset-0 z-[90] grid place-items-center bg-accent/15 p-6 backdrop-blur-[2px]">
		<div class="flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-accent bg-surface px-10 py-8 text-lg font-semibold shadow-2xl">
			<Icon name="upload" size={36} class="text-accent" />
			{t().drop.overlay}
		</div>
	</div>
{/if}

<PasteDialog bind:open={pasteOpen} onadd={(f) => addFiles([f])} />
