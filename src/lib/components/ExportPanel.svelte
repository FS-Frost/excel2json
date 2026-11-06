<script lang="ts">
	// Exportar: formato, modo (cómo se agrupan las hojas) y opciones. La conversión corre en el Worker.
	import { buildFile, cancelAll, converter } from '../client/converter.svelte.ts';
	import { copyText, downloadFile, fileText } from '../client/download.ts';
	import { exportOptions, prefs, resetExportPrefs, savePrefs } from '../client/prefs.svelte.ts';
	import { session } from '../client/session.svelte.ts';
	import { toast } from '../client/toast.svelte.ts';
	import { MODES_BY_FORMAT, type CsvDelimiter, type OutputFormat, type OutputMode } from '../core/types.ts';
	import { errorText, formatBytes, t } from '../i18n/index.svelte.ts';
	import Icon from './Icon.svelte';

	const FORMATS: OutputFormat[] = ['json', 'xlsx', 'csv'];

	let running = $state<'download' | 'copy' | null>(null);
	let progress = $state<{ done: number; total: number } | null>(null);
	let last = $state<string | null>(null);

	const e = $derived(prefs.export);
	const modes = $derived(MODES_BY_FORMAT[e.format]);
	const mode = $derived<OutputMode>(modes.includes(e.modes[e.format]) ? e.modes[e.format] : modes[0]);
	const tables = $derived.by(() => {
		void session.rev;
		return session.exportTables(e.onlyVisible);
	});
	const rowCount = $derived(tables.reduce((n, tb) => n + tb.rows.length, 0));
	const anyFilter = $derived.by(() => {
		void session.rev;
		return session.tables.some((tb) => session.selected[tb.id] && Object.values(session.views[tb.id]?.filters ?? {}).some((f) => f.query?.trim() || Array.isArray(f.values)));
	});
	const copyable = $derived(e.format !== 'xlsx' && (tables.length === 1 || mode === 'merged' || (e.format === 'json' && mode === 'combined')));

	function setFormat(f: OutputFormat) {
		prefs.export.format = f;
		savePrefs();
	}

	function setMode(m: OutputMode) {
		prefs.export.modes[e.format] = m;
		savePrefs();
	}

	async function run(kind: 'download' | 'copy') {
		if (running) return;
		running = kind;
		progress = null;
		try {
			const opts = { ...exportOptions(t().export.originName), mode };
			const res = await buildFile(tables, opts, (p) => (progress = p));
			if (kind === 'copy') {
				if (res.count > 1) {
					toast.warn(t().export.copyMany);
					return;
				}
				await copyText(fileText(res.file));
				toast.success(t().export.copied, { key: 'export' });
			} else {
				downloadFile(res.file);
				toast.success(t().export.downloaded(res.file.name), { key: 'export', detail: t().export.summary(res.count, formatBytes(res.file.bytes.length)) });
			}
			last = t().export.summary(res.count, formatBytes(res.file.bytes.length));
		} catch (err) {
			const code = (err as { code?: string })?.code;
			if (code === 'cancelled') toast.info(t().errors.cancelled, { key: 'export' });
			else toast.error(errorText(err), { key: 'export' });
		} finally {
			running = null;
			progress = null;
		}
	}

	const DELIMS: { v: CsvDelimiter; label: () => string }[] = [
		{ v: ',', label: () => t().export.delimiters.comma },
		{ v: ';', label: () => t().export.delimiters.semicolon },
		{ v: '\t', label: () => t().export.delimiters.tab }
	];
</script>

<section class="flex flex-col gap-4 p-3" aria-labelledby="export-title">
	<div class="flex items-center justify-between">
		<h2 id="export-title" class="panel-title">{t().export.heading}</h2>
		<button class="text-xs text-muted hover:text-accent" onclick={resetExportPrefs}>{t().export.reset}</button>
	</div>

	<div class="seg w-full" role="group" aria-label={t().export.format}>
		{#each FORMATS as f (f)}
			<button class="flex-1" aria-pressed={e.format === f} onclick={() => setFormat(f)}>{t().export.formats[f]}</button>
		{/each}
	</div>

	<fieldset class="flex flex-col gap-1">
		<legend class="mb-1 text-xs font-semibold text-muted">{t().export.mode}</legend>
		{#each modes as m (m)}
			<label class="flex cursor-pointer gap-2 rounded-md border px-2.5 py-2 {mode === m ? 'border-accent bg-accent-soft' : 'border-line hover:bg-surface-2'}">
				<input type="radio" class="check mt-0.5" name="mode" checked={mode === m} onchange={() => setMode(m)} />
				<span class="min-w-0">
					<span class="block text-sm font-medium">{t().export.modes[m]}</span>
					<span class="block text-xs text-muted">{t().export.modeHelp[e.format][m]}</span>
				</span>
			</label>
		{/each}
	</fieldset>

	<fieldset class="flex flex-col gap-2 text-sm">
		<legend class="mb-1 text-xs font-semibold text-muted">{t().export.options}</legend>
		{#if e.format === 'json'}
			<label class="flex items-center gap-2"><input type="checkbox" class="check" bind:checked={prefs.export.nested} onchange={savePrefs} />{t().export.nested}</label>
			<label class="flex items-center gap-2"><input type="checkbox" class="check" bind:checked={prefs.export.dropNulls} onchange={savePrefs} />{t().export.dropNulls}</label>
			<label class="flex items-center justify-between gap-2">
				<span>{t().export.indent}</span>
				<select class="select" bind:value={prefs.export.indent} onchange={savePrefs}>
					<option value={0}>{t().export.indentCompact}</option>
					<option value={2}>{t().export.indentSpaces(2)}</option>
					<option value={4}>{t().export.indentSpaces(4)}</option>
				</select>
			</label>
		{:else if e.format === 'csv'}
			<label class="flex flex-col gap-1">
				<span>{t().export.delimiter}</span>
				<select class="select w-full" bind:value={prefs.export.csvDelimiter} onchange={savePrefs}>
					{#each DELIMS as d (d.v)}<option value={d.v}>{d.label()}</option>{/each}
				</select>
			</label>
			<label class="flex items-center gap-2"><input type="checkbox" class="check" bind:checked={prefs.export.bom} onchange={savePrefs} />{t().export.bom}</label>
		{:else}
			<label class="flex items-start gap-2"><input type="checkbox" class="check mt-0.5" bind:checked={prefs.export.styled} onchange={savePrefs} />{t().export.styled}</label>
		{/if}
		<label class="flex items-start gap-2" title={t().export.allTextHelp}>
			<input type="checkbox" class="check mt-0.5" bind:checked={prefs.export.allText} onchange={savePrefs} />
			<span>{t().export.allText}<span class="block text-xs text-muted">{t().export.allTextHelp}</span></span>
		</label>
		{#if mode === 'merged'}
			<label class="flex items-center gap-2"><input type="checkbox" class="check" bind:checked={prefs.export.origin} onchange={savePrefs} />{t().export.origin}</label>
		{/if}
		{#if anyFilter}
			<label class="flex items-center gap-2"><input type="checkbox" class="check" bind:checked={prefs.export.onlyVisible} onchange={savePrefs} />{t().export.onlyVisible}</label>
		{/if}
	</fieldset>

	<div class="flex flex-col gap-2 border-t border-line pt-3">
		<p class="text-xs text-muted" aria-live="polite">{tables.length ? t().export.selection(tables.length, rowCount) : t().errors.empty}</p>
		{#if running}
			<div class="flex items-center gap-2">
				<div class="h-1.5 flex-1 overflow-hidden rounded bg-surface-2" role="progressbar" aria-valuemin={0} aria-valuemax={progress?.total ?? 1} aria-valuenow={progress?.done ?? 0}>
					<div class="h-full bg-accent transition-[width]" style="width:{progress && progress.total ? (100 * progress.done) / progress.total : 15}%"></div>
				</div>
				<button class="btn py-0.5 text-xs" onclick={cancelAll}>{t().export.cancel}</button>
			</div>
			<p class="text-xs text-muted" aria-live="polite">{progress ? t().export.building(progress.done, progress.total) : t().common.loading}</p>
		{/if}
		<div class="flex gap-2">
			<button class="btn btn-primary flex-1 py-2" disabled={!tables.length || !!running || converter.busy > 0} onclick={() => run('download')}>
				<Icon name="download" size={16} />{t().export.download}
			</button>
			{#if copyable}
				<button class="btn py-2" title={t().export.copyHelp} disabled={!tables.length || !!running || converter.busy > 0} onclick={() => run('copy')}>
					<Icon name="copy" size={16} />{t().export.copy}
				</button>
			{/if}
		</div>
		{#if last && !running}<p class="text-xs text-faint">{last}</p>{/if}
	</div>
</section>
