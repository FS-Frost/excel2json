<script lang="ts">
	// Buscar y reemplazar en la tabla activa (sobre las filas visibles). Enter = siguiente, Mayús+Enter = anterior.
	import { tick } from 'svelte';
	import { DEFAULT_FIND, findAll, matcher, replaceValue, type FindOptions, type Match } from '../../edit/find.ts';
	import { visibleRows } from '../../edit/view.ts';
	import type { CellChange } from '../../edit/ops.ts';
	import { session } from '../../client/session.svelte.ts';
	import { toast } from '../../client/toast.svelte.ts';
	import { t } from '../../i18n/index.svelte.ts';
	import Icon from '../Icon.svelte';

	interface Props {
		tableId: string;
		replace: boolean;
		onreveal: (r: number, c: number) => void;
		onmatches: (keys: Set<string> | null) => void;
		onclose: () => void;
	}

	let { tableId, replace = $bindable(), onreveal, onmatches, onclose }: Props = $props();

	let query = $state('');
	let replacement = $state('');
	let opts = $state<FindOptions>({ ...DEFAULT_FIND });
	let index = $state(0);
	let input = $state<HTMLInputElement>();

	const table = $derived.by(() => {
		void session.rev;
		return session.table(tableId);
	});
	const matches = $derived.by((): Match[] => {
		void session.rev;
		const tb = table;
		if (!tb || !query) return [];
		// La columna elegida puede no existir en otra tabla (o tras borrar columnas): entonces, todas.
		const column = opts.column !== null && opts.column < tb.headers.length ? opts.column : null;
		return findAll(tb, visibleRows(tb, session.views[tableId]?.filters ?? {}), query, { ...opts, column });
	});
	const invalid = $derived(!!query && opts.regex && !matcher(query, opts));

	$effect(() => {
		onmatches(matches.length ? new Set(matches.map((m) => `${m.r}:${m.c}`)) : null);
	});

	$effect(() => {
		return () => onmatches(null);
	});

	export function focusInput() {
		tick().then(() => {
			input?.focus();
			input?.select();
		});
	}

	$effect(() => {
		focusInput();
	});

	function go(delta: number) {
		if (!matches.length) return;
		index = (index + delta + matches.length) % matches.length;
		const m = matches[index];
		onreveal(m.r, m.c);
		tick().then(() => input?.focus());
	}

	function onQuery() {
		index = 0;
		const m = matches[0];
		if (m) onreveal(m.r, m.c);
		tick().then(() => input?.focus());
	}

	function replaceOne() {
		const tb = table;
		const m = matches[Math.min(index, matches.length - 1)];
		if (!tb || !m) return;
		const v = replaceValue(tb.rows[m.r][m.c], query, replacement, opts);
		if (v !== tb.rows[m.r][m.c]) session.exec(tableId, [{ t: 'cells', changes: [{ r: m.r, c: m.c, v }] }]);
		tick().then(() => {
			if (matches.length) {
				index = index % matches.length;
				const next = matches[index];
				onreveal(next.r, next.c);
			}
		});
	}

	function replaceAll() {
		const tb = table;
		if (!tb || !matches.length) return;
		const changes: CellChange[] = [];
		for (const m of matches) {
			const v = replaceValue(tb.rows[m.r][m.c], query, replacement, opts);
			if (v !== tb.rows[m.r][m.c]) changes.push({ r: m.r, c: m.c, v });
		}
		if (changes.length) session.exec(tableId, [{ t: 'cells', changes }]);
		toast.success(t().find.replaced(changes.length), { key: 'replace' });
	}

	function keydown(e: KeyboardEvent) {
		if (e.key === 'Escape') {
			e.preventDefault();
			onclose();
		} else if (e.key === 'Enter') {
			e.preventDefault();
			go(e.shiftKey ? -1 : 1);
		}
	}
</script>

<div class="flex flex-wrap items-center gap-x-2 gap-y-1.5 border-b border-line bg-surface px-3 py-2 text-sm" role="search">
	<div class="flex min-w-0 flex-1 items-center gap-1.5 sm:max-w-80">
		<Icon name="search" size={16} class="text-faint" />
		<input
			bind:this={input}
			class="input min-w-0 flex-1 py-1 {invalid ? 'border-bad' : ''}"
			placeholder={t().find.find}
			aria-label={t().find.find}
			bind:value={query}
			oninput={onQuery}
			onkeydown={keydown}
		/>
	</div>
	<span class="min-w-20 text-xs text-muted tabular-nums" aria-live="polite">{invalid ? t().find.invalidRegex : query ? t().find.results(matches.length ? index + 1 : 0, matches.length) : ''}</span>
	<button class="tb" title={t().find.prev} aria-label={t().find.prev} disabled={!matches.length} onclick={() => go(-1)}><Icon name="chevronUp" /></button>
	<button class="tb" title={t().find.next} aria-label={t().find.next} disabled={!matches.length} onclick={() => go(1)}><Icon name="chevronDown" /></button>
	<button class="tb" aria-pressed={replace} title={t().find.replaceOne} aria-label={t().find.replaceOne} onclick={() => (replace = !replace)}><Icon name="replace" /></button>
	<details class="relative">
		<summary class="tb cursor-pointer list-none" title={t().common.more} aria-label={t().common.more}><Icon name="settings" size={16} /></summary>
		<div class="menu absolute right-0 mt-1 w-56 p-2">
			<label class="flex items-center gap-2 py-1"><input type="checkbox" class="check" bind:checked={opts.matchCase} onchange={onQuery} />{t().find.matchCase}</label>
			<label class="flex items-center gap-2 py-1"><input type="checkbox" class="check" bind:checked={opts.wholeCell} onchange={onQuery} />{t().find.wholeCell}</label>
			<label class="flex items-center gap-2 py-1"><input type="checkbox" class="check" bind:checked={opts.regex} onchange={onQuery} />{t().find.regex}</label>
			<select class="select mt-1 w-full" bind:value={opts.column} onchange={onQuery} aria-label={t().grid.colType}>
				<option value={null}>{t().find.allColumns}</option>
				{#each table?.headers ?? [] as h, c (c)}
					<option value={c}>{h}</option>
				{/each}
			</select>
		</div>
	</details>
	<button class="tb ml-auto" title={t().common.close} aria-label={t().common.close} onclick={onclose}><Icon name="x" /></button>
	{#if replace}
		<div class="flex w-full flex-wrap items-center gap-2">
			<div class="flex min-w-0 flex-1 items-center gap-1.5 sm:max-w-80">
				<Icon name="replace" size={16} class="text-faint" />
				<input
					class="input min-w-0 flex-1 py-1"
					placeholder={t().find.replace}
					aria-label={t().find.replace}
					bind:value={replacement}
					onkeydown={(e) => {
						if (e.key === 'Enter') {
							e.preventDefault();
							replaceOne();
						} else if (e.key === 'Escape') onclose();
					}}
				/>
			</div>
			<button class="btn py-1" disabled={!matches.length} onclick={replaceOne}>{t().find.replaceOne}</button>
			<button class="btn py-1" disabled={!matches.length} onclick={replaceAll}>{t().find.replaceAll}</button>
		</div>
	{/if}
</div>
