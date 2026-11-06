<script lang="ts" module>
	export type GridSel = { r1: number; c1: number; r2: number; c2: number };
</script>

<script lang="ts">
	// Grid virtualizado y editable (diseño de FS-Frost/mai-docs, sin fórmulas ni estilos). Un contenedor con scroll
	// nativo da las barras; encima, un viewport sticky dibuja la cabecera, los números de fila y el cuerpo, cada panel
	// trasladado según el scroll. Solo se renderizan las filas/columnas visibles (en bloques, para no rehacer el DOM en
	// cada píxel). Todo el teclado pasa por un textarea («sink») siempre enfocado: escribir empieza a editar (compatible
	// con IME), las flechas navegan y copiar/cortar/pegar llegan como eventos del portapapeles.
	// Coordenadas: `vi` = índice en las filas visibles (con filtro); la fila real es `vis[vi]`.
	import { tick, untrack } from 'svelte';
	import { cellText, fits, parseTyped } from '../../core/cell.ts';
	import { COL_TYPES, type Cell, type ColType } from '../../core/types.ts';
	import { parseHTMLTable, parseTSV, toHTML, toTSV } from '../../edit/clipboard.ts';
	import { Axis } from '../../edit/layout.ts';
	import { freeHeader, moveOrder, type CellChange, type Op } from '../../edit/ops.ts';
	import { distinctValues, isActive, sortOrder, visibleRows, type ColumnFilter } from '../../edit/view.ts';
	import { session } from '../../client/session.svelte.ts';
	import { promptDialog } from '../../client/dialogs.svelte.ts';
	import { toast } from '../../client/toast.svelte.ts';
	import { decimalSep, t } from '../../i18n/index.svelte.ts';
	import ContextMenu, { type MenuAt } from '../ContextMenu.svelte';
	import Icon from '../Icon.svelte';

	interface Props {
		tableId: string;
		/** Celdas que coinciden con la búsqueda: claves `fila:columna` (fila real). */
		highlight?: Set<string> | null;
		onfind?: (replace: boolean) => void;
		onselect?: (sel: GridSel) => void;
		/** Deshacer/rehacer: la página lo resuelve (muestra la tabla afectada, que puede ser otra). */
		onundo?: (redo: boolean) => void;
	}

	let { tableId, highlight = null, onfind, onselect, onundo }: Props = $props();

	function undoRedo(redo: boolean) {
		if (onundo) onundo(redo);
		else if (redo) session.redo();
		else session.undo();
	}

	const ROW_H = 28;
	const HEAD_H = 34;
	const DEF_W = 140;
	const MIN_W = 48;
	const MAX_W = 900;
	const RB = 8;
	const CB = 4;

	// ---------- Datos ----------

	// Copia superficial por revisión: la tabla se muta en el lugar; una referencia nueva actualiza a quien la lee.
	const table = $derived.by(() => {
		void session.rev;
		const tb = session.table(tableId);
		return tb ? { ...tb } : null;
	});
	const filters = $derived(session.views[tableId]?.filters ?? {});
	const widths = $derived(session.views[tableId]?.widths ?? {});
	const vis = $derived(table ? visibleRows(table, filters) : []);
	const nRows = $derived(vis.length);
	const nCols = $derived(table?.headers.length ?? 0);
	const filtered = $derived(Object.values(filters).some(isActive));
	const dec = $derived(decimalSep());

	let live = $state<{ c: number; w: number } | null>(null);
	const cols = $derived.by(() => {
		const sizes = new Map<number, number>();
		table?.headers.forEach((h, c) => {
			const w = widths[h];
			if (w) sizes.set(c, w);
		});
		if (live) sizes.set(live.c, live.w);
		return new Axis(Math.max(nCols, 0), DEF_W, sizes);
	});
	const RN = $derived(Math.max(48, 18 + String(table?.rows.length ?? 0).length * 8));

	// ---------- Viewport ----------

	let scroller = $state<HTMLDivElement>();
	let sink = $state<HTMLTextAreaElement>();
	let vw = $state(800);
	let vh = $state(500);
	let sx = $state(0);
	let sy = $state(0);

	$effect(() => {
		if (!scroller) return;
		const el = scroller;
		const ro = new ResizeObserver(() => {
			vw = el.clientWidth;
			vh = el.clientHeight;
		});
		ro.observe(el);
		return () => ro.disconnect();
	});

	function onscroll() {
		if (!scroller) return;
		sx = scroller.scrollLeft;
		sy = scroller.scrollTop;
	}

	const r0 = $derived(Math.max(0, Math.floor(Math.floor(sy / ROW_H) / RB) * RB));
	const r1 = $derived(Math.min(nRows - 1, (Math.floor((sy + vh) / ROW_H / RB) + 1) * RB));
	const c0 = $derived(nCols ? Math.floor(cols.indexAt(sx) / CB) * CB : 0);
	const c1 = $derived(nCols ? Math.min(nCols - 1, (Math.floor(cols.indexAt(sx + vw - RN) / CB) + 1) * CB) : -1);
	const rowIdx = $derived(Array.from({ length: Math.max(0, r1 - r0 + 1) }, (_, i) => r0 + i));
	const colIdx = $derived(Array.from({ length: Math.max(0, c1 - c0 + 1) }, (_, i) => c0 + i));
	const totalW = $derived(RN + cols.total);
	const totalH = $derived(HEAD_H + nRows * ROW_H);

	// ---------- Selección ----------

	let anchor = $state({ r: 0, c: 0 });
	let focus = $state({ r: 0, c: 0 });
	/** Selección de filas o columnas completas (desde sus cabeceras). */
	let whole = $state<'rows' | 'cols' | 'all' | null>(null);

	const sel = $derived<GridSel>({
		r1: whole === 'cols' || whole === 'all' ? 0 : Math.min(anchor.r, focus.r),
		r2: whole === 'cols' || whole === 'all' ? Math.max(0, nRows - 1) : Math.max(anchor.r, focus.r),
		c1: whole === 'rows' || whole === 'all' ? 0 : Math.min(anchor.c, focus.c),
		c2: whole === 'rows' || whole === 'all' ? Math.max(0, nCols - 1) : Math.max(anchor.c, focus.c)
	});

	$effect(() => {
		onselect?.(sel);
	});

	// Al cambiar de tabla: selección al inicio y scroll arriba.
	let lastTable = '';
	$effect(() => {
		const id = tableId;
		untrack(() => {
			if (id === lastTable) return;
			lastTable = id;
			anchor = { r: 0, c: 0 };
			focus = { r: 0, c: 0 };
			whole = null;
			editing = null;
			if (scroller) scroller.scrollTo(0, 0);
		});
	});

	// Si la tabla se achica (filtro, borrar filas/columnas), la selección vuelve a un lugar válido.
	$effect(() => {
		const maxR = Math.max(0, nRows - 1);
		const maxC = Math.max(0, nCols - 1);
		untrack(() => {
			const clamp = (p: { r: number; c: number }) => ({ r: Math.min(p.r, maxR), c: Math.min(p.c, maxC) });
			if (anchor.r > maxR || anchor.c > maxC) anchor = clamp(anchor);
			if (focus.r > maxR || focus.c > maxC) focus = clamp(focus);
		});
	});

	function select(r: number, c: number, extend = false) {
		const p = { r: Math.max(0, Math.min(nRows - 1, r)), c: Math.max(0, Math.min(nCols - 1, c)) };
		whole = null;
		focus = p;
		if (!extend) anchor = p;
		ensureVisible(p.r, p.c);
	}

	function ensureVisible(r: number, c: number) {
		if (!scroller) return;
		const top = r * ROW_H;
		const bodyH = vh - HEAD_H;
		if (top < scroller.scrollTop) scroller.scrollTop = top;
		else if (top + ROW_H > scroller.scrollTop + bodyH) scroller.scrollTop = top + ROW_H - bodyH;
		if (c < 0 || c >= nCols) return;
		const left = cols.offset(c);
		const w = cols.size(c);
		const bodyW = vw - RN;
		if (left < scroller.scrollLeft) scroller.scrollLeft = left;
		else if (left + w > scroller.scrollLeft + bodyW) scroller.scrollLeft = Math.min(left, left + w - bodyW);
	}

	/** Muestra y selecciona una celda por fila real (buscar). */
	export function reveal(realRow: number, c: number) {
		const vi = vis.indexOf(realRow);
		if (vi < 0) return;
		select(vi, c);
		focusSink();
	}

	export function focusSink() {
		sink?.focus({ preventScroll: true });
	}

	// ---------- Edición ----------

	/** `quick`: empezó escribiendo (las flechas confirman y mueven, como en Excel). */
	let editing = $state<{ r: number; c: number; quick: boolean } | null>(null);
	let draft = $state('');

	const active = $derived(focus);
	const activeCell = $derived(table && nRows ? (table.rows[vis[active.r]]?.[active.c] ?? null) : null);

	function startEdit(quick: boolean, initial?: string) {
		if (!table || !nRows || !nCols) return;
		whole = null;
		anchor = { ...focus };
		editing = { r: focus.r, c: focus.c, quick };
		draft = initial ?? cellText(activeCell, dec);
		tick().then(() => {
			if (!sink) return;
			sink.focus({ preventScroll: true });
			if (initial === undefined) sink.setSelectionRange(draft.length, draft.length);
		});
	}

	function commit() {
		const e = editing;
		if (!e || !table) return;
		editing = null;
		const real = vis[e.r];
		const tb = session.table(tableId);
		if (!tb || real === undefined) return;
		const v = parseTyped(draft, tb.types[e.c] ?? 'auto', dec);
		const old = tb.rows[real]?.[e.c] ?? null;
		draft = '';
		if (v === old) return;
		session.exec(tableId, [{ t: 'cells', changes: [{ r: real, c: e.c, v }] }]);
	}

	function cancelEdit() {
		editing = null;
		draft = '';
	}

	// ---------- Operaciones ----------

	function exec(ops: Op[]) {
		session.exec(tableId, ops);
	}

	function selectedRealRows(): number[] {
		const out: number[] = [];
		for (let vi = sel.r1; vi <= sel.r2; vi++) if (vis[vi] !== undefined) out.push(vis[vi]);
		return out;
	}

	export function clearSelection() {
		const tb = session.table(tableId);
		if (!tb) return;
		const changes: CellChange[] = [];
		for (const r of selectedRealRows()) for (let c = sel.c1; c <= sel.c2; c++) if (tb.rows[r]?.[c] !== null) changes.push({ r, c, v: null });
		if (changes.length) exec([{ t: 'cells', changes }]);
	}

	export function insertRows(where: 'above' | 'below') {
		const tb = session.table(tableId);
		if (!tb) return;
		// Con columnas completas seleccionadas (cabecera de columna, Ctrl+A) se inserta una sola fila, no una por fila.
		const n = !nRows || whole === 'cols' || whole === 'all' ? 1 : sel.r2 - sel.r1 + 1;
		const at = !nRows ? tb.rows.length : where === 'above' ? vis[sel.r1] : vis[sel.r2] + 1;
		const c = sel.c1;
		exec([{ t: 'insertRows', at, rows: Array.from({ length: Math.min(n, 1000) }, () => []) }]);
		showNewRow(at, c);
	}

	/** Selecciona una fila recién insertada. Con un filtro activo quedaría oculta (y la selección caería sobre otra
	 *  fila con datos, que se sobrescribiría al escribir): se quitan los filtros para mostrarla. */
	function showNewRow(real: number, c: number) {
		if (filtered) clearFilters();
		whole = null;
		tick().then(() => {
			const vi = vis.indexOf(real);
			if (vi >= 0) select(vi, Math.max(0, c));
			focusSink();
		});
	}

	export function addRow() {
		const tb = session.table(tableId);
		if (!tb) return;
		const at = tb.rows.length;
		exec([{ t: 'insertRows', at, rows: [[]] }]);
		showNewRow(at, sel.c1);
	}

	export function insertCols(where: 'left' | 'right') {
		const tb = session.table(tableId);
		if (!tb) return;
		const n = !nCols || whole === 'rows' || whole === 'all' ? 1 : sel.c2 - sel.c1 + 1;
		const at = !nCols ? 0 : where === 'left' ? sel.c1 : sel.c2 + 1;
		const headers: string[] = [];
		for (let k = 0; k < n; k++) headers.push(freeHeader([...tb.headers, ...headers], t().grid.newColumn));
		exec([{ t: 'insertCols', at, headers, types: headers.map((): ColType => 'auto') }]);
		tick().then(() => select(sel.r1, at));
	}

	export function addCol() {
		const tb = session.table(tableId);
		if (!tb) return;
		const name = freeHeader(tb.headers, t().grid.newColumn);
		exec([{ t: 'insertCols', at: tb.headers.length, headers: [name], types: ['auto'] }]);
		tick().then(() => {
			select(sel.r1, nCols - 1);
			ensureVisible(sel.r1, nCols - 1);
		});
	}

	export function deleteRows() {
		const rows = selectedRealRows();
		if (rows.length) exec([{ t: 'deleteRows', at: rows }]);
		whole = null;
		anchor = { r: sel.r1, c: sel.c1 };
		focus = { r: sel.r1, c: sel.c1 };
	}

	export function deleteCols() {
		const at = Array.from({ length: sel.c2 - sel.c1 + 1 }, (_, k) => sel.c1 + k);
		if (at.length) exec([{ t: 'deleteCols', at }]);
		whole = null;
		anchor = { r: sel.r1, c: Math.min(sel.c1, Math.max(0, nCols - 1)) };
		focus = { ...anchor };
	}

	export const canMoveRows = () => !filtered && nRows > 0;

	export function moveRows(dir: -1 | 1) {
		if (!canMoveRows()) return;
		const count = sel.r2 - sel.r1 + 1;
		const to = sel.r1 + dir;
		if (to < 0 || sel.r2 + dir >= nRows) return;
		exec([{ t: 'permuteRows', order: moveOrder(nRows, sel.r1, count, to) }]);
		anchor = { r: anchor.r + dir, c: anchor.c };
		focus = { r: focus.r + dir, c: focus.c };
		ensureVisible(focus.r, focus.c);
	}

	export function moveCols(dir: -1 | 1) {
		const count = sel.c2 - sel.c1 + 1;
		const to = sel.c1 + dir;
		if (to < 0 || sel.c2 + dir >= nCols) return;
		exec([{ t: 'permuteCols', order: moveOrder(nCols, sel.c1, count, to) }]);
		anchor = { r: anchor.r, c: anchor.c + dir };
		focus = { r: focus.r, c: focus.c + dir };
		ensureVisible(focus.r, focus.c);
	}

	export function sortBy(c: number, dir: 'asc' | 'desc') {
		const tb = session.table(tableId);
		if (!tb || !tb.rows.length) return;
		exec([{ t: 'permuteRows', order: sortOrder(tb, c, dir) }]);
		toast.success(t().grid.sorted, { key: 'sort', duration: 2000 });
	}

	export async function renameCol(c: number) {
		const tb = session.table(tableId);
		if (!tb) return;
		const name = await promptDialog(t().grid.renamePrompt, tb.headers[c], {
			confirm: t().common.rename,
			validate: (v) => (!v.trim() ? t().grid.renameEmpty : tb.headers.some((h, i) => i !== c && h === v.trim()) ? t().grid.renameDup : null)
		});
		focusSink();
		if (name === null || name.trim() === tb.headers[c]) return;
		const old = tb.headers[c];
		const w = session.view(tableId).widths[old];
		exec([{ t: 'header', c, name: name.trim() }]);
		if (w) {
			session.setWidth(tableId, old, null);
			session.setWidth(tableId, name.trim(), w);
		}
	}

	export function setType(c: number, type: ColType) {
		const tb = session.table(tableId);
		if (!tb || tb.types[c] === type) return;
		exec([{ t: 'type', c, type }]);
	}

	export function selectAll() {
		whole = 'all';
		anchor = { r: 0, c: 0 };
		focus = { r: 0, c: 0 };
	}

	function setFilter(c: number, f: ColumnFilter | null) {
		const next = { ...filters };
		if (f && isActive(f)) next[c] = f;
		else delete next[c];
		session.setFilters(tableId, next);
	}

	export function clearFilters() {
		session.setFilters(tableId, {});
	}

	// ---------- Portapapeles ----------

	function selectionGrid(): Cell[][] {
		const tb = session.table(tableId);
		if (!tb) return [];
		return selectedRealRows().map((r) => tb.rows[r].slice(sel.c1, sel.c2 + 1));
	}

	function writeClipboard(e: ClipboardEvent) {
		const grid = selectionGrid();
		const textRows = grid.map((r) => r.map((v) => cellText(v, dec)));
		e.clipboardData?.setData('text/plain', toTSV(textRows));
		e.clipboardData?.setData('text/html', toHTML(textRows.map((r) => r.map((text) => ({ text })))));
		e.preventDefault();
	}

	function oncopy(e: ClipboardEvent) {
		if (editing) return;
		writeClipboard(e);
	}

	function oncut(e: ClipboardEvent) {
		if (editing) return;
		writeClipboard(e);
		clearSelection();
	}

	function onpaste(e: ClipboardEvent) {
		if (editing) return;
		e.preventDefault();
		const html = e.clipboardData?.getData('text/html') ?? '';
		const text = e.clipboardData?.getData('text/plain') ?? '';
		const grid = (html && parseHTMLTable(html)) || parseTSV(text);
		pasteGrid(grid);
	}

	export async function pasteFromClipboard() {
		try {
			pasteGrid(parseTSV(await navigator.clipboard.readText()));
		} catch {
			toast.warn(t().grid.pasteDenied);
		}
		focusSink();
	}

	export async function copySelection(cut = false) {
		const textRows = selectionGrid().map((r) => r.map((v) => cellText(v, dec)));
		try {
			await navigator.clipboard.writeText(toTSV(textRows));
			if (cut) clearSelection();
		} catch {
			/* sin permiso: Ctrl+C sigue funcionando */
		}
		focusSink();
	}

	/** Pega desde la esquina superior izquierda de la selección. Un solo valor llena toda la selección. Lo que no cabe
	 *  agrega filas al final y columnas a la derecha. */
	function pasteGrid(grid: string[][]) {
		const tb = session.table(tableId);
		if (!tb || !grid.length) return;
		const ops: Op[] = [];
		const fill = grid.length === 1 && grid[0].length === 1 && (sel.r2 > sel.r1 || sel.c2 > sel.c1);
		const h = fill ? sel.r2 - sel.r1 + 1 : grid.length;
		const w = fill ? sel.c2 - sel.c1 + 1 : Math.max(...grid.map((r) => r.length));
		const startC = nCols ? sel.c1 : 0;
		const extraCols = Math.max(0, startC + w - tb.headers.length);
		if (extraCols) {
			const headers: string[] = [];
			for (let k = 0; k < extraCols; k++) headers.push(freeHeader([...tb.headers, ...headers], t().grid.newColumn));
			ops.push({ t: 'insertCols', at: tb.headers.length, headers, types: headers.map((): ColType => 'auto') });
		}
		// Filas destino: las visibles desde la selección; las que falten se agregan al final de la tabla.
		const startV = nRows ? sel.r1 : 0;
		const targets: number[] = [];
		for (let k = 0; k < h; k++) {
			const real = vis[startV + k];
			if (real !== undefined) targets.push(real);
		}
		const missing = h - targets.length;
		if (missing > 0) {
			const at = tb.rows.length;
			ops.push({ t: 'insertRows', at, rows: Array.from({ length: missing }, () => []) });
			for (let k = 0; k < missing; k++) targets.push(at + k);
		}
		const changes: CellChange[] = [];
		const types = [...tb.types, ...Array(extraCols).fill('auto')];
		for (let k = 0; k < h; k++)
			for (let j = 0; j < w; j++) {
				const raw = fill ? grid[0][0] : (grid[k]?.[j] ?? '');
				changes.push({ r: targets[k], c: startC + j, v: parseTyped(raw, types[startC + j] ?? 'auto', dec) });
			}
		ops.push({ t: 'cells', changes });
		exec(ops);
		whole = null;
		anchor = { r: startV, c: startC };
		focus = { r: Math.min(startV + h - 1, startV + targets.length - 1), c: startC + w - 1 };
	}

	// ---------- Teclado ----------

	/** Última fila/columna con datos en una dirección (Ctrl+flecha). */
	function edge(r: number, c: number, dr: number, dc: number): { r: number; c: number } {
		const tb = session.table(tableId);
		if (!tb) return { r, c };
		const val = (vr: number, vc: number) => tb.rows[vis[vr]]?.[vc] ?? null;
		const inside = (vr: number, vc: number) => vr >= 0 && vr < nRows && vc >= 0 && vc < nCols;
		let nr = r + dr;
		let nc = c + dc;
		if (!inside(nr, nc)) return { r, c };
		const startFilled = val(r, c) !== null && val(nr, nc) !== null;
		if (startFilled) {
			while (inside(nr + dr, nc + dc) && val(nr + dr, nc + dc) !== null) {
				nr += dr;
				nc += dc;
			}
			return { r: nr, c: nc };
		}
		while (inside(nr, nc) && val(nr, nc) === null && inside(nr + dr, nc + dc)) {
			nr += dr;
			nc += dc;
		}
		return { r: nr, c: nc };
	}

	function onkeydown(e: KeyboardEvent) {
		const mod = e.ctrlKey || e.metaKey;
		if (editing) {
			if (e.key === 'Enter' && !e.altKey && !e.isComposing) {
				e.preventDefault();
				commit();
				select(focus.r + (e.shiftKey ? -1 : 1), focus.c);
			} else if (e.key === 'Tab') {
				e.preventDefault();
				commit();
				select(focus.r, focus.c + (e.shiftKey ? -1 : 1));
			} else if (e.key === 'Escape') {
				e.preventDefault();
				cancelEdit();
			} else if (e.key === 'Enter' && e.altKey) {
				e.preventDefault();
				const ta = e.currentTarget as HTMLTextAreaElement;
				ta.setRangeText('\n', ta.selectionStart, ta.selectionEnd, 'end');
				draft = ta.value;
			} else if (editing.quick && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && !e.isComposing) {
				e.preventDefault();
				commit();
				const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]!;
				select(focus.r + d[0], focus.c + d[1]);
			}
			return;
		}
		if (mod && e.key.toLowerCase() === 'z') {
			e.preventDefault();
			undoRedo(e.shiftKey);
			return;
		}
		if (mod && e.key.toLowerCase() === 'y') {
			e.preventDefault();
			undoRedo(true);
			return;
		}
		if (mod && (e.key.toLowerCase() === 'f' || e.key.toLowerCase() === 'h')) {
			e.preventDefault();
			onfind?.(e.key.toLowerCase() === 'h');
			return;
		}
		if (mod && e.key.toLowerCase() === 'a') {
			e.preventDefault();
			selectAll();
			return;
		}
		if (mod && e.key === ' ') {
			e.preventDefault();
			whole = 'cols';
			return;
		}
		if (e.shiftKey && e.key === ' ') {
			e.preventDefault();
			whole = 'rows';
			return;
		}
		if (e.altKey && e.key.startsWith('Arrow')) {
			e.preventDefault();
			if (e.key === 'ArrowUp') moveRows(-1);
			else if (e.key === 'ArrowDown') moveRows(1);
			else if (e.key === 'ArrowLeft') moveCols(-1);
			else moveCols(1);
			return;
		}
		const page = Math.max(1, Math.floor((vh - HEAD_H) / ROW_H) - 1);
		let move: { r: number; c: number } | null = null;
		switch (e.key) {
			case 'ArrowUp':
				move = mod ? edge(focus.r, focus.c, -1, 0) : { r: focus.r - 1, c: focus.c };
				break;
			case 'ArrowDown':
				move = mod ? edge(focus.r, focus.c, 1, 0) : { r: focus.r + 1, c: focus.c };
				break;
			case 'ArrowLeft':
				move = mod ? edge(focus.r, focus.c, 0, -1) : { r: focus.r, c: focus.c - 1 };
				break;
			case 'ArrowRight':
				move = mod ? edge(focus.r, focus.c, 0, 1) : { r: focus.r, c: focus.c + 1 };
				break;
			case 'PageUp':
				move = { r: focus.r - page, c: focus.c };
				break;
			case 'PageDown':
				move = { r: focus.r + page, c: focus.c };
				break;
			case 'Home':
				move = mod ? { r: 0, c: 0 } : { r: focus.r, c: 0 };
				break;
			case 'End':
				move = mod ? { r: nRows - 1, c: nCols - 1 } : { r: focus.r, c: nCols - 1 };
				break;
			case 'Enter':
				e.preventDefault();
				select(focus.r + (e.shiftKey ? -1 : 1), focus.c);
				return;
			case 'Tab':
				e.preventDefault();
				select(focus.r, focus.c + (e.shiftKey ? -1 : 1));
				return;
			case 'F2':
				e.preventDefault();
				startEdit(false);
				return;
			case 'Delete':
				e.preventDefault();
				clearSelection();
				return;
			case 'Backspace':
				e.preventDefault();
				clearSelection();
				startEdit(true, '');
				return;
			case 'Escape':
				whole = null;
				anchor = { ...focus };
				return;
			case 'ContextMenu':
				e.preventDefault();
				openCellMenuAtActive();
				return;
		}
		if (e.shiftKey && e.key === 'F10') {
			e.preventDefault();
			openCellMenuAtActive();
			return;
		}
		if (move) {
			e.preventDefault();
			select(move.r, move.c, e.shiftKey);
		}
	}

	/** Escribir sin estar editando: el carácter ya llegó al textarea (también con IME) → empieza la edición con él. */
	function oninput(e: Event) {
		const ta = e.currentTarget as HTMLTextAreaElement;
		if (editing) {
			draft = ta.value;
			return;
		}
		if (!ta.value || !nRows || !nCols) {
			ta.value = '';
			return;
		}
		editing = { r: focus.r, c: focus.c, quick: true };
		whole = null;
		anchor = { ...focus };
		draft = ta.value;
	}

	// ---------- Ratón ----------

	let body = $state<HTMLDivElement>();
	let drag: 'cells' | 'rows' | 'cols' | null = null;

	function cellAt(clientX: number, clientY: number) {
		const rect = body!.getBoundingClientRect();
		const y = clientY - rect.top + sy;
		const x = clientX - rect.left + sx;
		return { r: Math.max(0, Math.min(nRows - 1, Math.floor(y / ROW_H))), c: Math.max(0, Math.min(nCols - 1, cols.indexAt(x))) };
	}

	function autoScroll(e: PointerEvent) {
		if (!scroller || !body) return;
		const rect = body.getBoundingClientRect();
		if (e.clientY > rect.bottom - 8) scroller.scrollTop += Math.min(40, e.clientY - rect.bottom + 8);
		else if (e.clientY < rect.top + 8 && drag !== 'cols') scroller.scrollTop -= Math.min(40, rect.top + 8 - e.clientY);
		if (e.clientX > rect.right - 8) scroller.scrollLeft += Math.min(40, e.clientX - rect.right + 8);
		else if (e.clientX < rect.left + 8 && drag !== 'rows') scroller.scrollLeft -= Math.min(40, rect.left + 8 - e.clientX);
	}

	function bodyDown(e: PointerEvent) {
		if (e.button === 2) {
			// Clic derecho dentro de la selección: la conserva.
			const p = cellAt(e.clientX, e.clientY);
			if (!(p.r >= sel.r1 && p.r <= sel.r2 && p.c >= sel.c1 && p.c <= sel.c2)) {
				commit();
				select(p.r, p.c);
			}
			return;
		}
		if (e.button !== 0 || !nRows || !nCols) return;
		if (editing) commit();
		const p = cellAt(e.clientX, e.clientY);
		whole = null;
		focus = p;
		if (!e.shiftKey) anchor = p;
		drag = 'cells';
		(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
		e.preventDefault();
		focusSink();
	}

	function bodyMove(e: PointerEvent) {
		if (drag !== 'cells') return;
		autoScroll(e);
		focus = cellAt(e.clientX, e.clientY);
	}

	function rowHeadDown(e: PointerEvent) {
		if (e.button !== 0 || !nRows) return;
		if (editing) commit();
		const p = cellAt(e.clientX + 1e5, e.clientY);
		if (!e.shiftKey) anchor = { r: p.r, c: 0 };
		focus = { r: p.r, c: nCols - 1 };
		whole = 'rows';
		drag = 'rows';
		(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
		e.preventDefault();
		focusSink();
	}

	function rowHeadMove(e: PointerEvent) {
		if (drag !== 'rows') return;
		autoScroll(e);
		const rect = body!.getBoundingClientRect();
		const r = Math.max(0, Math.min(nRows - 1, Math.floor((e.clientY - rect.top + sy) / ROW_H)));
		focus = { r, c: focus.c };
	}

	function colHeadDown(e: PointerEvent, c: number) {
		if (e.button !== 0) return;
		if ((e.target as HTMLElement).closest('[data-noselect]')) return;
		if (editing) commit();
		if (!e.shiftKey) anchor = { r: 0, c };
		focus = { r: focus.r, c };
		whole = 'cols';
		drag = 'cols';
		(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
		focusSink();
	}

	function colHeadMove(e: PointerEvent) {
		if (drag !== 'cols') return;
		autoScroll(e);
		const rect = body!.getBoundingClientRect();
		focus = { r: focus.r, c: Math.max(0, Math.min(nCols - 1, cols.indexAt(e.clientX - rect.left + sx))) };
	}

	function up() {
		drag = null;
	}

	// ---------- Ancho de columnas ----------

	let resizing: { c: number; x0: number; w0: number } | null = null;

	function resizeDown(e: PointerEvent, c: number) {
		e.stopPropagation();
		e.preventDefault();
		resizing = { c, x0: e.clientX, w0: cols.size(c) };
		live = { c, w: cols.size(c) };
		(e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
	}

	function resizeMove(e: PointerEvent) {
		if (!resizing) return;
		live = { c: resizing.c, w: Math.max(MIN_W, Math.min(MAX_W, resizing.w0 + e.clientX - resizing.x0)) };
	}

	function resizeUp() {
		if (!resizing || !live || !table) return;
		session.setWidth(tableId, table.headers[resizing.c], live.w);
		resizing = null;
		live = null;
	}

	let measureCtx: CanvasRenderingContext2D | null = null;
	function textWidth(s: string, bold = false) {
		measureCtx ??= document.createElement('canvas').getContext('2d');
		if (!measureCtx) return s.length * 8;
		measureCtx.font = `${bold ? '600 ' : ''}13px system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif`;
		return measureCtx.measureText(s).width;
	}

	/** Ajusta el ancho al contenido (muestra de las primeras 500 filas visibles). */
	function autofit(c: number) {
		if (!table) return;
		let w = textWidth(table.headers[c], true) + 56;
		const n = Math.min(nRows, 500);
		for (let k = 0; k < n; k++) w = Math.max(w, textWidth(cellText(table.rows[vis[k]][c] ?? null, dec)) + 20);
		session.setWidth(tableId, table.headers[c], Math.max(MIN_W, Math.min(MAX_W, Math.ceil(w))));
	}

	// ---------- Menús ----------

	let cellMenu = $state<MenuAt | null>(null);
	let colMenu = $state<{ at: MenuAt; c: number } | null>(null);
	let filterQuery = $state('');
	let valueSearch = $state('');

	function openCellMenu(e: MouseEvent) {
		e.preventDefault();
		cellMenu = { x: e.clientX, y: e.clientY };
	}

	function openCellMenuAtActive() {
		if (!body) return;
		const rect = body.getBoundingClientRect();
		cellMenu = { x: rect.left + cols.offset(focus.c) - sx + 8, y: rect.top + (focus.r + 1) * ROW_H - sy };
	}

	function openColMenu(e: MouseEvent, c: number) {
		e.stopPropagation();
		e.preventDefault();
		const r = (e.currentTarget as HTMLElement).getBoundingClientRect();
		if (!(sel.c1 <= c && c <= sel.c2)) {
			anchor = { r: 0, c };
			focus = { r: focus.r, c };
			whole = 'cols';
		}
		filterQuery = filters[c]?.query ?? '';
		valueSearch = '';
		colMenu = { at: { x: Math.max(8, r.right - 260), y: r.bottom + 2 }, c };
	}

	const colValues = $derived(colMenu && table ? distinctValues(table, colMenu.c) : null);
	const shownValues = $derived(
		colValues ? (valueSearch.trim() ? colValues.values.filter((v) => v.toLowerCase().includes(valueSearch.trim().toLowerCase())) : colValues.values) : []
	);

	function toggleValue(c: number, value: string, on: boolean) {
		const all = colValues?.values ?? [];
		const current = new Set(filters[c]?.values ?? all);
		if (on) current.add(value);
		else current.delete(value);
		const values = current.size === all.length && !colValues?.truncated ? null : [...current];
		setFilter(c, { ...filters[c], values });
	}

	function setAllValues(c: number, on: boolean) {
		setFilter(c, { ...filters[c], values: on ? null : [] });
	}

	// ---------- Presentación ----------

	const typeLabel = (ty: ColType) => t().grid.types[ty];

	function display(v: Cell): string {
		return cellText(v, dec);
	}

	const selRect = $derived.by(() => {
		if (!nRows || !nCols) return null;
		const x = cols.offset(sel.c1);
		return { x, y: sel.r1 * ROW_H, w: cols.offset(sel.c2 + 1) - x, h: (sel.r2 - sel.r1 + 1) * ROW_H };
	});

	const editorBox = $derived(editing ? { x: cols.offset(editing.c), y: editing.r * ROW_H, w: cols.size(editing.c) } : { x: cols.offset(focus.c), y: focus.r * ROW_H, w: cols.size(focus.c) });

	const ariaActive = $derived(table && nCols && nRows ? `${table.headers[focus.c]} ${vis[focus.r] + 1}: ${display(activeCell)}` : t().grid.emptyTable);

	const inSelRows = (vi: number) => vi >= sel.r1 && vi <= sel.r2;
	const inSelCols = (c: number) => c >= sel.c1 && c <= sel.c2;
</script>

{#if table}
	<div
		bind:this={scroller}
		class="relative h-full w-full overflow-auto overscroll-contain bg-surface text-[13px] select-none"
		role="grid"
		aria-label={t().grid.label(table.name)}
		aria-rowcount={nRows + 1}
		aria-colcount={nCols}
		{onscroll}
		tabindex="-1"
	>
		<div class="sticky top-0 left-0 overflow-hidden" style="width:{vw}px;height:{vh}px;margin-bottom:-{vh}px">
			<!-- Cuerpo -->
			<div
				bind:this={body}
				class="absolute overflow-hidden"
				style="left:{RN}px;top:{HEAD_H}px;right:0;bottom:0"
				onpointerdown={bodyDown}
				onpointermove={bodyMove}
				onpointerup={up}
				onpointercancel={up}
				ondblclick={() => startEdit(false)}
				oncontextmenu={openCellMenu}
				role="presentation"
			>
				<div class="absolute top-0 left-0" style="transform:translate({-sx}px,{-sy}px)">
					{#each rowIdx as vi (vi)}
						{@const r = vis[vi]}
						{#each colIdx as c (c)}
							<!-- Leer desde `table` (referencia nueva por revisión): las filas se mutan en el lugar, un
							     {@const row} intermedio sería el mismo array y la celda no se repintaría. -->
							{@const v = table.rows[r]?.[c] ?? null}
							{@const ty = table.types[c] ?? 'auto'}
							{@const bad = !fits(v, ty)}
							<div
								class="cell absolute flex items-center overflow-hidden border-r border-b border-grid px-2 whitespace-pre {typeof v === 'number'
									? 'justify-end tabular-nums'
									: typeof v === 'boolean'
										? 'justify-center font-mono text-xs text-info'
										: ''} {highlight?.has(`${r}:${c}`) ? 'bg-match' : ''} {bad ? 'text-bad underline decoration-bad/60 decoration-wavy' : ''}"
								style="left:{cols.offset(c)}px;top:{vi * ROW_H}px;width:{cols.size(c)}px;height:{ROW_H}px"
								title={bad ? t().grid.typeHint(typeLabel(ty)) : undefined}
							>
								<span class="truncate">{display(v)}</span>
							</div>
						{/each}
					{/each}
					{#if selRect && !(whole === 'all' && !nRows)}
						<div
							class="pointer-events-none absolute border-2 border-sel bg-sel-fill"
							style="left:{selRect.x - 1}px;top:{selRect.y - 1}px;width:{selRect.w + 1}px;height:{selRect.h + 1}px"
						></div>
						{#if !editing && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2)}
							<div
								class="pointer-events-none absolute border border-sel bg-surface/0 outline outline-1 outline-sel/40"
								style="left:{cols.offset(focus.c)}px;top:{focus.r * ROW_H}px;width:{cols.size(focus.c) - 1}px;height:{ROW_H - 1}px"
							></div>
						{/if}
					{/if}
					<!-- Sink / editor -->
					<textarea
						bind:this={sink}
						class="absolute z-10 resize-none border-2 border-sel bg-surface px-[7px] py-[4px] text-[13px] leading-[16px] text-fg outline-none {editing
							? 'opacity-100 shadow-lg'
							: 'pointer-events-none opacity-0'}"
						style="left:{editorBox.x - 1}px;top:{editorBox.y - 1}px;min-width:{Math.max(editorBox.w + 1, editing ? 160 : 0)}px;height:{editing ? 'auto' : `${ROW_H + 1}px`};min-height:{ROW_H + 1}px;field-sizing:content;max-height:240px"
						value={editing ? draft : ''}
						aria-label={ariaActive}
						aria-multiline="false"
						spellcheck="false"
						autocomplete="off"
						{onkeydown}
						{oninput}
						{oncopy}
						{oncut}
						{onpaste}
						onblur={() => editing && commit()}
					></textarea>
				</div>
			</div>

			<!-- Números de fila -->
			<div
				class="absolute overflow-hidden border-r border-line bg-grid-head"
				style="left:0;top:{HEAD_H}px;width:{RN}px;bottom:0"
				onpointerdown={rowHeadDown}
				onpointermove={rowHeadMove}
				onpointerup={up}
				oncontextmenu={openCellMenu}
				role="presentation"
			>
				<div class="absolute top-0 left-0 w-full" style="transform:translateY({-sy}px)">
					{#each rowIdx as vi (vi)}
						<div
							class="absolute flex w-full items-center justify-end border-b border-grid pr-2 text-xs tabular-nums {inSelRows(vi) ? 'bg-grid-head-active font-semibold text-fg' : 'text-faint'}"
							style="top:{vi * ROW_H}px;height:{ROW_H}px"
						>
							{vis[vi] + 1}
						</div>
					{/each}
				</div>
			</div>

			<!-- Cabecera -->
			<div class="absolute overflow-hidden border-b border-line bg-grid-head" style="left:{RN}px;top:0;right:0;height:{HEAD_H}px">
				<div class="absolute top-0 left-0 h-full" style="transform:translateX({-sx}px)">
					{#each colIdx as c (c)}
						{@const name = table.headers[c]}
						{@const ty = table.types[c] ?? 'auto'}
						{@const active = isActive(filters[c])}
						<div
							class="group absolute flex h-full items-center gap-1 border-r border-grid pr-1 pl-2 font-semibold {inSelCols(c) ? 'bg-grid-head-active text-fg' : 'text-muted'}"
							style="left:{cols.offset(c)}px;width:{cols.size(c)}px"
							onpointerdown={(e) => colHeadDown(e, c)}
							onpointermove={colHeadMove}
							onpointerup={up}
							ondblclick={(e) => !(e.target as HTMLElement).closest('[data-noselect]') && renameCol(c)}
							oncontextmenu={(e) => openColMenu(e, c)}
							role="columnheader"
							aria-colindex={c + 1}
							title={name}
							tabindex="-1"
						>
							<span class="min-w-0 flex-1 truncate">{name}</span>
							{#if ty !== 'auto'}
								<span class="shrink-0 rounded bg-surface-2 px-1 text-[10px] font-medium text-faint">{typeLabel(ty)}</span>
							{/if}
							<button
								data-noselect
								class="grid size-6 shrink-0 place-items-center rounded hover:bg-surface-2 {active ? 'text-accent' : 'text-faint opacity-70 group-hover:opacity-100'}"
								aria-label={t().grid.colMenu(name)}
								title={t().grid.colMenu(name)}
								onclick={(e) => openColMenu(e, c)}
							>
								<Icon name={active ? 'filter' : 'chevronDown'} size={14} />
							</button>
							<div
								data-noselect
								class="absolute top-0 -right-1 z-10 h-full w-2 cursor-col-resize hover:bg-accent/40"
								onpointerdown={(e) => resizeDown(e, c)}
								onpointermove={resizeMove}
								onpointerup={resizeUp}
								ondblclick={(e) => {
									e.stopPropagation();
									autofit(c);
								}}
								title={t().grid.widthReset}
								role="separator"
								aria-orientation="vertical"
							></div>
						</div>
					{/each}
				</div>
			</div>

			<!-- Esquina: seleccionar todo -->
			<button
				class="absolute top-0 left-0 border-r border-b border-line bg-grid-head hover:bg-surface-2"
				style="width:{RN}px;height:{HEAD_H}px"
				aria-label={t().grid.selectAll}
				title={t().grid.selectAll}
				onclick={() => {
					selectAll();
					focusSink();
				}}
			></button>

			{#if !nRows}
				<div class="pointer-events-none absolute inset-x-0 grid place-items-center text-sm text-faint" style="top:{HEAD_H + 24}px">
					{nCols ? (filtered ? t().find.results(0, 0) : t().grid.emptyTable) : t().grid.emptyTable}
				</div>
			{/if}
		</div>
		<div style="width:{totalW + 40}px;height:{totalH + 40}px"></div>
	</div>

	<ContextMenu at={cellMenu} onclose={() => (cellMenu = null)} label={t().grid.cellMenu}>
		{#snippet children(close)}
			{@const item = (fn: () => void) => () => {
				close();
				fn();
			}}
			<button class="menu-item" onclick={item(() => copySelection(true))}><Icon name="x" size={16} />{t().grid.cut}</button>
			<button class="menu-item" onclick={item(() => copySelection())}><Icon name="copy" size={16} />{t().grid.copy}</button>
			<button class="menu-item" onclick={item(() => pasteFromClipboard())}><Icon name="clipboard" size={16} />{t().grid.paste}</button>
			<button class="menu-item" onclick={item(clearSelection)}><Icon name="clear" size={16} />{t().grid.clearCells}</button>
			<div class="menu-sep"></div>
			<button class="menu-item" onclick={item(() => insertRows('above'))}><Icon name="rowAdd" size={16} />{t().grid.insertAbove}</button>
			<button class="menu-item" onclick={item(() => insertRows('below'))}><Icon name="rowAdd" size={16} />{t().grid.insertBelow}</button>
			<button class="menu-item" disabled={!nRows} onclick={item(deleteRows)}><Icon name="rowDel" size={16} />{t().grid.deleteRows(sel.r2 - sel.r1 + 1)}</button>
			<button class="menu-item" disabled={!canMoveRows() || sel.r1 === 0} onclick={item(() => moveRows(-1))}><Icon name="arrowUp" size={16} />{t().grid.moveUp}</button>
			<button class="menu-item" disabled={!canMoveRows() || sel.r2 >= nRows - 1} onclick={item(() => moveRows(1))}><Icon name="arrowDown" size={16} />{t().grid.moveDown}</button>
			<div class="menu-sep"></div>
			<button class="menu-item" onclick={item(() => insertCols('left'))}><Icon name="colAdd" size={16} />{t().grid.insertLeft}</button>
			<button class="menu-item" onclick={item(() => insertCols('right'))}><Icon name="colAdd" size={16} />{t().grid.insertRight}</button>
			<button class="menu-item" disabled={!nCols} onclick={item(deleteCols)}><Icon name="colDel" size={16} />{t().grid.deleteCols(sel.c2 - sel.c1 + 1)}</button>
			<div class="menu-sep"></div>
			<button class="menu-item" disabled={!nRows} onclick={item(() => sortBy(focus.c, 'asc'))}><Icon name="sortAsc" size={16} />{t().grid.sortAsc}</button>
			<button class="menu-item" disabled={!nRows} onclick={item(() => sortBy(focus.c, 'desc'))}><Icon name="sortDesc" size={16} />{t().grid.sortDesc}</button>
		{/snippet}
	</ContextMenu>

	<ContextMenu at={colMenu?.at ?? null} onclose={() => (colMenu = null)} label={colMenu ? t().grid.colMenu(table.headers[colMenu.c] ?? '') : ''}>
		{#snippet children(close)}
			{#if colMenu}
				{@const c = colMenu.c}
				{@const f = filters[c]}
				{@const item = (fn: () => void) => () => {
					close();
					fn();
				}}
				<div class="w-64">
					<button class="menu-item" disabled={!nRows} onclick={item(() => sortBy(c, 'asc'))}><Icon name="sortAsc" size={16} />{t().grid.sortAsc}</button>
					<button class="menu-item" disabled={!nRows} onclick={item(() => sortBy(c, 'desc'))}><Icon name="sortDesc" size={16} />{t().grid.sortDesc}</button>
					<div class="menu-sep"></div>
					<div class="px-3 py-1.5">
						<div class="mb-1 flex items-center justify-between text-xs font-semibold text-muted">
							<span class="flex items-center gap-1"><Icon name="filter" size={13} />{t().grid.filter}</span>
							{#if isActive(f)}
								<button class="text-accent hover:underline" onclick={() => setFilter(c, null)}>{t().grid.clearFilter}</button>
							{/if}
						</div>
						<input
							class="input mb-2 w-full py-1"
							placeholder={t().grid.filterQuery}
							bind:value={filterQuery}
							oninput={() => setFilter(c, { ...f, query: filterQuery })}
							onkeydown={(e) => e.key === 'Enter' && close()}
						/>
						<div class="mb-1 flex items-center justify-between text-xs text-muted">
							<span>{t().grid.filterValues}</span>
							<span class="flex gap-2">
								<button class="hover:text-accent" onclick={() => setAllValues(c, true)}>{t().common.all}</button>
								<button class="hover:text-accent" onclick={() => setAllValues(c, false)}>{t().common.none}</button>
							</span>
						</div>
						{#if (colValues?.values.length ?? 0) > 12}
							<input class="input mb-1 w-full py-0.5 text-xs" placeholder={t().find.find} bind:value={valueSearch} />
						{/if}
						<div class="max-h-44 overflow-y-auto rounded border border-line">
							{#each shownValues.slice(0, 300) as v (v)}
								<label class="flex cursor-pointer items-center gap-2 px-2 py-0.5 text-sm hover:bg-surface-2">
									<input type="checkbox" class="check" checked={!f?.values || f.values.includes(v)} onchange={(e) => toggleValue(c, v, e.currentTarget.checked)} />
									<span class="truncate {v === '' ? 'text-faint italic' : ''}">{v === '' ? t().grid.filterEmpty : v}</span>
								</label>
							{/each}
						</div>
						{#if colValues?.truncated}<p class="mt-1 text-[11px] text-faint">{t().grid.filterTruncated}</p>{/if}
					</div>
					<div class="menu-sep"></div>
					<div class="px-3 py-1.5">
						<div class="mb-1 flex items-center gap-1 text-xs font-semibold text-muted"><Icon name="type" size={13} />{t().grid.colType}</div>
						<div class="flex flex-wrap gap-1">
							{#each COL_TYPES as ty (ty)}
								<button
									class="rounded border px-2 py-0.5 text-xs {table.types[c] === ty ? 'border-accent bg-accent-soft text-accent' : 'border-line hover:bg-surface-2'}"
									aria-pressed={table.types[c] === ty}
									onclick={() => setType(c, ty)}>{typeLabel(ty)}</button
								>
							{/each}
						</div>
					</div>
					<div class="menu-sep"></div>
					<button class="menu-item" onclick={item(() => renameCol(c))}><Icon name="rename" size={16} />{t().grid.renameCol}</button>
					<button class="menu-item" onclick={item(() => insertCols('left'))}><Icon name="colAdd" size={16} />{t().grid.insertLeft}</button>
					<button class="menu-item" onclick={item(() => insertCols('right'))}><Icon name="colAdd" size={16} />{t().grid.insertRight}</button>
					<button class="menu-item" disabled={sel.c1 === 0} onclick={item(() => moveCols(-1))}><Icon name="arrowLeft" size={16} />{t().grid.moveLeft}</button>
					<button class="menu-item" disabled={sel.c2 >= nCols - 1} onclick={item(() => moveCols(1))}><Icon name="arrowRight" size={16} />{t().grid.moveRight}</button>
					<button class="menu-item text-bad" onclick={item(deleteCols)}><Icon name="trash" size={16} />{t().grid.deleteCols(sel.c2 - sel.c1 + 1)}</button>
				</div>
			{/if}
		{/snippet}
	</ContextMenu>
{/if}

<style>
	.cell {
		contain: layout paint style;
	}
</style>
