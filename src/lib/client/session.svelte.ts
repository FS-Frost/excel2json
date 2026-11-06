// Sesión de trabajo: archivos abiertos, tablas, selección para exportar, tabla activa, filtros/anchos de la vista e
// historial de edición. Las tablas NO son proxies reactivos (pueden tener millones de celdas): se mutan en el lugar y
// `rev` avisa a quien las lee. Cada cambio llama a `onChange` (autoguardado).
import { untrack } from 'svelte';
import type { SourceFile, Table } from '../core/types.ts';
import { History, type Step } from '../edit/history.ts';
import { applyAll, type Op } from '../edit/ops.ts';
import { isActive, visibleRows, type Filters } from '../edit/view.ts';
import type { SessionMeta, TableViewState } from '../storage/db.ts';

const COL_STRUCT = new Set<Op['t']>(['insertCols', 'deleteCols', 'restoreCols', 'permuteCols']);

class Session {
	files = $state.raw<SourceFile[]>([]);
	/** Sube con cualquier cambio de datos. Leerlo registra la dependencia. */
	rev = $state(0);
	selected = $state<Record<string, boolean>>({});
	activeId = $state<string | null>(null);
	views = $state<Record<string, TableViewState>>({});
	/** Tablas con ediciones (para marcarlas en la lista). */
	edited = $state.raw<Set<string>>(new Set());
	/** Archivos restaurados de una sesión guardada: no tienen el original para reimportar. */
	restored = $state.raw<Set<string>>(new Set());
	histRev = $state(0);

	readonly history = new History();
	private originals = new Map<string, File>();
	private dirty = new Set<string>();
	onChange: (() => void) | null = null;

	// ---------- Lectura ----------

	get tables(): Table[] {
		void this.rev;
		return this.files.flatMap((f) => f.tables);
	}

	table(id: string | null): Table | null {
		if (!id) return null;
		for (const f of this.files) for (const t of f.tables) if (t.id === id) return t;
		return null;
	}

	fileOf(tableId: string): SourceFile | null {
		return this.files.find((f) => f.tables.some((t) => t.id === tableId)) ?? null;
	}

	get active(): Table | null {
		void this.rev;
		return this.table(this.activeId);
	}

	view(id: string): TableViewState {
		return this.views[id] ?? { filters: {}, widths: {} };
	}

	get canUndo() {
		void this.histRev;
		return this.history.canUndo;
	}
	get canRedo() {
		void this.histRev;
		return this.history.canRedo;
	}

	original(fileId: string): File | undefined {
		return this.originals.get(fileId);
	}

	/** Tablas marcadas para exportar, en el orden de los archivos. Con `onlyVisible`, solo las filas que pasan el filtro. */
	exportTables(onlyVisible: boolean): Table[] {
		const out: Table[] = [];
		for (const t of this.tables) {
			if (!this.selected[t.id]) continue;
			const filters = this.view(t.id).filters;
			if (onlyVisible && Object.values(filters).some(isActive)) {
				const rows = visibleRows(t, filters).map((r) => t.rows[r]);
				out.push({ ...t, rows });
			} else out.push(t);
		}
		return out;
	}

	// ---------- Cambios ----------

	private bump(tableIds: Iterable<string> = []) {
		for (const id of tableIds) this.dirty.add(id);
		untrack(() => {
			this.rev++;
		});
		this.onChange?.();
	}

	private bumpHistory() {
		untrack(() => {
			this.histRev++;
		});
	}

	private markEdited(id: string) {
		if (this.edited.has(id)) return;
		this.edited = new Set([...this.edited, id]);
	}

	/** Cambia algo que no son datos (selección, vista): se guarda, sin marcar tablas. */
	touch() {
		this.onChange?.();
	}

	private resetColumnFilters(tableId: string, ops: Op[]) {
		if (!ops.some((o) => COL_STRUCT.has(o.t))) return;
		const v = this.views[tableId];
		if (v && Object.keys(v.filters).length) v.filters = {};
	}

	/** Aplica operaciones como un paso de deshacer. */
	exec(tableId: string, ops: Op[]) {
		const t = this.table(tableId);
		if (!t || !ops.length) return;
		const inverse = applyAll(t, ops);
		this.history.push({ tableId, ops, inverse });
		this.resetColumnFilters(tableId, ops);
		this.markEdited(tableId);
		this.bumpHistory();
		this.bump([tableId]);
	}

	private replay(step: Step, toRedo: boolean): string | null {
		const t = this.table(step.tableId);
		if (!t) return null;
		const back = applyAll(t, step.inverse);
		const next: Step = { tableId: step.tableId, ops: step.inverse, inverse: back };
		if (toRedo) this.history.pushRedo(next);
		else this.history.pushUndoKeepRedo(next);
		this.resetColumnFilters(step.tableId, step.inverse);
		this.markEdited(step.tableId);
		this.bumpHistory();
		this.bump([step.tableId]);
		return step.tableId;
	}

	/** Deshace el último paso; devuelve la tabla afectada (para mostrarla). */
	undo(): string | null {
		const s = this.history.popUndo();
		return s ? this.replay(s, true) : null;
	}

	redo(): string | null {
		const s = this.history.popRedo();
		return s ? this.replay(s, false) : null;
	}

	setFilters(tableId: string, filters: Filters) {
		const v = this.views[tableId] ?? (this.views[tableId] = { filters: {}, widths: {} });
		v.filters = filters;
		this.touch();
	}

	setWidth(tableId: string, header: string, w: number | null) {
		const v = this.views[tableId] ?? (this.views[tableId] = { filters: {}, widths: {} });
		if (w === null) delete v.widths[header];
		else v.widths[header] = Math.round(w);
		this.touch();
	}

	setActive(id: string | null) {
		this.activeId = id;
		this.touch();
	}

	setSelected(id: string, on: boolean) {
		this.selected[id] = on;
		this.touch();
	}

	selectAll(on: boolean) {
		for (const t of this.tables) this.selected[t.id] = on;
		this.touch();
	}

	// ---------- Archivos ----------

	addFiles(list: { file: SourceFile; original?: File }[]) {
		if (!list.length) return;
		for (const { file, original } of list) {
			if (original) this.originals.set(file.id, original);
			for (const t of file.tables) this.selected[t.id] = true;
		}
		this.files = [...this.files, ...list.map((x) => x.file)];
		if (!this.table(this.activeId)) this.activeId = list[0].file.tables[0]?.id ?? null;
		this.bump(list.flatMap((x) => x.file.tables.map((t) => t.id)));
	}

	private forgetTables(tables: Table[]) {
		const ids = new Set(tables.map((t) => t.id));
		for (const id of ids) {
			this.history.forget(id);
			delete this.selected[id];
			delete this.views[id];
		}
		this.edited = new Set([...this.edited].filter((id) => !ids.has(id)));
		this.bumpHistory();
	}

	removeFile(fileId: string) {
		const f = this.files.find((x) => x.id === fileId);
		if (!f) return;
		this.forgetTables(f.tables);
		this.originals.delete(fileId);
		this.restored = new Set([...this.restored].filter((id) => id !== fileId));
		this.files = this.files.filter((x) => x.id !== fileId);
		if (!this.table(this.activeId)) this.activeId = this.files[0]?.tables[0]?.id ?? null;
		this.bump();
	}

	/** Reemplaza un archivo por su versión releída (mismas posiciones; selección por nombre de hoja). */
	replaceFile(fileId: string, next: SourceFile) {
		const i = this.files.findIndex((x) => x.id === fileId);
		if (i < 0) return;
		const old = this.files[i];
		const wasSelected = new Map(old.tables.map((t) => [t.name, this.selected[t.id] ?? true]));
		const activeName = old.tables.find((t) => t.id === this.activeId)?.name;
		this.forgetTables(old.tables);
		const original = this.originals.get(fileId);
		this.originals.delete(fileId);
		if (original) this.originals.set(next.id, original);
		for (const t of next.tables) this.selected[t.id] = wasSelected.get(t.name) ?? true;
		const files = [...this.files];
		files[i] = next;
		this.files = files;
		if (activeName !== undefined) this.activeId = next.tables.find((t) => t.name === activeName)?.id ?? next.tables[0]?.id ?? null;
		this.bump(next.tables.map((t) => t.id));
	}

	clear() {
		this.forgetTables(this.tables);
		this.history.clear();
		this.originals.clear();
		this.restored = new Set();
		this.files = [];
		this.activeId = null;
		this.bump();
	}

	// ---------- Guardado ----------

	toMeta(): SessionMeta {
		return {
			v: 1,
			savedAt: Date.now(),
			files: this.files.map((f) => ({ id: f.id, name: f.name, kind: f.kind, size: f.size, warnings: f.warnings, tableIds: f.tables.map((t) => t.id) })),
			selected: $state.snapshot(this.selected),
			activeId: this.activeId,
			views: $state.snapshot(this.views),
			edited: [...this.edited]
		};
	}

	/** Tablas cambiadas desde el último guardado (y las vacía de la lista). */
	takeDirty(): Table[] {
		const out: Table[] = [];
		for (const id of this.dirty) {
			const t = this.table(id);
			if (t) out.push(t);
		}
		this.dirty.clear();
		return out;
	}

	/** Si falla el guardado, las tablas vuelven a la cola. */
	requeue(tables: Table[]) {
		for (const t of tables) this.dirty.add(t.id);
	}

	restore(meta: SessionMeta, tables: Map<string, Table>) {
		const files: SourceFile[] = meta.files.map((f) => ({
			id: f.id,
			name: f.name,
			kind: f.kind,
			size: f.size,
			warnings: f.warnings,
			tables: f.tableIds.map((id) => tables.get(id)).filter((t): t is Table => !!t)
		}));
		this.history.clear();
		this.originals.clear();
		this.files = files.filter((f) => f.tables.length);
		this.restored = new Set(this.files.map((f) => f.id));
		this.selected = { ...meta.selected };
		this.views = { ...meta.views };
		this.edited = new Set(meta.edited);
		this.activeId = this.table(meta.activeId) ? meta.activeId : (this.files[0]?.tables[0]?.id ?? null);
		this.dirty.clear();
		untrack(() => {
			this.rev++;
		});
		this.bumpHistory();
	}

	/** Todas las tablas pendientes (tras restaurar no hay ninguna; al empezar de cero, todas). */
	markAllDirty() {
		for (const t of this.tables) this.dirty.add(t.id);
	}
}

export const session = new Session();
