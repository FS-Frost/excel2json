// Sesión de trabajo en IndexedDB: `meta` (archivos, selección, vista) y `tables` (una entrada por tabla), para que el
// autoguardado escriba solo las tablas que cambiaron. Solo se guarda la última sesión.
import { openDB, type DBSchema, type IDBPDatabase } from 'idb';
import type { ImportWarning, SourceKind, Table } from '../core/types.ts';
import type { Filters } from '../edit/view.ts';

export interface SavedFile {
	id: string;
	name: string;
	kind: SourceKind;
	size: number;
	warnings: ImportWarning[];
	tableIds: string[];
}

export interface TableViewState {
	filters: Filters;
	/** Anchos de columna por nombre de cabecera (sobreviven a mover/insertar columnas). */
	widths: Record<string, number>;
}

export interface SessionMeta {
	v: 1;
	savedAt: number;
	files: SavedFile[];
	selected: Record<string, boolean>;
	activeId: string | null;
	views: Record<string, TableViewState>;
	edited: string[];
}

interface Schema extends DBSchema {
	meta: { key: string; value: SessionMeta };
	tables: { key: string; value: Table };
}

const KEY = 'current';
let dbp: Promise<IDBPDatabase<Schema>> | null = null;

function db() {
	dbp ??= openDB<Schema>('excel2json', 1, {
		upgrade(d) {
			d.createObjectStore('meta');
			d.createObjectStore('tables');
		}
	});
	return dbp;
}

export async function loadMeta(): Promise<SessionMeta | null> {
	try {
		const m = await (await db()).get('meta', KEY);
		return m?.v === 1 ? m : null;
	} catch {
		return null;
	}
}

export async function loadTables(ids: string[]): Promise<Map<string, Table>> {
	const tx = (await db()).transaction('tables');
	const out = new Map<string, Table>();
	for (const id of ids) {
		const t = await tx.store.get(id);
		if (t) out.set(id, t);
	}
	await tx.done;
	return out;
}

/** Escribe la meta, las tablas cambiadas y borra las que ya no existen, en una transacción. */
export async function saveSession(meta: SessionMeta, dirty: Table[], keep: Set<string>) {
	const tx = (await db()).transaction(['meta', 'tables'], 'readwrite');
	const tables = tx.objectStore('tables');
	// Primero los put (en una descarga de página, lo que viene tras un await puede no ejecutarse).
	void tx.objectStore('meta').put(meta, KEY);
	for (const t of dirty) void tables.put(t, t.id);
	for (const id of await tables.getAllKeys()) if (!keep.has(id)) void tables.delete(id);
	await tx.done;
}

export async function clearSession() {
	const tx = (await db()).transaction(['meta', 'tables'], 'readwrite');
	void tx.objectStore('meta').clear();
	void tx.objectStore('tables').clear();
	await tx.done;
}
