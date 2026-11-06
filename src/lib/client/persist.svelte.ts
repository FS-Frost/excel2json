// Autoguardado de la sesión en IndexedDB y restauración al abrir. Mientras haya una sesión guardada sin decidir
// (restaurar/descartar), no se guarda nada para no pisarla.
import { createAutosave, type SaveStatus } from '../storage/autosave.svelte.ts';
import { clearSession, loadMeta, loadTables, saveSession, type SessionMeta } from '../storage/db.ts';
import { session } from './session.svelte.ts';

export const saved = $state({ meta: null as SessionMeta | null, lowSpace: false });

let autosave: ReturnType<typeof createAutosave> | null = null;

let markReady: () => void = () => {};
/** Se resuelve cuando ya se sabe si hay una sesión guardada. Abrir archivos debe esperarlo: si no, `saved.meta` podría
 *  llegar después y bloquear el autoguardado de lo recién abierto. */
export const persistReady = new Promise<void>((r) => (markReady = r));

export function saveStatus(): SaveStatus {
	return autosave?.state.status ?? 'saved';
}

export async function initPersistence() {
	try {
		const meta = await loadMeta();
		saved.meta = meta && meta.files.length ? meta : null;
	} finally {
		markReady();
	}
	autosave = createAutosave(async () => {
		const dirty = session.takeDirty();
		try {
			await saveSession(session.toMeta(), dirty, new Set(session.tables.map((t) => t.id)));
		} catch (e) {
			session.requeue(dirty);
			throw e;
		}
	});
	session.onChange = () => {
		if (!saved.meta) autosave?.schedule();
	};
	try {
		const est = await navigator.storage?.estimate?.();
		if (est?.quota && est.usage !== undefined && est.quota - est.usage < 50 * 1024 * 1024) saved.lowSpace = true;
	} catch {
		/* sin API de almacenamiento */
	}
}

export async function restoreSaved() {
	const meta = saved.meta;
	if (!meta) return;
	const tables = await loadTables(meta.files.flatMap((f) => f.tableIds));
	session.restore(meta, tables);
	saved.meta = null;
}

/** Descarta la sesión guardada (al empezar una nueva). Lo que ya esté abierto se guarda completo. */
export async function discardSaved() {
	if (!saved.meta) return;
	saved.meta = null;
	try {
		await clearSession();
	} catch {
		/* se sobrescribe en el próximo guardado */
	}
	session.markAllDirty();
	autosave?.schedule();
}

export function flushSave() {
	return autosave?.flush();
}
