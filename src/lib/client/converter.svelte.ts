// Puente con el Worker de conversión: llamadas con id, avance reactivo y cancelación (termina el Worker).
import type { BuildResult, Request, Response, SerializedError } from '../worker/protocol.ts';
import type { ExportOptions, ImportOptions, SourceFile, Table } from '../core/types.ts';

export type Progress = { done: number; total: number; label: string };

type Pending = { resolve: (v: unknown) => void; reject: (e: SerializedError) => void; onProgress?: (p: Progress) => void };

/** Estado visible: hay trabajo en curso y su avance (el último que informó). */
export const converter = $state({ busy: 0, progress: null as Progress | null });

let worker: Worker | null = null;
let nextId = 1;
const pending = new Map<number, Pending>();

function failAll(error: SerializedError) {
	for (const p of pending.values()) p.reject(error);
	pending.clear();
	converter.busy = 0;
	converter.progress = null;
}

function getWorker(): Worker {
	if (worker) return worker;
	const w = new Worker(new URL('../worker/convert.worker.ts', import.meta.url), { type: 'module', name: 'excel2json-convert' });
	w.onmessage = (ev: MessageEvent<Response>) => {
		const msg = ev.data;
		const p = pending.get(msg.id);
		if (!p) return;
		if (msg.type === 'progress') {
			const prog = { done: msg.done, total: msg.total, label: msg.label };
			converter.progress = prog;
			p.onProgress?.(prog);
			return;
		}
		pending.delete(msg.id);
		converter.busy = Math.max(0, converter.busy - 1);
		if (!converter.busy) converter.progress = null;
		if (msg.type === 'ok') p.resolve(msg.result);
		else p.reject(msg.error);
	};
	w.onerror = (e) => {
		console.error('[excel2json] worker error', e);
		e.preventDefault();
		w.terminate();
		worker = null;
		failAll({ code: 'worker', detail: '', message: e.message ?? '' });
	};
	worker = w;
	return w;
}

type Body = Request extends infer R ? (R extends { id: number } ? Omit<R, 'id'> : never) : never;

function call<T>(body: Body, transfer: Transferable[] = [], onProgress?: (p: Progress) => void): Promise<T> {
	const w = getWorker();
	const id = nextId++;
	converter.busy++;
	return new Promise<T>((resolve, reject) => {
		pending.set(id, { resolve: resolve as (v: unknown) => void, reject, onProgress });
		w.postMessage({ ...body, id } as Request, transfer);
	});
}

/** Lee un archivo. Los bytes se transfieren al Worker (el buffer queda vacío aquí). */
export function parseFile(name: string, bytes: ArrayBuffer, opts: ImportOptions, onProgress?: (p: Progress) => void) {
	return call<SourceFile>({ type: 'parse', name, bytes, opts: { ...opts } }, [bytes], onProgress);
}

/** Genera la salida (un archivo o un zip). Las tablas se copian (structured clone). */
export function buildFile(tables: Table[], opts: ExportOptions, onProgress?: (p: Progress) => void) {
	return call<BuildResult>({ type: 'build', tables, opts: { ...opts } }, [], onProgress);
}

export function sampleWorkbook(locale: 'es' | 'en') {
	return call<ArrayBuffer>({ type: 'sample', locale });
}

/** Cancela todo lo pendiente: el Worker se termina y el próximo pedido crea uno nuevo. */
export function cancelAll() {
	worker?.terminate();
	worker = null;
	failAll({ code: 'cancelled', detail: '', message: '' });
}
