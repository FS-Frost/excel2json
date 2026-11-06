// Claves anidadas: `{a: {b: 1}}` ⇄ columna `a.b`. Los arrays nunca se aplanan: viajan como texto JSON.
import type { Cell } from './types.ts';

const isPlainObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v);

/** Valor JSON → celda: escalares tal cual; objetos y arrays como texto JSON. */
export function jsonToCell(v: unknown): Cell {
	if (v === null || v === undefined) return null;
	if (typeof v === 'string' || typeof v === 'boolean') return v;
	if (typeof v === 'number') return Number.isFinite(v) ? v : null;
	return JSON.stringify(v);
}

/**
 * Objeto → pares columna/valor. Con `nested`, los objetos se recorren (`a.b.c`); un objeto vacío queda como `{}` en
 * su propia columna para no perderlo. Sin `nested`, cada clave de primer nivel es una columna.
 */
export function flatten(obj: Record<string, unknown>, nested: boolean, prefix = '', out: [string, Cell][] = []): [string, Cell][] {
	for (const [k, v] of Object.entries(obj)) {
		const key = prefix ? `${prefix}.${k}` : k;
		if (nested && isPlainObject(v) && Object.keys(v).length) flatten(v, true, key, out);
		else out.push([key, jsonToCell(v)]);
	}
	return out;
}

/**
 * Columnas → objeto. Con `nested`, `a.b` crea `{a: {b}}`. Si una ruta choca con un valor ya puesto (`a` = 1 y
 * `a.b` = 2), esa columna se deja plana (`"a.b": 2`) para no perder datos.
 */
export function unflatten(keys: string[], values: Cell[], nested: boolean, dropNulls = false): Record<string, unknown> {
	const out: Record<string, unknown> = {};
	for (let i = 0; i < keys.length; i++) {
		const v = values[i] ?? null;
		if (dropNulls && v === null) continue;
		const key = keys[i];
		if (!nested || !key.includes('.')) {
			out[key] = v;
			continue;
		}
		const parts = key.split('.');
		if (parts.some((p) => p === '')) {
			out[key] = v;
			continue;
		}
		let node: Record<string, unknown> = out;
		let ok = true;
		for (let j = 0; j < parts.length - 1; j++) {
			const p = parts[j];
			if (!(p in node)) node[p] = {};
			else if (!isPlainObject(node[p])) {
				ok = false;
				break;
			}
			node = node[p] as Record<string, unknown>;
		}
		const last = parts[parts.length - 1];
		if (ok && !(last in node && isPlainObject(node[last]))) node[last] = v;
		else out[key] = v;
	}
	return out;
}
