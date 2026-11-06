// Filtro de la vista (no modifica datos) y orden (sí modifica: produce una permutación que se aplica como operación).
import { cellText } from '../core/cell.ts';
import type { Cell, Table } from '../core/types.ts';

/** Criterio de una columna: texto contenido (sin mayúsculas ni tildes) y/o lista de valores permitidos. */
export interface ColumnFilter {
	query?: string;
	/** Textos permitidos (como se muestran); null/undefined = todos. */
	values?: string[] | null;
}

export type Filters = Record<number, ColumnFilter>;

/** Minúsculas y sin tildes, para buscar «canción» con «cancion». */
export const fold = (s: string) => s.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase();

export function isActive(f: ColumnFilter | undefined): boolean {
	return !!f && (!!f.query?.trim() || Array.isArray(f.values));
}

export function activeFilters(filters: Filters): [number, ColumnFilter][] {
	return Object.entries(filters)
		.filter(([, f]) => isActive(f))
		.map(([c, f]) => [Number(c), f]);
}

/** Índices reales de las filas visibles. Sin filtros activos: todas. */
export function visibleRows(t: Table, filters: Filters): number[] {
	const act = activeFilters(filters).map(([c, f]) => ({ c, q: f.query?.trim() ? fold(f.query.trim()) : '', values: f.values ? new Set(f.values) : null }));
	const out: number[] = [];
	for (let r = 0; r < t.rows.length; r++) {
		const row = t.rows[r];
		let ok = true;
		for (const f of act) {
			const text = cellText(row[f.c] ?? null);
			if (f.values && !f.values.has(text)) ok = false;
			else if (f.q && !fold(text).includes(f.q)) ok = false;
			if (!ok) break;
		}
		if (ok) out.push(r);
	}
	return out;
}

/** Valores distintos de una columna (texto mostrado), en orden, como mucho `limit`. */
export function distinctValues(t: Table, c: number, limit = 1000): { values: string[]; truncated: boolean } {
	const seen = new Set<string>();
	for (const row of t.rows) {
		seen.add(cellText(row[c] ?? null));
		if (seen.size > limit) break;
	}
	const values = [...seen].slice(0, limit).sort((a, b) => (a === '' ? 1 : b === '' ? -1 : collator.compare(a, b)));
	return { values, truncated: seen.size > limit };
}

const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });

/** Orden: vacíos al final; números < booleanos < textos; textos con orden natural ("2" antes que "10"). */
export function compareCells(a: Cell, b: Cell): number {
	if (a === b) return 0;
	if (a === null) return 1;
	if (b === null) return -1;
	const rank = (v: Cell) => (typeof v === 'number' ? 0 : typeof v === 'boolean' ? 1 : 2);
	const ra = rank(a);
	const rb = rank(b);
	if (ra !== rb) return ra - rb;
	if (typeof a === 'number' && typeof b === 'number') return a - b;
	if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
	return collator.compare(String(a), String(b));
}

/** Permutación estable que ordena las filas por la columna `c` (los vacíos siempre al final). */
export function sortOrder(t: Table, c: number, dir: 'asc' | 'desc'): number[] {
	const idx = t.rows.map((_, i) => i);
	const sign = dir === 'asc' ? 1 : -1;
	idx.sort((i, j) => {
		const a = t.rows[i][c] ?? null;
		const b = t.rows[j][c] ?? null;
		if (a === null || b === null) return a === b ? i - j : a === null ? 1 : -1;
		return sign * compareCells(a, b) || i - j;
	});
	return idx;
}
