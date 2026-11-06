// Operaciones de edición sobre una Table (mutan en el lugar). `apply` devuelve la operación inversa exacta, así
// deshacer/rehacer no guarda copias de la tabla. Índices de fila = posición real en `table.rows` (no la de la vista).
import type { Cell, ColType, Table } from '../core/types.ts';

export type CellChange = { r: number; c: number; v: Cell };

export type Op =
	| { t: 'cells'; changes: CellChange[] }
	| { t: 'insertRows'; at: number; rows: Cell[][] }
	| { t: 'deleteRows'; at: number[] }
	| { t: 'restoreRows'; at: number[]; rows: Cell[][] }
	| { t: 'insertCols'; at: number; headers: string[]; types: ColType[]; cols?: Cell[][] }
	| { t: 'deleteCols'; at: number[] }
	| { t: 'restoreCols'; at: number[]; headers: string[]; types: ColType[]; cols: Cell[][] }
	| { t: 'header'; c: number; name: string }
	| { t: 'name'; name: string }
	| { t: 'type'; c: number; type: ColType }
	| { t: 'permuteRows'; order: number[] }
	| { t: 'permuteCols'; order: number[] };

const sortAsc = (a: number[]) => [...new Set(a)].sort((x, y) => x - y);

/** Inversa de una permutación: si nuevo[i] = viejo[order[i]], entonces viejo[j] = nuevo[inv[j]]. */
export function invert(order: number[]): number[] {
	const inv = new Array<number>(order.length);
	order.forEach((o, i) => (inv[o] = i));
	return inv;
}

export function apply(t: Table, op: Op): Op {
	const width = t.headers.length;
	switch (op.t) {
		case 'cells': {
			const before: CellChange[] = [];
			for (const { r, c, v } of op.changes) {
				const row = t.rows[r];
				if (!row || c < 0 || c >= width) continue;
				before.push({ r, c, v: row[c] });
				row[c] = v;
			}
			return { t: 'cells', changes: before.reverse() };
		}
		case 'insertRows': {
			const rows = op.rows.map((r) => Array.from({ length: width }, (_, c) => r[c] ?? null));
			t.rows.splice(op.at, 0, ...rows);
			return { t: 'deleteRows', at: rows.map((_, i) => op.at + i) };
		}
		case 'deleteRows': {
			const at = sortAsc(op.at).filter((i) => i >= 0 && i < t.rows.length);
			const drop = new Set(at);
			const rows = at.map((i) => t.rows[i]);
			t.rows = t.rows.filter((_, i) => !drop.has(i));
			return { t: 'restoreRows', at, rows };
		}
		case 'restoreRows': {
			op.at.forEach((idx, k) => t.rows.splice(idx, 0, op.rows[k]));
			return { t: 'deleteRows', at: [...op.at] };
		}
		case 'insertCols': {
			const n = op.headers.length;
			t.headers.splice(op.at, 0, ...op.headers);
			t.types.splice(op.at, 0, ...op.types);
			t.rows.forEach((row, r) => row.splice(op.at, 0, ...Array.from({ length: n }, (_, k) => op.cols?.[k]?.[r] ?? null)));
			return { t: 'deleteCols', at: Array.from({ length: n }, (_, k) => op.at + k) };
		}
		case 'deleteCols': {
			const at = sortAsc(op.at).filter((i) => i >= 0 && i < width);
			const drop = new Set(at);
			const inv: Op = {
				t: 'restoreCols',
				at,
				headers: at.map((c) => t.headers[c]),
				types: at.map((c) => t.types[c]),
				cols: at.map((c) => t.rows.map((r) => r[c]))
			};
			const keep = (_: unknown, i: number) => !drop.has(i);
			t.headers = t.headers.filter(keep);
			t.types = t.types.filter(keep);
			t.rows = t.rows.map((r) => r.filter(keep));
			return inv;
		}
		case 'restoreCols': {
			op.at.forEach((c, k) => {
				t.headers.splice(c, 0, op.headers[k]);
				t.types.splice(c, 0, op.types[k]);
				t.rows.forEach((row, r) => row.splice(c, 0, op.cols[k][r] ?? null));
			});
			return { t: 'deleteCols', at: [...op.at] };
		}
		case 'header': {
			const old = t.headers[op.c];
			t.headers[op.c] = op.name;
			return { t: 'header', c: op.c, name: old };
		}
		case 'name': {
			const old = t.name;
			t.name = op.name;
			return { t: 'name', name: old };
		}
		case 'type': {
			const old = t.types[op.c];
			t.types[op.c] = op.type;
			return { t: 'type', c: op.c, type: old };
		}
		case 'permuteRows': {
			const old = t.rows;
			t.rows = op.order.map((i) => old[i]);
			return { t: 'permuteRows', order: invert(op.order) };
		}
		case 'permuteCols': {
			const o = op.order;
			const h = t.headers;
			const ty = t.types;
			t.headers = o.map((i) => h[i]);
			t.types = o.map((i) => ty[i]);
			t.rows = t.rows.map((r) => o.map((i) => r[i]));
			return { t: 'permuteCols', order: invert(o) };
		}
	}
}

/** Aplica varias operaciones; devuelve las inversas en el orden en que hay que aplicarlas para deshacer. */
export function applyAll(t: Table, ops: Op[]): Op[] {
	return ops.map((op) => apply(t, op)).reverse();
}

/** Permutación que mueve el bloque [from, from+count) para que empiece en `to` (índice en el arreglo final). */
export function moveOrder(length: number, from: number, count: number, to: number): number[] {
	const idx = Array.from({ length }, (_, i) => i);
	const block = idx.splice(from, count);
	const dest = Math.max(0, Math.min(to, idx.length));
	idx.splice(dest, 0, ...block);
	return idx;
}

/** Nombre libre para una columna nueva: `columna`, `columna_2`… */
export function freeHeader(headers: string[], base: string): string {
	const used = new Set(headers);
	if (!used.has(base)) return base;
	let k = 2;
	while (used.has(`${base}_${k}`)) k++;
	return `${base}_${k}`;
}
