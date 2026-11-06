// Unir tablas: unión de columnas por nombre (orden de aparición), celdas faltantes en null.
import { baseName } from './naming.ts';
import { newId, type Cell, type ColType, type Table } from './types.ts';

/** Etiqueta del origen de una tabla: `archivo › hoja` (o solo el archivo si la tabla se llama como él). */
export function originLabel(t: Table): string {
	const file = baseName(t.source);
	return t.name === file ? file : `${file} › ${t.name}`;
}

/**
 * Une las tablas en una. `originColumn` (si no es null) agrega al inicio una columna con el origen de cada fila;
 * si choca con una columna existente se le agrega `_2`. El tipo de una columna se conserva solo si todas las tablas
 * que la tienen coinciden; si no, `auto`.
 */
export function mergeTables(tables: Table[], name: string, originColumn: string | null): Table {
	const index = new Map<string, number>();
	const headers: string[] = [];
	const types: ColType[] = [];
	for (const t of tables)
		t.headers.forEach((h, c) => {
			const i = index.get(h);
			if (i === undefined) {
				index.set(h, headers.length);
				headers.push(h);
				types.push(t.types[c] ?? 'auto');
			} else if (types[i] !== (t.types[c] ?? 'auto')) types[i] = 'auto';
		});
	let origin = originColumn;
	if (origin !== null) while (index.has(origin)) origin = `${origin}_2`;
	const extra = origin !== null ? 1 : 0;
	const rows: Cell[][] = [];
	for (const t of tables) {
		const map = t.headers.map((h) => index.get(h)! + extra);
		const label = originLabel(t);
		for (const r of t.rows) {
			const row: Cell[] = new Array(headers.length + extra).fill(null);
			if (extra) row[0] = label;
			for (let c = 0; c < map.length; c++) row[map[c]] = r[c] ?? null;
			rows.push(row);
		}
	}
	return {
		id: newId(),
		name,
		source: tables[0]?.source ?? name,
		headers: origin !== null ? [origin, ...headers] : headers,
		types: origin !== null ? ['text', ...types] : types,
		rows
	};
}
