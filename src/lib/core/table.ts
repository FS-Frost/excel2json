// Grilla cruda (todas las filas, incluida la cabecera) → Table, según las opciones de importación.
import { cellText } from './cell.ts';
import { colLetter, normalizeHeaders } from './naming.ts';
import { newId, type Cell, type ImportOptions, type ImportWarning, type Table } from './types.ts';

export const DEFAULT_IMPORT: ImportOptions = { headerRow: 1, skipEmptyRows: true, nested: false };

const isEmptyRow = (r: Cell[] | undefined) => !r || r.every((v) => v === null || v === '');

/**
 * `grid` = filas tal como vienen del archivo (pueden tener largos distintos y huecos `undefined`).
 * Cabecera en `opts.headerRow` (1-based; 0 = sin cabecera). Las filas anteriores a la cabecera se descartan.
 */
export function gridToTable(grid: (Cell | undefined)[][], name: string, source: string, opts: ImportOptions): { table: Table; warnings: ImportWarning[] } {
	const warnings: ImportWarning[] = [];
	// Array.from (y no map) para recorrer también los huecos de arrays dispersos (filas/celdas vacías de exceljs).
	const rowsRaw: Cell[][] = Array.from(grid, (r) => (r ? Array.from(r, (v) => (v === undefined ? null : v)) : []));
	// Filas vacías del final: fuera siempre.
	let end = rowsRaw.length;
	while (end > 0 && isEmptyRow(rowsRaw[end - 1])) end--;
	const hasHeader = opts.headerRow > 0;
	const headerIdx = hasHeader ? opts.headerRow - 1 : -1;
	const headerRaw = hasHeader ? (rowsRaw[headerIdx] ?? []) : [];
	let data = rowsRaw.slice(hasHeader ? headerIdx + 1 : 0, Math.max(end, hasHeader ? headerIdx + 1 : 0));
	if (opts.skipEmptyRows) data = data.filter((r) => !isEmptyRow(r));

	// Ancho = la columna más a la derecha con algo (cabecera o datos).
	let width = 0;
	for (let c = headerRaw.length - 1; c >= 0; c--)
		if (headerRaw[c] !== null && headerRaw[c] !== '') {
			width = c + 1;
			break;
		}
	for (const r of data)
		for (let c = r.length - 1; c >= width; c--)
			if (r[c] !== null && r[c] !== '') {
				width = c + 1;
				break;
			}

	let headers: string[];
	if (hasHeader) {
		const norm = normalizeHeaders(Array.from({ length: width }, (_, c) => cellText(headerRaw[c] ?? null)));
		headers = norm.headers;
		if (norm.empty) warnings.push({ kind: 'emptyHeaders', sheet: name, count: norm.empty });
		if (norm.duplicates.length) warnings.push({ kind: 'duplicateHeaders', sheet: name, names: norm.duplicates });
	} else headers = Array.from({ length: width }, (_, c) => colLetter(c));

	const rows = data.map((r) => {
		const out: Cell[] = new Array(width);
		for (let c = 0; c < width; c++) {
			const v = r[c];
			out[c] = v === undefined || v === '' ? null : v;
		}
		return out;
	});
	if (!width) warnings.push({ kind: 'emptySheet', sheet: name });
	return { table: { id: newId(), name, source, headers, types: headers.map(() => 'auto'), rows }, warnings };
}

/** Tabla vacía con columnas dadas (pegar, nueva tabla). */
export function emptyTable(name: string, source: string, headers: string[]): Table {
	return { id: newId(), name, source, headers, types: headers.map(() => 'auto'), rows: [] };
}

/** Copia profunda (filas incluidas). */
export function cloneTable(t: Table): Table {
	return { ...t, headers: [...t.headers], types: [...t.types], rows: t.rows.map((r) => [...r]) };
}

/** Cantidad de celdas no vacías (para resúmenes). */
export function countCells(t: Table): number {
	let n = 0;
	for (const r of t.rows) for (const v of r) if (v !== null) n++;
	return n;
}
