// JSON ⇄ tablas. Entradas aceptadas:
//   [{…}]            → 1 tabla (columnas = unión de claves de TODAS las filas, en orden de aparición)
//   [[…]]            → 1 tabla, la primera fila es la cabecera (según headerRow)
//   [escalar]        → 1 tabla con la columna `value`
//   {"Hoja": [...]}  → varias tablas (todos los valores son arrays)
//   {…}              → 1 tabla de 1 fila
import { flatten, jsonToCell, unflatten } from './nested.ts';
import { gridToTable } from './table.ts';
import { ConvertError, newId, type Cell, type ImportOptions, type ImportWarning, type Table } from './types.ts';

const isPlainObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v);

export function parseJsonText(text: string): unknown {
	try {
		return JSON.parse(text.replace(/^﻿/, ''));
	} catch (e) {
		throw new ConvertError('jsonSyntax', (e as Error).message);
	}
}

/** Array de objetos → tabla con la unión de columnas. */
function objectsToTable(items: unknown[], name: string, source: string, nested: boolean): Table {
	const index = new Map<string, number>();
	const headers: string[] = [];
	const pairsList = items.map((item) => {
		const pairs: [string, Cell][] = isPlainObject(item) ? flatten(item, nested) : [['value', jsonToCell(item)]];
		for (const [k] of pairs)
			if (!index.has(k)) {
				index.set(k, headers.length);
				headers.push(k);
			}
		return pairs;
	});
	const rows = pairsList.map((pairs) => {
		const row: Cell[] = new Array(headers.length).fill(null);
		for (const [k, v] of pairs) row[index.get(k)!] = v;
		return row;
	});
	return { id: newId(), name, source, headers, types: headers.map(() => 'auto'), rows };
}

function arrayToTable(arr: unknown[], name: string, source: string, opts: ImportOptions): { table: Table; warnings: ImportWarning[] } {
	if (arr.length && arr.every((x) => Array.isArray(x))) {
		const grid = (arr as unknown[][]).map((r) => r.map(jsonToCell));
		return gridToTable(grid, name, source, opts);
	}
	return { table: objectsToTable(arr, name, source, opts.nested), warnings: [] };
}

/** Valor JSON ya parseado → tablas. `name` = nombre para la tabla única (normalmente el archivo sin extensión). */
export function jsonToTables(data: unknown, name: string, source: string, opts: ImportOptions): { tables: Table[]; warnings: ImportWarning[] } {
	if (Array.isArray(data)) {
		const { table, warnings } = arrayToTable(data, name, source, opts);
		return { tables: [table], warnings };
	}
	if (isPlainObject(data)) {
		const entries = Object.entries(data);
		if (entries.length && entries.every(([, v]) => Array.isArray(v))) {
			const tables: Table[] = [];
			const warnings: ImportWarning[] = [];
			for (const [sheet, v] of entries) {
				const r = arrayToTable(v as unknown[], sheet, source, opts);
				tables.push(r.table);
				warnings.push(...r.warnings);
			}
			return { tables, warnings };
		}
		return { tables: [objectsToTable([data], name, source, opts.nested)], warnings: [] };
	}
	throw new ConvertError('jsonShape', typeof data);
}

/** Tabla → array de objetos (la clave es la cabecera). */
export function tableToObjects(headers: string[], rows: Cell[][], nested: boolean, dropNulls: boolean): Record<string, unknown>[] {
	return rows.map((r) => unflatten(headers, r, nested, dropNulls));
}

export function stringify(value: unknown, indent: number): string {
	return JSON.stringify(value, null, indent > 0 ? indent : undefined);
}
