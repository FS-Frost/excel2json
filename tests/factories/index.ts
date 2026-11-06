// Generadores de datos para los tests: tablas, libros .xlsx en memoria (exceljs, build de Node) y textos.
import ExcelJS from 'exceljs';
import type { Cell, ColType, ImportOptions, Table } from '../../src/lib/core/types.ts';

export const IMPORT: ImportOptions = { headerRow: 1, skipEmptyRows: true, nested: false };

let n = 0;
export function makeTable(headers: string[], rows: Cell[][], extra: Partial<Table> = {}): Table {
	return { id: `t${++n}`, name: 'Hoja1', source: 'datos.xlsx', headers, types: headers.map((): ColType => 'auto'), rows, ...extra };
}

/** Libro con hojas dadas como filas de valores de exceljs (incluida la cabecera). */
export async function makeXlsx(sheets: Record<string, unknown[][]>, setup?: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
	const wb = new ExcelJS.Workbook();
	for (const [name, rows] of Object.entries(sheets)) {
		const ws = wb.addWorksheet(name);
		rows.forEach((r, i) => {
			if (r.length) ws.getRow(i + 1).values = r as ExcelJS.CellValue[];
		});
	}
	setup?.(wb);
	return new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

export async function readXlsx(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.load(bytes.slice().buffer);
	return wb;
}

export const text = (b: Uint8Array) => new TextDecoder().decode(b);
export const bytes = (s: string) => new TextEncoder().encode(s);
