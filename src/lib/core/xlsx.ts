// .xlsx ⇄ tablas con exceljs (cargado bajo demanda: build de navegador en la app/Worker, build de Node en los tests).
import type * as ExcelNS from 'exceljs';
import { cellText, isoToDate, normalizeExcel } from './cell.ts';
import { uniqueSheetName } from './naming.ts';
import { gridToTable } from './table.ts';
import { ConvertError, XLSX_MAX_COLS, XLSX_MAX_ROWS, type Cell, type ColType, type ImportOptions, type ImportWarning, type Table } from './types.ts';

type ExcelJS = typeof ExcelNS;

async function lib(): Promise<ExcelJS> {
	// El build de navegador de exceljs es UMD: según el bundler llega como `default` o como el módulo mismo.
	const m = (await import('exceljs')) as unknown as { default?: ExcelJS } & ExcelJS;
	return m.default ?? m;
}

/** Libro .xlsx → una tabla por hoja (en el orden del libro). `onSheet` informa el avance. */
export async function xlsxToTables(
	bytes: ArrayBuffer | Uint8Array,
	source: string,
	opts: ImportOptions,
	onSheet?: (done: number, total: number, name: string) => void
): Promise<{ tables: Table[]; warnings: ImportWarning[] }> {
	const ExcelJS = await lib();
	const wb = new ExcelJS.Workbook();
	try {
		// exceljs acepta ArrayBuffer; una vista se copia a su propio buffer.
		const buf = bytes instanceof Uint8Array ? bytes.slice().buffer : bytes;
		await wb.xlsx.load(buf);
	} catch (e) {
		throw new ConvertError('xlsxRead', String((e as Error)?.message ?? e));
	}
	const tables: Table[] = [];
	const warnings: ImportWarning[] = [];
	const sheets = wb.worksheets;
	sheets.forEach((ws, i) => {
		onSheet?.(i, sheets.length, ws.name);
		const grid: (Cell | undefined)[][] = [];
		let errors = 0;
		ws.eachRow({ includeEmpty: false }, (row, rn) => {
			const arr: (Cell | undefined)[] = [];
			row.eachCell({ includeEmpty: false }, (cell, cn) => {
				// Celdas combinadas: el valor queda solo en la celda maestra.
				if (cell.isMerged && cell.master.address !== cell.address) return;
				const n = normalizeExcel(cell.value);
				if (n.error) errors++;
				arr[cn - 1] = n.value;
			});
			grid[rn - 1] = arr;
		});
		const r = gridToTable(grid, ws.name, source, opts);
		tables.push(r.table);
		warnings.push(...r.warnings);
		if (errors) warnings.push({ kind: 'cellErrors', sheet: ws.name, count: errors });
	});
	onSheet?.(sheets.length, sheets.length, '');
	return { tables, warnings };
}

export interface SheetOut {
	name: string;
	headers: string[];
	types: ColType[];
	rows: Cell[][];
}

export function checkXlsxLimits(s: SheetOut) {
	if (s.rows.length + 1 > XLSX_MAX_ROWS) throw new ConvertError('xlsxTooManyRows', `${s.name}: ${s.rows.length + 1}`);
	if (s.headers.length > XLSX_MAX_COLS) throw new ConvertError('xlsxTooManyCols', `${s.name}: ${s.headers.length}`);
}

const DATE_FMT = 'yyyy-mm-dd';
const DATETIME_FMT = 'yyyy-mm-dd hh:mm:ss';

/**
 * Tablas → libro .xlsx (una hoja por tabla). Con `allText` todo se escribe como texto; si no, los textos ISO de
 * columnas `auto`/`date` se escriben como fecha de Excel. `styled`: cabecera en negrita con fondo, autofiltro, fila 1
 * congelada y ancho de columna según el contenido.
 */
export async function tablesToXlsx(sheets: SheetOut[], opts: { styled: boolean; allText: boolean }, onSheet?: (done: number, total: number) => void): Promise<Uint8Array> {
	for (const s of sheets) checkXlsxLimits(s);
	const ExcelJS = await lib();
	const wb = new ExcelJS.Workbook();
	wb.creator = 'excel2json';
	wb.created = new Date();
	const used = new Set<string>();
	sheets.forEach((s, i) => {
		onSheet?.(i, sheets.length);
		const ws = wb.addWorksheet(uniqueSheetName(s.name, used), opts.styled ? { views: [{ state: 'frozen', ySplit: 1 }] } : undefined);
		ws.addRow(s.headers);
		const dateCols = s.types.map((t) => !opts.allText && (t === 'date' || t === 'auto'));
		for (const r of s.rows) {
			const values = r.map((v, c) => {
				if (opts.allText) return v === null ? null : cellText(v);
				if (typeof v === 'string' && dateCols[c]) return isoToDate(v) ?? v;
				return v;
			});
			const row = ws.addRow(values);
			values.forEach((v, c) => {
				if (v instanceof Date) row.getCell(c + 1).numFmt = v.getUTCHours() || v.getUTCMinutes() || v.getUTCSeconds() ? DATETIME_FMT : DATE_FMT;
			});
		}
		if (opts.styled && s.headers.length) {
			const head = ws.getRow(1);
			head.font = { bold: true };
			head.eachCell((cell) => {
				cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8ECF4' } };
				cell.border = { bottom: { style: 'thin', color: { argb: 'FF9AA3B5' } } };
			});
			ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: s.headers.length } };
			// Ancho según el contenido (muestra de las primeras 2000 filas para no recorrer libros enormes).
			const sample = s.rows.length > 2000 ? s.rows.slice(0, 2000) : s.rows;
			s.headers.forEach((h, c) => {
				let w = h.length + 2;
				for (const r of sample) {
					const v = r[c];
					const len = v === null ? 0 : typeof v === 'string' && dateCols[c] && isoToDate(v) ? v.length : cellText(v).length;
					if (len + 1 > w) w = len + 1;
				}
				ws.getColumn(c + 1).width = Math.min(60, Math.max(8, w));
			});
		}
	});
	onSheet?.(sheets.length, sheets.length);
	const buf = await wb.xlsx.writeBuffer();
	return new Uint8Array(buf as ArrayBuffer);
}
