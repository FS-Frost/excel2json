// CSV/TSV ⇄ tablas (papaparse). Lectura: delimitador detectado; con `;` los números usan coma decimal (CSV de Excel
// en español: "1,5;2"), con `,` o tab, punto. Escritura: RFC 4180, CRLF, BOM opcional; con `;` la coma decimal.
import Papa from 'papaparse';
import { cellText, inferText } from './cell.ts';
import { gridToTable } from './table.ts';
import type { Cell, CsvDelimiter, ImportOptions, ImportWarning, Table } from './types.ts';

/** Bytes → texto: UTF-8 (con o sin BOM); si no es UTF-8 válido, Windows-1252 (CSV antiguos de Excel). */
export function decodeText(bytes: Uint8Array): { text: string; encoding: 'utf-8' | 'windows-1252' } {
	try {
		return { text: new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, ''), encoding: 'utf-8' };
	} catch {
		return { text: new TextDecoder('windows-1252').decode(bytes), encoding: 'windows-1252' };
	}
}

export function parseDelimited(text: string, delimiter?: CsvDelimiter): { rows: string[][]; delimiter: string } {
	const res = Papa.parse<string[]>(text.replace(/^﻿/, ''), {
		delimiter: delimiter ?? '',
		delimitersToGuess: [',', ';', '\t', '|'],
		skipEmptyLines: false
	});
	return { rows: res.data, delimiter: res.meta.delimiter };
}

/**
 * Texto CSV/TSV → tabla. `typed` infiere números y booleanos (si no, todo texto).
 * La cabecera nunca se infiere (siempre texto).
 */
export function csvToTable(
	text: string,
	name: string,
	source: string,
	opts: ImportOptions,
	delimiter?: CsvDelimiter
): { table: Table; warnings: ImportWarning[]; delimiter: string } {
	const parsed = parseDelimited(text, delimiter);
	const decimal = parsed.delimiter === ';' ? ',' : '.';
	const headerIdx = opts.headerRow - 1;
	const grid: Cell[][] = parsed.rows.map((row, r) => row.map((raw) => (r === headerIdx ? raw : inferText(raw, decimal))));
	const { table, warnings } = gridToTable(grid, name, source, opts);
	return { table, warnings, delimiter: parsed.delimiter };
}

/** Tabla → texto CSV/TSV (cabecera incluida). Sin BOM: lo agrega quien escribe el archivo. */
export function tableToCsv(headers: string[], rows: Cell[][], delimiter: CsvDelimiter): string {
	const decimal = delimiter === ';' ? ',' : '.';
	const data = [headers, ...rows.map((r) => r.map((v) => cellText(v, decimal)))];
	return Papa.unparse(data, { delimiter, newline: '\r\n', quotes: false, escapeFormulae: false });
}
