// Modelo único de la app: toda entrada (xlsx, csv, tsv, json) se convierte en tablas y toda salida se escribe desde ellas.
// Puro (sin DOM): lo usan el Worker, la página y los tests de Bun.

/** Valor de una celda. Las fechas viajan como texto ISO 8601 local (`2024-01-31` o `2024-01-31T08:30:00`). */
export type Cell = string | number | boolean | null;

/** Tipo declarado de una columna. `auto` respeta el valor de cada celda; el resto se aplica al exportar. */
export type ColType = 'auto' | 'text' | 'number' | 'boolean' | 'date';
export const COL_TYPES: readonly ColType[] = ['auto', 'text', 'number', 'boolean', 'date'];

export type SourceKind = 'xlsx' | 'csv' | 'tsv' | 'json';

export interface Table {
	id: string;
	/** Nombre de la hoja (o del archivo si no tiene hojas: CSV, JSON de una tabla). */
	name: string;
	/** Nombre del archivo de origen, con extensión. */
	source: string;
	headers: string[];
	types: ColType[];
	/** Filas de datos (sin la cabecera). Cada fila mide exactamente `headers.length`. */
	rows: Cell[][];
}

/** Aviso de importación (el texto lo pone la interfaz según el idioma). */
export type ImportWarning =
	| { kind: 'cellErrors'; sheet: string; count: number }
	| { kind: 'emptyHeaders'; sheet: string; count: number }
	| { kind: 'duplicateHeaders'; sheet: string; names: string[] }
	| { kind: 'emptySheet'; sheet: string }
	| { kind: 'decoding'; encoding: string };

export interface SourceFile {
	id: string;
	name: string;
	kind: SourceKind;
	size: number;
	tables: Table[];
	warnings: ImportWarning[];
}

export interface ImportOptions {
	/** Fila de la cabecera (1 = primera). 0 = sin cabecera: las columnas se llaman A, B, C… */
	headerRow: number;
	/** Omitir filas completamente vacías en medio de los datos (las del final se omiten siempre). */
	skipEmptyRows: boolean;
	/** JSON → tabla: aplanar objetos anidados a columnas `a.b.c` (si no, van como texto JSON). */
	nested: boolean;
}

export type OutputFormat = 'json' | 'xlsx' | 'csv';

/**
 * Cómo se agrupan las tablas en la salida:
 * - split: un archivo por tabla.
 * - perFile: un archivo por archivo de entrada (JSON `{hoja: filas}`, xlsx con sus hojas).
 * - combined: un solo archivo con todas las tablas (JSON `{hoja: filas}`, xlsx multihoja).
 * - merged: todas las filas unidas en una sola tabla (unión de columnas).
 */
export type OutputMode = 'split' | 'perFile' | 'combined' | 'merged';

export const MODES_BY_FORMAT: Record<OutputFormat, readonly OutputMode[]> = {
	json: ['split', 'perFile', 'combined', 'merged'],
	xlsx: ['combined', 'perFile', 'split', 'merged'],
	csv: ['split', 'merged']
};

export type CsvDelimiter = ',' | ';' | '\t';

export interface ExportOptions {
	format: OutputFormat;
	mode: OutputMode;
	/** Todo como texto (comportamiento histórico de excel2json). */
	allText: boolean;
	/** JSON: columnas `a.b.c` → objetos anidados. */
	nested: boolean;
	/** JSON: espacios de indentación (0 = compacto). */
	indent: number;
	/** JSON: omitir claves con valor null. */
	dropNulls: boolean;
	csvDelimiter: CsvDelimiter;
	/** CSV: BOM UTF-8 (Excel lo necesita para leer acentos). */
	bom: boolean;
	/** xlsx: cabecera con formato, autofiltro, fila congelada y anchos automáticos. */
	styled: boolean;
	/** Unir: nombre de la columna con el origen de cada fila (null = sin columna). */
	originColumn: string | null;
}

export interface OutputFile {
	name: string;
	mime: string;
	bytes: Uint8Array;
}

export const MIME = {
	json: 'application/json',
	xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	csv: 'text/csv',
	tsv: 'text/tab-separated-values',
	zip: 'application/zip'
} as const;

/** Límites de una hoja de Excel. */
export const XLSX_MAX_ROWS = 1_048_576;
export const XLSX_MAX_COLS = 16_384;

export class ConvertError extends Error {
	constructor(
		readonly code: 'xlsxTooManyRows' | 'xlsxTooManyCols' | 'jsonShape' | 'jsonSyntax' | 'xlsxRead' | 'unsupported' | 'empty',
		readonly detail = ''
	) {
		super(detail ? `${code}: ${detail}` : code);
	}
}

let seq = 0;
/** Id corto y único en la sesión (no necesita ser criptográfico). */
export function newId(prefix = 't') {
	seq = (seq + 1) % 1e9;
	return `${prefix}${Date.now().toString(36)}${seq.toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}
