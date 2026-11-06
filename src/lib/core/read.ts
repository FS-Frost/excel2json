// Archivo (nombre + bytes) → SourceFile con sus tablas, según la extensión (o el contenido si no la tiene).
import { csvToTable, decodeText } from './csv.ts';
import { jsonToTables, parseJsonText } from './json.ts';
import { baseName, extension } from './naming.ts';
import { xlsxToTables } from './xlsx.ts';
import { ConvertError, newId, type ImportOptions, type SourceFile, type SourceKind } from './types.ts';

export const ACCEPT = '.xlsx,.xlsm,.csv,.tsv,.tab,.txt,.json,application/json,text/csv,text/tab-separated-values,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const isZip = (b: Uint8Array) => b.length > 3 && b[0] === 0x50 && b[1] === 0x4b && b[2] === 0x03 && b[3] === 0x04;

/** Tipo de entrada por extensión; sin extensión conocida, por contenido (zip → xlsx, `[`/`{` → json, si no CSV). */
export function sourceKind(name: string, bytes?: Uint8Array): SourceKind | null {
	const ext = extension(name);
	if (ext === 'xlsx' || ext === 'xlsm') return 'xlsx';
	if (ext === 'csv') return 'csv';
	if (ext === 'tsv' || ext === 'tab') return 'tsv';
	if (ext === 'json') return 'json';
	// Una extensión conocida que no es de las nuestras (.xls, .pdf, .docx…) no se adivina por contenido.
	if (ext && ext !== 'txt') return null;
	if (!bytes) return ext === 'txt' ? 'csv' : null;
	if (isZip(bytes)) return 'xlsx';
	const head = new TextDecoder().decode(bytes.subarray(0, 64)).replace(/^﻿/, '').trimStart();
	if (head.startsWith('[') || head.startsWith('{')) return 'json';
	return 'csv';
}

export async function readSource(
	name: string,
	bytes: Uint8Array,
	opts: ImportOptions,
	onProgress?: (done: number, total: number, label: string) => void
): Promise<SourceFile> {
	const kind = sourceKind(name, bytes);
	if (!kind) throw new ConvertError('unsupported', name);
	const base: Omit<SourceFile, 'tables' | 'warnings'> = { id: newId('f'), name, kind, size: bytes.length };
	if (kind === 'xlsx') {
		const { tables, warnings } = await xlsxToTables(bytes, name, opts, onProgress);
		return { ...base, tables, warnings };
	}
	const { text, encoding } = decodeText(bytes);
	const warnings: SourceFile['warnings'] = encoding === 'utf-8' ? [] : [{ kind: 'decoding', encoding }];
	if (kind === 'json') {
		const r = jsonToTables(parseJsonText(text), baseName(name), name, opts);
		return { ...base, tables: r.tables, warnings: [...warnings, ...r.warnings] };
	}
	const r = csvToTable(text, baseName(name), name, opts, kind === 'tsv' ? '\t' : undefined);
	return { ...base, tables: [r.table], warnings: [...warnings, ...r.warnings] };
}
