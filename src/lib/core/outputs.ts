// Tablas → archivos de salida según formato y modo. Más de un archivo → un .zip.
import { zipSync } from 'fflate';
import { coerce } from './cell.ts';
import { tableToCsv } from './csv.ts';
import { stringify, tableToObjects } from './json.ts';
import { mergeTables, originLabel } from './merge.ts';
import { baseName, uniqueFileName, uniqueName } from './naming.ts';
import { tablesToXlsx, type SheetOut } from './xlsx.ts';
import { ConvertError, MIME, MODES_BY_FORMAT, type Cell, type ExportOptions, type OutputFile, type OutputMode, type Table } from './types.ts';

export const DEFAULT_EXPORT: ExportOptions = {
	format: 'json',
	mode: 'split',
	allText: false,
	nested: false,
	indent: 2,
	dropNulls: false,
	csvDelimiter: ',',
	bom: true,
	styled: true,
	originColumn: '_source'
};

/** Modo válido para el formato (CSV no tiene «por archivo» ni «combinado»). */
export function effectiveMode(opts: Pick<ExportOptions, 'format' | 'mode'>): OutputMode {
	return MODES_BY_FORMAT[opts.format].includes(opts.mode) ? opts.mode : MODES_BY_FORMAT[opts.format][0];
}

/** Aplica los tipos de columna (y «todo como texto») a una tabla. */
export function prepare(t: Table, opts: Pick<ExportOptions, 'allText' | 'format'>): SheetOut {
	const typed = t.types.some((ty) => ty !== 'auto');
	const rows: Cell[][] = typed || opts.allText ? t.rows.map((r) => r.map((v, c) => applyType(v, t.types[c] ?? 'auto', opts))) : t.rows;
	return { name: t.name, headers: t.headers, types: t.types, rows };
}

function applyType(v: Cell, type: Table['types'][number], opts: Pick<ExportOptions, 'allText' | 'format'>): Cell {
	const value = coerce(v, type).value;
	if (!opts.allText) return value;
	// Histórico de excel2json: en JSON todo es string, también las celdas vacías ("").
	if (value === null) return opts.format === 'json' ? '' : null;
	return coerce(value, 'text').value;
}

const enc = new TextEncoder();
const BOM = new Uint8Array([0xef, 0xbb, 0xbf]);

function textFile(name: string, mime: string, text: string, bom = false): OutputFile {
	const body = enc.encode(text);
	if (!bom) return { name, mime, bytes: body };
	const bytes = new Uint8Array(BOM.length + body.length);
	bytes.set(BOM);
	bytes.set(body, BOM.length);
	return { name, mime, bytes };
}

type Group = { base: string; tables: Table[]; label: (t: Table) => string };

/** Reparte las tablas en archivos según el modo. */
function groups(tables: Table[], mode: OutputMode, originColumn: string | null): Group[] {
	const sources = [...new Set(tables.map((t) => t.source))];
	const oneSource = sources.length === 1;
	const commonBase = oneSource ? baseName(sources[0]) : 'excel2json';
	switch (mode) {
		case 'split': {
			const perSource = new Map<string, number>();
			for (const t of tables) perSource.set(t.source, (perSource.get(t.source) ?? 0) + 1);
			// Como excel2json clásico: `archivo.json` si el archivo tiene una tabla, `archivo - Hoja.json` si tiene varias.
			return tables.map((t) => {
				const b = baseName(t.source);
				return { base: (perSource.get(t.source) ?? 0) > 1 ? `${b} - ${t.name}` : b, tables: [t], label: (x) => x.name };
			});
		}
		case 'perFile':
			return sources.map((s) => ({ base: baseName(s), tables: tables.filter((t) => t.source === s), label: (x) => x.name }));
		case 'combined':
			return [{ base: commonBase, tables, label: oneSource ? (x) => x.name : originLabel }];
		case 'merged': {
			// Cada tabla se convierte con SUS tipos de columna antes de unir: si dos tablas difieren en el tipo de una
			// columna, la unida queda `auto` y la conversión se perdería.
			const typed = tables.map((t) => ({ ...t, rows: prepare(t, { allText: false, format: 'json' }).rows }));
			const merged = mergeTables(typed, oneSource && tables.length === 1 ? tables[0].name : commonBase, originColumn);
			return [{ base: oneSource ? commonBase : 'merged', tables: [merged], label: (x) => x.name }];
		}
	}
}

/** Genera los archivos. `onProgress(done, total)` cuenta tablas procesadas. */
export async function buildOutputs(tables: Table[], opts: ExportOptions, onProgress?: (done: number, total: number) => void): Promise<OutputFile[]> {
	if (!tables.length) throw new ConvertError('empty');
	const mode = effectiveMode(opts);
	const gs = groups(tables, mode, opts.originColumn);
	const total = gs.reduce((n, g) => n + g.tables.length, 0);
	let done = 0;
	const tick = () => onProgress?.(++done, total);
	const used = new Set<string>();
	const files: OutputFile[] = [];
	const single = mode === 'split' || mode === 'merged';

	for (const g of gs) {
		const prepared = g.tables.map((t) => ({ t, out: prepare(t, opts) }));
		if (opts.format === 'json') {
			let value: unknown;
			if (single) {
				const { out } = prepared[0];
				value = tableToObjects(out.headers, out.rows, opts.nested, opts.dropNulls);
				tick();
			} else {
				const obj: Record<string, unknown> = {};
				const keys = new Set<string>();
				for (const { t, out } of prepared) {
					obj[uniqueName(g.label(t), keys)] = tableToObjects(out.headers, out.rows, opts.nested, opts.dropNulls);
					tick();
				}
				value = obj;
			}
			files.push(textFile(uniqueFileName(g.base, 'json', used), MIME.json, stringify(value, opts.indent)));
		} else if (opts.format === 'csv') {
			const tsv = opts.csvDelimiter === '\t';
			for (const { out } of prepared) {
				files.push(textFile(uniqueFileName(g.base, tsv ? 'tsv' : 'csv', used), tsv ? MIME.tsv : MIME.csv, tableToCsv(out.headers, out.rows, opts.csvDelimiter), opts.bom));
				tick();
			}
		} else {
			const sheets = prepared.map(({ t, out }) => ({ ...out, name: g.label(t) }));
			const bytes = await tablesToXlsx(sheets, { styled: opts.styled, allText: opts.allText });
			done += sheets.length;
			onProgress?.(done, total);
			files.push({ name: uniqueFileName(g.base, 'xlsx', used), mime: MIME.xlsx, bytes });
		}
	}
	return files;
}

/** Un archivo tal cual; varios → zip (`json.zip`, `xlsx.zip`, `csv.zip`, `tsv.zip`). */
export function bundle(files: OutputFile[], opts: Pick<ExportOptions, 'format' | 'csvDelimiter'>): OutputFile {
	if (files.length === 1) return files[0];
	const entries: Record<string, Uint8Array> = {};
	for (const f of files) entries[f.name] = f.bytes;
	// xlsx ya está comprimido: guardarlo sin volver a comprimir.
	const level = opts.format === 'xlsx' ? 0 : 6;
	const name = opts.format === 'csv' && opts.csvDelimiter === '\t' ? 'tsv' : opts.format;
	return { name: `${name}.zip`, mime: MIME.zip, bytes: zipSync(entries, { level }) };
}
