import { describe, expect, test } from 'bun:test';
import { unzipSync } from 'fflate';
import { cellText, coerce, dateToIso, fits, inferText, isoToDate, normalizeExcel, parseTyped } from '../src/lib/core/cell.ts';
import { csvToTable, decodeText, tableToCsv } from '../src/lib/core/csv.ts';
import { jsonToTables, parseJsonText, tableToObjects } from '../src/lib/core/json.ts';
import { mergeTables } from '../src/lib/core/merge.ts';
import { flatten, unflatten } from '../src/lib/core/nested.ts';
import { baseName, colLetter, normalizeHeaders, safeSheetName, uniqueFileName, uniqueSheetName } from '../src/lib/core/naming.ts';
import { bundle, buildOutputs, DEFAULT_EXPORT, effectiveMode, prepare } from '../src/lib/core/outputs.ts';
import { readSource, sourceKind } from '../src/lib/core/read.ts';
import { gridToTable } from '../src/lib/core/table.ts';
import { ConvertError, type ExportOptions } from '../src/lib/core/types.ts';
import { tablesToXlsx, xlsxToTables } from '../src/lib/core/xlsx.ts';
import { bytes, IMPORT, makeTable, makeXlsx, readXlsx, text } from './factories/index.ts';

describe('cell', () => {
	test('normalizes exceljs values', () => {
		expect(normalizeExcel(null).value).toBeNull();
		expect(normalizeExcel('').value).toBeNull();
		expect(normalizeExcel(3.5).value).toBe(3.5);
		expect(normalizeExcel(true).value).toBe(true);
		expect(normalizeExcel({ formula: 'A1+1', result: 4 }).value).toBe(4);
		expect(normalizeExcel({ sharedFormula: 'A1', result: 'x' }).value).toBe('x');
		expect(normalizeExcel({ formula: 'NOW()' }).value).toBeNull();
		expect(normalizeExcel({ richText: [{ text: 'Ho' }, { text: 'la' }] }).value).toBe('Hola');
		expect(normalizeExcel({ text: 'sitio', hyperlink: 'https://x.cl' }).value).toBe('sitio');
		expect(normalizeExcel({ error: '#N/A' })).toEqual({ value: null, error: true });
		expect(normalizeExcel({ formula: 'X', result: { error: '#DIV/0!' } })).toEqual({ value: null, error: true });
		expect(normalizeExcel(new Date(Date.UTC(2024, 0, 31))).value).toBe('2024-01-31');
		expect(normalizeExcel(new Date(Date.UTC(2024, 0, 31, 8, 5, 9))).value).toBe('2024-01-31T08:05:09');
	});
	test('iso dates round trip and reject impossible dates', () => {
		expect(dateToIso(isoToDate('2024-02-29')!)).toBe('2024-02-29');
		expect(dateToIso(isoToDate('2024-02-29T23:59:00')!)).toBe('2024-02-29T23:59:00');
		expect(isoToDate('2023-02-29')).toBeNull();
		expect(isoToDate('2024-13-01')).toBeNull();
		expect(isoToDate('31/01/2024')).toBeNull();
	});
	test('infers only clean numbers', () => {
		expect(inferText('42')).toBe(42);
		expect(inferText('-0.5')).toBe(-0.5);
		expect(inferText('1e3')).toBe(1000);
		expect(inferText('00123')).toBe('00123');
		expect(inferText('+56 9 1234')).toBe('+56 9 1234');
		expect(inferText('1,5')).toBe('1,5');
		expect(inferText('1,5', ',')).toBe(1.5);
		expect(inferText('1.000,5', ',')).toBe('1.000,5');
		expect(inferText('TRUE')).toBe(true);
		expect(inferText('false')).toBe(false);
		expect(inferText('  ')).toBeNull();
	});
	test('typed input', () => {
		expect(parseTyped("'42", 'auto')).toBe('42');
		expect(parseTyped('42', 'text')).toBe('42');
		expect(parseTyped('', 'text')).toBeNull();
		expect(parseTyped('42', 'number')).toBe(42);
	});
	test('coercion by column type', () => {
		expect(coerce('12', 'number')).toEqual({ value: 12, ok: true });
		expect(coerce('1,5', 'number')).toEqual({ value: 1.5, ok: true });
		expect(coerce('abc', 'number')).toEqual({ value: 'abc', ok: false });
		expect(coerce('sí', 'boolean').value).toBe(true);
		expect(coerce('No', 'boolean').value).toBe(false);
		expect(coerce('31/01/2024', 'date').value).toBe('2024-01-31');
		expect(coerce(45322, 'date').value).toBe('2024-01-31');
		expect(coerce(5, 'text').value).toBe('5');
		expect(fits('x', 'number')).toBe(false);
		expect(fits(null, 'number')).toBe(true);
		expect(cellText(0.1 + 0.2)).toBe('0.3');
		expect(cellText(1.5, ',')).toBe('1,5');
	});
});

describe('naming', () => {
	test('letters, base names, headers', () => {
		expect([0, 25, 26, 701, 702].map(colLetter)).toEqual(['A', 'Z', 'AA', 'ZZ', 'AAA']);
		expect(baseName('datos.2024.xlsx')).toBe('datos.2024');
		expect(baseName('sin_ext')).toBe('sin_ext');
		const h = normalizeHeaders(['id', '', 'id', 'id_2', 'id']);
		expect(h.headers).toEqual(['id', 'B', 'id_3', 'id_2', 'id_4']);
		expect(h.empty).toBe(1);
		expect(h.duplicates).toEqual(['id']);
	});
	test('sheet and file names', () => {
		expect(safeSheetName("'a/b:c*'")).toBe('a b c');
		expect(safeSheetName('x'.repeat(40)).length).toBe(31);
		const used = new Set<string>();
		expect(uniqueSheetName('Hoja', used)).toBe('Hoja');
		expect(uniqueSheetName('hoja', used)).toBe('hoja (2)');
		const long = new Set<string>();
		uniqueSheetName('y'.repeat(31), long);
		expect(uniqueSheetName('y'.repeat(31), long)).toBe(`${'y'.repeat(27)} (2)`);
		const files = new Set<string>();
		expect(uniqueFileName('a', 'json', files)).toBe('a.json');
		expect(uniqueFileName('a', 'json', files)).toBe('a (2).json');
		expect(uniqueFileName('a/b', 'csv', files)).toBe('a_b.csv');
	});
});

describe('table', () => {
	test('header row, empty rows, width', () => {
		const grid = [['título'], [], ['id', 'nombre', null], [1, 'a'], [], [2, 'b', null, 'extra'], [], []];
		const { table, warnings } = gridToTable(grid, 'H', 'f.xlsx', { ...IMPORT, headerRow: 3 });
		expect(table.headers).toEqual(['id', 'nombre', 'C', 'D']);
		expect(table.rows).toEqual([
			[1, 'a', null, null],
			[2, 'b', null, 'extra']
		]);
		expect(warnings).toEqual([{ kind: 'emptyHeaders', sheet: 'H', count: 2 }]);
		const keep = gridToTable(grid, 'H', 'f', { ...IMPORT, headerRow: 3, skipEmptyRows: false }).table;
		expect(keep.rows.length).toBe(3);
	});
	test('no header and sparse rows', () => {
		const sparse: (number | undefined)[][] = [];
		sparse[1] = [];
		sparse[1][2] = 7;
		const { table } = gridToTable(sparse, 'H', 'f', { ...IMPORT, headerRow: 0, skipEmptyRows: false });
		expect(table.headers).toEqual(['A', 'B', 'C']);
		expect(table.rows).toEqual([
			[null, null, null],
			[null, null, 7]
		]);
	});
	test('empty sheet warns', () => {
		expect(gridToTable([], 'V', 'f', IMPORT).warnings).toEqual([{ kind: 'emptySheet', sheet: 'V' }]);
	});
});

describe('nested', () => {
	test('flatten and unflatten round trip', () => {
		const obj = { a: { b: 1, c: { d: 'x' } }, e: [1, 2], f: {}, g: null };
		const pairs = flatten(obj, true);
		expect(pairs).toEqual([
			['a.b', 1],
			['a.c.d', 'x'],
			['e', '[1,2]'],
			['f', '{}'],
			['g', null]
		]);
		const back = unflatten(
			pairs.map(([k]) => k),
			pairs.map(([, v]) => v),
			true
		);
		expect(back).toEqual({ a: { b: 1, c: { d: 'x' } }, e: '[1,2]', f: '{}', g: null });
		expect(flatten(obj, false)[0]).toEqual(['a', '{"b":1,"c":{"d":"x"}}']);
	});
	test('conflicts stay flat, nulls can be dropped', () => {
		expect(unflatten(['a', 'a.b'], [1, 2], true)).toEqual({ a: 1, 'a.b': 2 });
		expect(unflatten(['a.b', 'a'], [2, 1], true)).toEqual({ a: 1 });
		expect(unflatten(['x.', 'y'], [1, null], true, true)).toEqual({ 'x.': 1 });
	});
});

describe('csv', () => {
	test('detects delimiter and decimal comma', () => {
		const t = csvToTable('id;precio;nota\r\n1;1,5;00123\r\n2;3;\r\n', 'p', 'p.csv', IMPORT).table;
		expect(t.headers).toEqual(['id', 'precio', 'nota']);
		expect(t.rows).toEqual([
			[1, 1.5, '00123'],
			[2, 3, null]
		]);
		const c = csvToTable('a,b\n"x, y","línea\nnueva"\n', 'p', 'p.csv', IMPORT).table;
		expect(c.rows).toEqual([['x, y', 'línea\nnueva']]);
		const tab = csvToTable('a\tb\n1\t2', 'p', 'p.tsv', IMPORT, '\t').table;
		expect(tab.rows).toEqual([[1, 2]]);
	});
	test('writes with quoting and decimal comma for ;', () => {
		expect(tableToCsv(['a', 'b'], [['x,y', 1.5]], ',')).toBe('a,b\r\n"x,y",1.5');
		expect(tableToCsv(['a', 'b'], [['x"y', 1.5]], ';')).toBe('a;b\r\n"x""y";1,5');
	});
	test('decodes utf-8 with BOM and falls back to windows-1252', () => {
		expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x61]))).toEqual({ text: 'a', encoding: 'utf-8' });
		expect(decodeText(new Uint8Array([0x61, 0xf1]))).toEqual({ text: 'añ', encoding: 'windows-1252' });
	});
});

describe('json', () => {
	test('array of objects: union of keys from every row', () => {
		const { tables } = jsonToTables([{ a: 1 }, { b: 2, a: 3 }, { c: { d: true } }], 'x', 'x.json', IMPORT);
		expect(tables[0].headers).toEqual(['a', 'b', 'c']);
		expect(tables[0].rows).toEqual([
			[1, null, null],
			[3, 2, null],
			[null, null, '{"d":true}']
		]);
		const nested = jsonToTables([{ c: { d: true } }], 'x', 'x.json', { ...IMPORT, nested: true }).tables[0];
		expect(nested.headers).toEqual(['c.d']);
	});
	test('other shapes', () => {
		expect(jsonToTables([['id', 'n'], [1, 'a']], 'x', 'x.json', IMPORT).tables[0].rows).toEqual([[1, 'a']]);
		expect(jsonToTables([1, 'b'], 'x', 'x.json', IMPORT).tables[0]).toMatchObject({ headers: ['value'], rows: [[1], ['b']] });
		const multi = jsonToTables({ Ventas: [{ a: 1 }], Stock: [{ b: 2 }] }, 'x', 'x.json', IMPORT).tables;
		expect(multi.map((t) => t.name)).toEqual(['Ventas', 'Stock']);
		expect(jsonToTables({ a: 1, b: 'z' }, 'x', 'x.json', IMPORT).tables[0].rows).toEqual([[1, 'z']]);
		expect(() => jsonToTables(5, 'x', 'x.json', IMPORT)).toThrow(ConvertError);
		expect(() => parseJsonText('{')).toThrow(ConvertError);
	});
	test('table to objects', () => {
		expect(tableToObjects(['a.b', 'c'], [[1, null]], true, true)).toEqual([{ a: { b: 1 } }]);
	});
});

describe('merge', () => {
	test('union of columns with origin', () => {
		const a = makeTable(['id', 'n'], [[1, 'a']], { name: 'A', source: 'uno.xlsx' });
		const b = makeTable(['n', 'z'], [['b', true]], { name: 'B', source: 'uno.xlsx', types: ['text', 'auto'] });
		const m = mergeTables([a, b], 'm', '_source');
		expect(m.headers).toEqual(['_source', 'id', 'n', 'z']);
		expect(m.types).toEqual(['text', 'auto', 'auto', 'auto']);
		expect(m.rows).toEqual([
			['uno › A', 1, 'a', null],
			['uno › B', null, 'b', true]
		]);
		const clash = mergeTables([makeTable(['_source'], [[1]])], 'm', '_source');
		expect(clash.headers).toEqual(['_source_2', '_source']);
		expect(mergeTables([a], 'm', null).headers).toEqual(['id', 'n']);
	});
});

describe('xlsx', () => {
	test('reads every sheet with types, formulas, merges and errors', async () => {
		const file = await makeXlsx(
			{
				Datos: [
					['id', 'fecha', 'total', 'ok', 'link'],
					[1, new Date(Date.UTC(2024, 0, 31)), { formula: 'A2*2', result: 2 }, true, { text: 'web', hyperlink: 'https://x.cl' }],
					[2, new Date(Date.UTC(2024, 1, 1, 10, 30)), { error: '#N/A' }, false, null]
				],
				Vacía: []
			},
			(wb) => {
				const ws = wb.getWorksheet('Datos')!;
				ws.getCell('A5').value = 'combinada';
				ws.mergeCells('A5:B5');
			}
		);
		const { tables, warnings } = await xlsxToTables(file, 'libro.xlsx', IMPORT);
		expect(tables.map((t) => t.name)).toEqual(['Datos', 'Vacía']);
		expect(tables[0].rows).toEqual([
			[1, '2024-01-31', 2, true, 'web'],
			[2, '2024-02-01T10:30:00', null, false, null],
			['combinada', null, null, null, null]
		]);
		expect(warnings).toContainEqual({ kind: 'cellErrors', sheet: 'Datos', count: 1 });
		expect(warnings).toContainEqual({ kind: 'emptySheet', sheet: 'Vacía' });
	});
	test('writes styled workbook with real dates and reads it back', async () => {
		const sheet = { name: 'S/1', headers: ['id', 'cuando', 'txt'], types: ['auto', 'auto', 'text'] as const, rows: [[1, '2024-01-31', '2024-01-31']] };
		const out = await tablesToXlsx([{ ...sheet, types: [...sheet.types] }, { ...sheet, types: [...sheet.types] }], { styled: true, allText: false });
		const wb = await readXlsx(out);
		expect(wb.worksheets.map((w) => w.name)).toEqual(['S 1', 'S 1 (2)']);
		const ws = wb.worksheets[0];
		expect(ws.getCell('B2').value).toBeInstanceOf(Date);
		expect(ws.getCell('C2').value).toBe('2024-01-31');
		expect(ws.getRow(1).font?.bold).toBe(true);
		expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
		const back = await xlsxToTables(out, 'x.xlsx', IMPORT);
		expect(back.tables[0].rows).toEqual([[1, '2024-01-31', '2024-01-31']]);
	});
	test('all text mode', async () => {
		const out = await tablesToXlsx([{ name: 'a', headers: ['n'], types: ['auto'], rows: [[1.5], [null]] }], { styled: false, allText: true });
		const ws = (await readXlsx(out)).worksheets[0];
		expect(ws.getCell('A2').value).toBe('1.5');
		expect(ws.getCell('A3').value).toBeNull();
	});
	test('limits', async () => {
		const huge = { name: 'a', headers: Array.from({ length: 16_385 }, (_, i) => `c${i}`), types: [], rows: [] };
		await expect(tablesToXlsx([huge], { styled: false, allText: false })).rejects.toThrow(ConvertError);
	});
});

describe('read', () => {
	test('kind detection', () => {
		expect(sourceKind('a.XLSX')).toBe('xlsx');
		expect(sourceKind('a.xls')).toBeNull();
		expect(sourceKind('a.pdf', bytes('%PDF-1.4'))).toBeNull();
		expect(sourceKind('notas.txt', bytes('a;b'))).toBe('csv');
		expect(sourceKind('pegado', bytes(' [1]'))).toBe('json');
		expect(sourceKind('pegado', bytes('a,b'))).toBe('csv');
		expect(sourceKind('x', new Uint8Array([0x50, 0x4b, 3, 4]))).toBe('xlsx');
	});
	test('reads files', async () => {
		const j = await readSource('a.b.json', bytes('{"Uno":[{"x":1}],"Dos":[{"y":2}]}'), IMPORT);
		expect(j.kind).toBe('json');
		expect(j.tables.map((t) => [t.name, t.source])).toEqual([
			['Uno', 'a.b.json'],
			['Dos', 'a.b.json']
		]);
		const c = await readSource('lista.csv', bytes('a;b\n1;2'), IMPORT);
		expect(c.tables[0]).toMatchObject({ name: 'lista', rows: [[1, 2]] });
		await expect(readSource('x.xls', bytes(''), IMPORT)).rejects.toThrow(ConvertError);
	});
});

describe('outputs', () => {
	const opts = (o: Partial<ExportOptions>): ExportOptions => ({ ...DEFAULT_EXPORT, ...o });
	const a = makeTable(['id', 'n'], [[1, 'a']], { name: 'A', source: 'libro.xlsx' });
	const b = makeTable(['id', 'z'], [[2, null]], { name: 'B', source: 'libro.xlsx' });
	const c = makeTable(['k'], [[true]], { name: 'otro', source: 'otro.csv' });

	test('json split keeps the classic names', async () => {
		const files = await buildOutputs([a, b, c], opts({ format: 'json', mode: 'split' }));
		expect(files.map((f) => f.name)).toEqual(['libro - A.json', 'libro - B.json', 'otro.json']);
		expect(JSON.parse(text(files[0].bytes))).toEqual([{ id: 1, n: 'a' }]);
	});
	test('json per file, combined and merged', async () => {
		const per = await buildOutputs([a, b, c], opts({ mode: 'perFile' }));
		expect(per.map((f) => f.name)).toEqual(['libro.json', 'otro.json']);
		expect(JSON.parse(text(per[0].bytes))).toEqual({ A: [{ id: 1, n: 'a' }], B: [{ id: 2, z: null }] });
		const comb = await buildOutputs([a, c], opts({ mode: 'combined', indent: 0 }));
		expect(text(comb[0].bytes)).toBe('{"libro › A":[{"id":1,"n":"a"}],"otro":[{"k":true}]}');
		expect(comb[0].name).toBe('excel2json.json');
		const merged = await buildOutputs([a, b], opts({ mode: 'merged', originColumn: null }));
		expect(merged[0].name).toBe('libro.json');
		expect(JSON.parse(text(merged[0].bytes))).toEqual([
			{ id: 1, n: 'a', z: null },
			{ id: 2, n: null, z: null }
		]);
	});
	test('merged keeps each table column type conversion', async () => {
		const x = makeTable(['monto'], [['1,5']], { name: 'X', types: ['number'] });
		const y = makeTable(['monto'], [['2']], { name: 'Y' });
		const files = await buildOutputs([x, y], opts({ mode: 'merged', originColumn: null }));
		expect(JSON.parse(text(files[0].bytes))).toEqual([{ monto: 1.5 }, { monto: '2' }]);
	});
	test('all text and column types', async () => {
		const t = makeTable(['n', 'm'], [[1, '2'], [null, 'x']], { types: ['auto', 'number'] });
		expect(prepare(t, { allText: false, format: 'json' }).rows).toEqual([
			[1, 2],
			[null, 'x']
		]);
		const files = await buildOutputs([t], opts({ allText: true }));
		expect(JSON.parse(text(files[0].bytes))).toEqual([
			{ n: '1', m: '2' },
			{ n: '', m: 'x' }
		]);
	});
	test('csv with BOM, tsv, zip', async () => {
		const files = await buildOutputs([a, b], opts({ format: 'csv', mode: 'combined' }));
		expect(effectiveMode({ format: 'csv', mode: 'combined' })).toBe('split');
		expect(files.map((f) => f.name)).toEqual(['libro - A.csv', 'libro - B.csv']);
		expect([...files[0].bytes.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
		const zip = bundle(files, { format: 'csv', csvDelimiter: ',' });
		expect(zip.name).toBe('csv.zip');
		expect(Object.keys(unzipSync(zip.bytes))).toEqual(['libro - A.csv', 'libro - B.csv']);
		const tsv = await buildOutputs([c], opts({ format: 'csv', csvDelimiter: '\t', bom: false }));
		expect(tsv[0].name).toBe('otro.tsv');
		expect(text(tsv[0].bytes)).toBe('k\r\ntrue');
	});
	test('xlsx combined is one workbook with a sheet per table', async () => {
		const files = await buildOutputs([a, b, c], opts({ format: 'xlsx', mode: 'combined' }));
		expect(files.length).toBe(1);
		const wb = await readXlsx(files[0].bytes);
		expect(wb.worksheets.map((w) => w.name)).toEqual(['libro › A', 'libro › B', 'otro']);
		const per = await buildOutputs([a, b, c], opts({ format: 'xlsx', mode: 'perFile' }));
		expect(per.map((f) => f.name)).toEqual(['libro.xlsx', 'otro.xlsx']);
	});
	test('nothing to export', async () => {
		await expect(buildOutputs([], DEFAULT_EXPORT)).rejects.toThrow(ConvertError);
	});
});
