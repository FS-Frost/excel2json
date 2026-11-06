import { describe, expect, test } from 'bun:test';
import { parseHTMLTable, parseTSV, toTSV } from '../src/lib/edit/clipboard.ts';
import { findAll, replaceValue, DEFAULT_FIND } from '../src/lib/edit/find.ts';
import { History } from '../src/lib/edit/history.ts';
import { Axis } from '../src/lib/edit/layout.ts';
import { apply, applyAll, freeHeader, invert, moveOrder, type Op } from '../src/lib/edit/ops.ts';
import { compareCells, distinctValues, fold, sortOrder, visibleRows } from '../src/lib/edit/view.ts';
import { cloneTable } from '../src/lib/core/table.ts';
import { makeTable } from './factories/index.ts';

const base = () =>
	makeTable(
		['id', 'nombre', 'n'],
		[
			[1, 'Ana', 10],
			[2, 'Bruno', null],
			[3, 'Canción', 2]
		]
	);

/** Aplica, deshace con la inversa y comprueba que la tabla vuelve exactamente a su estado original. */
function roundTrip(ops: Op[]) {
	const t = base();
	const before = cloneTable(t);
	const inverse = applyAll(t, ops);
	const after = cloneTable(t);
	applyAll(t, inverse);
	expect({ h: t.headers, ty: t.types, r: t.rows, n: t.name }).toEqual({ h: before.headers, ty: before.types, r: before.rows, n: before.name });
	return after;
}

describe('ops', () => {
	test('cells', () => {
		const after = roundTrip([{ t: 'cells', changes: [{ r: 0, c: 1, v: 'Ani' }, { r: 0, c: 1, v: 'Anita' }] }]);
		expect(after.rows[0][1]).toBe('Anita');
	});
	test('insert and delete rows', () => {
		expect(roundTrip([{ t: 'insertRows', at: 1, rows: [[9], []] }]).rows.map((r) => r[0])).toEqual([1, 9, null, 2, 3]);
		expect(roundTrip([{ t: 'deleteRows', at: [2, 0] }]).rows).toEqual([[2, 'Bruno', null]]);
	});
	test('insert and delete columns', () => {
		const ins = roundTrip([{ t: 'insertCols', at: 1, headers: ['x'], types: ['text'], cols: [['a', 'b']] }]);
		expect(ins.headers).toEqual(['id', 'x', 'nombre', 'n']);
		expect(ins.rows.map((r) => r[1])).toEqual(['a', 'b', null]);
		const del = roundTrip([{ t: 'deleteCols', at: [0, 2] }]);
		expect(del.headers).toEqual(['nombre']);
		expect(del.rows).toEqual([['Ana'], ['Bruno'], ['Canción']]);
	});
	test('rename, type, sheet name', () => {
		const a = roundTrip([
			{ t: 'header', c: 0, name: 'código' },
			{ t: 'type', c: 2, type: 'number' },
			{ t: 'name', name: 'Clientes' }
		]);
		expect([a.headers[0], a.types[2], a.name]).toEqual(['código', 'number', 'Clientes']);
	});
	test('permutations and moves', () => {
		expect(invert([2, 0, 1])).toEqual([1, 2, 0]);
		expect(moveOrder(5, 1, 2, 3)).toEqual([0, 3, 4, 1, 2]);
		expect(moveOrder(5, 3, 1, 0)).toEqual([3, 0, 1, 2, 4]);
		expect(roundTrip([{ t: 'permuteRows', order: [2, 0, 1] }]).rows.map((r) => r[0])).toEqual([3, 1, 2]);
		expect(roundTrip([{ t: 'permuteCols', order: [2, 0, 1] }]).headers).toEqual(['n', 'id', 'nombre']);
	});
	test('compound step undoes in reverse order', () => {
		const after = roundTrip([
			{ t: 'insertRows', at: 3, rows: [[]] },
			{ t: 'insertCols', at: 3, headers: ['extra'], types: ['auto'] },
			{ t: 'cells', changes: [{ r: 3, c: 3, v: 'pegado' }] }
		]);
		expect(after.rows[3]).toEqual([null, null, null, 'pegado']);
	});
	test('out of range cells are ignored', () => {
		const t = base();
		const inv = apply(t, { t: 'cells', changes: [{ r: 99, c: 0, v: 1 }, { r: 0, c: 9, v: 1 }] });
		expect(inv).toEqual({ t: 'cells', changes: [] });
	});
	test('free header names', () => {
		expect(freeHeader(['a', 'columna', 'columna_2'], 'columna')).toBe('columna_3');
		expect(freeHeader(['a'], 'columna')).toBe('columna');
	});
});

describe('history', () => {
	test('push clears redo, limit drops oldest', () => {
		const h = new History(2);
		const step = (id: string) => ({ tableId: id, ops: [], inverse: [] });
		h.push(step('a'));
		h.push(step('b'));
		h.push(step('c'));
		expect(h.undoStack.map((s) => s.tableId)).toEqual(['b', 'c']);
		const s = h.popUndo()!;
		h.pushRedo(s);
		expect(h.canRedo).toBe(true);
		h.push(step('d'));
		expect(h.canRedo).toBe(false);
		h.forget('b');
		expect(h.undoStack.map((x) => x.tableId)).toEqual(['d']);
	});
});

describe('view', () => {
	test('filters by text (accent-insensitive) and values', () => {
		const t = base();
		expect(visibleRows(t, {})).toEqual([0, 1, 2]);
		expect(visibleRows(t, { 1: { query: 'cancion' } })).toEqual([2]);
		expect(visibleRows(t, { 2: { values: ['10', ''] } })).toEqual([0, 1]);
		expect(visibleRows(t, { 2: { values: [] } })).toEqual([]);
		expect(fold('ÁrbOl')).toBe('arbol');
	});
	test('distinct values with empty last', () => {
		expect(distinctValues(base(), 2)).toEqual({ values: ['2', '10', ''], truncated: false });
	});
	test('stable sort, empties last in both directions', () => {
		const t = makeTable(['v'], [[3], [null], ['b'], [1], [true], ['a10'], ['a9'], [1]]);
		expect(sortOrder(t, 0, 'asc')).toEqual([3, 7, 0, 4, 6, 5, 2, 1]);
		expect(sortOrder(t, 0, 'desc')).toEqual([2, 5, 6, 4, 0, 3, 7, 1]);
		expect(compareCells(null, 1)).toBeGreaterThan(0);
	});
});

describe('find', () => {
	const t = base();
	test('finds ignoring case and accents, whole cell, regex, column', () => {
		expect(findAll(t, [0, 1, 2], 'CANCION', DEFAULT_FIND)).toEqual([{ r: 2, c: 1 }]);
		expect(findAll(t, [0, 1, 2], '1', DEFAULT_FIND)).toEqual([{ r: 0, c: 0 }, { r: 0, c: 2 }]);
		expect(findAll(t, [0, 1, 2], '1', { ...DEFAULT_FIND, wholeCell: true })).toEqual([{ r: 0, c: 0 }]);
		expect(findAll(t, [0, 1, 2], '^B', { ...DEFAULT_FIND, regex: true })).toEqual([{ r: 1, c: 1 }]);
		expect(findAll(t, [0, 1, 2], '1', { ...DEFAULT_FIND, column: 2 })).toEqual([{ r: 0, c: 2 }]);
		expect(findAll(t, [0, 1], 'Canción', DEFAULT_FIND)).toEqual([]);
		expect(findAll(t, [0, 1, 2], '(', { ...DEFAULT_FIND, regex: true })).toEqual([]);
	});
	test('replaces keeping the original accents and types', () => {
		expect(replaceValue('Canción triste', 'cancion', 'Tema', DEFAULT_FIND)).toBe('Tema triste');
		expect(replaceValue('aAa', 'a', 'x', DEFAULT_FIND)).toBe('xxx');
		expect(replaceValue('aAa', 'a', 'x', { ...DEFAULT_FIND, matchCase: true })).toBe('xAx');
		expect(replaceValue(10, '1', '2', DEFAULT_FIND)).toBe(20);
		expect(replaceValue('2024-01-01', '(\\d+)-(\\d+)-(\\d+)', '$3/$2/$1', { ...DEFAULT_FIND, regex: true })).toBe('01/01/2024');
		expect(replaceValue(null, 'a', 'b', DEFAULT_FIND)).toBeNull();
	});
});

describe('clipboard and layout', () => {
	test('tsv round trip with quotes and newlines', () => {
		const grid = [
			['a', 'b\tc'],
			['línea\nnueva', 'x"y']
		];
		expect(parseTSV(toTSV(grid))).toEqual(grid);
		expect(parseHTMLTable('<table></table>')).toBeNull(); // sin DOMParser en Bun
	});
	test('axis offsets with overrides', () => {
		const a = new Axis(5, 100, new Map([[1, 50], [3, 200]]));
		expect(a.total).toBe(550);
		expect([0, 1, 2, 3, 4].map((i) => a.offset(i))).toEqual([0, 100, 150, 250, 450]);
		expect(a.indexAt(149)).toBe(1);
		expect(a.indexAt(450)).toBe(4);
		expect(a.indexAt(9999)).toBe(4);
	});
});
