import { describe, expect, test } from 'bun:test';
import en from '../src/lib/i18n/en.ts';
import es from '../src/lib/i18n/es.ts';
import { detect, readPref, resolveLocale } from '../src/lib/i18n/detect.ts';
import { noteVersion } from '../src/lib/client/version.ts';
import { defaultPrefs, parsePrefs, toExportOptions } from '../src/lib/client/prefs-schema.ts';

type Tree = { [k: string]: unknown };

function leaves(t: Tree, prefix = ''): [string, unknown][] {
	return Object.entries(t).flatMap(([k, v]) =>
		typeof v === 'object' && v !== null && !Array.isArray(v) ? leaves(v as Tree, `${prefix}${k}.`) : [[`${prefix}${k}`, v] as [string, unknown]]
	);
}

describe('i18n', () => {
	const a = leaves(en as unknown as Tree);
	const b = leaves(es as unknown as Tree);
	test('same keys in both languages', () => {
		expect(b.map(([k]) => k).sort()).toEqual(a.map(([k]) => k).sort());
	});
	test('no empty strings, same kinds', () => {
		const map = new Map(b);
		// Las ayudas de modos que no existen para CSV quedan vacías a propósito.
		const allowedEmpty = new Set(['export.modeHelp.csv.perFile', 'export.modeHelp.csv.combined']);
		for (const [k, v] of a) {
			const w = map.get(k);
			expect(typeof w, k).toBe(typeof v);
			if (typeof v === 'string' && !allowedEmpty.has(k)) {
				expect(v.trim().length, k).toBeGreaterThan(0);
				expect((w as string).trim().length, k).toBeGreaterThan(0);
			}
		}
	});
	test('functions produce text', () => {
		expect(es.files.rows(1)).toBe('1 fila');
		expect(es.files.rows(1500)).toBe('1500 filas'.replace('1500', (1500).toLocaleString('es')));
		expect(en.files.sheets(2)).toBe('2 sheets');
		expect(es.grid.rowsOf(3, 10)).toBe('3 de 10 filas');
		expect(en.find.results(0, 0)).toBe('No results');
	});
	test('detection', () => {
		expect(detect(['es-CL', 'en'])).toBe('es');
		expect(detect(['en-US', 'es'])).toBe('es');
		expect(detect(['fr-FR', 'de'])).toBe('en');
		expect(resolveLocale('en', ['es-ES'])).toBe('en');
		expect(readPref({ getItem: () => 'es' })).toBe('es');
		expect(readPref({ getItem: () => 'xx' })).toBe('system');
		expect(readPref(null)).toBe('system');
	});
});

describe('version', () => {
	test('noteVersion reports changes once', () => {
		const m = new Map<string, string>();
		const store = { getItem: (k: string) => m.get(k) ?? null, setItem: (k: string, v: string) => void m.set(k, v) };
		expect(noteVersion('abc', store)).toBeNull();
		expect(noteVersion('abc', store)).toBeNull();
		expect(noteVersion('def', store)).toBe('abc');
		expect(noteVersion('dev', store)).toBeNull();
	});
});

describe('prefs', () => {
	test('defaults and per-field recovery', () => {
		const d = defaultPrefs();
		expect(d.import).toEqual({ headerRow: 1, skipEmptyRows: true, nested: false });
		expect(d.export.modes).toEqual({ json: 'split', xlsx: 'combined', csv: 'split' });
		const p = parsePrefs(JSON.stringify({ import: { headerRow: -3, nested: true }, export: { format: 'xml', indent: 4, modes: { csv: 'combined' } } }));
		expect(p.import).toEqual({ headerRow: 1, skipEmptyRows: true, nested: true });
		expect(p.export.format).toBe('json');
		expect(p.export.indent).toBe(4);
		expect(p.export.modes.csv).toBe('combined');
		expect(parsePrefs('{nope')).toEqual(d);
		expect(parsePrefs('null')).toEqual(d);
	});
	test('export options', () => {
		const p = defaultPrefs();
		p.export.format = 'xlsx';
		p.export.origin = false;
		expect(toExportOptions(p, '_origen')).toMatchObject({ format: 'xlsx', mode: 'combined', originColumn: null });
	});
});
