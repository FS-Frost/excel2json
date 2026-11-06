// Buscar y reemplazar en una tabla (sobre las filas visibles).
import { cellText, inferText } from '../core/cell.ts';
import type { Cell, Table } from '../core/types.ts';
import { fold } from './view.ts';

export interface FindOptions {
	matchCase: boolean;
	/** La celda completa debe coincidir. */
	wholeCell: boolean;
	regex: boolean;
	/** Solo esta columna (null = todas). */
	column: number | null;
}

export const DEFAULT_FIND: FindOptions = { matchCase: false, wholeCell: false, regex: false, column: null };

export type Match = { r: number; c: number };

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Expresión de búsqueda; null si la consulta está vacía o la regex no es válida. */
export function matcher(query: string, o: FindOptions): RegExp | null {
	if (!query) return null;
	try {
		const src = o.regex ? query : escapeRe(o.matchCase ? query : fold(query));
		return new RegExp(o.wholeCell ? `^(?:${src})$` : src, o.matchCase ? 'gu' : 'giu');
	} catch {
		return null;
	}
}

/** Texto contra el que se compara (sin tildes si no importan las mayúsculas, salvo en modo regex). */
const subject = (v: Cell, o: FindOptions) => {
	const s = cellText(v);
	return o.matchCase || o.regex ? s : fold(s);
};

/** Coincidencias en orden de lectura (fila visible, luego columna). `rows` = índices reales visibles. */
export function findAll(t: Table, rows: number[], query: string, o: FindOptions): Match[] {
	const re = matcher(query, o);
	if (!re) return [];
	const out: Match[] = [];
	const cols = o.column === null ? t.headers.map((_, c) => c) : [o.column];
	for (const r of rows) {
		const row = t.rows[r];
		for (const c of cols) {
			if (row[c] === null) continue;
			re.lastIndex = 0;
			if (re.test(subject(row[c], o))) out.push({ r, c });
		}
	}
	return out;
}

/**
 * Valor nuevo de una celda tras reemplazar. Con tildes/mayúsculas ignoradas el reemplazo se hace sobre el texto
 * original buscando la misma posición (fold no cambia el largo para letras latinas con tilde compuesta).
 * Si el valor original no era texto, el resultado se vuelve a inferir (un número sigue siendo número).
 */
export function replaceValue(v: Cell, query: string, replacement: string, o: FindOptions): Cell {
	const re = matcher(query, o);
	if (!re || v === null) return v;
	const text = cellText(v);
	let out: string;
	if (o.matchCase || o.regex) out = text.replace(re, replacement);
	else {
		// Buscar en la versión sin tildes y cortar el original en las mismas posiciones (NFC: 1 char ↔ 1 char).
		const nfc = text.normalize('NFC');
		const folded = [...nfc].map((ch) => fold(ch)).join('');
		if (folded.length !== nfc.length) out = nfc.replace(re, replacement);
		else {
			let res = '';
			let last = 0;
			re.lastIndex = 0;
			for (const m of folded.matchAll(re)) {
				const i = m.index ?? 0;
				res += nfc.slice(last, i) + replacement;
				last = i + m[0].length;
				if (m[0].length === 0) break;
			}
			out = res + nfc.slice(last);
		}
	}
	if (out === text) return v;
	return typeof v === 'string' ? out : inferText(out);
}
