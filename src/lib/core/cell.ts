// Valores de celda: normalización desde exceljs, fechas ISO, coerción por tipo de columna y texto escrito por el usuario.
import type { Cell, ColType } from './types.ts';

const pad = (n: number, w = 2) => String(n).padStart(w, '0');

/**
 * Fecha de exceljs → ISO local sin zona. exceljs entrega la hora "de pared" del libro en los componentes UTC
 * (el serial de Excel no tiene zona), así que se leen con getUTC*. Medianoche exacta → solo la fecha.
 */
export function dateToIso(d: Date): string | null {
	if (Number.isNaN(d.getTime())) return null;
	const date = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
	const h = d.getUTCHours();
	const m = d.getUTCMinutes();
	const s = d.getUTCSeconds();
	return h || m || s ? `${date}T${pad(h)}:${pad(m)}:${pad(s)}` : date;
}

const ISO_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?)?$/;

/** ISO estricto (sin zona) → Date con la hora de pared en UTC, como la escribe exceljs. null si no es una fecha válida. */
export function isoToDate(s: string): Date | null {
	const m = ISO_RE.exec(s);
	if (!m) return null;
	const [y, mo, d, h = '0', mi = '0', se = '0'] = m.slice(1);
	const date = new Date(Date.UTC(+y, +mo - 1, +d, +h, +mi, +se));
	// Rechaza 2024-02-31 y similares (Date los desborda al mes siguiente).
	if (date.getUTCMonth() !== +mo - 1 || date.getUTCDate() !== +d || +h > 23 || +mi > 59 || +se > 59) return null;
	return date;
}

export const isIsoDate = (s: string) => isoToDate(s) !== null;

/** Resultado de normalizar: `error` marca celdas con error de Excel (#N/A, #DIV/0!…), que quedan en null. */
export type Normalized = { value: Cell; error?: boolean };

function richText(v: { richText: { text: string }[] }) {
	return v.richText.map((r) => r.text).join('');
}

/** Valor de exceljs (`cell.value`) → Cell. Fórmula → resultado; richText/hipervínculo → texto; error → null. */
export function normalizeExcel(v: unknown): Normalized {
	if (v === null || v === undefined) return { value: null };
	if (typeof v === 'string') return { value: v === '' ? null : v };
	if (typeof v === 'number') return { value: Number.isFinite(v) ? v : null };
	if (typeof v === 'boolean') return { value: v };
	if (v instanceof Date) return { value: dateToIso(v) };
	if (typeof v === 'object') {
		const o = v as Record<string, unknown>;
		if ('error' in o) return { value: null, error: true };
		if ('formula' in o || 'sharedFormula' in o) {
			if (o.result === undefined) return { value: null };
			return normalizeExcel(o.result);
		}
		if (Array.isArray(o.richText)) return { value: richText(o as { richText: { text: string }[] }) || null };
		if ('hyperlink' in o || 'text' in o) {
			const t = o.text;
			if (typeof t === 'string') return { value: t || null };
			if (t && typeof t === 'object' && Array.isArray((t as { richText?: unknown }).richText)) return { value: richText(t as { richText: { text: string }[] }) || null };
			return { value: typeof o.hyperlink === 'string' ? o.hyperlink : null };
		}
	}
	return { value: String(v) };
}

// ---------- Texto → valor ----------

const NUM_DOT = /^-?(0|[1-9]\d*)(\.\d+)?([eE][+-]?\d+)?$/;
const NUM_COMMA = /^-?(0|[1-9]\d*)(,\d+)?$/;

/**
 * Infiere el tipo de un texto (CSV, celda escrita a mano). Solo números "limpios": `00123`, `1.000,5` o teléfonos
 * quedan como texto. `decimal` = separador decimal aceptado además del punto.
 */
export function inferText(raw: string, decimal: '.' | ',' = '.'): Cell {
	const s = raw.trim();
	if (s === '') return null;
	if (NUM_DOT.test(s)) return Number(s);
	if (decimal === ',' && NUM_COMMA.test(s)) return Number(s.replace(',', '.'));
	const lower = s.toLowerCase();
	if (lower === 'true') return true;
	if (lower === 'false') return false;
	return raw;
}

/** Texto escrito en una celda del grid. `'` inicial fuerza texto (como en Excel). */
export function parseTyped(raw: string, type: ColType, decimal: '.' | ',' = '.'): Cell {
	if (raw.startsWith("'")) return raw.slice(1);
	if (type === 'text') return raw === '' ? null : raw;
	return inferText(raw, decimal);
}

// ---------- Coerción por tipo de columna ----------

const TRUE = new Set(['true', '1', 'yes', 'y', 'si', 'sí', 's', 'verdadero', 'x']);
const FALSE = new Set(['false', '0', 'no', 'n', 'falso']);

/** ¿El valor calza con el tipo de la columna? (para marcarlo en el grid). */
export function fits(v: Cell, type: ColType): boolean {
	if (v === null || type === 'auto' || type === 'text') return true;
	return coerce(v, type).ok;
}

/** Convierte una celda al tipo de su columna. Si no se puede, devuelve el valor original con ok = false. */
export function coerce(v: Cell, type: ColType): { value: Cell; ok: boolean } {
	if (v === null || type === 'auto') return { value: v, ok: true };
	switch (type) {
		case 'text':
			return { value: typeof v === 'string' ? v : String(v), ok: true };
		case 'number': {
			if (typeof v === 'number') return { value: v, ok: true };
			if (typeof v === 'boolean') return { value: v ? 1 : 0, ok: true };
			const s = v.trim().replace(/\s/g, '');
			const n = NUM_DOT.test(s) || /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : /^-?\d+(,\d+)?$/.test(s) ? Number(s.replace(',', '.')) : NaN;
			return Number.isFinite(n) ? { value: n, ok: true } : { value: v, ok: false };
		}
		case 'boolean': {
			if (typeof v === 'boolean') return { value: v, ok: true };
			const k = String(v).trim().toLowerCase();
			if (TRUE.has(k)) return { value: true, ok: true };
			if (FALSE.has(k)) return { value: false, ok: true };
			return { value: v, ok: false };
		}
		case 'date': {
			if (typeof v === 'string') {
				const s = v.trim();
				if (isIsoDate(s)) return { value: s, ok: true };
				// dd/mm/aaaa y dd-mm-aaaa (formato habitual en español).
				const m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(s);
				if (m) {
					const iso = `${m[3]}-${pad(+m[2])}-${pad(+m[1])}`;
					if (isIsoDate(iso)) return { value: iso, ok: true };
				}
				return { value: v, ok: false };
			}
			if (typeof v === 'number' && v > 0 && v < 2958466) {
				// Serial de Excel (sistema 1900): días desde 1899-12-30.
				return { value: dateToIso(new Date(Math.round((v - 25569) * 86400000))), ok: true };
			}
			return { value: v, ok: false };
		}
	}
}

/** Texto para mostrar/copiar/CSV. `decimal` cambia el separador decimal de los números. */
export function cellText(v: Cell, decimal: '.' | ',' = '.'): string {
	if (v === null) return '';
	if (typeof v === 'number') {
		const s = String(Number(v.toPrecision(15)));
		return decimal === ',' ? s.replace('.', ',') : s;
	}
	return String(v);
}
