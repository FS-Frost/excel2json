// Idioma de la interfaz (origen: FS-Frost/mai-docs). `t()` devuelve el diccionario del idioma actual; leerlo en una
// plantilla o un efecto es reactivo:  {t().drop.title}   {t().files.rows(n)}
import en, { type Dict } from './en.ts';
import es from './es.ts';
import { LANG_KEY, readPref, resolveLocale, type LangPref, type Locale } from './detect.ts';
import type { ImportWarning } from '../core/types.ts';

export type { Dict, Locale, LangPref };

const dicts: Record<Locale, Dict> = { en, es };

function storage() {
	try {
		return localStorage;
	} catch {
		return null;
	}
}

function systemLangs(): readonly string[] {
	if (typeof navigator === 'undefined') return ['en'];
	return navigator.languages?.length ? navigator.languages : [navigator.language || 'en'];
}

const initialPref = typeof window === 'undefined' ? 'system' : readPref(storage());

export const i18n = $state({
	pref: initialPref as LangPref,
	locale: resolveLocale(initialPref, systemLangs())
});

export function t(): Dict {
	return dicts[i18n.locale];
}

/** Etiqueta BCP 47 para Intl.*. */
export function intlLocale() {
	return i18n.locale === 'es' ? 'es' : 'en';
}

/** Separador decimal para mostrar y escribir números en el grid. */
export function decimalSep(): '.' | ',' {
	return i18n.locale === 'es' ? ',' : '.';
}

export function setLangPref(pref: LangPref) {
	i18n.pref = pref;
	i18n.locale = resolveLocale(pref, systemLangs());
	try {
		if (pref === 'system') localStorage.removeItem(LANG_KEY);
		else localStorage.setItem(LANG_KEY, pref);
	} catch {
		/* solo esta sesión */
	}
	document.documentElement.lang = i18n.locale;
}

/** Sigue los cambios de idioma del SO mientras la preferencia sea «sistema». */
export function initI18n() {
	document.documentElement.lang = i18n.locale;
	addEventListener('languagechange', () => {
		if (i18n.pref === 'system') {
			i18n.locale = resolveLocale('system', systemLangs());
			document.documentElement.lang = i18n.locale;
		}
	});
}

const STEPS: [Intl.RelativeTimeFormatUnit, number][] = [
	['year', 365 * 86400],
	['month', 30 * 86400],
	['week', 7 * 86400],
	['day', 86400],
	['hour', 3600],
	['minute', 60]
];

/** «hace 3 minutos», «ayer»… en el idioma de la interfaz. */
export function ago(ms: number, now = Date.now()) {
	const s = Math.round((now - ms) / 1000);
	if (!Number.isFinite(s)) return '';
	if (s < 60) return t().time.justNow;
	const rtf = new Intl.RelativeTimeFormat(intlLocale(), { numeric: 'auto' });
	for (const [unit, secs] of STEPS) if (s >= secs) return rtf.format(-Math.floor(s / secs), unit);
	return '';
}

export function formatBytes(n: number) {
	const units = ['B', 'KB', 'MB', 'GB', 'TB'];
	let i = 0;
	while (n >= 1024 && i < units.length - 1) {
		n /= 1024;
		i++;
	}
	return `${new Intl.NumberFormat(intlLocale(), { maximumFractionDigits: i ? 1 : 0 }).format(n)} ${units[i]}`;
}

/** Texto de un aviso de importación. */
export function warningText(w: ImportWarning): string {
	const d = t().warnings;
	switch (w.kind) {
		case 'cellErrors':
			return d.cellErrors(w.sheet, w.count);
		case 'emptyHeaders':
			return d.emptyHeaders(w.sheet, w.count);
		case 'duplicateHeaders':
			return d.duplicateHeaders(w.sheet, w.names);
		case 'emptySheet':
			return d.emptySheet(w.sheet);
		case 'decoding':
			return d.decoding(w.encoding);
	}
}

/** Mensaje de un error del conversor (ConvertError serializado desde el Worker o lanzado aquí). */
export function errorText(e: unknown): string {
	const d = t().errors;
	const o = (e ?? {}) as { code?: string; detail?: string; message?: string };
	switch (o.code) {
		case 'xlsxTooManyRows':
			return d.xlsxTooManyRows(o.detail ?? '');
		case 'xlsxTooManyCols':
			return d.xlsxTooManyCols(o.detail ?? '');
		case 'jsonShape':
			return d.jsonShape;
		case 'jsonSyntax':
			return d.jsonSyntax(o.detail ?? '');
		case 'xlsxRead':
			return d.xlsxRead;
		case 'unsupported':
			return d.unsupported;
		case 'empty':
			return d.empty;
		case 'cancelled':
			return d.cancelled;
		case 'worker':
			return d.worker;
		default:
			return o.message || d.generic;
	}
}
