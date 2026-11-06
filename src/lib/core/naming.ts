// Nombres: archivos, hojas de Excel, cabeceras.

/** 0 → A, 25 → Z, 26 → AA (letras de columna de Excel). */
export function colLetter(i: number): string {
	let n = i + 1;
	let s = '';
	while (n > 0) {
		const r = (n - 1) % 26;
		s = String.fromCharCode(65 + r) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s;
}

/** `datos.2024.xlsx` → `datos.2024` (solo se quita la última extensión). */
export function baseName(file: string): string {
	const name = file.split(/[/\\]/).pop() ?? file;
	const i = name.lastIndexOf('.');
	return i > 0 ? name.slice(0, i) : name || 'file';
}

export function extension(file: string): string {
	const i = file.lastIndexOf('.');
	return i > 0 ? file.slice(i + 1).toLowerCase() : '';
}

/** `name`, `name (2)`, `name (3)`… evitando los usados (comparación sin mayúsculas, como Excel y los sistemas de archivos). */
export function uniqueName(name: string, used: Set<string>, fmt = (n: string, k: number) => `${n} (${k})`): string {
	let out = name;
	for (let k = 2; used.has(out.toLowerCase()); k++) out = fmt(name, k);
	used.add(out.toLowerCase());
	return out;
}

/** Nombre de archivo con extensión, único dentro del conjunto (para el zip). */
export function uniqueFileName(base: string, ext: string, used: Set<string>): string {
	const safe = base.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim() || 'file';
	return uniqueName(`${safe}.${ext}`, used, (n, k) => `${n.slice(0, -(ext.length + 1))} (${k}).${ext}`);
}

const SHEET_MAX = 31;

/** Nombre de hoja válido para Excel: sin `[]:*?/\`, sin comillas simples en los bordes, ≤ 31 caracteres, no vacío. */
export function safeSheetName(name: string): string {
	const s = name
		.replace(/[[\]:*?/\\]/g, ' ')
		.replace(/^'+|'+$/g, '')
		.trim()
		.slice(0, SHEET_MAX)
		.trim();
	return s || 'Sheet';
}

/** Nombre de hoja válido y único en el libro (`Hoja (2)` recortando para no pasar de 31). */
export function uniqueSheetName(name: string, used: Set<string>): string {
	const base = safeSheetName(name);
	return uniqueName(base, used, (n, k) => {
		const suffix = ` (${k})`;
		return n.slice(0, SHEET_MAX - suffix.length) + suffix;
	});
}

/**
 * Cabeceras finales: vacías → letra de la columna (`C`), repetidas → `nombre_2`, `nombre_3`…
 * Devuelve también qué se corrigió, para avisar.
 */
export function normalizeHeaders(raw: string[]): { headers: string[]; empty: number; duplicates: string[] } {
	const filled = raw.map((h, i) => {
		const t = h.trim();
		return t === '' ? colLetter(i) : t;
	});
	const empty = raw.filter((h) => h.trim() === '').length;
	const taken = new Set(filled);
	const seen = new Set<string>();
	const duplicates = new Set<string>();
	const headers = filled.map((h) => {
		if (!seen.has(h)) {
			seen.add(h);
			return h;
		}
		duplicates.add(h);
		let k = 2;
		while (taken.has(`${h}_${k}`) || seen.has(`${h}_${k}`)) k++;
		const name = `${h}_${k}`;
		seen.add(name);
		return name;
	});
	return { headers, empty, duplicates: [...duplicates] };
}
