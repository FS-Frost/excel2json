// Opciones de importación y exportación recordadas por dispositivo (localStorage['excel2json.prefs']).
import type { ExportOptions, ImportOptions } from '../core/types.ts';
import { defaultPrefs, parsePrefs, toExportOptions, type Prefs } from './prefs-schema.ts';

const KEY = 'excel2json.prefs';

function load(): Prefs {
	try {
		return parsePrefs(localStorage.getItem(KEY));
	} catch {
		return defaultPrefs();
	}
}

export const prefs = $state<Prefs>(typeof window === 'undefined' ? defaultPrefs() : load());

export function savePrefs() {
	try {
		localStorage.setItem(KEY, JSON.stringify($state.snapshot(prefs)));
	} catch {
		/* sin almacenamiento: vale solo para esta sesión */
	}
}

export function resetExportPrefs() {
	prefs.export = defaultPrefs().export;
	savePrefs();
}

export function importOptions(): ImportOptions {
	return { ...prefs.import };
}

export function exportOptions(originName: string): ExportOptions {
	return toExportOptions(prefs, originName);
}
