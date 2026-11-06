// Esquema de las preferencias (puro, con tests). Cada campo inválido o ausente vuelve a su valor por defecto.
import { z } from 'zod';
import { DEFAULT_EXPORT } from '../core/outputs.ts';
import { DEFAULT_IMPORT } from '../core/table.ts';
import type { ExportOptions, OutputFormat, OutputMode } from '../core/types.ts';

const mode = z.enum(['split', 'perFile', 'combined', 'merged']);

const ImportSchema = z.object({
	headerRow: z.number().int().min(0).max(10_000).catch(DEFAULT_IMPORT.headerRow),
	skipEmptyRows: z.boolean().catch(DEFAULT_IMPORT.skipEmptyRows),
	nested: z.boolean().catch(DEFAULT_IMPORT.nested)
});

export const DEFAULT_MODES: Record<OutputFormat, OutputMode> = { json: 'split', xlsx: 'combined', csv: 'split' };

const ExportSchema = z.object({
	format: z.enum(['json', 'xlsx', 'csv']).catch(DEFAULT_EXPORT.format),
	modes: z
		.object({ json: mode.catch(DEFAULT_MODES.json), xlsx: mode.catch(DEFAULT_MODES.xlsx), csv: mode.catch(DEFAULT_MODES.csv) })
		.catch(() => ({ ...DEFAULT_MODES })),
	allText: z.boolean().catch(DEFAULT_EXPORT.allText),
	nested: z.boolean().catch(DEFAULT_EXPORT.nested),
	indent: z.number().int().min(0).max(8).catch(DEFAULT_EXPORT.indent),
	dropNulls: z.boolean().catch(DEFAULT_EXPORT.dropNulls),
	csvDelimiter: z.enum([',', ';', '\t']).catch(DEFAULT_EXPORT.csvDelimiter),
	bom: z.boolean().catch(DEFAULT_EXPORT.bom),
	styled: z.boolean().catch(DEFAULT_EXPORT.styled),
	origin: z.boolean().catch(true),
	onlyVisible: z.boolean().catch(false)
});

const PrefsSchema = z.object({
	import: ImportSchema.catch(() => ImportSchema.parse({})),
	export: ExportSchema.catch(() => ExportSchema.parse({}))
});

export type Prefs = z.infer<typeof PrefsSchema>;

export const defaultPrefs = (): Prefs => PrefsSchema.parse({ import: {}, export: { modes: {} } });

export function parsePrefs(raw: string | null): Prefs {
	if (!raw) return defaultPrefs();
	try {
		const data: unknown = JSON.parse(raw);
		return PrefsSchema.parse(typeof data === 'object' && data !== null ? data : {});
	} catch {
		return defaultPrefs();
	}
}

/** Opciones completas de exportación (el nombre de la columna de origen depende del idioma). */
export function toExportOptions(p: Prefs, originName: string): ExportOptions {
	const e = p.export;
	return {
		format: e.format,
		mode: e.modes[e.format],
		allText: e.allText,
		nested: e.nested,
		indent: e.indent,
		dropNulls: e.dropNulls,
		csvDelimiter: e.csvDelimiter,
		bom: e.bom,
		styled: e.styled,
		originColumn: e.origin ? originName : null
	};
}
