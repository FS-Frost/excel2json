// Descargar y copiar resultados.
import type { OutputFile } from '../core/types.ts';

export function downloadFile(file: OutputFile) {
	const blob = new Blob([file.bytes as BlobPart], { type: file.mime });
	const url = URL.createObjectURL(blob);
	const a = document.createElement('a');
	a.href = url;
	a.download = file.name;
	a.rel = 'noopener';
	document.body.appendChild(a);
	a.click();
	a.remove();
	// Revocar en seguida puede cortar la descarga en algunos navegadores.
	setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

/** Texto de un archivo de salida (JSON/CSV) sin el BOM. */
export function fileText(file: OutputFile): string {
	return new TextDecoder().decode(file.bytes).replace(/^﻿/, '');
}

export async function copyText(text: string) {
	await navigator.clipboard.writeText(text);
}
