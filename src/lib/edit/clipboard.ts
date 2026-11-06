// Copiado de FS-Frost/mai-docs (src/lib/sheets/clipboard.ts).
// Clipboard interop: TSV + HTML table out (Excel/Sheets read both), TSV/HTML in.

const escHtml = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** TSV with Excel quoting (cells containing tab, newline or quote are quoted). */
export function toTSV(rows: string[][]): string {
	return rows.map((r) => r.map((s) => (/[\t\n\r"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s)).join('\t')).join('\n');
}

export function toHTML(rows: { text: string; style?: string }[][]): string {
	const body = rows
		.map((r) => `<tr>${r.map((c) => `<td${c.style ? ` style="${c.style.replace(/"/g, "'")}"` : ''}>${escHtml(c.text).replace(/\n/g, '<br>')}</td>`).join('')}</tr>`)
		.join('');
	return `<meta charset="utf-8"><table>${body}</table>`;
}

/** Parses TSV (Excel quoting rules). Trailing newline ignored. */
export function parseTSV(text: string): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let cur = '';
	let i = 0;
	let quoted = false;
	const t = text.replace(/\r\n?/g, '\n');
	while (i < t.length) {
		const ch = t[i];
		if (quoted) {
			if (ch === '"') {
				if (t[i + 1] === '"') {
					cur += '"';
					i += 2;
					continue;
				}
				quoted = false;
				i++;
				continue;
			}
			cur += ch;
			i++;
			continue;
		}
		if (ch === '"' && cur === '') {
			quoted = true;
			i++;
			continue;
		}
		if (ch === '\t') {
			row.push(cur);
			cur = '';
		} else if (ch === '\n') {
			row.push(cur);
			rows.push(row);
			row = [];
			cur = '';
		} else cur += ch;
		i++;
	}
	if (cur !== '' || row.length) {
		row.push(cur);
		rows.push(row);
	}
	return rows;
}

/**
 * Texto de una celda HTML del portapapeles. En un documento de DOMParser (no renderizado) innerText se comporta como
 * textContent y pierde los <br>: se convierten a \n antes. Sin trim (los espacios de los bordes son datos); solo se
 * quitan los saltos de línea que el HTML pone alrededor del contenido.
 */
function cellHtmlText(td: HTMLTableCellElement): string {
	const copy = td.cloneNode(true) as HTMLElement;
	for (const br of Array.from(copy.querySelectorAll('br'))) br.replaceWith('\n');
	return (copy.textContent ?? '').replace(/ /g, ' ').replace(/^[\r\n]+|[\r\n]+$/g, '');
}

/** First <table> of clipboard HTML → text grid (browser only). Honors colspan/rowspan by padding. */
export function parseHTMLTable(html: string): string[][] | null {
	if (typeof DOMParser === 'undefined') return null;
	const doc = new DOMParser().parseFromString(html, 'text/html');
	const table = doc.querySelector('table');
	if (!table) return null;
	const grid: string[][] = [];
	const taken: boolean[][] = [];
	Array.from(table.rows).forEach((tr, r) => {
		grid[r] ??= [];
		taken[r] ??= [];
		let c = 0;
		for (const td of Array.from(tr.cells)) {
			while (taken[r][c]) c++;
			const text = cellHtmlText(td);
			const cs = Math.max(1, td.colSpan || 1);
			const rs = Math.max(1, td.rowSpan || 1);
			for (let dr = 0; dr < rs; dr++)
				for (let dc = 0; dc < cs; dc++) {
					grid[r + dr] ??= [];
					taken[r + dr] ??= [];
					grid[r + dr][c + dc] = dr === 0 && dc === 0 ? text : '';
					taken[r + dr][c + dc] = true;
				}
			c += cs;
		}
	});
	return grid.map((row) => Array.from(row, (x) => x ?? ''));
}
