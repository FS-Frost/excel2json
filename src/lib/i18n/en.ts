// UI strings. English defines the shape (`Dict`); es.ts must have exactly the same keys (tests/i18n.test.ts).
import type { ColType, OutputFormat, OutputMode } from '../core/types.ts';

const plural = (n: number, one: string, many: string) => `${n.toLocaleString('en')} ${n === 1 ? one : many}`;

const en = {
	common: {
		close: 'Close',
		cancel: 'Cancel',
		ok: 'OK',
		retry: 'Retry',
		undo: 'Undo',
		redo: 'Redo',
		delete: 'Delete',
		rename: 'Rename',
		clear: 'Clear',
		apply: 'Apply',
		reset: 'Reset',
		all: 'All',
		none: 'None',
		more: 'More',
		loading: 'Loading…'
	},
	app: {
		name: 'excel2json',
		tagline: 'Convert and edit Excel, CSV and JSON right in your browser. Nothing is uploaded.',
		privacy: 'Your files never leave this device.'
	},
	header: {
		language: 'Language',
		langSystem: 'System language',
		themeLight: 'Light',
		themeDark: 'Dark',
		themeSystem: 'System',
		themeToggle: (current: string, next: string) => `Theme: ${current} (switch to ${next})`,
		github: 'Source code on GitHub'
	},
	net: {
		offline: 'You are offline',
		offlineDetail: 'excel2json keeps working: everything happens on this device.',
		online: 'Back online'
	},
	update: {
		available: 'A new version of excel2json is available',
		action: 'Update',
		updating: 'Updating…',
		updatingDetail: 'Your work is saved on this device.',
		publishing: 'The new version is still being published',
		publishingDetail: 'Try again in a minute.',
		offline: 'Could not check for updates (offline?)',
		failed: (v: string) => `The update to ${v} did not load`,
		failedDetail: 'Force a reload from the network (your session is kept).',
		force: 'Force',
		done: (v: string) => `excel2json updated (${v})`,
		publishedAgo: (when: string) => `Published ${when}`
	},
	time: { justNow: 'just now' },
	drop: {
		title: 'Drop your files here',
		subtitle: 'Excel (.xlsx), CSV, TSV or JSON. Several at once, with all their sheets.',
		choose: 'Choose files',
		paste: 'Paste data',
		sample: 'Try an example',
		overlay: 'Drop to add',
		reading: (name: string) => `Reading ${name}…`,
		readingSheet: (name: string, sheet: string) => `Reading ${name} › ${sheet}`,
		unsupported: (name: string) => `${name}: format not supported (use .xlsx, .csv, .tsv or .json)`,
		readFailed: (name: string) => `Could not read ${name}`,
		added: (n: number) => (n === 1 ? 'File added' : `${n} files added`),
		bigFile: (name: string, size: string) => `${name} is large (${size}): it may take a while`
	},
	paste: {
		title: 'Paste JSON, CSV or a table copied from Excel',
		placeholder: '[{"id": 1, "name": "Ana"}]  or  id,name…',
		name: 'Name',
		defaultName: 'pasted',
		add: 'Add',
		empty: 'Paste something first'
	},
	restore: {
		title: 'Previous session',
		detail: (files: number, when: string) => `${plural(files, 'file', 'files')}, saved ${when}.`,
		restore: 'Restore',
		discard: 'Discard',
		replacedWarn: 'Adding files starts a new session and replaces the saved one.',
		restored: 'Session restored',
		noOriginal: 'Restored from the saved session: import options can no longer be reapplied to this file.'
	},
	files: {
		heading: 'Files',
		addMore: 'Add files',
		sheets: (n: number) => plural(n, 'sheet', 'sheets'),
		rows: (n: number) => plural(n, 'row', 'rows'),
		cols: (n: number) => plural(n, 'column', 'columns'),
		remove: 'Remove file',
		removeConfirm: (name: string) => `Remove ${name} and its edits?`,
		clearAll: 'Close all',
		clearConfirm: 'Close every file? Unexported edits will be lost.',
		selectAll: 'Select all sheets',
		selectNone: 'Deselect all',
		include: (name: string) => `Include ${name} in the export`,
		selected: (n: number, total: number) => `${n} of ${total} selected`,
		edited: 'Edited',
		warnings: (n: number) => plural(n, 'warning', 'warnings')
	},
	importOpts: {
		heading: 'Import options',
		headerRow: 'Header row',
		noHeader: 'No header (columns A, B, C…)',
		skipEmpty: 'Skip empty rows',
		nested: 'JSON: nested objects as columns a.b.c',
		reapply: 'Reapply',
		reapplyConfirm: 'Reapplying reads the files again and discards their edits. Continue?',
		reapplied: 'Files read again'
	},
	warnings: {
		cellErrors: (sheet: string, n: number) => `${sheet}: ${plural(n, 'cell', 'cells')} with Excel errors (#N/A, #DIV/0!…) left empty`,
		emptyHeaders: (sheet: string, n: number) => `${sheet}: ${plural(n, 'empty header', 'empty headers')} named after their column letter`,
		duplicateHeaders: (sheet: string, names: string[]) => `${sheet}: repeated headers renamed (${names.join(', ')})`,
		emptySheet: (sheet: string) => `${sheet}: the sheet is empty`,
		decoding: (enc: string) => `Not UTF-8: read as ${enc}`
	},
	errors: {
		xlsxTooManyRows: (d: string) => `Too many rows for an Excel sheet (max 1,048,576): ${d}`,
		xlsxTooManyCols: (d: string) => `Too many columns for an Excel sheet (max 16,384): ${d}`,
		jsonShape: 'The JSON must be an array, an object of arrays or an object',
		jsonSyntax: (d: string) => `Invalid JSON: ${d}`,
		xlsxRead: 'The file is not a valid .xlsx workbook (an old .xls? save it as .xlsx)',
		unsupported: 'Format not supported',
		empty: 'Nothing to export: select at least one sheet',
		generic: 'Something went wrong',
		cancelled: 'Cancelled',
		worker: 'The converter stopped unexpectedly'
	},
	grid: {
		label: (name: string) => `Data of ${name}`,
		rowsOf: (visible: number, total: number) => (visible === total ? plural(total, 'row', 'rows') : `${visible.toLocaleString('en')} of ${plural(total, 'row', 'rows')}`),
		emptyTable: 'This sheet has no data yet',
		addRow: 'Add row',
		addCol: 'Add column',
		insertAbove: 'Insert row above',
		insertBelow: 'Insert row below',
		insertLeft: 'Insert column left',
		insertRight: 'Insert column right',
		deleteRows: (n: number) => (n === 1 ? 'Delete row' : `Delete ${n.toLocaleString('en')} rows`),
		deleteCols: (n: number) => (n === 1 ? 'Delete column' : `Delete ${n} columns`),
		moveLeft: 'Move left',
		moveRight: 'Move right',
		moveUp: 'Move up',
		moveDown: 'Move down',
		renameCol: 'Rename column',
		renamePrompt: 'Column name',
		renameEmpty: 'The name cannot be empty',
		renameDup: 'Another column already has that name',
		newColumn: 'column',
		colType: 'Column type',
		types: { auto: 'Automatic', text: 'Text', number: 'Number', boolean: 'Yes/no', date: 'Date' } satisfies Record<ColType, string>,
		typeHint: (type: string) => `Doesn't match the column type (${type}): it will be exported as is`,
		sortAsc: 'Sort A → Z',
		sortDesc: 'Sort Z → A',
		sorted: 'Rows sorted',
		filter: 'Filter',
		filterQuery: 'Contains…',
		filterValues: 'Values',
		filterTruncated: 'Showing the first 1,000 values',
		filterEmpty: '(empty)',
		clearFilter: 'Clear filter',
		clearFilters: 'Clear filters',
		filtered: (n: number) => `${plural(n, 'filter', 'filters')} active`,
		copy: 'Copy',
		cut: 'Cut',
		paste: 'Paste',
		pasteDenied: 'The browser blocked reading the clipboard: use Ctrl+V',
		clearCells: 'Clear contents',
		selectAll: 'Select all',
		find: 'Find and replace',
		cellMenu: 'Cell actions',
		colMenu: (name: string) => `Actions for column ${name}`,
		sel: (rows: number, cols: number) => `${rows.toLocaleString('en')} × ${cols}`,
		shortcuts: 'Shortcuts',
		shortcutList: [
			['Type / F2 / Enter', 'Edit the cell'],
			['Esc', 'Cancel editing'],
			['Ctrl+Z / Ctrl+Y', 'Undo / redo'],
			['Ctrl+C / X / V', 'Copy, cut, paste (works with Excel)'],
			['Ctrl+F / Ctrl+H', 'Find / replace'],
			['Delete', 'Clear the selection'],
			['Ctrl+Space / Shift+Space', 'Select column / row'],
			['Alt+↑ ↓ ← →', 'Move rows or columns']
		] as [string, string][],
		widthReset: 'Double-click to fit'
	},
	find: {
		find: 'Find',
		replace: 'Replace with',
		matchCase: 'Match case',
		wholeCell: 'Whole cell',
		regex: 'Regular expression',
		allColumns: 'All columns',
		results: (i: number, n: number) => (n ? `${i} of ${n.toLocaleString('en')}` : 'No results'),
		invalidRegex: 'Invalid regular expression',
		next: 'Next',
		prev: 'Previous',
		replaceOne: 'Replace',
		replaceAll: 'Replace all',
		replaced: (n: number) => (n === 1 ? '1 replacement' : `${n.toLocaleString('en')} replacements`)
	},
	export: {
		heading: 'Export',
		format: 'Format',
		formats: { json: 'JSON', xlsx: 'Excel', csv: 'CSV / TSV' } satisfies Record<OutputFormat, string>,
		mode: 'Output',
		modes: {
			split: 'One file per sheet',
			perFile: 'One file per input file',
			combined: 'Everything in one file',
			merged: 'Merge into one table'
		} satisfies Record<OutputMode, string>,
		modeHelp: {
			json: {
				split: 'An array of objects per sheet. Several files go in a .zip.',
				perFile: 'One JSON per file: {"Sheet": [rows]}.',
				combined: 'A single JSON with every sheet: {"Sheet": [rows]}.',
				merged: 'A single array with the rows of every sheet (columns are joined by name).'
			},
			xlsx: {
				split: 'One workbook per sheet.',
				perFile: 'One workbook per input file, with its sheets.',
				combined: 'A single workbook with one sheet per table.',
				merged: 'A single sheet with the rows of every table (columns are joined by name).'
			},
			csv: {
				split: 'One file per sheet. Several files go in a .zip.',
				perFile: '',
				combined: '',
				merged: 'A single file with the rows of every sheet (columns are joined by name).'
			}
		} satisfies Record<OutputFormat, Record<OutputMode, string>>,
		options: 'Options',
		allText: 'Everything as text',
		allTextHelp: 'Like the classic excel2json: numbers, dates and booleans become strings.',
		nested: 'Columns a.b.c → nested objects',
		indent: 'Indentation',
		indentCompact: 'Compact',
		indentSpaces: (n: number) => `${n} spaces`,
		dropNulls: 'Leave out empty values',
		delimiter: 'Separator',
		delimiters: { comma: 'Comma (,)', semicolon: 'Semicolon (;) · decimal comma', tab: 'Tab (TSV)' },
		bom: 'Add BOM (Excel reads accents correctly)',
		styled: 'Formatted: bold header, filters, frozen row, column widths',
		origin: 'Add a column with the origin of each row',
		originName: '_source',
		onlyVisible: 'Only filtered rows',
		selection: (n: number, rows: number) => `${plural(n, 'table', 'tables')} · ${plural(rows, 'row', 'rows')}`,
		download: 'Download',
		copy: 'Copy',
		copyHelp: 'Copy the result to the clipboard',
		copied: 'Copied to the clipboard',
		copyMany: 'The result has several files: download it instead',
		building: (done: number, total: number) => `Converting ${done} / ${total}…`,
		cancel: 'Cancel',
		downloaded: (name: string) => `Downloaded ${name}`,
		summary: (files: number, size: string) => `${plural(files, 'file', 'files')} · ${size}`,
		reset: 'Reset options'
	},
	autosave: {
		saved: 'Saved on this device',
		saving: 'Saving…',
		dirty: 'Unsaved changes',
		error: 'Could not save the session (storage full?)',
		quota: 'Little storage left on this device: the session may not be saved'
	},
	footer: {
		version: (v: string) => `Version ${v}`,
		source: 'Source code',
		offline: 'Works offline'
	},
	sample: {
		file: 'example.xlsx'
	}
};

export type Dict = typeof en;
export default en;
