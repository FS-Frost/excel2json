// Textos en español. Misma forma que en.ts (`Dict`): TypeScript falla si falta o sobra una clave.
import type { Dict } from './en.ts';

const num = (n: number) => n.toLocaleString('es');
const plural = (n: number, one: string, many: string) => `${num(n)} ${n === 1 ? one : many}`;

const es: Dict = {
	common: {
		close: 'Cerrar',
		cancel: 'Cancelar',
		ok: 'Aceptar',
		retry: 'Reintentar',
		undo: 'Deshacer',
		redo: 'Rehacer',
		delete: 'Eliminar',
		rename: 'Renombrar',
		clear: 'Limpiar',
		apply: 'Aplicar',
		reset: 'Restablecer',
		all: 'Todas',
		none: 'Ninguna',
		more: 'Más',
		loading: 'Cargando…'
	},
	app: {
		name: 'excel2json',
		tagline: 'Convierte y edita Excel, CSV y JSON en tu navegador. No se sube nada.',
		privacy: 'Tus archivos nunca salen de este dispositivo.'
	},
	header: {
		language: 'Idioma',
		langSystem: 'Idioma del sistema',
		themeLight: 'Claro',
		themeDark: 'Oscuro',
		themeSystem: 'Sistema',
		themeToggle: (current, next) => `Tema: ${current} (cambiar a ${next})`,
		github: 'Código fuente en GitHub'
	},
	net: {
		offline: 'Sin conexión',
		offlineDetail: 'excel2json sigue funcionando: todo ocurre en este dispositivo.',
		online: 'Conexión recuperada'
	},
	update: {
		available: 'Hay una versión nueva de excel2json',
		action: 'Actualizar',
		updating: 'Actualizando…',
		updatingDetail: 'Tu trabajo está guardado en este dispositivo.',
		publishing: 'La versión nueva aún se está publicando',
		publishingDetail: 'Vuelve a intentarlo en un minuto.',
		offline: 'No se pudo buscar la actualización (¿sin conexión?)',
		failed: (v) => `La actualización a ${v} no se cargó`,
		failedDetail: 'Fuerza la recarga desde la red (tu sesión se conserva).',
		force: 'Forzar',
		done: (v) => `excel2json actualizado (${v})`,
		publishedAgo: (when) => `Publicado ${when}`
	},
	time: { justNow: 'hace un momento' },
	drop: {
		title: 'Suelta tus archivos aquí',
		subtitle: 'Excel (.xlsx), CSV, TSV o JSON. Varios a la vez, con todas sus hojas.',
		choose: 'Elegir archivos',
		paste: 'Pegar datos',
		sample: 'Probar con un ejemplo',
		overlay: 'Suelta para agregar',
		reading: (name) => `Leyendo ${name}…`,
		readingSheet: (name, sheet) => `Leyendo ${name} › ${sheet}`,
		unsupported: (name) => `${name}: formato no soportado (usa .xlsx, .csv, .tsv o .json)`,
		readFailed: (name) => `No se pudo leer ${name}`,
		added: (n) => (n === 1 ? 'Archivo agregado' : `${n} archivos agregados`),
		bigFile: (name, size) => `${name} es grande (${size}): puede tardar un poco`
	},
	paste: {
		title: 'Pega JSON, CSV o una tabla copiada de Excel',
		placeholder: '[{"id": 1, "nombre": "Ana"}]  o  id;nombre…',
		name: 'Nombre',
		defaultName: 'pegado',
		add: 'Agregar',
		empty: 'Primero pega algo'
	},
	restore: {
		title: 'Sesión anterior',
		detail: (files, when) => `${plural(files, 'archivo', 'archivos')}, guardada ${when}.`,
		restore: 'Restaurar',
		discard: 'Descartar',
		replacedWarn: 'Si agregas archivos se empieza una sesión nueva y reemplaza la guardada.',
		restored: 'Sesión restaurada',
		noOriginal: 'Restaurado de la sesión guardada: ya no se pueden reaplicar las opciones de importación a este archivo.'
	},
	files: {
		heading: 'Archivos',
		addMore: 'Agregar archivos',
		sheets: (n) => plural(n, 'hoja', 'hojas'),
		rows: (n) => plural(n, 'fila', 'filas'),
		cols: (n) => plural(n, 'columna', 'columnas'),
		remove: 'Quitar archivo',
		removeConfirm: (name) => `¿Quitar ${name} y sus ediciones?`,
		clearAll: 'Cerrar todo',
		clearConfirm: '¿Cerrar todos los archivos? Se perderán las ediciones no exportadas.',
		selectAll: 'Seleccionar todas las hojas',
		selectNone: 'Quitar selección',
		include: (name) => `Incluir ${name} en la exportación`,
		selected: (n, total) => `${n} de ${total} seleccionadas`,
		edited: 'Editado',
		warnings: (n) => plural(n, 'aviso', 'avisos')
	},
	importOpts: {
		heading: 'Opciones de importación',
		headerRow: 'Fila de cabecera',
		noHeader: 'Sin cabecera (columnas A, B, C…)',
		skipEmpty: 'Omitir filas vacías',
		nested: 'JSON: objetos anidados como columnas a.b.c',
		reapply: 'Reaplicar',
		reapplyConfirm: 'Reaplicar vuelve a leer los archivos y descarta sus ediciones. ¿Continuar?',
		reapplied: 'Archivos leídos de nuevo'
	},
	warnings: {
		cellErrors: (sheet, n) => `${sheet}: ${plural(n, 'celda', 'celdas')} con error de Excel (#N/A, #DIV/0!…) quedaron vacías`,
		emptyHeaders: (sheet, n) => `${sheet}: ${plural(n, 'cabecera vacía', 'cabeceras vacías')}, nombradas con la letra de su columna`,
		duplicateHeaders: (sheet, names) => `${sheet}: cabeceras repetidas renombradas (${names.join(', ')})`,
		emptySheet: (sheet) => `${sheet}: la hoja está vacía`,
		decoding: (enc) => `No es UTF-8: se leyó como ${enc}`
	},
	errors: {
		xlsxTooManyRows: (d) => `Demasiadas filas para una hoja de Excel (máx. 1.048.576): ${d}`,
		xlsxTooManyCols: (d) => `Demasiadas columnas para una hoja de Excel (máx. 16.384): ${d}`,
		jsonShape: 'El JSON debe ser un array, un objeto de arrays o un objeto',
		jsonSyntax: (d) => `JSON inválido: ${d}`,
		xlsxRead: 'El archivo no es un libro .xlsx válido (¿un .xls antiguo? guárdalo como .xlsx)',
		unsupported: 'Formato no soportado',
		empty: 'Nada que exportar: selecciona al menos una hoja',
		generic: 'Algo salió mal',
		cancelled: 'Cancelado',
		worker: 'El conversor se detuvo inesperadamente'
	},
	grid: {
		label: (name) => `Datos de ${name}`,
		rowsOf: (visible, total) => (visible === total ? plural(total, 'fila', 'filas') : `${num(visible)} de ${plural(total, 'fila', 'filas')}`),
		emptyTable: 'Esta hoja aún no tiene datos',
		addRow: 'Agregar fila',
		addCol: 'Agregar columna',
		insertAbove: 'Insertar fila arriba',
		insertBelow: 'Insertar fila abajo',
		insertLeft: 'Insertar columna a la izquierda',
		insertRight: 'Insertar columna a la derecha',
		deleteRows: (n) => (n === 1 ? 'Eliminar fila' : `Eliminar ${num(n)} filas`),
		deleteCols: (n) => (n === 1 ? 'Eliminar columna' : `Eliminar ${n} columnas`),
		moveLeft: 'Mover a la izquierda',
		moveRight: 'Mover a la derecha',
		moveUp: 'Mover arriba',
		moveDown: 'Mover abajo',
		renameCol: 'Renombrar columna',
		renamePrompt: 'Nombre de la columna',
		renameEmpty: 'El nombre no puede quedar vacío',
		renameDup: 'Otra columna ya tiene ese nombre',
		newColumn: 'columna',
		colType: 'Tipo de columna',
		types: { auto: 'Automático', text: 'Texto', number: 'Número', boolean: 'Sí/no', date: 'Fecha' },
		typeHint: (type) => `No calza con el tipo de la columna (${type}): se exportará tal cual`,
		sortAsc: 'Ordenar A → Z',
		sortDesc: 'Ordenar Z → A',
		sorted: 'Filas ordenadas',
		filter: 'Filtrar',
		filterQuery: 'Contiene…',
		filterValues: 'Valores',
		filterTruncated: 'Se muestran los primeros 1.000 valores',
		filterEmpty: '(vacío)',
		clearFilter: 'Quitar filtro',
		clearFilters: 'Quitar filtros',
		filtered: (n) => (n === 1 ? '1 filtro activo' : `${n} filtros activos`),
		copy: 'Copiar',
		cut: 'Cortar',
		paste: 'Pegar',
		pasteDenied: 'El navegador bloqueó la lectura del portapapeles: usa Ctrl+V',
		clearCells: 'Borrar contenido',
		selectAll: 'Seleccionar todo',
		find: 'Buscar y reemplazar',
		cellMenu: 'Acciones de la celda',
		colMenu: (name) => `Acciones de la columna ${name}`,
		sel: (rows, cols) => `${num(rows)} × ${cols}`,
		shortcuts: 'Atajos',
		shortcutList: [
			['Escribir / F2 / Enter', 'Editar la celda'],
			['Esc', 'Cancelar la edición'],
			['Ctrl+Z / Ctrl+Y', 'Deshacer / rehacer'],
			['Ctrl+C / X / V', 'Copiar, cortar, pegar (compatible con Excel)'],
			['Ctrl+F / Ctrl+H', 'Buscar / reemplazar'],
			['Supr', 'Borrar la selección'],
			['Ctrl+Espacio / Mayús+Espacio', 'Seleccionar columna / fila'],
			['Alt+↑ ↓ ← →', 'Mover filas o columnas']
		],
		widthReset: 'Doble clic para ajustar'
	},
	find: {
		find: 'Buscar',
		replace: 'Reemplazar por',
		matchCase: 'Mayúsculas/minúsculas',
		wholeCell: 'Celda completa',
		regex: 'Expresión regular',
		allColumns: 'Todas las columnas',
		results: (i, n) => (n ? `${i} de ${num(n)}` : 'Sin resultados'),
		invalidRegex: 'Expresión regular inválida',
		next: 'Siguiente',
		prev: 'Anterior',
		replaceOne: 'Reemplazar',
		replaceAll: 'Reemplazar todo',
		replaced: (n) => (n === 1 ? '1 reemplazo' : `${num(n)} reemplazos`)
	},
	export: {
		heading: 'Exportar',
		format: 'Formato',
		formats: { json: 'JSON', xlsx: 'Excel', csv: 'CSV / TSV' },
		mode: 'Salida',
		modes: {
			split: 'Un archivo por hoja',
			perFile: 'Un archivo por archivo de entrada',
			combined: 'Todo en un archivo',
			merged: 'Unir en una sola tabla'
		},
		modeHelp: {
			json: {
				split: 'Un array de objetos por hoja. Si son varios archivos, van en un .zip.',
				perFile: 'Un JSON por archivo: {"Hoja": [filas]}.',
				combined: 'Un solo JSON con todas las hojas: {"Hoja": [filas]}.',
				merged: 'Un solo array con las filas de todas las hojas (las columnas se juntan por nombre).'
			},
			xlsx: {
				split: 'Un libro por hoja.',
				perFile: 'Un libro por archivo de entrada, con sus hojas.',
				combined: 'Un solo libro con una hoja por tabla.',
				merged: 'Una sola hoja con las filas de todas las tablas (las columnas se juntan por nombre).'
			},
			csv: {
				split: 'Un archivo por hoja. Si son varios, van en un .zip.',
				perFile: '',
				combined: '',
				merged: 'Un solo archivo con las filas de todas las hojas (las columnas se juntan por nombre).'
			}
		},
		options: 'Opciones',
		allText: 'Todo como texto',
		allTextHelp: 'Como el excel2json clásico: números, fechas y booleanos pasan a texto.',
		nested: 'Columnas a.b.c → objetos anidados',
		indent: 'Indentación',
		indentCompact: 'Compacto',
		indentSpaces: (n) => `${n} espacios`,
		dropNulls: 'Omitir valores vacíos',
		delimiter: 'Separador',
		delimiters: { comma: 'Coma (,)', semicolon: 'Punto y coma (;) · coma decimal', tab: 'Tabulación (TSV)' },
		bom: 'Agregar BOM (Excel lee bien los acentos)',
		styled: 'Con formato: cabecera en negrita, filtros, fila fija y anchos',
		origin: 'Agregar una columna con el origen de cada fila',
		originName: '_origen',
		onlyVisible: 'Solo filas filtradas',
		selection: (n, rows) => `${plural(n, 'tabla', 'tablas')} · ${plural(rows, 'fila', 'filas')}`,
		download: 'Descargar',
		copy: 'Copiar',
		copyHelp: 'Copiar el resultado al portapapeles',
		copied: 'Copiado al portapapeles',
		copyMany: 'El resultado tiene varios archivos: descárgalo',
		building: (done, total) => `Convirtiendo ${done} / ${total}…`,
		cancel: 'Cancelar',
		downloaded: (name) => `Descargado ${name}`,
		summary: (files, size) => `${plural(files, 'archivo', 'archivos')} · ${size}`,
		reset: 'Restablecer opciones'
	},
	autosave: {
		saved: 'Guardado en este dispositivo',
		saving: 'Guardando…',
		dirty: 'Cambios sin guardar',
		error: 'No se pudo guardar la sesión (¿almacenamiento lleno?)',
		quota: 'Queda poco espacio en este dispositivo: la sesión podría no guardarse'
	},
	footer: {
		version: (v) => `Versión ${v}`,
		source: 'Código fuente',
		offline: 'Funciona sin conexión'
	},
	sample: {
		file: 'ejemplo.xlsx'
	}
};

export default es;
