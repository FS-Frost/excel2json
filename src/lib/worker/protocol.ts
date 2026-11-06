// Mensajes entre la página y el Worker de conversión.
import type { ExportOptions, ImportOptions, OutputFile, SourceFile, Table } from '../core/types.ts';

export type Request =
	| { id: number; type: 'parse'; name: string; bytes: ArrayBuffer; opts: ImportOptions }
	| { id: number; type: 'build'; tables: Table[]; opts: ExportOptions }
	| { id: number; type: 'sample'; locale: 'es' | 'en' };

export type BuildResult = { file: OutputFile; count: number };

export type Results = { parse: SourceFile; build: BuildResult; sample: ArrayBuffer };

export type SerializedError = { code: string; detail: string; message: string };

export type Response =
	| { id: number; type: 'progress'; done: number; total: number; label: string }
	| { id: number; type: 'ok'; result: Results[keyof Results] }
	| { id: number; type: 'error'; error: SerializedError };
