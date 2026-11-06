/// <reference lib="webworker" />
// Worker de conversión: lee archivos y genera salidas fuera del hilo principal (la página nunca se congela).
// Cancelar = la página termina el Worker y crea otro.
import { ConvertError } from '../core/types.ts';
import { readSource } from '../core/read.ts';
import { buildOutputs, bundle } from '../core/outputs.ts';
import { sampleSheets } from '../core/sample.ts';
import { tablesToXlsx } from '../core/xlsx.ts';
import type { Request, Response, SerializedError } from './protocol.ts';

const ctx = self as unknown as DedicatedWorkerGlobalScope;

function post(msg: Response, transfer: Transferable[] = []) {
	ctx.postMessage(msg, transfer);
}

function serialize(e: unknown): SerializedError {
	if (e instanceof ConvertError) return { code: e.code, detail: e.detail, message: e.message };
	return { code: 'generic', detail: '', message: String((e as Error)?.message ?? e) };
}

ctx.onmessage = async (ev: MessageEvent<Request>) => {
	const req = ev.data;
	const progress = (done: number, total: number, label = '') => post({ id: req.id, type: 'progress', done, total, label });
	try {
		switch (req.type) {
			case 'parse': {
				const file = await readSource(req.name, new Uint8Array(req.bytes), req.opts, progress);
				post({ id: req.id, type: 'ok', result: file });
				break;
			}
			case 'build': {
				const files = await buildOutputs(req.tables, req.opts, progress);
				const file = bundle(files, req.opts);
				post({ id: req.id, type: 'ok', result: { file, count: files.length } }, [file.bytes.buffer as ArrayBuffer]);
				break;
			}
			case 'sample': {
				const bytes = await tablesToXlsx(sampleSheets(req.locale), { styled: true, allText: false });
				const buf = bytes.buffer as ArrayBuffer;
				post({ id: req.id, type: 'ok', result: buf }, [buf]);
				break;
			}
		}
	} catch (e) {
		post({ id: req.id, type: 'error', error: serialize(e) });
	}
};
