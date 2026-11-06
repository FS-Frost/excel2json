// Deshacer/rehacer: cada paso guarda las operaciones aplicadas y sus inversas (no copias de la tabla).
import type { Op } from './ops.ts';

export interface Step {
	tableId: string;
	/** Operaciones tal como se aplicaron (para rehacer). */
	ops: Op[];
	/** Inversas, en el orden en que se aplican para deshacer. */
	inverse: Op[];
}

export const HISTORY_LIMIT = 200;

export class History {
	undoStack: Step[] = [];
	redoStack: Step[] = [];

	constructor(readonly limit = HISTORY_LIMIT) {}

	push(step: Step) {
		this.undoStack.push(step);
		if (this.undoStack.length > this.limit) this.undoStack.shift();
		this.redoStack = [];
	}

	/** Paso a deshacer (sale de la pila); quien lo aplica llama a `pushRedo` con las inversas nuevas. */
	popUndo(): Step | undefined {
		return this.undoStack.pop();
	}

	popRedo(): Step | undefined {
		return this.redoStack.pop();
	}

	pushRedo(step: Step) {
		this.redoStack.push(step);
	}

	pushUndoKeepRedo(step: Step) {
		this.undoStack.push(step);
		if (this.undoStack.length > this.limit) this.undoStack.shift();
	}

	/** Olvida los pasos de una tabla (al quitarla o reimportarla). */
	forget(tableId: string) {
		this.undoStack = this.undoStack.filter((s) => s.tableId !== tableId);
		this.redoStack = this.redoStack.filter((s) => s.tableId !== tableId);
	}

	clear() {
		this.undoStack = [];
		this.redoStack = [];
	}

	get canUndo() {
		return this.undoStack.length > 0;
	}
	get canRedo() {
		return this.redoStack.length > 0;
	}
}
