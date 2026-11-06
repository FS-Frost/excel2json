import 'fake-indexeddb/auto';
import { describe, expect, test } from 'bun:test';
import { clearSession, loadMeta, loadTables, saveSession, type SessionMeta } from '../src/lib/storage/db.ts';
import { makeTable } from './factories/index.ts';

const meta = (ids: string[]): SessionMeta => ({
	v: 1,
	savedAt: 1,
	files: [{ id: 'f1', name: 'a.xlsx', kind: 'xlsx', size: 10, warnings: [], tableIds: ids }],
	selected: Object.fromEntries(ids.map((id) => [id, true])),
	activeId: ids[0] ?? null,
	views: {},
	edited: []
});

describe('session storage', () => {
	test('saves only dirty tables and prunes removed ones', async () => {
		const a = makeTable(['x'], [[1]], { id: 'A' });
		const b = makeTable(['y'], [['z']], { id: 'B' });
		await saveSession(meta(['A', 'B']), [a, b], new Set(['A', 'B']));
		expect((await loadMeta())?.files[0].tableIds).toEqual(['A', 'B']);
		expect([...(await loadTables(['A', 'B'])).keys()]).toEqual(['A', 'B']);

		a.rows[0][0] = 2;
		await saveSession(meta(['A']), [a], new Set(['A']));
		const tables = await loadTables(['A', 'B']);
		expect([...tables.keys()]).toEqual(['A']);
		expect(tables.get('A')?.rows).toEqual([[2]]);

		await clearSession();
		expect(await loadMeta()).toBeNull();
	});
});
