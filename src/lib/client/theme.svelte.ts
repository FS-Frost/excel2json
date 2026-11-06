// Tema por dispositivo (localStorage['excel2json.theme']; ausente = sistema). app.html lo aplica antes de pintar.
// data-theme siempre es concreto ('light' | 'dark'): «sistema» se resuelve aquí y sigue al SO. Origen: FS-Frost/mai-docs.

export type ThemeMode = 'light' | 'dark' | 'system';

export const THEME_KEY = 'excel2json.theme';
const META_COLOR = { light: '#ffffff', dark: '#1a1d23' } as const;

export const theme = $state({ mode: 'system' as ThemeMode, resolved: 'light' as 'light' | 'dark' });

function read(): ThemeMode {
	try {
		const v = localStorage.getItem(THEME_KEY);
		return v === 'light' || v === 'dark' ? v : 'system';
	} catch {
		return 'system';
	}
}

const systemDark = () => matchMedia('(prefers-color-scheme: dark)').matches;

function apply() {
	theme.resolved = theme.mode === 'system' ? (systemDark() ? 'dark' : 'light') : theme.mode;
	document.documentElement.dataset.theme = theme.resolved;
	document.querySelector('meta[name="theme-color"]')?.setAttribute('content', META_COLOR[theme.resolved]);
}

let listening = false;

export function initTheme() {
	theme.mode = read();
	apply();
	if (!listening) {
		listening = true;
		matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => theme.mode === 'system' && apply());
	}
}

export function setTheme(mode: ThemeMode) {
	theme.mode = mode;
	try {
		if (mode === 'system') localStorage.removeItem(THEME_KEY);
		else localStorage.setItem(THEME_KEY, mode);
	} catch {
		/* no storage: applies to this session only */
	}
	apply();
}
