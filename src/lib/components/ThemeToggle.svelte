<script lang="ts">
	// Rota el tema: sistema → claro → oscuro → sistema. Origen: FS-Frost/mai-docs.
	import { setTheme, theme, type ThemeMode } from '../client/theme.svelte.ts';
	import { t } from '../i18n/index.svelte.ts';
	import Icon from './Icon.svelte';

	const NEXT: Record<ThemeMode, ThemeMode> = { system: 'light', light: 'dark', dark: 'system' };
	const ICON = { system: 'monitor', light: 'sun', dark: 'moon' } as const;
	const name = (m: ThemeMode) => (m === 'light' ? t().header.themeLight : m === 'dark' ? t().header.themeDark : t().header.themeSystem);
	const label = $derived(t().header.themeToggle(name(theme.mode), name(NEXT[theme.mode])));
</script>

<button class="tb" title={label} aria-label={label} onclick={() => setTheme(NEXT[theme.mode])}>
	<Icon name={ICON[theme.mode]} />
</button>
