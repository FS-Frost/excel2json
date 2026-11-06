<script lang="ts">
	import type { Snippet } from 'svelte';
	import { t } from '../i18n/index.svelte.ts';
	import Icon from './Icon.svelte';
	import LangToggle from './LangToggle.svelte';
	import Logo from './Logo.svelte';
	import ThemeToggle from './ThemeToggle.svelte';

	let { status }: { status?: Snippet } = $props();
</script>

<header class="z-30 shrink-0 border-b border-line bg-surface/95 pt-[env(safe-area-inset-top)] backdrop-blur">
	<div class="flex h-12 items-center gap-2 px-3 sm:px-4">
		<a href="./" class="flex items-center gap-2 font-semibold text-fg" aria-label={t().app.name}>
			<Logo size={26} />
			<span class="text-[1.05rem] tracking-tight">{t().app.name}</span>
		</a>
		<div class="ml-2 hidden min-w-0 flex-1 truncate text-xs text-faint md:block">{@render status?.()}</div>
		<nav class="ml-auto flex items-center gap-0.5">
			<LangToggle />
			<ThemeToggle />
			<a class="tb" href="https://github.com/FS-Frost/excel2json" target="_blank" rel="noopener noreferrer" title={t().header.github} aria-label={t().header.github}>
				<Icon name="github" />
			</a>
		</nav>
	</div>
</header>
