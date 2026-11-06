<script lang="ts">
	// Pegar JSON, CSV o una tabla copiada de Excel (TSV) como un archivo nuevo.
	import { tick } from 'svelte';
	import { t } from '../i18n/index.svelte.ts';

	interface Props {
		open: boolean;
		onadd: (file: File) => void;
	}

	let { open = $bindable(), onadd }: Props = $props();

	let text = $state('');
	let name = $state('');
	let error = $state('');
	let area = $state<HTMLTextAreaElement>();

	$effect(() => {
		if (!open) return;
		name = t().paste.defaultName;
		error = '';
		tick().then(() => area?.focus());
	});

	/** Extensión según el contenido: JSON si empieza con [ o {, TSV si trae tabulaciones (copiado de Excel), si no CSV. */
	function extFor(s: string) {
		const head = s.trimStart();
		if (head.startsWith('[') || head.startsWith('{')) return 'json';
		return s.includes('\t') ? 'tsv' : 'csv';
	}

	function submit(e: Event) {
		e.preventDefault();
		if (!text.trim()) {
			error = t().paste.empty;
			return;
		}
		const base = name.trim() || t().paste.defaultName;
		onadd(new File([text], `${base}.${extFor(text)}`, { type: 'text/plain' }));
		text = '';
		open = false;
	}
</script>

{#if open}
	<div class="fixed inset-0 z-[80] grid place-items-center bg-black/40 p-4" role="presentation" onmousedown={(e) => e.target === e.currentTarget && (open = false)}>
		<div class="w-full max-w-2xl" role="dialog" aria-modal="true" aria-labelledby="paste-title">
		<form class="panel flex flex-col gap-3 p-5 shadow-2xl" onsubmit={submit}>
			<h2 id="paste-title" class="text-base font-semibold">{t().paste.title}</h2>
			<textarea
				bind:this={area}
				bind:value={text}
				class="input h-64 resize-y font-mono text-xs"
				placeholder={t().paste.placeholder}
				spellcheck="false"
				aria-invalid={!!error}
				onkeydown={(e) => e.key === 'Escape' && (open = false)}
			></textarea>
			{#if error}<p class="text-xs text-bad">{error}</p>{/if}
			<div class="flex flex-wrap items-center gap-2">
				<label class="flex items-center gap-2 text-sm">
					{t().paste.name}
					<input class="input w-48 py-1" bind:value={name} />
				</label>
				<div class="ml-auto flex gap-2">
					<button type="button" class="btn" onclick={() => (open = false)}>{t().common.cancel}</button>
					<button type="submit" class="btn btn-primary">{t().paste.add}</button>
				</div>
			</div>
		</form>
		</div>
	</div>
{/if}
