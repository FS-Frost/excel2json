// Compartido por la página y el service worker (sin DOM). Origen: FS-Frost/mai-docs.

export const CACHE_PREFIX = 'excel2json-';
/** Cachés que sobreviven a una versión nueva (el SW borra las demás excel2json-* al activarse). */
export const KEPT_CACHES: string[] = [];
