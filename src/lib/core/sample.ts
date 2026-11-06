// Libro de ejemplo («Probar con un ejemplo»): muestra tipos, fechas, columnas anidadas a.b y varias hojas.
import type { SheetOut } from './xlsx.ts';

export function sampleSheets(locale: 'es' | 'en'): SheetOut[] {
	const es = locale === 'es';
	const cities = es ? ['Santiago', 'Valparaíso', 'Concepción', 'La Serena', 'Temuco'] : ['Boston', 'Denver', 'Austin', 'Seattle', 'Miami'];
	const names = es ? ['Ana', 'Bruno', 'Carla', 'Diego', 'Elena', 'Fabián', 'Gloria', 'Héctor'] : ['Ann', 'Ben', 'Carla', 'Dan', 'Eve', 'Frank', 'Grace', 'Hank'];
	const products = es
		? [['P-01', 'Teclado', 24990], ['P-02', 'Ratón', 12990], ['P-03', 'Monitor 24"', 149990], ['P-04', 'Audífonos', 39990], ['P-05', 'Cable USB-C', 6990]]
		: [['P-01', 'Keyboard', 49.9], ['P-02', 'Mouse', 19.5], ['P-03', '24" monitor', 179], ['P-04', 'Headphones', 59.99], ['P-05', 'USB-C cable', 9.99]];
	const orders: SheetOut['rows'] = [];
	for (let i = 0; i < 60; i++) {
		const p = products[(i * 7) % products.length];
		const qty = 1 + ((i * 5) % 4);
		const day = 1 + ((i * 3) % 28);
		const month = 1 + (i % 6);
		orders.push([
			1000 + i,
			`2025-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
			names[i % names.length],
			cities[(i * 3) % cities.length],
			p[0] as string,
			qty,
			Math.round(qty * (p[2] as number) * 100) / 100,
			i % 3 !== 0,
			i % 11 === 0 ? (es ? 'Entregar por la tarde' : 'Deliver in the afternoon') : null
		]);
	}
	return [
		{
			name: es ? 'Pedidos' : 'Orders',
			headers: es
				? ['id', 'fecha', 'cliente.nombre', 'cliente.ciudad', 'producto', 'cantidad', 'total', 'pagado', 'nota']
				: ['id', 'date', 'customer.name', 'customer.city', 'product', 'quantity', 'total', 'paid', 'note'],
			types: ['auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto', 'auto'],
			rows: orders
		},
		{
			name: es ? 'Productos' : 'Products',
			headers: es ? ['código', 'nombre', 'precio'] : ['code', 'name', 'price'],
			types: ['auto', 'auto', 'auto'],
			rows: products
		}
	];
}
