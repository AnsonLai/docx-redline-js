/**
 * Accuracy-oriented capability record for the current reconciliation routes.
 * This is internal instrumentation, not a routing policy API.
 */
export const RECONCILIATION_CAPABILITY_MATRIX = Object.freeze({
    formatOnly: Object.freeze({ paragraphs: true, formatting: true, tables: 'scoped', hyperlinks: 'preserved', fields: 'preserved', comments: 'preserved' }),
    surgical: Object.freeze({ paragraphs: true, tables: 'cell-scoped', hyperlinks: 'preserved', fields: 'preserved', comments: 'preserved', notes: 'preserved' }),
    reconstruction: Object.freeze({ paragraphs: true, hyperlinks: 'sentinel-preserved', fields: 'sentinel-preserved', comments: 'marker-preserved', notes: 'reference-preserved' }),
    table: Object.freeze({ tables: true, paragraphs: true, formatting: 'cell-dependent', numbering: false }),
    listDirect: Object.freeze({ lists: true, numbering: true, paragraphs: 'single-source expansion', tables: 'embedded-markdown blocks', formatting: 'markdown hints' }),
    listCompatibilityPipeline: Object.freeze({ lists: true, numbering: true, paragraphs: 'multi-source patching', compatibility: true })
});

export function recordRouteSelection(options, route, context = {}) {
    const callback = options?._routeInstrumentation?.onRoute;
    if (typeof callback !== 'function') return;
    callback(Object.freeze({
        route,
        capabilities: RECONCILIATION_CAPABILITY_MATRIX[route] || null,
        ...context
    }));
}

export function createRouteFrequencyCollector() {
    const counts = new Map();
    return {
        onRoute(event) {
            const route = event?.route || 'unknown';
            counts.set(route, (counts.get(route) || 0) + 1);
        },
        snapshot() {
            return Object.fromEntries(Array.from(counts).sort(([a], [b]) => a.localeCompare(b)));
        }
    };
}
