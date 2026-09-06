/**
 * Backwards-compatible public facade for standalone document operations.
 *
 * Keep this module intentionally small: implementation lives behind focused
 * internal boundaries while this path and its declarations remain stable.
 */

export { preflightOperations } from './operation-preflight.js';
export { applyOperationToDocumentXml } from './document-operation-applier.js';
export {
    applyOperationsToDocumentXml,
    orderOperationsForStableTargets,
    buildOperationDependencyPlan
} from './batch-operation-orchestrator.js';
