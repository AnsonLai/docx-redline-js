/**
 * `@ansonlai/docx-redline-js/word`: the Word host surface. A Word add-in reads
 * `body.getOoxml()`, passes it here, and writes back the one returned package.
 */
export { inspectWordFlatOpc, prepareWordFlatOpcBatch } from './flat-opc-body.js';
export {
    detectWordExportTailIndex,
    computeWordSourceRevisionToken,
    computeWordRestorationToken
} from './export-identity.js';
