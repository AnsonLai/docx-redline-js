/**
 * Node.js compatibility layer for DocxDocument.
 *
 * Re-exports the universal DocxDocument and openDocx implementations from
 * document/docx-document.js for downstream callers importing @ansonlai/docx-redline-js/node.
 */

export {
    DocxDocument,
    openDocx,
    computePackageRevisionToken
} from '../document/docx-document.js';
