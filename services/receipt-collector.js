/**
 * Internal mutation receipt collector for tracking allocated revision,
 * comment, numbering, and relationship IDs per operation.
 */

export class ReceiptCollector {
    constructor() {
        this.receipts = [];
        this.activeReceipt = null;
    }

    beginOperation(operationIndex, operationId = null, authorUsed = null) {
        this.activeReceipt = {
            operationIndex: typeof operationIndex === 'number' ? operationIndex : 1,
            ...(operationId ? { operationId: String(operationId) } : {}),
            attemptedDisposition: 'applied',
            finalDisposition: 'applied',
            committed: true,
            ...(authorUsed ? { authorUsed: String(authorUsed) } : {}),
            revisionItems: [],
            commentIds: [],
            numberingIds: [],
            relationshipIds: [],
            affectedTargets: [],
            warnings: []
        };
    }

    recordRevision(id, kind = 'structural', partName = 'word/document.xml') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.revisionItems.some(item => item.id === strId && item.kind === kind && item.partName === partName)) {
            this.activeReceipt.revisionItems.push({
                id: strId,
                kind,
                partName
            });
        }
    }

    recordComment(id, _partName = 'word/comments.xml') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.commentIds.includes(strId)) {
            this.activeReceipt.commentIds.push(strId);
        }
    }

    recordNumbering(id, _partName = 'word/numbering.xml') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.numberingIds.includes(strId)) {
            this.activeReceipt.numberingIds.push(strId);
        }
    }

    recordRelationship(id, _partName = 'word/_rels/document.xml.rels') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.relationshipIds.includes(strId)) {
            this.activeReceipt.relationshipIds.push(strId);
        }
    }

    recordAffectedTarget(target) {
        if (!this.activeReceipt || !target) return;
        this.activeReceipt.affectedTargets.push(JSON.parse(JSON.stringify(target)));
    }

    recordWarning(warning) {
        if (!this.activeReceipt || !warning) return;
        this.activeReceipt.warnings.push(String(warning));
    }

    commitOperation(disposition = 'applied') {
        if (!this.activeReceipt) return null;
        this.activeReceipt.attemptedDisposition = disposition;
        this.activeReceipt.finalDisposition = disposition;
        this.activeReceipt.committed = disposition === 'applied';
        const committed = JSON.parse(JSON.stringify(this.activeReceipt));
        this.receipts.push(committed);
        this.activeReceipt = null;
        return committed;
    }

    abortOperation(disposition = 'refused') {
        if (!this.activeReceipt) return null;
        this.activeReceipt.attemptedDisposition = disposition;
        this.activeReceipt.finalDisposition = disposition;
        this.activeReceipt.committed = false;
        const aborted = JSON.parse(JSON.stringify(this.activeReceipt));
        this.activeReceipt = null;
        return aborted;
    }

    createSavepoint() {
        return {
            receipts: JSON.parse(JSON.stringify(this.receipts)),
            activeReceipt: this.activeReceipt ? JSON.parse(JSON.stringify(this.activeReceipt)) : null
        };
    }

    restoreSavepoint(savepoint) {
        if (!savepoint) return;
        this.receipts = Array.isArray(savepoint.receipts)
            ? JSON.parse(JSON.stringify(savepoint.receipts))
            : [];
        this.activeReceipt = savepoint.activeReceipt
            ? JSON.parse(JSON.stringify(savepoint.activeReceipt))
            : null;
    }

    clear() {
        this.receipts = [];
        this.activeReceipt = null;
    }

    markRolledBack() {
        for (const receipt of this.receipts) {
            if (receipt.attemptedDisposition === 'applied') {
                receipt.finalDisposition = 'rolled_back';
                receipt.committed = false;
            }
        }
        this.activeReceipt = null;
    }

    getReceipts() {
        return JSON.parse(JSON.stringify(this.receipts));
    }

    getCurrentReceipt() {
        return this.activeReceipt ? JSON.parse(JSON.stringify(this.activeReceipt)) : null;
    }
}

/**
 * Creates an empty receipt structure for unattempted, refused, or no-op operations.
 *
 * @param {number} operationIndex
 * @param {string|null} [operationId]
 * @param {string|null} [authorUsed]
 * @param {'applied'|'no_change'|'refused'|'not_attempted'} [disposition]
 * @returns {Object}
 */
export function createEmptyReceipt(operationIndex, operationId = null, authorUsed = null, disposition = 'not_attempted') {
    return {
        operationIndex: typeof operationIndex === 'number' ? operationIndex : 1,
        ...(operationId ? { operationId: String(operationId) } : {}),
        attemptedDisposition: disposition,
        finalDisposition: disposition,
        committed: false,
        ...(authorUsed ? { authorUsed: String(authorUsed) } : {}),
        revisionItems: [],
        commentIds: [],
        numberingIds: [],
        relationshipIds: [],
        affectedTargets: [],
        warnings: []
    };
}

/**
 * Reconciles committed mutation receipts against the generated/committed output parts.
 *
 * @param {Object} parts
 * @param {string} parts.documentXml - Serialized word/document.xml
 * @param {string|null} [parts.commentsXml] - Serialized word/comments.xml
 * @param {string|null} [parts.numberingXml] - Serialized word/numbering.xml
 * @param {string[]} [parts.numberingXmlParts] - Additional numbering XML parts
 * @param {string|null} [parts.relationshipsXml] - Serialized word/_rels/document.xml.rels
 * @param {Array<Object>} receipts - Array of MutationReceipt objects
 * @returns {{ valid: boolean, error?: { code: string, message: string } }}
 */
export function reconcileReceiptsAgainstOutput(parts, receipts) {
    if (!Array.isArray(receipts) || receipts.length === 0) {
        return { valid: true };
    }

    const committedReceipts = receipts.filter(r => r && r.committed === true && r.finalDisposition === 'applied');
    if (committedReceipts.length === 0) {
        return { valid: true };
    }

    const revisionIdSet = new Set();
    if (parts?.documentXml && typeof parts.documentXml === 'string') {
        const revRegex = /<(?:w:)?(?:ins|del|rPrChange|pPrChange|moveFrom|moveTo)\b[^>]*?\b(?:w:)?id="([^"]+)"/g;
        let m;
        while ((m = revRegex.exec(parts.documentXml)) !== null) {
            revisionIdSet.add(m[1]);
        }
    }

    const commentIdSet = new Set();
    if (parts?.commentsXml && typeof parts.commentsXml === 'string') {
        const comRegex = /<(?:w:)?comment\b[^>]*?\b(?:w:)?id="([^"]+)"/g;
        let m;
        while ((m = comRegex.exec(parts.commentsXml)) !== null) {
            commentIdSet.add(m[1]);
        }
    }

    const numberingIdSet = new Set();
    const combinedNumbering = [parts?.numberingXml, ...(parts?.numberingXmlParts || [])].filter(Boolean).join('\n');
    if (combinedNumbering) {
        const numRegex = /<(?:w:)?num\b[^>]*?\b(?:w:)?numId="([^"]+)"/g;
        let m;
        while ((m = numRegex.exec(combinedNumbering)) !== null) {
            numberingIdSet.add(m[1]);
        }
    }
    if (parts?.documentXml && typeof parts.documentXml === 'string') {
        const docNumRegex = /<(?:w:)?numId\b[^>]*?\b(?:w:)?val="([^"]+)"/g;
        let m;
        while ((m = docNumRegex.exec(parts.documentXml)) !== null) {
            numberingIdSet.add(m[1]);
        }
    }

    const relIdSet = new Set();
    if (parts?.relationshipsXml && typeof parts.relationshipsXml === 'string') {
        const relRegex = /<Relationship\b[^>]*?\bId="([^"]+)"/g;
        let m;
        while ((m = relRegex.exec(parts.relationshipsXml)) !== null) {
            relIdSet.add(m[1]);
        }
    }

    for (const receipt of committedReceipts) {
        if (Array.isArray(receipt.revisionItems)) {
            for (const item of receipt.revisionItems) {
                if (item.partName === 'word/document.xml' && !revisionIdSet.has(String(item.id))) {
                    return {
                        valid: false,
                        error: {
                            code: 'RECEIPT_RECONCILIATION_FAILED',
                            message: `Committed revision id '${item.id}' (kind: ${item.kind}) was not found in word/document.xml.`
                        }
                    };
                }
            }
        }

        if (Array.isArray(receipt.commentIds)) {
            for (const id of receipt.commentIds) {
                if (!commentIdSet.has(String(id))) {
                    return {
                        valid: false,
                        error: {
                            code: 'RECEIPT_RECONCILIATION_FAILED',
                            message: `Committed comment id '${id}' was not found in word/comments.xml.`
                        }
                    };
                }
            }
        }

        if (Array.isArray(receipt.numberingIds)) {
            for (const id of receipt.numberingIds) {
                if (!numberingIdSet.has(String(id))) {
                    return {
                        valid: false,
                        error: {
                            code: 'RECEIPT_RECONCILIATION_FAILED',
                            message: `Committed numbering id '${id}' was not found in numbering parts or document.`
                        }
                    };
                }
            }
        }

        if (parts?.relationshipsXml && Array.isArray(receipt.relationshipIds)) {
            for (const id of receipt.relationshipIds) {
                if (!relIdSet.has(String(id))) {
                    return {
                        valid: false,
                        error: {
                            code: 'RECEIPT_RECONCILIATION_FAILED',
                            message: `Committed relationship id '${id}' was not found in document.xml.rels.`
                        }
                    };
                }
            }
        }
    }

    return { valid: true };
}
