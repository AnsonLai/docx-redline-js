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

    recordComment(id, partName = 'word/comments.xml') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.commentIds.includes(strId)) {
            this.activeReceipt.commentIds.push(strId);
        }
    }

    recordNumbering(id, partName = 'word/numbering.xml') {
        if (!this.activeReceipt || id == null) return;
        const strId = String(id);
        if (!this.activeReceipt.numberingIds.includes(strId)) {
            this.activeReceipt.numberingIds.push(strId);
        }
    }

    recordRelationship(id, partName = 'word/_rels/document.xml.rels') {
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
