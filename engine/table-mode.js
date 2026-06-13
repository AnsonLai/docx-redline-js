/**
 * Table-specific reconciliation and transformation flows.
 */

import { NS_W, createRevisionMetadata } from '../core/types.js';
import {
    getElementsByTagNS,
    getElementsByTagNSOrTag,
    getFirstElementByTag,
    getFirstElementByTagNS,
    getFirstElementByTagNSOrTag,
    getXmlParseError
} from '../core/xml-query.js';
import { createParser } from '../adapters/xml-adapter.js';
import { log, error } from '../adapters/logger.js';
import { diffTablesWithVirtualGrid, serializeVirtualGridToOoxml, generateTableOoxml } from '../services/table-reconciliation.js';
import { parseTable } from '../pipeline/pipeline.js';
import { ingestTableToVirtualGrid } from '../pipeline/ingestion.js';
import { createWordElement, withOoxmlSourceType } from '../core/word-xml.js';
import { markParagraphMarkDeleted } from './run-builders.js';

function noChanges(serializer, xmlDoc) {
    return withOoxmlSourceType({ oxml: serializer.serializeToString(xmlDoc), hasChanges: false });
}

/**
 * Applies structural reconciliation to tables using Virtual Grid.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Markdown/target table text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {DOMParser} parser - Parser instance
 * @param {string} author - Author name
 * @param {boolean} [generateRedlines=true] - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyTableReconciliation(xmlDoc, modifiedText, serializer, parser, author, generateRedlines = true) {
    const tableNodes = getElementsByTagNSOrTag(xmlDoc, NS_W, 'tbl');
    const newTableData = parseTable(modifiedText);
    const hasNewContent = newTableData.rows.length > 0 || newTableData.headers.length > 0;

    if (tableNodes.length === 0 || !hasNewContent) {
        return noChanges(serializer, xmlDoc);
    }

    const targetTable = tableNodes[0];
    const oldGrid = ingestTableToVirtualGrid(targetTable);
    const operations = diffTablesWithVirtualGrid(oldGrid, newTableData);

    if (operations.length === 0) {
        return noChanges(serializer, xmlDoc);
    }

    const options = { generateRedlines, author };
    const reconciledOxml = serializeVirtualGridToOoxml(oldGrid, operations, options);
    const wrappedOxml = `<root xmlns:w="${NS_W}">${reconciledOxml}</root>`;
    const reconcileParser = parser || createParser();
    const reconciledDoc = reconcileParser.parseFromString(wrappedOxml, 'application/xml');

    const parseError = getXmlParseError(reconciledDoc);
    if (parseError) {
        error('[OxmlEngine] Failed to parse reconciled table OOXML:', parseError.textContent);
        return noChanges(serializer, xmlDoc);
    }

    const newTableNode = getFirstElementByTagNSOrTag(reconciledDoc, NS_W, 'tbl');
    if (!newTableNode) {
        error('[OxmlEngine] No table found in reconciled OOXML');
        return noChanges(serializer, xmlDoc);
    }

    const importedTable = xmlDoc.importNode(newTableNode, true);
    targetTable.parentNode.replaceChild(importedTable, targetTable);

    return withOoxmlSourceType({ oxml: serializer.serializeToString(xmlDoc), hasChanges: true });
}

/**
 * Transforms paragraph content into a new table from Markdown text.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string} modifiedText - Markdown table text
 * @param {XMLSerializer} serializer - Serializer instance
 * @param {DOMParser} parser - Parser instance
 * @param {string} author - Author name
 * @param {boolean} generateRedlines - Track change toggle
 * @returns {{ oxml: string, hasChanges: boolean }}
 */
export function applyTextToTableTransformation(xmlDoc, modifiedText, serializer, parser, author, generateRedlines) {
    const tableData = parseTable(modifiedText);
    if (!tableData || (tableData.rows.length === 0 && tableData.headers.length === 0)) {
        log('[OxmlEngine] Failed to parse table data from Markdown');
        return noChanges(serializer, xmlDoc);
    }

    const tableOoxml = generateTableOoxml(tableData, { generateRedlines, author });
    const activeParser = parser || createParser();
    const tableDoc = activeParser.parseFromString(`<root xmlns:w="${NS_W}">${tableOoxml}</root>`, 'application/xml');

    const tableParseError = getXmlParseError(tableDoc);
    if (tableParseError) {
        error('[OxmlEngine] Failed to parse generated table OOXML:', tableParseError.textContent);
        return noChanges(serializer, xmlDoc);
    }

    let newTableElement = getFirstElementByTagNS(tableDoc, NS_W, 'tbl');
    if (!newTableElement) {
        newTableElement = getFirstElementByTagNS(tableDoc, NS_W, 'ins');
    }
    if (!newTableElement) {
        error('[OxmlEngine] No table element found in generated OOXML');
        return noChanges(serializer, xmlDoc);
    }

    let workingDoc = xmlDoc;
    let paragraphs = getElementsByTagNS(workingDoc, NS_W, 'p');
    if (paragraphs.length === 0) {
        log('[OxmlEngine] No paragraphs found to replace');
        return noChanges(serializer, workingDoc);
    }

    let firstParagraph = paragraphs[0];
    let parent = firstParagraph.parentNode;

    if (parent && parent.nodeType === 9) {
        const wrappedDoc = activeParser.parseFromString(
            `<w:document xmlns:w="${NS_W}"><w:body/></w:document>`,
            'application/xml'
        );
        const wrappedBody = getFirstElementByTagNS(wrappedDoc, NS_W, 'body');
        paragraphs.forEach(p => wrappedBody.appendChild(wrappedDoc.importNode(p, true)));

        workingDoc = wrappedDoc;
        paragraphs = getElementsByTagNS(workingDoc, NS_W, 'p');
        firstParagraph = paragraphs[0];
        parent = firstParagraph.parentNode;
    }

    const importedTable = workingDoc.importNode(newTableElement, true);

    if (generateRedlines) {
        paragraphs.forEach(p => {
            markParagraphMarkDeleted(workingDoc, p, author);
            const runs = getElementsByTagNS(p, NS_W, 'r');
            runs.forEach(run => {
                const textNodes = getElementsByTagNS(run, NS_W, 't');
                textNodes.forEach(t => {
                    const text = t.textContent || '';
                    if (text.trim()) {
                        const delText = createWordElement(workingDoc, 'w:delText');
                        delText.textContent = text;
                        t.parentNode.replaceChild(delText, t);
                    }
                });

                const del = createWordElement(workingDoc, 'w:del');
                const metadata = createRevisionMetadata(author);
                del.setAttribute('w:id', String(metadata.id));
                del.setAttribute('w:author', metadata.author);
                del.setAttribute('w:date', metadata.date);
                run.parentNode.insertBefore(del, run);
                del.appendChild(run);
            });
        });
    } else {
        paragraphs.slice(1).forEach(p => p.parentNode.removeChild(p));
    }

    parent.insertBefore(importedTable, firstParagraph);
    if (!generateRedlines) {
        parent.removeChild(firstParagraph);
    }

    log('[OxmlEngine] Text-to-table transformation complete');
    return withOoxmlSourceType({ oxml: serializer.serializeToString(workingDoc), hasChanges: true });
}
