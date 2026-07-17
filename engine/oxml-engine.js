/**
 * OOXML Engine V5.1 - Hybrid Mode
 *
 * Router/orchestrator for formatting/text reconciliation modes.
 */

import { preprocessMarkdown } from '../pipeline/markdown-processor.js';
import { isListTargetLoose } from '../pipeline/list-markers.js';
import { ReconciliationPipeline } from '../pipeline/pipeline.js';
import { wrapInDocumentFragment } from '../pipeline/serialization.js';
import {
    getElementsByTagNSOrTag,
    getXmlParseError
} from '../core/xml-query.js';
import { createParser, createSerializer, parseXml, serializeXml } from '../adapters/xml-adapter.js';
import { log, error } from '../adapters/logger.js';
import { extractFormattingFromOoxml } from './format-extraction.js';
import {
    applyFormatRemovalAsSurgicalReplacement,
    applyFormatOnlyChangesSurgical
} from './format-application.js';
import { buildParagraphInfos, findMatchingParagraphInfo, getContainingParagraph } from './format-paragraph-targeting.js';
import { detectTableCellContext, serializeParagraphOnly } from './table-cell-context.js';
import { applySurgicalMode } from './surgical-mode.js';
import { applyReconstructionMode } from './reconstruction-mode.js';
import { applyTableReconciliation, applyTextToTableTransformation } from './table-mode.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { containsTrackedChanges, withOoxmlSourceType } from '../core/word-xml.js';
import { NS_W, seedRevisionIdsFromDocument } from '../core/types.js';
import { acceptTrackedChangesInOoxml } from '../services/revision-comment-management.js';

/**
 * Applies redline track changes to OOXML by modifying the DOM in-place.
 *
 * @param {string} oxml - Original OOXML string
 * @param {string} originalText - Original plain text
 * @param {string} modifiedText - New text (may contain markdown)
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI'] - Author for track changes
 * @param {string|null} [options.targetParagraphId=null] - Preferred paragraph identity for table wrappers
 * @param {'reject-input'|'accept-all-first'} [options.existingRevisions='reject-input'] - Policy for source OOXML with tracked changes
 * @param {boolean} [options.removeFormatting=false] - Remove existing core formatting when text is otherwise unchanged
 * @returns {Promise<{ oxml: string, hasChanges: boolean, sourceType?: 'package'|'document'|'fragment', status?: 'ok'|'no-op'|'error', error?: { code: string, message: string } }>}
 */
export async function applyRedlineToOxml(oxml, originalText, modifiedText, options = {}) {
    const generateRedlines = options.generateRedlines ?? true;
    const author = options.author || getDefaultAuthor();
    const parser = createParser();
    const serializer = createSerializer();
    const finalize = result => {
        const withStatus = { ...result };
        if (!withStatus.status) {
            withStatus.status = withStatus.hasChanges ? 'ok' : 'no-op';
        }
        return withOoxmlSourceType(withStatus);
    };
    const noChanges = () => finalize({ oxml, hasChanges: false });

    let xmlDoc;
    try {
        xmlDoc = parser.parseFromString(oxml, 'text/xml');
    } catch (e) {
        error('[OxmlEngine] Failed to parse OXML:', e);
        return finalize({
            oxml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: 'Could not parse OOXML input.' }
        });
    }

    const parseError = getXmlParseError(xmlDoc);
    if (parseError) {
        error('[OxmlEngine] XML parse error:', parseError.textContent);
        return finalize({
            oxml,
            hasChanges: false,
            status: 'error',
            error: { code: 'PARSE_ERROR', message: parseError.textContent || 'Could not parse OOXML input.' }
        });
    }
    seedRevisionIdsFromDocument(xmlDoc);

    if (containsTrackedChanges(xmlDoc)) {
        const existingRevisionsPolicy = options.existingRevisions || 'reject-input';
        if (existingRevisionsPolicy === 'accept-all-first') {
            log('[OxmlEngine] Existing revisions detected; accepting all input revisions before redlining');
            const accepted = acceptTrackedChangesInOoxml(oxml, { allAuthors: true });
            oxml = accepted.oxml;
            xmlDoc = parser.parseFromString(oxml, 'text/xml');
            const acceptedParseError = getXmlParseError(xmlDoc);
            if (acceptedParseError) {
                error('[OxmlEngine] XML parse error after accepting existing revisions:', acceptedParseError.textContent);
                return finalize({
                    oxml,
                    hasChanges: false,
                    status: 'error',
                    error: {
                        code: 'PARSE_ERROR',
                        message: 'Could not parse OOXML after accepting existing revisions.'
                    }
                });
            }
            seedRevisionIdsFromDocument(xmlDoc);
        } else {
            log('[OxmlEngine] Existing revisions detected; rejecting input per existingRevisions policy');
            return finalize({
                oxml,
                hasChanges: false,
                status: 'error',
                error: {
                    code: 'EXISTING_REVISIONS',
                    message: 'Input OOXML contains existing tracked changes. Pass existingRevisions: "accept-all-first" to normalize before redlining.'
                }
            });
        }
    }

    const initialTableCellContext = detectTableCellContext(xmlDoc, originalText, options);
    if (initialTableCellContext.hasTableWrapper && initialTableCellContext.targetParagraph && !options._isolatedTableCell) {
        log('[OxmlEngine] Isolating table-cell paragraph before diff');
        const isolatedOxml = serializeParagraphOnly(xmlDoc, initialTableCellContext.targetParagraph, serializer);
        return applyRedlineToOxml(isolatedOxml, originalText, modifiedText, {
            ...options,
            _isolatedTableCell: true
        });
    }

    const sanitizedText = sanitizeAiResponse(modifiedText);
    const { cleanText: cleanModifiedText, formatHints } = preprocessMarkdown(sanitizedText);

    const hasTextChanges = cleanModifiedText.trim() !== originalText.trim();
    const hasFormatHints = formatHints.length > 0;

    const { existingFormatHints, textSpans, paragraphs } = extractFormattingFromOoxml(xmlDoc);
    const hasExistingFormatting = existingFormatHints.length > 0;
    const visibleText = textSpans.map(span => textSpanVisibleText(span)).join('');
    if (
        hasTextChanges
        && typeof originalText === 'string'
        && originalText.trim()
        && !originalText.includes('\n')
        && !visibleText.includes(originalText.trim())
        && !visibleText.replace(/[\t\n\u2011]/g, '').includes(originalText.trim().replace(/[\t\n\u2011]/g, ''))
        && !normalizeTargetText(visibleText).includes(normalizeTargetText(originalText))
    ) {
        log('[OxmlEngine] Target text not found in OOXML');
        return finalize({
            oxml,
            hasChanges: false,
            status: 'error',
            error: {
                code: 'TARGET_NOT_FOUND',
                message: 'Original text was not found in the supplied OOXML.'
            }
        });
    }
    let paragraphInfos = null;
    const getParagraphInfos = () => {
        if (!paragraphInfos) {
            paragraphInfos = buildParagraphInfos(xmlDoc, paragraphs, textSpans);
        }
        return paragraphInfos;
    };
    const applyFormatOnlyWithOoxmlFallback = (precomputedContext = null) => {
        const formatResult = applyFormatOnlyChangesSurgical(
            xmlDoc,
            originalText,
            formatHints,
            serializer,
            author,
            generateRedlines,
            precomputedContext
        );
        if (!formatResult.useNativeApi) {
            return formatResult;
        }

        log('[OxmlEngine] Format-only surgical fallback signal encountered; retrying with OOXML reconstruction fallback');
        return applyReconstructionMode(
            xmlDoc,
            originalText,
            cleanModifiedText,
            serializer,
            author,
            formatHints,
            generateRedlines
        );
    };

    log(`[OxmlEngine] Text changes: ${hasTextChanges}, New format hints: ${formatHints.length}, Existing format hints: ${existingFormatHints.length}`);

    const needsFormatRemoval = options.removeFormatting === true
        && !hasTextChanges
        && !hasFormatHints
        && hasExistingFormatting;

    if (!hasTextChanges && !hasFormatHints && !hasExistingFormatting) {
        log('[OxmlEngine] No text changes, no format hints, and no existing formatting detected');
        return noChanges();
    }

    if (!hasTextChanges && !hasFormatHints && hasExistingFormatting && !needsFormatRemoval) {
        log('[OxmlEngine] No text or explicit formatting changes; preserving existing formatting');
        return noChanges();
    }

    if (needsFormatRemoval) {
        log('[OxmlEngine] Format REMOVAL detected: applying surgical replacement in OOXML');

        const tableCellCtx = initialTableCellContext;
        let targetParagraph = tableCellCtx.targetParagraph || null;

        if (!targetParagraph) {
            const matchedInfo = findMatchingParagraphInfo(getParagraphInfos(), originalText);
            if (matchedInfo) {
                targetParagraph = matchedInfo.paragraph;
            }
        }

        let filteredHints = existingFormatHints;
        if (targetParagraph) {
            filteredHints = existingFormatHints.filter(hint => {
                const hintParagraph = getContainingParagraph(hint.run);
                return hintParagraph === targetParagraph;
            });
        }

        const removalResult = applyFormatRemovalAsSurgicalReplacement(
            xmlDoc,
            textSpans,
            filteredHints,
            serializer,
            author,
            generateRedlines
        );

        if (tableCellCtx.hasTableWrapper && targetParagraph) {
            return finalize({
                oxml: serializeParagraphOnly(xmlDoc, targetParagraph, serializer),
                hasChanges: removalResult.hasChanges
            });
        }

        return finalize(removalResult);
    }

    if (!hasTextChanges && hasFormatHints) {
        log(`[OxmlEngine] Format-only change detected: ${formatHints.length} format hints`);

        const tableCellCtx = initialTableCellContext;
        const precomputedFormatContext = {
            textSpans,
            paragraphs,
            paragraphInfos: getParagraphInfos()
        };
        if (tableCellCtx.hasTableWrapper && tableCellCtx.targetParagraph) {
            log('[OxmlEngine] Table cell context: applying formatting to target paragraph only');

            const formatResult = applyFormatOnlyWithOoxmlFallback(precomputedFormatContext);

            log('[OxmlEngine] Stripping table wrapper for table cell paragraph (format-only)');
            return finalize({
                oxml: serializeParagraphOnly(xmlDoc, tableCellCtx.targetParagraph, serializer),
                hasChanges: formatResult.hasChanges
            });
        }

        return finalize(applyFormatOnlyWithOoxmlFallback(precomputedFormatContext));
    }

    const tables = getElementsByTagNSOrTag(xmlDoc, NS_W, 'tbl');
    const hasTables = tables.length > 0;
    const isMarkdownTable = /^\|.+\|/.test(cleanModifiedText.trim()) && cleanModifiedText.includes('\n');
    const isTargetList = isListTargetLoose(cleanModifiedText);
    const tableCellContext = initialTableCellContext;

    log(`[OxmlEngine] Mode: ${hasTables ? 'SURGICAL' : 'RECONSTRUCTION'}, formatHints: ${formatHints.length}, isMarkdownTable: ${isMarkdownTable}, isTargetList: ${isTargetList}, isTableCellParagraph: ${tableCellContext.isTableCellParagraph}`);

    if (isMarkdownTable && !hasTables) {
        log('[OxmlEngine] Text-to-table transformation: generating new table from Markdown');
        return finalize(applyTextToTableTransformation(xmlDoc, cleanModifiedText, serializer, parser, author, generateRedlines));
    }

    if (hasTables && isMarkdownTable) {
        return finalize(applyTableReconciliation(xmlDoc, cleanModifiedText, serializer, parser, author, generateRedlines));
    }
    if (hasTables) {
        const surgicalTarget = tableCellContext.hasTableWrapper && tableCellContext.targetParagraph
            ? tableCellContext.targetParagraph
            : null;
        if (surgicalTarget) {
            log('[OxmlEngine] Table cell edit: scoping surgical mode to target paragraph');
        }

        const result = applySurgicalMode(
            xmlDoc,
            originalText,
            cleanModifiedText,
            serializer,
            author,
            formatHints,
            generateRedlines,
            surgicalTarget
        );

        if (tableCellContext.hasTableWrapper && result.hasChanges && tableCellContext.targetParagraph) {
            log('[OxmlEngine] Stripping table wrapper for table cell paragraph (surgical mode)');
            return finalize({ oxml: serializeParagraphOnly(xmlDoc, tableCellContext.targetParagraph, serializer), hasChanges: true });
        }
        return finalize(result);
    }
    if (isTargetList) {
        log('[OxmlEngine] 🎯 Using reconciliation pipeline for list generation');
        const pipeline = new ReconciliationPipeline({ author, generateRedlines });
        const result = await pipeline.execute(oxml, modifiedText, { xmlDoc });

        if (result.isValid && result.ooxml && result.ooxml !== oxml) {
            const includeNumbering = result.includeNumbering === true;
            log(`[OxmlEngine] Wrapping list OOXML with numbering definitions, includeNumbering=${includeNumbering}`);
            const wrapped = wrapInDocumentFragment(result.ooxml, {
                includeNumbering,
                numberingXml: result.numberingXml
            });
            log(`[OxmlEngine] ✅ Wrapped OOXML length: ${wrapped.length}`);
            return finalize({ oxml: wrapped, hasChanges: true });
        }
        return noChanges();
    }

    return finalize(applyReconstructionMode(xmlDoc, originalText, cleanModifiedText, serializer, author, formatHints, generateRedlines));
}

function normalizeTargetText(text) {
    return String(text || '').replace(/[\t\n\u2011]/g, ' ').replace(/\s+/g, ' ').trim();
}

function textSpanVisibleText(span) {
    const node = span?.textElement;
    const localName = String(node?.localName || node?.nodeName || '').replace(/^.*:/, '');
    if (localName === 'tab') return '\t';
    if (localName === 'br' || localName === 'cr') return '\n';
    if (localName === 'noBreakHyphen') return '\u2011';
    return node?.textContent || '';
}

/**
 * Sanitizes AI response text by removing common prefixes.
 *
 * @param {string} text - AI response text
 * @returns {string}
 */
export function sanitizeAiResponse(text) {
    let cleaned = text;
    cleaned = cleaned.replace(/^(Here is the redline:|Here is the text:|Sure, I can help:|Here's the updated text:)\s*/i, '');
    cleaned = cleaned.replace(/\$\\text\{/g, '').replace(/\}\$/g, '');
    cleaned = cleaned.replace(/\$([^0-9\n]+?)\$/g, '$1');
    cleaned = cleaned.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
    return cleaned;
}

/**
 * Parses OOXML into a DOM document.
 *
 * @param {string} ooxmlString - OOXML text
 * @returns {Document}
 */
export function parseOoxml(ooxmlString) {
    return parseXml(ooxmlString, 'application/xml');
}

/**
 * Serializes a DOM document to OOXML text.
 *
 * @param {Node} doc - XML document/node
 * @returns {string}
 */
export function serializeOoxml(doc) {
    return serializeXml(doc);
}
