/**
 * OOXML Comment Engine
 *
 * Provides pure OOXML-based comment insertion without Word JS API calls.
 */

import {
    NS_W,
    createRevisionIdAllocator,
    getNextRevisionId,
    getRevisionTimestamp,
    resetRevisionIdCounter
} from '../core/types.js';
import { createSerializer, parseOoxmlSafe } from '../adapters/xml-adapter.js';
import { log, error as logError } from '../adapters/logger.js';
import { getElementsByTag, getFirstElementByTag, getXmlParseError } from '../core/xml-query.js';
import { buildCommentElement, buildCommentsPartXml, buildCommentMarkers } from './comment-builders.js';
import { getDefaultAuthor } from '../adapters/config.js';
import { createParagraphTextIndex, injectMarkersIntoParagraph, resolveTextInParagraphIndex } from './comment-locator.js';
import {
    injectCommentsIntoPackage as injectCommentsIntoExistingPackage,
    wrapParagraphWithComments,
    wrapWithCommentsPart
} from './comment-package.js';

export { getNextRevisionId, resetRevisionIdCounter };
export { buildCommentElement, buildCommentsPartXml, buildCommentMarkers };

/**
 * @typedef {Object} CommentRequest
 * @property {number} paragraphIndex - 1-based paragraph index
 * @property {string} textToFind - Text to attach comment to
 * @property {string} commentContent - The comment text
 */

/**
 * @typedef {Object} CommentInjectionResult
 * @property {string} oxml - Complete OOXML package with comments
 * @property {string} [commentsXml] - comments.xml content when comments are applied
 * @property {number} commentsApplied - Number of successfully placed comments
 * @property {string[]} warnings - Any issues encountered
 */

function parseDocumentOxml(oxml, parseFailureWarning) {
    const parsed = parseOoxmlSafe(oxml, 'text/xml');
    const parseError = parsed.doc ? getXmlParseError(parsed.doc) : null;
    if (parsed.error || parseError) {
        const message = parsed.error?.message || parseError?.textContent || 'parse error';
        return {
            xmlDoc: null,
            warning: parseFailureWarning(message),
            warnings: parsed.warnings,
            error: { code: 'PARSE_ERROR', message }
        };
    }
    return { xmlDoc: parsed.doc, warning: null, warnings: parsed.warnings, error: null };
}

/**
 * Injects comments into OOXML using pure XML manipulation.
 *
 * @param {string} oxml - Original document OOXML
 * @param {CommentRequest[]} comments - Comment requests
 * @param {Object} [options={}] - Options
 * @param {string} [options.author] - Author for comments (defaults to configured default author)
 * @returns {CommentInjectionResult}
 */
export function injectCommentsIntoOoxml(oxml, comments, options = {}) {
    const author = options?.author || getDefaultAuthor();
    const date = getRevisionTimestamp();
    const warnings = [];
    const placedComments = [];
    const resolvedAnchors = [];
    const errors = [];

    if (!comments || comments.length === 0) {
        return {
            oxml,
            hasChanges: false,
            commentsApplied: 0,
            warnings: ['No comments to inject']
        };
    }

    const serializer = createSerializer();
    const parseResult = parseDocumentOxml(
        oxml,
        warning => `Failed to parse OXML: ${warning}`
    );
    warnings.push(...(parseResult.warnings || []));

    if (!parseResult.xmlDoc) {
        logError('[CommentEngine] Parse failure:', parseResult.warning);
        return {
            oxml,
            hasChanges: false,
            commentsApplied: 0,
            status: 'error',
            error: parseResult.error,
            warnings: [...warnings, parseResult.warning]
        };
    }

    const xmlDoc = parseResult.xmlDoc;
    const revisionIdAllocator = createRevisionIdAllocator(xmlDoc);
    const paragraphs = getElementsByTag(xmlDoc, 'w:p');
    log(`[CommentEngine] Found ${paragraphs.length} paragraphs, processing ${comments.length} comment requests`);

    /** @type {Map<number, number>} */
    const remainingRequestsByParagraph = new Map();
    for (const request of comments) {
        const paragraphIndex = request.paragraphIndex - 1;
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
            const message = `Paragraph ${request.paragraphIndex} out of range (1-${paragraphs.length})`;
            warnings.push(message);
            errors.push({ code: 'TARGET_NOT_FOUND', message, paragraphIndex: request.paragraphIndex });
            continue;
        }
        remainingRequestsByParagraph.set(paragraphIndex, (remainingRequestsByParagraph.get(paragraphIndex) || 0) + 1);
    }

    /** @type {Map<number, { fullText: string, runOffsets: Array<{run: Element, start: number, end: number}> }>} */
    const paragraphIndexes = new Map();

    for (const [requestIndex, request] of comments.entries()) {
        const paragraphIndex = request.paragraphIndex - 1;
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.length) {
            continue;
        }

        const targetParagraph = paragraphs[paragraphIndex];
        let textIndex = paragraphIndexes.get(paragraphIndex);
        if (!textIndex) {
            textIndex = createParagraphTextIndex(targetParagraph);
            paragraphIndexes.set(paragraphIndex, textIndex);
        }

        const anchorText = String(request.textToFind ?? '');
        const resolution = resolveTextInParagraphIndex(textIndex, anchorText);
        const remaining = (remainingRequestsByParagraph.get(paragraphIndex) || 1) - 1;
        remainingRequestsByParagraph.set(paragraphIndex, remaining);
        if (!resolution.found) {
            const error = {
                ...resolution.error,
                requestIndex: requestIndex + 1,
                paragraphIndex: request.paragraphIndex
            };
            errors.push(error);
            warnings.push(error.message);
            if (remaining === 0) paragraphIndexes.delete(paragraphIndex);
            continue;
        }

        const commentId = typeof options.commentIdAllocator === 'function'
            ? options.commentIdAllocator()
            : getNextRevisionId();
        const success = injectMarkersIntoParagraph(
            xmlDoc,
            targetParagraph,
            anchorText,
            commentId,
            textIndex,
            revisionIdAllocator,
            resolution
        );

        if (!success) {
            const error = {
                code: 'ANCHOR_INSERTION_FAILED',
                message: `Resolved comment anchor could not be inserted in paragraph ${request.paragraphIndex}.`,
                requestIndex: requestIndex + 1,
                paragraphIndex: request.paragraphIndex
            };
            errors.push(error);
            warnings.push(error.message);
            if (remaining === 0) {
                paragraphIndexes.delete(paragraphIndex);
            }
            continue;
        }

        placedComments.push({
            id: commentId,
            content: request.commentContent,
            author,
            date
        });
        resolvedAnchors.push({
            requestIndex: requestIndex + 1,
            paragraphIndex: request.paragraphIndex,
            text: anchorText,
            resolvedBy: resolution.resolvedBy,
            start: resolution.start,
            end: resolution.end
        });

        if (remaining > 0) {
            // Rebuild only when another request still targets this paragraph.
            paragraphIndexes.set(paragraphIndex, createParagraphTextIndex(targetParagraph));
        } else {
            paragraphIndexes.delete(paragraphIndex);
        }
    }

    if (placedComments.length === 0) {
        return {
            oxml,
            hasChanges: false,
            commentsApplied: 0,
            warnings,
            resolvedAnchors,
            ...(errors.length > 0 ? { status: 'error', error: errors[0], errors } : {})
        };
    }

    return {
        oxml: serializer.serializeToString(xmlDoc),
        hasChanges: true,
        commentsXml: buildCommentsPartXml(placedComments),
        commentsApplied: placedComments.length,
        placedComments,
        warnings,
        resolvedAnchors,
        ...(errors.length > 0 ? { status: 'error', error: errors[0], errors } : {})
    };
}

/**
 * Injects a comment into a single paragraph OOXML and returns a complete mini-package.
 *
 * @param {string} paragraphOoxml - Paragraph OOXML (raw paragraph or pkg:package)
 * @param {string} textToFind - Target text
 * @param {string} commentContent - Comment body
 * @param {Object} [options={}] - Options
 * @param {string} [options.author='AI Assistant'] - Comment author
 * @returns {{ success: boolean, package?: string, warning?: string, commentId?: number }}
 */
export function injectCommentIntoParagraphOoxml(paragraphOoxml, textToFind, commentContent, options = {}) {
    const { author = 'AI Assistant' } = options;
    const date = getRevisionTimestamp();

    const serializer = createSerializer();
    const parseResult = parseDocumentOxml(
        paragraphOoxml,
        warning => `Failed to parse paragraph OOXML: ${warning}`
    );

    if (!parseResult.xmlDoc) {
        return { success: false, warning: parseResult.warning };
    }

    const xmlDoc = parseResult.xmlDoc;
    const revisionIdAllocator = createRevisionIdAllocator(xmlDoc);
    const paragraphs = getElementsByTag(xmlDoc, 'w:p');
    if (paragraphs.length === 0) {
        return { success: false, warning: 'No paragraph found in OOXML' };
    }

    const paragraph = paragraphs[0];
    const paragraphIndex = createParagraphTextIndex(paragraph);
    const resolution = resolveTextInParagraphIndex(paragraphIndex, textToFind);
    if (!resolution.found) {
        return { success: false, warning: resolution.error.message, error: resolution.error };
    }
    const commentId = getNextRevisionId();
    const success = injectMarkersIntoParagraph(
        xmlDoc,
        paragraph,
        textToFind,
        commentId,
        paragraphIndex,
        revisionIdAllocator,
        resolution
    );
    if (!success) {
        return { success: false, warning: `Could not find "${textToFind.substring(0, 30)}..." in paragraph` };
    }

    const commentElement = buildCommentElement(commentId, author, commentContent, date);
    const commentsXml = `<w:comments xmlns:w="${NS_W}">${commentElement}</w:comments>`;
    const pkgPackage = getFirstElementByTag(xmlDoc, 'pkg:package');

    if (pkgPackage) {
        const withComments = injectCommentsIntoExistingPackage(serializer.serializeToString(xmlDoc), commentsXml);
        return { success: true, package: withComments, commentId };
    }

    const modifiedParagraphXml = serializer.serializeToString(xmlDoc);
    return {
        success: true,
        package: wrapParagraphWithComments(modifiedParagraphXml, commentsXml),
        commentId
    };
}

/**
 * Injects comments part into an existing OOXML package from getOoxml().
 *
 * @param {string} packageOxml - Existing pkg:package
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export function injectCommentsIntoPackage(packageOxml, commentsXml) {
    return injectCommentsIntoExistingPackage(packageOxml, commentsXml);
}

/**
 * @deprecated Use injectCommentsIntoPackage instead.
 *
 * @param {string} documentXml - Document XML
 * @param {string} commentsXml - comments.xml payload
 * @returns {string}
 */
export { wrapWithCommentsPart };
