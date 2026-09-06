/**
 * Shared paragraph-targeting helpers for standalone/add-in consumers.
 *
 * This module centralizes target parsing and matching used by callers that
 * apply per-paragraph operations (for example chat redlines/comments/highlights).
 */

function toArray(nodeList) {
    return Array.from(nodeList || []);
}

function createTargetNotFoundError(message) {
    const error = new Error(message);
    error.code = 'TARGET_NOT_FOUND';
    return error;
}

function createTargetError(code, message, candidates = null) {
    const error = new Error(message);
    error.code = code;
    if (Array.isArray(candidates)) error.candidates = candidates;
    return error;
}

export const WORD_MAIN_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getElementsByLocalName(node, localName) {
    if (!node) return [];

    if (typeof node.getElementsByTagNameNS === 'function') {
        const namespaced = toArray(node.getElementsByTagNameNS('*', localName));
        if (namespaced.length > 0) return namespaced;
    }

    if (typeof node.getElementsByTagName !== 'function') return [];

    const prefixed = toArray(node.getElementsByTagName(`w:${localName}`));
    if (prefixed.length > 0) return prefixed;

    return toArray(node.getElementsByTagName(localName));
}

import { extractCanonicalParagraphText } from './paragraph-text.js';
import { isWordElement } from './word-xml.js';

/**
 * Reads visible text from a paragraph by concatenating `w:t` nodes and mapping
 * structural `w:tab` elements to a literal tab character.
 *
 * @param {Element|null|undefined} paragraph - OOXML paragraph node
 * @returns {string}
 */
export function getParagraphText(paragraph) {
    if (!paragraph) return '';
    return extractCanonicalParagraphText(paragraph);
}

/**
 * Returns body paragraphs for a document, or all paragraphs as fallback.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @returns {Element[]}
 */
export function getDocumentParagraphNodes(xmlDoc) {
    if (!xmlDoc) return [];
    const bodies = getElementsByLocalName(xmlDoc, 'body');
    const searchRoot = bodies.length > 0 ? bodies[0] : xmlDoc;
    return getElementsByLocalName(searchRoot, 'p');
}

export function getParagraphId(paragraph) {
    if (!paragraph) return null;
    const attribute = toArray(paragraph.attributes).find(candidate =>
        String(candidate?.localName || candidate?.name || '').replace(/^.*:/, '') === 'paraId'
    );
    return attribute?.value || null;
}

export function createParagraphFingerprint(paragraph, metadata = {}) {
    if (!paragraph) return null;
    const revisionView = metadata.revisionView === 'rejected' ? 'rejected' : 'accepted';
    const text = typeof metadata.text === 'string'
        ? metadata.text
        : extractCanonicalParagraphText(paragraph, { revisionView });
    const documentIndex = Number.isInteger(metadata.index)
        ? metadata.index
        : getDocumentParagraphNodes(paragraph.ownerDocument || paragraph).indexOf(paragraph) + 1;
    const paragraphId = metadata.paragraphId === undefined ? getParagraphId(paragraph) : metadata.paragraphId;
    const inTable = typeof metadata.inTable === 'boolean' ? metadata.inTable : isParagraphInTable(paragraph);
    const viewPart = revisionView === 'rejected' ? 'rejected\u001f' : '';
    const identity = `${paragraphId || ''}\u001f${documentIndex}\u001f${inTable ? 'table' : 'body'}\u001f${viewPart}${text}`;
    let hash = 0x811c9dc5;
    for (let i = 0; i < identity.length; i++) {
        hash ^= identity.charCodeAt(i);
        hash = Math.imul(hash, 0x01000193) >>> 0;
    }
    return `fnv1a32:${hash.toString(16).padStart(8, '0')}`;
}

/**
 * Builds immutable paragraph metadata and grouped lookup maps for one document
 * state. The index is intentionally session-scoped because it retains DOM
 * nodes and must be discarded after a mutation.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @param {Object} [options={}] - Options
 * @param {'accepted'|'rejected'} [options.revisionView='accepted'] - Revision view
 * @returns {{revisionView: 'accepted'|'rejected', entries: ReadonlyArray<Object>, byParagraph: Map<Element,Object>, byId: Map<string,Object>, byNormalizedText: Map<string,ReadonlyArray<Object>>}}
 */
export function buildParagraphMetadataIndex(xmlDoc, options = {}) {
    const revisionView = options.revisionView === 'rejected' ? 'rejected' : 'accepted';
    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const entries = [];
    const byParagraph = new Map();
    const byId = new Map();
    const grouped = new Map();

    for (let offset = 0; offset < paragraphs.length; offset++) {
        const paragraph = paragraphs[offset];
        const text = extractCanonicalParagraphText(paragraph, { revisionView });
        const paragraphId = getParagraphId(paragraph);
        const inTable = isParagraphInTable(paragraph);
        const normalizedText = normalizeWhitespaceForTargeting(text);
        const entry = Object.freeze({
            paragraph,
            index: offset + 1,
            paragraphId,
            text,
            normalizedText,
            revisionView,
            fingerprint: createParagraphFingerprint(paragraph, {
                text,
                index: offset + 1,
                paragraphId,
                inTable,
                revisionView
            }),
            inTable
        });
        entries.push(entry);
        byParagraph.set(paragraph, entry);
        if (paragraphId && !byId.has(paragraphId)) byId.set(paragraphId, entry);
        if (normalizedText) {
            if (!grouped.has(normalizedText)) grouped.set(normalizedText, []);
            grouped.get(normalizedText).push(entry);
        }
    }

    for (const [key, values] of grouped) grouped.set(key, Object.freeze(values));
    return Object.freeze({
        revisionView,
        entries: Object.freeze(entries),
        byParagraph,
        byId,
        byNormalizedText: grouped
    });
}

/**
 * Normalizes whitespace for paragraph-comparison matching.
 *
 * @param {string} text - Input text
 * @returns {string}
 */
export function normalizeWhitespaceForTargeting(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
}

/**
 * Detects markdown table syntax used for table reconciliation.
 *
 * @param {string} text - Candidate markdown text
 * @returns {boolean}
 */
export function isMarkdownTableText(text) {
    const trimmed = String(text || '').trim();
    return /^\|.+\|/.test(trimmed) && trimmed.includes('\n');
}

/**
 * Parses paragraph references such as `P12`, `[P12]`, `12`, or `P12.3`.
 *
 * @param {string|number|null|undefined} rawValue - Reference input
 * @returns {number|null}
 */
export function parseParagraphReference(rawValue) {
    if (rawValue == null) return null;
    if (typeof rawValue === 'number' && Number.isInteger(rawValue) && rawValue > 0) return rawValue;

    const text = String(rawValue).trim();
    if (!text) return null;

    const prefixed = text.match(/^\[?P(\d+)(?:\.\d+)?\]?$/i);
    if (prefixed) return Number.parseInt(prefixed[1], 10);

    const numeric = text.match(/^(\d+)$/);
    if (numeric) return Number.parseInt(numeric[1], 10);

    return null;
}

/**
 * Removes leading paragraph labels (for example `[P12]`) from text fields.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {string}
 */
export function stripLeadingParagraphMarker(text) {
    if (text == null) return '';
    return String(text).replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim();
}

/**
 * Splits a leading paragraph label from text.
 *
 * @param {string|null|undefined} text - Input text
 * @returns {{ text: string, targetRef: number|null }}
 */
export function splitLeadingParagraphMarker(text) {
    const raw = String(text || '');
    const marker = raw.match(/^\s*\[P(\d+)(?:\.\d+)?\]\s*/i);
    if (!marker) return { text: raw.trim(), targetRef: null };

    return {
        text: raw.replace(/^\s*\[P\d+(?:\.\d+)?\]\s*/i, '').trim(),
        targetRef: Number.parseInt(marker[1], 10)
    };
}

/**
 * Resolves a paragraph by 1-based paragraph index.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {number|null|undefined} targetRef - 1-based paragraph number
 * @returns {Element|null}
 */
export function findParagraphByReference(xmlDoc, targetRef, paragraphMetadataIndex = null) {
    if (!Number.isInteger(targetRef) || targetRef < 1) return null;
    if (paragraphMetadataIndex?.entries) {
        return paragraphMetadataIndex.entries[targetRef - 1]?.paragraph || null;
    }
    return getDocumentParagraphNodes(xmlDoc)[targetRef - 1] || null;
}

/**
 * Finds the closest ancestor matching Word namespace + localName.
 *
 * @param {Node|null|undefined} node - Start node
 * @param {string} localName - WordprocessingML local element name (for example `tbl`, `tc`)
 * @param {string} [namespaceUri] - Namespace URI to match
 * @returns {Element|null}
 */
export function findContainingWordElement(node, localName, namespaceUri = WORD_MAIN_NS) {
    let current = node;
    while (current) {
        if (
            current.nodeType === 1 &&
            current.namespaceURI === namespaceUri &&
            current.localName === localName
        ) {
            return current;
        }
        current = current.parentNode;
    }
    return null;
}

/**
 * Finds paragraph by exact/normalized text equality.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @returns {Element|null}
 */
export function findParagraphByStrictText(xmlDoc, targetText, options = {}) {
    const metadataIndex = options.paragraphMetadataIndex || null;
    const entries = metadataIndex?.entries || null;
    const paragraphs = entries ? null : getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const exact = entries
        ? entries.find(entry => entry.text.trim() === normalizedTarget)?.paragraph
        : paragraphs.find(p => getParagraphText(p).trim() === normalizedTarget);
    if (exact) return exact;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);
    if (metadataIndex?.byNormalizedText) {
        return metadataIndex.byNormalizedText.get(normTarget)?.[0]?.paragraph || null;
    }
    return paragraphs.find(p => normalizeWhitespaceForTargeting(getParagraphText(p)) === normTarget) || null;
}

/**
 * Finds paragraph by strict match, then fuzzy fallback heuristics.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {string} targetText - Target paragraph text
 * @param {{ onInfo?: (msg:string)=>void }} [options] - Optional logger callbacks
 * @returns {Element|null}
 */
export function findParagraphByBestTextMatch(xmlDoc, targetText, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const metadataIndex = options.paragraphMetadataIndex || null;
    const entries = metadataIndex?.entries || null;
    const paragraphs = entries ? null : getDocumentParagraphNodes(xmlDoc);
    const normalizedTarget = String(targetText || '').trim();
    if (!normalizedTarget) return null;

    const strictMatch = findParagraphByStrictText(xmlDoc, normalizedTarget, { paragraphMetadataIndex: metadataIndex });
    if (strictMatch) return strictMatch;

    const normTarget = normalizeWhitespaceForTargeting(normalizedTarget);

    const startsWithMatch = entries
        ? (entries.find(entry => entry.normalizedText.length > 10 && normTarget.startsWith(entry.normalizedText))?.paragraph || null)
        : paragraphs.find(p => {
            const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
            return paragraphText.length > 10 && normTarget.startsWith(paragraphText);
        });
    if (startsWithMatch) {
        onInfo(`[Fuzzy] Prefix match (target starts with paragraph): "${getParagraphText(startsWithMatch).trim().slice(0, 60)}..."`);
        return startsWithMatch;
    }

    const containsMatch = entries
        ? (entries.find(entry => entry.normalizedText.length > 15 && normTarget.includes(entry.normalizedText))?.paragraph || null)
        : paragraphs.find(p => {
            const paragraphText = normalizeWhitespaceForTargeting(getParagraphText(p));
            return paragraphText.length > 15 && normTarget.includes(paragraphText);
        });
    if (containsMatch) {
        onInfo(`[Fuzzy] Contains match: "${getParagraphText(containsMatch).trim().slice(0, 60)}..."`);
        return containsMatch;
    }

    let bestScore = 0;
    let bestParagraph = null;
    const targetWords = new Set(normTarget.toLowerCase().split(/\s+/).filter(word => word.length > 2));
    const candidateCount = entries?.length ?? paragraphs.length;
    for (let index = 0; index < candidateCount; index++) {
        const paragraph = entries?.[index]?.paragraph ?? paragraphs[index];
        const paragraphText = (entries?.[index]?.text ?? getParagraphText(paragraph)).trim();
        if (!paragraphText) continue;

        const paragraphWords = normalizeWhitespaceForTargeting(paragraphText)
            .toLowerCase()
            .split(/\s+/)
            .filter(word => word.length > 2);
        const overlap = paragraphWords.filter(word => targetWords.has(word)).length;
        const score = overlap / Math.max(targetWords.size, 1);
        if (score > bestScore && score > 0.5) {
            bestScore = score;
            bestParagraph = paragraph;
        }
    }

    if (bestParagraph) {
        onInfo(`[Fuzzy] Best word-overlap match (${(bestScore * 100).toFixed(0)}%): "${getParagraphText(bestParagraph).trim().slice(0, 60)}..."`);
        return bestParagraph;
    }

    return null;
}

/**
 * Resolves a target paragraph from `targetRef` + `targetText`.
 *
 * Resolution order:
 * 1) `targetRef` when provided and valid
 * 2) strict text match
 * 3) fuzzy text match
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {{
 *   targetText?: string,
 *   targetRef?: string|number|null,
 *   opType?: string,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} options - Resolution options
 * @returns {{ paragraph: Element, resolvedBy: 'ref'|'strict_text'|'fuzzy_text'|'strict_text_after_ref_drift'|'fuzzy_text_after_ref_drift' }}
 */
export function resolveTargetParagraph(xmlDoc, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const onWarn = typeof options.onWarn === 'function' ? options.onWarn : () => {};
    const opType = options.opType || 'operation';
    const descriptor = options.targetDescriptor && typeof options.targetDescriptor === 'object'
        ? options.targetDescriptor
        : null;
    const revisionView = descriptor?.revisionView === 'rejected' ? 'rejected' : 'accepted';
    const cleanTargetText = String(descriptor?.exactText ?? descriptor?.text ?? options.targetText ?? '').trim();
    const parsedRef = parseParagraphReference(descriptor?.index ?? descriptor?.paragraphIndex ?? options.targetRef);
    const strictAmbiguity = options.strictAmbiguity === true;
    let paragraphMetadataIndex = options.paragraphMetadataIndex || null;
    if (paragraphMetadataIndex && paragraphMetadataIndex.revisionView !== revisionView) {
        paragraphMetadataIndex = options.metadataIndices?.[revisionView]
            || buildParagraphMetadataIndex(xmlDoc, { revisionView });
    } else if (!paragraphMetadataIndex) {
        paragraphMetadataIndex = buildParagraphMetadataIndex(xmlDoc, { revisionView });
    }

    if (descriptor?.paragraphId) {
        const byId = findParagraphById(xmlDoc, descriptor.paragraphId, paragraphMetadataIndex);
        if (!byId) {
            throw createTargetError(
                'TARGET_NOT_FOUND',
                `Target paragraphId not found: "${descriptor.paragraphId}".`
            );
        }
        const cachedEntry = paragraphMetadataIndex?.byParagraph?.get(byId) || null;
        const actualIndex = cachedEntry?.index ?? getDocumentParagraphNodes(xmlDoc).indexOf(byId) + 1;
        const actualFingerprint = cachedEntry?.fingerprint
            || createParagraphFingerprint(byId, { revisionView });
        const actualInTable = cachedEntry?.inTable ?? isParagraphInTable(byId);
        const actualText = cachedEntry?.normalizedText
            || normalizeWhitespaceForTargeting(extractCanonicalParagraphText(byId, { revisionView }));

        if (parsedRef != null && parsedRef !== actualIndex) {
            throw createTargetError(
                'TARGET_INDEX_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" (index ${actualIndex}) does not match requested index ${parsedRef}.`,
                cachedEntry ? [serializeTargetCandidate(cachedEntry)] : null
            );
        }
        if (descriptor.fingerprint && descriptor.fingerprint !== actualFingerprint) {
            throw createTargetError(
                'TARGET_FINGERPRINT_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" no longer matches its source fingerprint.`,
                cachedEntry ? [serializeTargetCandidate(cachedEntry)] : null
            );
        }
        if (typeof descriptor.inTable === 'boolean' && descriptor.inTable !== actualInTable) {
            throw createTargetError(
                'TARGET_CONTEXT_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" does not match the requested table context.`,
                cachedEntry ? [serializeTargetCandidate(cachedEntry)] : null
            );
        }
        if (cleanTargetText && actualText !== normalizeWhitespaceForTargeting(cleanTargetText)) {
            throw createTargetError(
                'TARGET_TEXT_MISMATCH',
                `Target paragraphId "${descriptor.paragraphId}" no longer matches the supplied text.`,
                cachedEntry ? [serializeTargetCandidate(cachedEntry)] : null
            );
        }
        if (descriptor.occurrence != null) {
            const textToFind = cleanTargetText || actualText;
            const textCandidates = findStrictTargetCandidates(xmlDoc, textToFind, paragraphMetadataIndex);
            const actualOccurrence = textCandidates.findIndex(c => c.paragraph === byId) + 1;
            if (actualOccurrence === 0 || actualOccurrence !== descriptor.occurrence) {
                throw createTargetError(
                    'TARGET_OCCURRENCE_MISMATCH',
                    `Target paragraphId "${descriptor.paragraphId}" matches occurrence ${actualOccurrence}, not requested occurrence ${descriptor.occurrence}.`,
                    textCandidates.map(serializeTargetCandidate)
                );
            }
        }
        return { paragraph: byId, resolvedBy: 'paragraph_id' };
    }

    let candidates = [];
    if (cleanTargetText) {
        const unfilteredCandidates = findStrictTargetCandidates(xmlDoc, cleanTargetText, paragraphMetadataIndex);
        candidates = filterTargetCandidates(unfilteredCandidates, descriptor);

        if (descriptor?.fingerprint && unfilteredCandidates.length > 0 && candidates.length === 0) {
            throw createTargetError(
                'TARGET_FINGERPRINT_MISMATCH',
                'Target text matched, but no paragraph matched the supplied source fingerprint.',
                unfilteredCandidates.map(serializeTargetCandidate)
            );
        }

        if (
            typeof descriptor?.inTable === 'boolean'
            && unfilteredCandidates.length > 0
            && candidates.length === 0
        ) {
            throw createTargetError(
                'TARGET_CONTEXT_MISMATCH',
                'Target text matched, but no paragraph matched the requested table context.',
                unfilteredCandidates.map(serializeTargetCandidate)
            );
        }


        if (descriptor?.occurrence) {
            const occurrenceMatch = candidates[descriptor.occurrence - 1] || null;
            if (!occurrenceMatch) {
                throw createTargetError(
                    'TARGET_NOT_FOUND',
                    `Target occurrence ${descriptor.occurrence} was not found.`,
                    candidates.map(serializeTargetCandidate)
                );
            }
            return { paragraph: occurrenceMatch.paragraph, resolvedBy: 'occurrence' };
        }

        if (strictAmbiguity) {
            if (parsedRef) {
                const byReference = candidates.find(candidate => candidate.index === parsedRef) || null;
                if (byReference) return { paragraph: byReference.paragraph, resolvedBy: 'ref' };
                if (descriptor?.fingerprint && candidates.length > 0) {
                    throw createTargetError(
                        'TARGET_FINGERPRINT_MISMATCH',
                        `Target fingerprint does not match paragraph reference [P${parsedRef}].`,
                        candidates.map(serializeTargetCandidate)
                    );
                }
                if (candidates.length === 1) {
                    return { paragraph: candidates[0].paragraph, resolvedBy: 'strict_text_after_ref_drift' };
                }
            }
            if (candidates.length > 1) {
                throw createTargetError(
                    'AMBIGUOUS_TARGET',
                    `Target text matched ${candidates.length} paragraphs; provide paragraphId, index, occurrence, or fingerprint.`,
                    candidates.map(serializeTargetCandidate)
                );
            }
            if (candidates.length === 0) {
                throw createTargetNotFoundError(`Target paragraph not found: "${cleanTargetText}"`);
            }
        }

        if (!parsedRef && candidates.length === 1) {
            const candidate = candidates[0];
            return {
                paragraph: candidate.paragraph,
                resolvedBy: descriptor?.fingerprint ? 'fingerprint' : 'strict_text'
            };
        }
    }

    if (parsedRef) {
        const byRef = findParagraphByReference(xmlDoc, parsedRef, paragraphMetadataIndex);
        if (byRef) {
            const cached = paragraphMetadataIndex?.byParagraph?.get(byRef) || null;
            if (descriptor?.fingerprint && descriptor.fingerprint !== cached?.fingerprint) {
                throw createTargetError(
                    'TARGET_FINGERPRINT_MISMATCH',
                    `Target fingerprint does not match paragraph reference [P${parsedRef}].`
                );
            }
            if (typeof descriptor?.inTable === 'boolean' && descriptor.inTable !== cached?.inTable) {
                throw createTargetError(
                    'TARGET_CONTEXT_MISMATCH',
                    `Target paragraph reference [P${parsedRef}] does not match requested table context.`
                );
            }
            if (cleanTargetText) {
                const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText, { paragraphMetadataIndex });
                const byRefText = (cached?.text || extractCanonicalParagraphText(byRef, { revisionView })).trim();
                const byRefNorm = normalizeWhitespaceForTargeting(byRefText);
                const targetNorm = normalizeWhitespaceForTargeting(cleanTargetText);
                const hasDrift = byRefNorm !== targetNorm;

                if (hasDrift && strictMatch && strictMatch !== byRef) {
                    onInfo(`[Target] [P${parsedRef}] drifted for ${opType}; using strict text rematch.`);
                    return { paragraph: strictMatch, resolvedBy: 'strict_text_after_ref_drift' };
                }

                if (hasDrift) {
                    const fuzzyMatch = findParagraphByBestTextMatch(xmlDoc, cleanTargetText, { onInfo, paragraphMetadataIndex });
                    if (fuzzyMatch && fuzzyMatch !== byRef) {
                        onInfo(`[Target] [P${parsedRef}] drifted for ${opType}; using fuzzy text rematch.`);
                        return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text_after_ref_drift' };
                    }
                    onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}; target text drifted.`);
                } else if (strictMatch && strictMatch !== byRef) {
                    onInfo(`[Target] [P${parsedRef}] disambiguated duplicate target text for ${opType}.`);
                }
            } else {
                onInfo(`[Target] Using [P${parsedRef}] fallback for ${opType}.`);
            }
            return { paragraph: byRef, resolvedBy: 'ref' };
        }

        onWarn(`[WARN] Target reference [P${parsedRef}] not found; falling back to text matching for ${opType}.`);
    }

    if (cleanTargetText && !strictAmbiguity) {
        const strictMatch = findParagraphByStrictText(xmlDoc, cleanTargetText, { paragraphMetadataIndex });
        if (strictMatch) {
            const candidateCount = candidates.length;
            if (candidateCount > 1) {
                const warningMsg = `AMBIGUOUS_TARGET_HEURISTIC_USED: Target text matched ${candidateCount} paragraphs; permissive resolution chose candidate 1. Migrate to strict targeting (e.g. strictTargets: true with paragraphId, index, occurrence, or fingerprint) before v1.0.0.`;
                onWarn(warningMsg);
                return {
                    paragraph: strictMatch,
                    resolvedBy: 'strict_text',
                    warnings: [warningMsg]
                };
            }
            return { paragraph: strictMatch, resolvedBy: 'strict_text' };
        }

        const fuzzyMatch = findParagraphByBestTextMatch(xmlDoc, cleanTargetText, { onInfo, paragraphMetadataIndex });
        if (fuzzyMatch) return { paragraph: fuzzyMatch, resolvedBy: 'fuzzy_text' };
    }

    if (cleanTargetText) throw createTargetNotFoundError(`Target paragraph not found: "${cleanTargetText}"`);
    if (parsedRef) throw createTargetNotFoundError(`Target paragraph reference not found: [P${parsedRef}]`);
    throw createTargetNotFoundError('Operation target missing: provide "target" text or "targetRef" ([P#]).');
}

function isParagraphInTable(paragraph) {
    return !!findContainingWordElement(paragraph, 'tbl');
}

export function findStrictTargetCandidates(xmlDoc, targetText, optionsOrIndex = null) {
    const normalizedTarget = normalizeWhitespaceForTargeting(targetText);
    if (!normalizedTarget) return [];

    const metadataIndex = optionsOrIndex?.byNormalizedText
        ? optionsOrIndex
        : (optionsOrIndex?.paragraphMetadataIndex || null);
    const revisionView = optionsOrIndex?.revisionView
        || metadataIndex?.revisionView
        || 'accepted';

    if (metadataIndex?.byNormalizedText && metadataIndex.revisionView === revisionView) {
        return Array.from(metadataIndex.byNormalizedText.get(normalizedTarget) || []);
    }

    const paragraphs = getDocumentParagraphNodes(xmlDoc);
    const candidates = [];
    for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphText = extractCanonicalParagraphText(paragraph, { revisionView }).trim();
        if (!paragraphText) continue;
        if (normalizeWhitespaceForTargeting(paragraphText) !== normalizedTarget) continue;
        candidates.push({
            paragraph,
            index: i + 1,
            inTable: isParagraphInTable(paragraph),
            paragraphId: getParagraphId(paragraph),
            fingerprint: createParagraphFingerprint(paragraph, {
                text: paragraphText,
                index: i + 1,
                revisionView
            }),
            text: paragraphText,
            revisionView
        });
    }
    return candidates;
}

function serializeTargetCandidate(candidate) {
    return {
        index: candidate.index,
        paragraphId: candidate.paragraphId || null,
        text: candidate.text,
        inTable: candidate.inTable,
        fingerprint: candidate.fingerprint,
        revisionView: candidate.revisionView || 'accepted'
    };
}

function filterTargetCandidates(candidates, descriptor) {
    let scoped = candidates.slice();
    if (descriptor?.paragraphId) {
        scoped = scoped.filter(candidate => candidate.paragraphId === descriptor.paragraphId);
    }
    if (typeof descriptor?.inTable === 'boolean') {
        scoped = scoped.filter(candidate => candidate.inTable === descriptor.inTable);
    }
    if (descriptor?.fingerprint) {
        scoped = scoped.filter(candidate => candidate.fingerprint === descriptor.fingerprint);
    }
    return scoped;
}

function findParagraphById(xmlDoc, paragraphId, paragraphMetadataIndex = null) {
    if (!paragraphId) return null;
    if (paragraphMetadataIndex?.byId) return paragraphMetadataIndex.byId.get(paragraphId)?.paragraph || null;
    return getDocumentParagraphNodes(xmlDoc).find(paragraph => getParagraphId(paragraph) === paragraphId) || null;
}

function selectBestTargetCandidate(candidates, parsedRef, expectedInTable = null) {
    if (!Array.isArray(candidates) || candidates.length === 0) return null;

    let scoped = candidates.slice();
    if (typeof expectedInTable === 'boolean') {
        const sameContext = scoped.filter(candidate => candidate.inTable === expectedInTable);
        if (sameContext.length > 0) scoped = sameContext;
    }

    if (Number.isInteger(parsedRef) && parsedRef > 0) {
        scoped.sort((a, b) => Math.abs(a.index - parsedRef) - Math.abs(b.index - parsedRef));
    }

    return scoped[0] || null;
}

/**
 * Builds a turn-start paragraph snapshot keyed by 1-based paragraph index.
 *
 * Intended for callers that apply multiple operations sequentially and need to
 * detect `targetRef` drift after earlier structural edits.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document root
 * @returns {Map<number, { text: string, normalizedText: string, inTable: boolean }>}
 */
export function buildTargetReferenceSnapshot(xmlDoc, paragraphMetadataIndex = null) {
    const entries = paragraphMetadataIndex?.entries || null;
    const paragraphs = entries ? null : getDocumentParagraphNodes(xmlDoc);
    const snapshot = new Map();
    const paragraphCount = entries?.length ?? paragraphs.length;
    for (let i = 0; i < paragraphCount; i++) {
        const paragraph = entries?.[i]?.paragraph ?? paragraphs[i];
        const text = (entries?.[i]?.text ?? getParagraphText(paragraph)).trim();
        snapshot.set(i + 1, {
            text,
            normalizedText: entries?.[i]?.normalizedText ?? normalizeWhitespaceForTargeting(text),
            inTable: entries?.[i]?.inTable ?? isParagraphInTable(paragraph)
        });
    }
    return snapshot;
}

/**
 * Resolves a paragraph using the standard resolver, then corrects stale
 * `targetRef` mappings via strict rematch when a turn-start snapshot is provided.
 *
 * @param {Document|Element|null|undefined} xmlDoc - OOXML document
 * @param {{
 *   targetText?: string,
 *   targetRef?: string|number|null,
 *   opType?: string,
 *   targetRefSnapshot?: Map<number, { text?: string, inTable?: boolean }>|null,
 *   onInfo?: (msg:string)=>void,
 *   onWarn?: (msg:string)=>void
 * }} options - Resolution options
 * @returns {{ paragraph: Element, resolvedBy: 'ref'|'strict_text'|'fuzzy_text'|'strict_text_after_ref_drift' }}
 */
export function resolveTargetParagraphWithSnapshot(xmlDoc, options = {}) {
    const onInfo = typeof options.onInfo === 'function' ? options.onInfo : () => {};
    const resolved = resolveTargetParagraph(xmlDoc, options);

    const parsedRef = parseParagraphReference(options.targetRef);
    if (!parsedRef || resolved?.resolvedBy !== 'ref') return resolved;

    const snapshotEntry = options.targetRefSnapshot instanceof Map
        ? (options.targetRefSnapshot.get(parsedRef) || null)
        : null;
    if (!snapshotEntry) return resolved;

    const cleanTargetText = String(options.targetText || '').trim();
    const expectedText = cleanTargetText || snapshotEntry.text || '';
    const expectedNorm = normalizeWhitespaceForTargeting(expectedText);
    if (!expectedNorm) return resolved;

    const resolvedNorm = options.paragraphMetadataIndex?.byParagraph?.get(resolved.paragraph)?.normalizedText
        || normalizeWhitespaceForTargeting(getParagraphText(resolved.paragraph));
    if (resolvedNorm === expectedNorm) return resolved;

    const candidateTexts = [];
    if (cleanTargetText) candidateTexts.push(cleanTargetText);
    if (snapshotEntry.text) {
        const snapshotNorm = normalizeWhitespaceForTargeting(snapshotEntry.text);
        if (snapshotNorm && !candidateTexts.some(text => normalizeWhitespaceForTargeting(text) === snapshotNorm)) {
            candidateTexts.push(snapshotEntry.text);
        }
    }

    let bestCandidate = null;
    for (const candidateText of candidateTexts) {
        const candidates = findStrictTargetCandidates(xmlDoc, candidateText, options.paragraphMetadataIndex || null);
        const selected = selectBestTargetCandidate(candidates, parsedRef, snapshotEntry.inTable);
        if (!selected) continue;
        if (!bestCandidate) bestCandidate = selected;
        if (selected.paragraph !== resolved.paragraph) {
            bestCandidate = selected;
            break;
        }
    }

    if (bestCandidate && bestCandidate.paragraph !== resolved.paragraph) {
        const opType = options.opType || 'operation';
        onInfo(`[Target] [P${parsedRef}] appears stale after prior edits; using strict text rematch for ${opType}.`);
        return { paragraph: bestCandidate.paragraph, resolvedBy: 'strict_text_after_ref_drift' };
    }

    throw createTargetNotFoundError(
        `Target paragraph [P${parsedRef}] no longer matches its batch-start anchor.`
    );

}

/**
 * Resolves a contiguous paragraph range using paragraph references.
 *
 * @param {Document} xmlDoc - XML document
 * @param {string|number|null} startRef - Start paragraph reference (e.g. P12)
 * @param {string|number|null} endRef - End paragraph reference (e.g. P15)
 * @param {Object} [options={}] - Resolution options
 * @param {string} [options.opType='redline'] - Operation type hint
 * @param {Array|null} [options.targetRefSnapshot=null] - Optional target snapshot
 * @param {(message: string) => void} [options.onInfo] - Optional info logger
 * @param {(message: string) => void} [options.onWarn] - Optional warn logger
 * @returns {Element[]|null}
 */
export function resolveParagraphRangeByRefs(xmlDoc, startRef, endRef, options = {}) {
    if (!xmlDoc || !startRef || !endRef) return null;

    const opType = options?.opType || 'redline';
    const targetRefSnapshot = options?.targetRefSnapshot || null;
    const onInfo = typeof options?.onInfo === 'function' ? options.onInfo : () => { };
    const onWarn = typeof options?.onWarn === 'function' ? options.onWarn : () => { };

    const start = resolveTargetParagraphWithSnapshot(xmlDoc, {
        targetRef: startRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!start) return null;

    const end = resolveTargetParagraphWithSnapshot(xmlDoc, {
        targetRef: endRef,
        opType,
        targetRefSnapshot,
        onInfo,
        onWarn
    })?.paragraph;
    if (!end) return null;

    const allParagraphs = Array.from(xmlDoc.getElementsByTagNameNS('*', 'p'));
    const startIdx = allParagraphs.indexOf(start);
    const endIdx = allParagraphs.indexOf(end);
    if (startIdx < 0 || endIdx < startIdx) return null;

    const range = allParagraphs.slice(startIdx, endIdx + 1);
    if (range.length === 0) return null;

    const parent = range[0]?.parentNode || null;
    if (!parent) return null;
    if (!range.every(node => node && node.parentNode === parent)) return null;
    return range;
}

/**
 * Validates a paragraph boundary operation (split, delete, join) before mutation.
 * Returns { valid: true } or { valid: false, code: 'UNSAFE_PARAGRAPH_BOUNDARY', message: string }.
 *
 * @param {Element} targetParagraph - Target paragraph node
 * @param {string} modifiedText - Replacement text
 * @param {Object} [options={}] - Options
 * @returns {{ valid: boolean, code?: string, message?: string }}
 */
export function validateParagraphBoundaryMutation(targetParagraph, modifiedText, options = {}) {
    if (!targetParagraph || !targetParagraph.parentNode) {
        return { valid: true };
    }

    const isDelete = modifiedText === '' || options.operationKind === 'delete';
    const isSplit = typeof modifiedText === 'string' && modifiedText.includes('\n');
    const targetEndParagraph = options.targetEndParagraph || null;

    // 1. Table cell boundary checks
    const containingTc = findContainingWordElement(targetParagraph, 'tc');
    if (containingTc) {
        if (isDelete) {
            const cellParagraphs = Array.from(containingTc.childNodes).filter(node => isWordElement(node, 'p'));
            if (cellParagraphs.length <= 1) {
                return {
                    valid: false,
                    code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                    message: 'Refusing to delete the sole terminal paragraph of a table cell.'
                };
            }
        }
        if (targetEndParagraph) {
            const endTc = findContainingWordElement(targetEndParagraph, 'tc');
            if (endTc !== containingTc) {
                return {
                    valid: false,
                    code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                    message: 'Refusing paragraph boundary join across different table cells.'
                };
            }
        }
    } else if (targetEndParagraph) {
        const endTc = findContainingWordElement(targetEndParagraph, 'tc');
        if (endTc) {
            return {
                valid: false,
                code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                message: 'Refusing paragraph boundary join across table boundaries.'
            };
        }
    }

    // 2. Section break (w:sectPr) checks
    const hasSectPr = targetParagraph.getElementsByTagNameNS(WORD_MAIN_NS, 'sectPr').length > 0;
    if (hasSectPr && isDelete) {
        return {
            valid: false,
            code: 'UNSAFE_PARAGRAPH_BOUNDARY',
            message: 'Refusing to delete paragraph containing section properties (w:sectPr).'
        };
    }

    if (targetEndParagraph) {
        let cursor = targetParagraph;
        while (cursor && cursor !== targetEndParagraph) {
            if (isWordElement(cursor, 'p') && cursor.getElementsByTagNameNS(WORD_MAIN_NS, 'sectPr').length > 0) {
                return {
                    valid: false,
                    code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                    message: 'Refusing paragraph boundary join across section break.'
                };
            }
            cursor = cursor.nextSibling;
        }
    }

    // 3. Field instruction checks during split
    if (isSplit) {
        const fldSimples = targetParagraph.getElementsByTagNameNS(WORD_MAIN_NS, 'fldSimple');
        if (fldSimples.length > 0) {
            return {
                valid: false,
                code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                message: 'Refusing to split paragraph containing a simple field instruction (w:fldSimple).'
            };
        }

        const fldChars = Array.from(targetParagraph.getElementsByTagNameNS(WORD_MAIN_NS, 'fldChar'));
        if (fldChars.length > 0) {
            let activeFields = 0;
            for (const fc of fldChars) {
                const fldCharType = fc.getAttributeNS(WORD_MAIN_NS, 'fldCharType') || fc.getAttribute('w:fldCharType');
                if (fldCharType === 'begin') activeFields++;
                else if (fldCharType === 'end') activeFields = Math.max(0, activeFields - 1);
            }
            if (activeFields > 0) {
                return {
                    valid: false,
                    code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                    message: 'Refusing to split paragraph across an unclosed field instruction.'
                };
            }
        }

        // 4. Bookmark range boundary checks
        const bookmarkStarts = Array.from(targetParagraph.getElementsByTagNameNS(WORD_MAIN_NS, 'bookmarkStart'));
        const bookmarkEnds = Array.from(targetParagraph.getElementsByTagNameNS(WORD_MAIN_NS, 'bookmarkEnd'));
        const bStartIds = new Set(bookmarkStarts.map(b => b.getAttributeNS(WORD_MAIN_NS, 'id') || b.getAttribute('w:id')));
        const bEndIds = new Set(bookmarkEnds.map(b => b.getAttributeNS(WORD_MAIN_NS, 'id') || b.getAttribute('w:id')));
        for (const id of bStartIds) {
            if (id && !bEndIds.has(id)) {
                return {
                    valid: false,
                    code: 'UNSAFE_PARAGRAPH_BOUNDARY',
                    message: `Refusing to split paragraph across open bookmark range (ID: ${id}).`
                };
            }
        }
    }

    return { valid: true };
}
