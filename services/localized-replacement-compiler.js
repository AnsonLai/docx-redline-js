import { normalizeErrorWithRecovery } from './error-recovery.js';

function invalidReplacement(message, field, invalidCode) {
    return normalizeErrorWithRecovery({
        code: invalidCode,
        message,
        ...(field ? { field } : {}),
        recovery: {
            action: 'change_request',
            sameArgumentsSafe: false,
            requiresReinspection: false,
            requiresUserAuthorization: false
        }
    });
}

function patchError(code, message, field, recoveryAction, details = {}) {
    return normalizeErrorWithRecovery({
        code,
        message,
        field,
        ...details,
        recovery: {
            action: recoveryAction,
            sameArgumentsSafe: false,
            requiresReinspection: false,
            requiresUserAuthorization: false
        }
    });
}

function occurrenceOffsets(sourceText, find) {
    const offsets = [];
    let from = 0;
    while (from <= sourceText.length - find.length) {
        const start = sourceText.indexOf(find, from);
        if (start < 0) break;
        offsets.push(start);
        from = start + 1;
    }
    return offsets;
}

function equivalentSpace(left, right) {
    return (left === ' ' || left === '\u00a0')
        && (right === ' ' || right === '\u00a0');
}

function spaceEquivalentOffsets(sourceText, find) {
    const offsets = [];
    for (let start = 0; start <= sourceText.length - find.length; start += 1) {
        let matches = true;
        for (let offset = 0; offset < find.length; offset += 1) {
            const sourceCharacter = sourceText[start + offset];
            const findCharacter = find[offset];
            if (sourceCharacter !== findCharacter && !equivalentSpace(sourceCharacter, findCharacter)) {
                matches = false;
                break;
            }
        }
        if (matches) offsets.push(start);
    }
    return offsets;
}

/** Resolve exact offsets first, then the narrow ASCII-space/NBSP equivalent fallback. */
export function localizedReplacementOffsets(sourceText, find) {
    const exact = occurrenceOffsets(sourceText, find);
    if (exact.length > 0) return { offsets: exact, matchMode: 'exact' };
    return { offsets: spaceEquivalentOffsets(sourceText, find), matchMode: 'space_equivalent' };
}

function preserveEquivalentBoundarySpaces(sourceSlice, find, replace) {
    const effective = replace.split('');
    let prefix = 0;
    while (
        prefix < find.length
        && prefix < replace.length
        && (find[prefix] === replace[prefix] || equivalentSpace(find[prefix], replace[prefix]))
    ) {
        if (equivalentSpace(sourceSlice[prefix], replace[prefix])) effective[prefix] = sourceSlice[prefix];
        prefix += 1;
    }
    let sourceOffset = find.length - 1;
    let replacementOffset = replace.length - 1;
    while (
        sourceOffset >= prefix
        && replacementOffset >= prefix
        && (
            find[sourceOffset] === replace[replacementOffset]
            || equivalentSpace(find[sourceOffset], replace[replacementOffset])
        )
    ) {
        if (equivalentSpace(sourceSlice[sourceOffset], replace[replacementOffset])) {
            effective[replacementOffset] = sourceSlice[sourceOffset];
        }
        sourceOffset -= 1;
        replacementOffset -= 1;
    }
    return effective.join('');
}

function offsetCandidate(sourceText, start, length) {
    const excerptStart = Math.max(0, start - 40);
    const excerptEnd = Math.min(sourceText.length, start + length + 40);
    return {
        start,
        end: start + length,
        excerpt: sourceText.slice(excerptStart, excerptEnd)
    };
}

function boundedExcerpt(text, start, end, limit = 120) {
    const source = String(text ?? '');
    const safeStart = Math.max(0, Math.min(source.length, start));
    const safeEnd = Math.max(safeStart, Math.min(source.length, end));
    const center = Math.floor((safeStart + safeEnd) / 2);
    let excerptStart = Math.max(0, center - Math.floor(limit / 2));
    let excerptEnd = Math.min(source.length, excerptStart + limit);
    excerptStart = Math.max(0, excerptEnd - limit);
    const prefix = excerptStart > 0 ? '…' : '';
    const suffix = excerptEnd < source.length ? '…' : '';
    const available = Math.max(0, limit - prefix.length - suffix.length);
    return `${prefix}${source.slice(excerptStart, excerptStart + available)}${suffix}`;
}

/** Validate and normalize exact replacement requests without resolving source ranges. */
export function validateExactReplacementRequests(
    replacements,
    field = 'replacements',
    { invalidCode = 'INVALID_OPERATION' } = {}
) {
    if (!Array.isArray(replacements) || replacements.length === 0) {
        return {
            ok: false,
            error: invalidReplacement('replacements must be a non-empty array.', field, invalidCode)
        };
    }

    const normalized = [];
    for (let index = 0; index < replacements.length; index += 1) {
        const replacement = replacements[index];
        const itemField = `${field}[${index}]`;
        if (!replacement || typeof replacement !== 'object' || Array.isArray(replacement)) {
            return {
                ok: false,
                error: invalidReplacement(`${itemField} must be an object.`, itemField, invalidCode)
            };
        }
        if (typeof replacement.find !== 'string' || replacement.find.length === 0) {
            return {
                ok: false,
                error: invalidReplacement(
                    `${itemField}.find must be a non-empty string.`,
                    `${itemField}.find`,
                    invalidCode
                )
            };
        }
        if (typeof replacement.replace !== 'string') {
            return {
                ok: false,
                error: invalidReplacement(
                    `${itemField}.replace must be a string.`,
                    `${itemField}.replace`,
                    invalidCode
                )
            };
        }
        if (replacement.find === replacement.replace) {
            return {
                ok: false,
                error: invalidReplacement(
                    `${itemField} must change the matched text.`,
                    itemField,
                    invalidCode
                )
            };
        }
        const occurrence = replacement.occurrence == null ? null : Number(replacement.occurrence);
        if (occurrence != null && (!Number.isInteger(occurrence) || occurrence < 1)) {
            return {
                ok: false,
                error: invalidReplacement(
                    `${itemField}.occurrence must be a positive integer.`,
                    `${itemField}.occurrence`,
                    invalidCode
                )
            };
        }
        normalized.push({
            find: replacement.find,
            replace: replacement.replace,
            ...(occurrence == null ? {} : { occurrence })
        });
    }

    return { ok: true, replacements: normalized };
}

/**
 * Compile exact, source-relative replacement intents into complete desired text.
 * All source ranges are resolved before mutation, so replacements are simultaneous.
 */
export function compileExactReplacements(
    sourceText,
    replacements,
    field = 'replacements',
    options = {}
) {
    const invalidCode = options.invalidCode || 'INVALID_OPERATION';
    if (typeof sourceText !== 'string') {
        return {
            ok: false,
            error: invalidReplacement(
                'Localized replacements require string source text.',
                field,
                invalidCode
            )
        };
    }
    const validation = validateExactReplacementRequests(replacements, field, { invalidCode });
    if (!validation.ok) return validation;

    const resolved = [];
    for (let index = 0; index < validation.replacements.length; index += 1) {
        const replacement = validation.replacements[index];
        const itemField = `${field}[${index}]`;
        const occurrence = replacement.occurrence ?? null;
        const { offsets, matchMode } = localizedReplacementOffsets(sourceText, replacement.find);
        if (offsets.length === 0 || (occurrence != null && occurrence > offsets.length)) {
            return {
                ok: false,
                error: patchError(
                    'PATCH_SOURCE_NOT_FOUND',
                    occurrence == null
                        ? `Patch source was not found: "${replacement.find}".`
                        : `Occurrence ${occurrence} of patch source was not found: "${replacement.find}".`,
                    itemField,
                    'change_patch',
                    { matchCount: offsets.length }
                )
            };
        }
        if (occurrence == null && offsets.length > 1) {
            return {
                ok: false,
                error: patchError(
                    'AMBIGUOUS_PATCH_SOURCE',
                    `Patch source matched ${offsets.length} locations; provide occurrence.`,
                    itemField,
                    'choose_occurrence',
                    { candidates: offsets.map(start => offsetCandidate(sourceText, start, replacement.find.length)) }
                )
            };
        }

        const start = offsets[(occurrence || 1) - 1];
        const sourceSlice = sourceText.slice(start, start + replacement.find.length);
        resolved.push({
            requestIndex: index,
            start,
            end: start + replacement.find.length,
            find: replacement.find,
            replace: replacement.replace,
            effectiveReplace: matchMode === 'space_equivalent'
                ? preserveEquivalentBoundarySpaces(sourceSlice, replacement.find, replacement.replace)
                : replacement.replace,
            matchMode,
            occurrence: occurrence || 1
        });
    }

    resolved.sort((left, right) => left.start - right.start || left.end - right.end || left.requestIndex - right.requestIndex);
    const unique = [];
    for (const replacement of resolved) {
        const previous = unique.at(-1);
        if (previous && replacement.start === previous.start && replacement.end === previous.end) {
            if (replacement.effectiveReplace !== previous.effectiveReplace) {
                return {
                    ok: false,
                    error: patchError(
                        'CONFLICTING_PATCHES',
                        `Localized replacements ${previous.requestIndex + 1} and ${replacement.requestIndex + 1} assign different text to the same source range.`,
                        field,
                        'combine_patches',
                        { replacementIndexes: [previous.requestIndex + 1, replacement.requestIndex + 1] }
                    )
                };
            }
            continue;
        }
        if (previous && replacement.start < previous.end) {
            return {
                ok: false,
                error: patchError(
                    'OVERLAPPING_PATCHES',
                    `Localized replacements ${previous.requestIndex + 1} and ${replacement.requestIndex + 1} overlap.`,
                    field,
                    'combine_patches',
                    { replacementIndexes: [previous.requestIndex + 1, replacement.requestIndex + 1] }
                )
            };
        }
        unique.push(replacement);
    }

    let cursor = 0;
    let desiredText = '';
    for (const replacement of unique) {
        desiredText += sourceText.slice(cursor, replacement.start);
        desiredText += replacement.effectiveReplace;
        cursor = replacement.end;
    }
    desiredText += sourceText.slice(cursor);

    return {
        ok: true,
        desiredText,
        replacements: unique.map(replacement => ({
            requestIndex: replacement.requestIndex + 1,
            start: replacement.start,
            end: replacement.end,
            find: replacement.find,
            replace: replacement.replace,
            matchMode: replacement.matchMode,
            occurrence: replacement.occurrence,
            removedLength: replacement.end - replacement.start,
            insertedLength: replacement.effectiveReplace.length
        }))
    };
}

/** Build a compact receipt-like summary from actual source and output text. */
export function buildLocalizedReplacementChange(compilation, outputText, target = {}, context = null) {
    if (!compilation || typeof compilation !== 'object') return null;
    const sourceText = String(compilation.sourceText ?? '');
    const desiredText = String(compilation.desiredText ?? '');
    const replacements = Array.isArray(compilation.replacements) ? compilation.replacements : [];
    let delta = 0;
    const applied = replacements.map(replacement => {
        const afterStart = replacement.start + delta;
        const afterEnd = afterStart + replacement.replace.length;
        delta += replacement.replace.length - (replacement.end - replacement.start);
        return {
            find: replacement.find,
            replace: replacement.replace,
            ...(replacement.matchMode === 'space_equivalent' ? { matchMode: replacement.matchMode } : {}),
            occurrence: replacement.occurrence,
            beforeExcerpt: boundedExcerpt(sourceText, replacement.start, replacement.end),
            afterExcerpt: boundedExcerpt(desiredText, afterStart, afterEnd)
        };
    });
    return {
        kind: 'localized_replacement',
        committed: true,
        finalDisposition: 'applied',
        target: {
            ...(target.paragraphId ? { paragraphId: target.paragraphId } : {}),
            ...(Number.isInteger(target.index) ? { index: target.index } : {}),
            ...(target.fingerprint ? { fingerprint: target.fingerprint } : {}),
            ...(context?.humanReference ? { humanReference: context.humanReference } : {})
        },
        ...(context ? { context: { ...context } } : {}),
        replacements: applied,
        verification: {
            acceptedViewMatchesCompiledText: typeof outputText === 'string' && outputText === desiredText
        }
    };
}
