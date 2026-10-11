import type { ExistingRevisionsPolicy, RedlineError, RedlineStatus, RevisionToken } from '../index.js';

export interface ParagraphTargetDescriptor {
  text?: string;
  exactText?: string;
  index?: number | string;
  paragraphIndex?: number | string;
  paragraphId?: string;
  occurrence?: number;
  inTable?: boolean;
  fingerprint?: string;
  sourceFingerprint?: string;
  /** Defaults to rejected for restore operations and accepted for all other operations. */
  revisionView?: 'accepted' | 'rejected';
  captureRef?: string;
  select?: string;
  createdByOperation?: number;
}

export interface InsertionAffinity {
  formatting?: 'left' | 'right' | 'none';
  hyperlink?: 'inside' | 'outside' | 'preserve';
  revision?: 'coalesce_same_author' | 'separate';
  bookmark?: 'inside' | 'outside';
  comment?: 'inside' | 'outside';
}

export interface RejectedTextInsertionAnchor {
  exactText: string;
  occurrence?: number;
  offset: number;
}

export interface LocalizedReplacement {
  find: string;
  replace: string;
  occurrence?: number;
}

/** A header/footer part path from `inspect().headersFooters`, or a selector. */
export type HeaderFooterPartSelector =
  | string
  | { kind: 'header' | 'footer'; type?: 'default' | 'first' | 'even'; section?: number };

/**
 * Structured list-conversion diagnostic (`list-conversion-diagnostics-v1`). Warnings appear in
 * `diagnostics` (and as `CODE: message` strings in `warnings`) on an operation that succeeded;
 * the same details appear on `error` when the strict policy refuses the operation. They are
 * unrelated to `REVISION_FIDELITY_VIOLATION`.
 */
export interface ListConversionDiagnostic {
  code:
    | 'LIST_MARKER_NOT_BOUND'
    | 'UNSUPPORTED_LIST_CONVERSION'
    | 'LIST_FORMAT_NOT_APPLIED'
    | 'UNSUPPORTED_LIST_FORMAT_CHANGE'
    | 'UNKNOWN_PARAGRAPH_FORMAT_PROPERTY';
  severity: 'warning' | 'error';
  message: string;
  /** The literal marker text (for example `A.`) the edit introduced. */
  listMarker?: string;
  requestedList?: { kind: 'bullet' | 'numbered'; format: string };
  /** The numbering definition the list item keeps (format-change diagnostics). */
  sourceList?: { numId: string; level: number; kind: 'bullet' | 'numbered'; format: string };
  targetExcerpt?: string;
  operationIndex?: number;
  /** Rejected `paragraph-format.properties` keys (UNKNOWN_PARAGRAPH_FORMAT_PROPERTY). */
  properties?: string[];
  recovery?: {
    action: string;
    hint?: string;
    field?: string;
    allowedValue?: string;
    /**
     * The list-format operation to send (add a target): convert-plain for LIST_MARKER_NOT_BOUND /
     * UNSUPPORTED_LIST_CONVERSION, change-format (scope "list") for LIST_FORMAT_NOT_APPLIED / UNSUPPORTED_LIST_FORMAT_CHANGE.
     */
    suggestedOperation?: { type: 'list-format'; list: ListFormatRequest };
  };
}

export interface DocumentOperationBase {
  operationId?: string;
  /** Target a header or footer instead of the body. Comments are not supported there. */
  part?: HeaderFooterPartSelector;
  captureKey?: string;
  target?: string | ParagraphTargetDescriptor;
  targetRef?: number | string | null;
  author?: string;
  generateRedlines?: boolean;
  existingRevisions?: ExistingRevisionsPolicy;
  pairReplacements?: boolean;
  insertionAffinity?: InsertionAffinity;
  /**
   * `'literal'` accepts list-marker text as literal text and suppresses LIST_MARKER_NOT_BOUND,
   * UNSUPPORTED_LIST_CONVERSION, LIST_FORMAT_NOT_APPLIED and UNSUPPORTED_LIST_FORMAT_CHANGE for this operation.
   */
  listMarkers?: 'literal';
}

interface RedlineDocumentOperationCommon extends DocumentOperationBase {
  structuredContent?: boolean;
  targetEnd?: ParagraphTargetDescriptor;
  targetEndRef?: number | string | null;
  /** Required when type is insert and target.revisionView is rejected. */
  anchor?: RejectedTextInsertionAnchor;
}

export type RedlineDocumentOperation =
  | (RedlineDocumentOperationCommon & {
      type: 'redline' | 'replace';
      modified: string;
      replacements?: never;
    })
  | (RedlineDocumentOperationCommon & {
      type: 'redline' | 'replace';
      modified?: never;
      /** Exact, case-sensitive replacements in one batch-start accepted-view paragraph. */
      replacements: LocalizedReplacement[];
    })
  | (RedlineDocumentOperationCommon & {
      type: 'format' | 'list-change' | 'table-reconciliation' | 'insert';
      modified: string;
      replacements?: never;
    });

interface RestoreDocumentOperationCommon extends DocumentOperationBase {
  type: 'restore';
}

export type RestoreDocumentOperation =
  | (RestoreDocumentOperationCommon & {
      modified: string | string[];
      replacements?: never;
      targetEnd?: ParagraphTargetDescriptor;
      targetEndRef?: number | string | null;
    })
  | (RestoreDocumentOperationCommon & {
      modified?: never;
      /** Exact, case-sensitive replacements in one strong rejected-view paragraph. */
      replacements: LocalizedReplacement[];
      targetEnd?: never;
      targetEndRef?: never;
    })
  | (RestoreDocumentOperationCommon & {
      /** Omit modified to restore one strong rejected-view paragraph verbatim. */
      modified?: never;
      replacements?: never;
      targetEnd?: never;
      targetEndRef?: never;
    });

export interface DeleteDocumentOperation extends DocumentOperationBase {
  type: 'delete';
  modified?: '';
}

export interface CommentDocumentOperation extends DocumentOperationBase {
  type: 'comment';
  textToComment?: string;
  commentContent: string;
}

export interface CommentReplyDocumentOperation {
  type: 'comment_reply';
  operationId?: string;
  parentCommentId: number | string;
  commentContent: string;
  author?: string;
  date?: string;
}

/** Resolves (or with `resolved: false`, reopens) the whole thread containing `commentId`. Body-only. */
export interface CommentResolveDocumentOperation {
  type: 'comment_resolve';
  operationId?: string;
  commentId: number | string;
  resolved?: boolean;
}

export interface CharacterFormatProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  strikethrough?: boolean;
  highlight?: string | null;
  color?: string | null;
  fontSize?: number | string | null;
  fontFamily?: string | null;
}

/** Closed set. Unknown keys are ignored with an UNKNOWN_PARAGRAPH_FORMAT_PROPERTY warning and will be rejected in a later minor release. */
export interface ParagraphFormatProperties {
  alignment?: 'left' | 'center' | 'right' | 'both';
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  style?: string | null;
  /**
   * Binds the paragraph to an EXISTING list (`numId` and `level` as reported by inspect), or `null` to remove its direct
   * list binding. Tracked as `w:pPrChange`; refuses `PARAGRAPH_NUMBERING_NOT_FOUND` when the numId/level is not defined.
   * Capability `paragraph-numbering-binding-v1`.
   */
  numbering?: { numId: number | string; level?: number } | null;
}

export interface HighlightDocumentOperation extends DocumentOperationBase {
  type: 'highlight';
  textToHighlight: string;
  color?: string;
}

export interface CharacterFormatDocumentOperation extends DocumentOperationBase {
  type: 'format' | 'character-format';
  textToFormat: string;
  /** 1-based match of `textToFormat` inside the resolved paragraph (default 1). Independent of `target.occurrence`. */
  textOccurrence?: number;
  properties: CharacterFormatProperties;
  formattingRevisionPolicy?: 'always' | 'coalesce-own-insertion';
}

export interface ParagraphFormatDocumentOperation extends DocumentOperationBase {
  type: 'paragraph-format';
  properties: ParagraphFormatProperties;
}

export type ListNumberFormat = 'decimal' | 'upperLetter' | 'lowerLetter' | 'upperRoman' | 'lowerRoman' | 'bullet';

/** `list-format` request. `convert-plain`: `list-format-v1` (CLI contract version 9); `change-format` and `join`: `list-format-change-v1`. */
export type ListFormatRequest = ListFormatConvertPlainRequest | ListFormatChangeFormatRequest | ListFormatJoinRequest;

/** `convert-plain` (`list-format-v1`): bind plain paragraphs to a NEW list. */
export interface ListFormatConvertPlainRequest {
  mode: 'convert-plain';
  kind?: 'bullet' | 'numbered';
  /** Required with `strip: 'none'`; with `strip: 'marker'` it must match the typed markers (and disambiguates `I.`/`C.`). */
  format?: 'decimal' | 'upperLetter' | 'lowerLetter' | 'upperRoman' | 'lowerRoman' | 'bullet';
  /** 0-8, default 0. */
  level?: number;
  /** `'marker'` (default) removes the typed marker as a tracked deletion; `'none'` keeps the text and only binds. */
  strip?: 'marker' | 'none';
  /** `strip: 'none'` only. Defaults to continuing the batch's previous converted list of the same format, else 1. */
  start?: number;
}

/**
 * `change-format` (`list-format-change-v1`): change the numbering format of the target's list. Allocates a NEW
 * definition (copying the source levels and geometry, one level redefined) and rebinds the items with tracked
 * `w:pPrChange`s; the shared definition is never edited, so Reject All restores the original labels.
 */
export interface ListFormatChangeFormatRequest {
  mode: 'change-format';
  kind?: 'bullet' | 'numbered';
  /** The new format of the level (`kind: 'numbered'` alone is refused: name the format). */
  format?: ListNumberFormat;
  /** Level text: numbered `%1`..`%9` with the level's own placeholder (`"%1."`, `"(%2)"`); bullet: the bullet character. */
  lvlText?: string;
  /** Restart value of the level (the new definition's `w:start`); scope `'list'` only. */
  start?: number;
  /** The level to redefine; default: the one level the target paragraphs use. */
  level?: number;
  /** `'list'` (default): every live item of the list; `'range'`: the target range, which must run to the end of the list. */
  scope?: 'list' | 'range';
}

/** `join` (`list-format-change-v1`): the target list (or a range running to its end) continues an earlier list. */
export interface ListFormatJoinRequest {
  mode: 'join';
  /** An item of the earlier list to continue; bound at batch start like `target`. */
  continueFrom: ParagraphTargetDescriptor;
  scope?: 'list' | 'range';
}

/**
 * Binds plain paragraphs (target, or target..targetEnd in one container) to a NEW list definition in place. The
 * binding is a tracked `w:pPrChange`; the stripped marker is a tracked `w:del`. Each typed marker's value is the
 * item's value, so Accept All renders the labels the text showed. Body only (no `part`); text changes belong in a
 * separate `redline`. Refusals: LIST_MARKER_NOT_FOUND, LIST_MARKER_AMBIGUOUS, LIST_MARKER_UNSUPPORTED,
 * LIST_FORMAT_TARGET_IS_LIST, LIST_FORMAT_STYLE_NUMBERED, PARAGRAPH_FORMAT_REVISION_CONFLICT.
 *
 * `change-format` / `join` (`list-format-change-v1`) rebind existing list items with tracked `w:pPrChange`s.
 * Refusals: LIST_FORMAT_TARGET_NOT_LIST, LIST_FORMAT_MIXED_RANGE, LIST_FORMAT_STYLE_LINKED, LIST_FORMAT_STYLE_NUMBERED,
 * LIST_FORMAT_SPLIT_UNSUPPORTED, LIST_FORMAT_DEFINITION_UNSUPPORTED, LIST_FORMAT_JOIN_INVALID,
 * PARAGRAPH_FORMAT_REVISION_CONFLICT; a no-op is `ok` with a LIST_FORMAT_UNCHANGED warning.
 */
export interface ListFormatDocumentOperation extends Omit<DocumentOperationBase, 'part'> {
  type: 'list-format';
  list: ListFormatRequest;
  targetEnd?: ParagraphTargetDescriptor;
  targetEndRef?: number | string | null;
}

/** What a `list-format` operation bound or rebound (`result.list`). */
export type ListFormatResult = ListFormatConvertPlainResult | ListFormatRebindResult;

/** `convert-plain` result. */
export interface ListFormatConvertPlainResult {
  mode: 'convert-plain';
  format: 'decimal' | 'upperLetter' | 'lowerLetter' | 'upperRoman' | 'lowerRoman' | 'bullet';
  level: number;
  strip: 'marker' | 'none';
  items: Array<{ paragraphIndex: number; paragraphId: string | null; marker: string | null; value: number | null; numId: string }>;
  /** Paragraphs of the range left as they were: empty, or (strip 'marker') body text without a marker. */
  skipped: Array<{ paragraphIndex: number; reason: 'empty' | 'no-marker' }>;
}

/** `change-format` / `join` result: every paragraph rebound (scope `'list'` includes items outside the target). */
export interface ListFormatRebindResult {
  mode: 'change-format' | 'join';
  scope: 'list' | 'range';
  sourceNumId: string;
  /** change-format: the newly allocated numId; join: the continued list's numId. */
  numId: string;
  format?: ListNumberFormat;
  level?: number;
  lvlText?: string;
  /** change-format: per-level `w:start` values of the new definition (carried restarts, split continuation, `start`). */
  levelStarts?: Record<string, number>;
  continueFrom?: { paragraphIndex: number | null; numId: string };
  items: Array<{ paragraphIndex: number | null; paragraphId: string | null; level: number; numId: string; sourceNumId: string; inTarget: boolean }>;
  /** Tracked-deleted paragraphs in the range keep their binding. */
  skipped: Array<{ paragraphIndex: number | null; reason: 'tracked-deleted' }>;
}

export type DocumentOperation =
  | RedlineDocumentOperation
  | RestoreDocumentOperation
  | DeleteDocumentOperation
  | CommentDocumentOperation
  | CommentReplyDocumentOperation
  | CommentResolveDocumentOperation
  | HighlightDocumentOperation
  | CharacterFormatDocumentOperation
  | ParagraphFormatDocumentOperation
  | ListFormatDocumentOperation;

export interface ResolvedDocumentTarget {
  index: number;
  paragraphId: string | null;
  text: string;
  fingerprint?: string;
  revisionView?: 'accepted' | 'rejected';
  inTable?: boolean;
  targetTextMatch?: {
    mode: 'exact' | 'space_equivalent' | 'normalized';
    differences?: Array<{
      offset: number;
      sourceCodePoint: string;
      requestedCodePoint: string;
    }>;
    sourceExcerpt?: string;
    requestedExcerpt?: string;
  };
}

export interface ResolvedCommentAnchor {
  requestIndex: number;
  paragraphIndex: number;
  text: string;
  resolvedBy: 'exact_anchor' | 'space_equivalent_anchor';
  start: number;
  end: number;
}

export interface StandaloneRunnerOptions {
  atomic?: boolean;
  continueOnError?: boolean;
  generateRedlines?: boolean;
  existingRevisions?: ExistingRevisionsPolicy;
  strictTargets?: boolean;
  /**
   * Make list-conversion diagnostics errors (UNSUPPORTED_LIST_CONVERSION, UNSUPPORTED_LIST_FORMAT_CHANGE)
   * instead of warnings. Defaults to `atomic`; the CLI agent profile sets it. An operation's
   * `listMarkers: 'literal'` always wins.
   */
  strictListConversion?: boolean;
  /**
   * Post-apply check that Word's Reject All of each tracked operation's output equals the source
   * (paragraph count, text, properties, numbering, wrapper membership). Default `'enforce'` refuses
   * the operation with `REVISION_FIDELITY_VIOLATION` (per-invariant `violations[]` and
   * `recovery.retryGuidance`); `'warn'` applies it and appends the same message to `warnings`;
   * `'off'` is for diagnosis only.
   */
  verifyRevisionFidelity?: 'enforce' | 'warn' | 'off';
  expectedRevision?: RevisionToken;
  pairReplacements?: boolean;
  insertionAffinity?: InsertionAffinity;
  onInfo?: (message: string) => void;
  onWarn?: (message: string) => void;
  [key: string]: unknown;
}

export interface MutationReceiptRevisionItem {
  id: string;
  kind: 'ins' | 'del' | 'move_from' | 'move_to' | 'rPrChange' | 'pPrChange' | 'structural';
  partName: string;
}

export interface MutationReceipt {
  operationIndex: number;
  operationId?: string;
  attemptedDisposition: 'applied' | 'no_change' | 'refused' | 'not_attempted';
  finalDisposition: 'applied' | 'no_change' | 'refused' | 'rolled_back' | 'not_attempted';
  committed: boolean;
  authorUsed?: string;
  revisionItems: MutationReceiptRevisionItem[];
  commentIds: string[];
  numberingIds: string[];
  relationshipIds: string[];
  affectedTargets: ResolvedDocumentTarget[];
  warnings: string[];
}

export interface LocalizedReplacementChange {
  kind: 'localized_replacement';
  committed: boolean;
  finalDisposition: 'applied' | 'rolled_back';
  target: {
    paragraphId?: string;
    index?: number;
    fingerprint?: string;
    humanReference?: string;
  };
  context?: {
    search?: string;
    range?: string;
    anchorMatchCount?: number;
    humanReference?: string;
  };
  replacements: Array<{
    find: string;
    replace: string;
    occurrence?: number;
    matchMode?: 'space_equivalent';
    beforeExcerpt: string;
    afterExcerpt: string;
  }>;
  verification: {
    acceptedViewMatchesCompiledText: boolean;
  };
}

export interface DocumentOperationResult {
  documentXml: string;
  hasChanges: boolean;
  numberingXml?: string | null;
  commentsXml?: string | null;
  commentsExtendedXml?: string | null;
  commentsXmlMode?: 'merge' | 'replace';
  commentsExtendedXmlMode?: 'merge' | 'replace';
  warnings?: string[];
  diagnostics?: ListConversionDiagnostic[];
  status?: RedlineStatus;
  error?: RedlineError;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'comment_resolve' | 'highlight' | 'format' | 'paragraph-format' | 'list-format';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
  receipt?: MutationReceipt;
  change?: LocalizedReplacementChange;
  /** `list-format` operations: what was bound. */
  list?: ListFormatResult;
}

export interface BatchOperationItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'comment_resolve' | 'highlight' | 'format' | 'paragraph-format' | 'list-format';
  status: 'applied' | 'no_change' | 'error';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
  warnings?: string[];
  diagnostics?: ListConversionDiagnostic[];
  error?: RedlineError;
  receipt?: MutationReceipt;
  change?: LocalizedReplacementChange;
  /** `list-format` operations: what was bound. */
  list?: ListFormatResult;
}

export interface DocumentOperationBatchResult {
  documentXml: string;
  hasChanges: boolean;
  commentsXml: string | null;
  commentsExtendedXml?: string | null;
  commentsXmlMode?: 'merge' | 'replace';
  commentsExtendedXmlMode?: 'merge' | 'replace';
  numberingXmlParts: string[];
  results: BatchOperationItemResult[];
  receipts?: MutationReceipt[];
  executionOrder: number[];
  /** Authors attached to committed changes. Empty when an atomic batch rolls back. */
  authorsUsed: string[];
  rolledBack?: boolean;
  status?: RedlineStatus;
  error?: RedlineError;
  warnings?: string[];
  conflicts?: OperationConflict[];
  retryPlan?: MutationRetryPlan;
}

export interface MutationRetryPlan {
  base: 'original' | 'output';
  committedIndexes: number[];
  failedIndexes: number[];
  unattemptedIndexes: number[];
  replayWholeBatch: boolean;
  sameArgumentsSafe: false;
}

export interface OperationPreflightItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'comment_resolve' | 'highlight' | 'format' | 'paragraph-format' | 'list-format';
  status: 'ready' | 'deferred' | 'error';
  authorUsed: string;
  resolvedBy?: string;
  captureRef?: string;
  select?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  matchDiagnostics?: {
    exactTextMatch: boolean;
    normalizedTextMatch: boolean;
    suppliedText: string;
    actualText: string;
  };
  anchor?: {
    text: string;
    found: boolean;
    resolvedBy?: 'exact_anchor' | 'space_equivalent_anchor';
    start?: number;
    end?: number;
    candidates?: Array<{ start: number; end: number }>;
  } | null;
  hasRevisions?: boolean;
  existingRevisions?: ExistingRevisionsPolicy;
  warnings?: string[];
  error?: RedlineError & { candidates?: ResolvedDocumentTarget[] };
}

export interface OperationConflict {
  code: 'OVERLAPPING_SOURCE_TARGETS' | 'OVERLAPPING_TEXT_EDITS' | 'REVISION_ORDER_CONFLICT' | 'CAPTURE_FANOUT_CONFLICT' | string;
  message: string;
  operationIndexes: number[];
  target: ResolvedDocumentTarget;
  recoveryVersion?: number;
  category?: string;
  recovery?: RedlineError['recovery'];
}

export interface OperationPreflightResult {
  valid: boolean;
  status: 'ok' | 'error';
  results: OperationPreflightItemResult[];
  conflicts: OperationConflict[];
  authorsUsed: string[];
  requiredArtifacts: { comments: boolean; numbering: boolean };
  error?: RedlineError;
}

export interface OperationDependencyPlan {
  valid: boolean;
  scheduled?: Array<{ operation: DocumentOperation; index: number }>;
  captureProducers?: Map<string, number>;
  error?: RedlineError;
}

export function buildOperationDependencyPlan(operations?: DocumentOperation[]): OperationDependencyPlan;

export function orderOperationsForStableTargets(operations?: DocumentOperation[]): DocumentOperation[];

export function preflightOperations(
  documentXml: string,
  operations: DocumentOperation[],
  author?: string,
  options?: StandaloneRunnerOptions
): OperationPreflightResult;

export function applyOperationToDocumentXml(
  documentXml: string,
  operation: DocumentOperation,
  author?: string,
  runtimeContext?: Record<string, unknown> | null,
  options?: StandaloneRunnerOptions
): Promise<DocumentOperationResult>;

export function applyOperationsToDocumentXml(
  documentXml: string,
  operations: DocumentOperation[],
  author?: string,
  runtimeContext?: Record<string, unknown> | null,
  options?: StandaloneRunnerOptions
): Promise<DocumentOperationBatchResult>;
