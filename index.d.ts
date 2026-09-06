export type OoxmlSourceType = 'package' | 'document' | 'fragment';
export type RedlineStatus = 'ok' | 'no-op' | 'error';
export type ExistingRevisionsPolicy = 'reject-input' | 'accept-all-first' | 'accept-all-first-keep-normalized';
export type RevisionView = 'accepted' | 'rejected';

export interface RevisionTextSegment {
  text: string;
  kind: 'baseline' | 'insertion' | 'deletion' | 'move_from' | 'move_to';
  author?: string;
  revisionId?: string;
  acceptedStart: number | null;
  rejectedStart: number | null;
}

import type { InsertionAffinity } from './services/standalone-operation-runner.js';

export type {
  BatchOperationItemResult,
  CommentDocumentOperation,
  CommentReplyDocumentOperation,
  DeleteDocumentOperation,
  DocumentOperation,
  DocumentOperationBatchResult,
  HighlightDocumentOperation,
  InsertionAffinity,
  OperationConflict,
  OperationDependencyPlan,
  OperationPreflightItemResult,
  OperationPreflightResult,
  ParagraphTargetDescriptor,
  RedlineDocumentOperation,
  ResolvedCommentAnchor,
  ResolvedDocumentTarget,
  StandaloneRunnerOptions
} from './services/standalone-operation-runner.js';

export interface RedlineError {
  code:
    | 'PARSE_ERROR'
    | 'TARGET_NOT_FOUND'
    | 'EXISTING_REVISIONS'
    | 'DIFF_TOKEN_LIMIT'
    | 'UNSUPPORTED_REVISION_VIEW_MUTATION'
    | 'UNSUPPORTED_INSERTION_AFFINITY'
    | 'UNSAFE_PARAGRAPH_BOUNDARY'
    | 'TARGET_INDEX_MISMATCH'
    | 'REVISION_MISMATCH'
    | 'REVISION_TOKEN_SCOPE_MISMATCH'
    | 'INVALID_REVISION_TOKEN'
    | 'DUPLICATE_CAPTURE_KEY'
    | 'CAPTURE_NOT_FOUND'
    | 'CAPTURE_DEPENDENCY_CYCLE'
    | 'AMBIGUOUS_CAPTURE_SELECTION'
    | 'CAPTURE_STALE'
    | string;
  message: string;
  commentIds?: string[];
  comments?: Array<{ id: string; author: string; text: string }>;
}

export interface RedlineOptions {
  generateRedlines?: boolean;
  author?: string;
  targetParagraphId?: string | null;
  existingRevisions?: ExistingRevisionsPolicy;
  removeFormatting?: boolean;
  sanitizeInput?: boolean;
  structuredContent?: boolean;
  pairReplacements?: boolean;
  insertionAffinity?: InsertionAffinity;
  [key: string]: unknown;
}

export interface RedlineResult {
  oxml: string;
  hasChanges: boolean;
  sourceType?: OoxmlSourceType;
  status?: RedlineStatus;
  error?: RedlineError;
  warnings?: string[];
  numberingXml?: string;
  useNativeApi?: boolean;
  [key: string]: unknown;
}

export interface TableReconciliationResult extends RedlineResult {
  isMarkdownTable: boolean;
  tableData?: {
    headers?: string[];
    rows?: string[][];
    [key: string]: unknown;
  };
}

export interface RevisionFilterOptions {
  author?: string;
  allAuthors?: boolean;
}

export interface AcceptTrackedChangesResult {
  oxml: string;
  hasChanges: boolean;
  acceptedCount: number;
  warnings: string[];
  status?: RedlineStatus;
  error?: RedlineError;
}

export interface RejectTrackedChangesResult {
  oxml: string;
  hasChanges: boolean;
  rejectedCount: number;
  warnings: string[];
  status?: RedlineStatus;
  error?: RedlineError;
}

export interface DeleteCommentsResult {
  oxml: string;
  hasChanges: boolean;
  commentsRemoved: number;
  referencesRemoved: number;
  warnings: string[];
  status?: RedlineStatus;
  error?: RedlineError;
}

export interface IngestionTextResult {
  text: string;
  status: 'ok' | 'error';
  error?: RedlineError;
  warnings?: string[];
}

export interface ParagraphTargetCandidate {
  paragraph: Element;
  index: number;
  inTable: boolean;
  paragraphId: string | null;
  fingerprint: string | null;
  text: string;
}

export interface ParagraphTargetResolution {
  paragraph: Element;
  resolvedBy:
    | 'ref'
    | 'paragraph_id'
    | 'occurrence'
    | 'fingerprint'
    | 'strict_text'
    | 'fuzzy_text'
    | 'strict_text_after_ref_drift'
    | 'fuzzy_text_after_ref_drift';
}

export interface XmlProvider {
  DOMParser: typeof DOMParser;
  XMLSerializer: typeof XMLSerializer;
}

export interface LoggerConfig {
  log?: (...args: unknown[]) => void;
  warn?: (...args: unknown[]) => void;
  error?: (...args: unknown[]) => void;
}

export interface DocumentOperationResult {
  documentXml: string;
  hasChanges: boolean;
  numberingXml?: string | null;
  commentsXml?: string | null;
  commentsExtendedXml?: string | null;
  warnings?: string[];
  status?: RedlineStatus;
  error?: RedlineError;
}

export function configureXmlProvider(provider: XmlProvider): void;
export function parseOoxmlSafe(oxml: unknown, contentType?: string): { doc: Document | null; error: RedlineError | null; warnings: string[] };
export function configureLogger(logger: LoggerConfig): void;
export function setDefaultAuthor(name: string): void;
export function getDefaultAuthor(): string;
export function setPlatform(label: string): void;
export function getPlatform(): string;

export function applyRedlineToOxml(
  oxml: string,
  originalText: string,
  modifiedText: string,
  options?: RedlineOptions
): Promise<RedlineResult>;

export function applyRedlineToOxmlWithListFallback(
  oxml: string,
  originalText: string,
  modifiedText: string,
  options?: RedlineOptions
): Promise<RedlineResult>;

export function reconcileMarkdownTableOoxml(
  oxml: string,
  originalText: string,
  markdownTable: string,
  options?: RedlineOptions
): Promise<TableReconciliationResult>;

export interface StructuredContentIssue {
  severity: 'error' | 'warning';
  code: string;
  line: number;
  message: string;
}
export interface StructuredContentBlock {
  type: 'heading' | 'paragraph' | 'list' | 'table';
  markdown: string;
  text?: string;
  level?: number;
  items?: number;
  columns?: number;
  rows?: number;
  hasHeader?: boolean;
}
export interface StructuredContentAnalysis {
  valid: boolean;
  normalizedMarkdown: string;
  blocks: StructuredContentBlock[];
  issues: StructuredContentIssue[];
  counts: Record<'heading' | 'paragraph' | 'list' | 'table', number>;
  requiresStructuredContent: boolean;
}
export function analyzeStructuredContent(markdown: string): StructuredContentAnalysis;
export function planStructuredReplacement(
  target: string | import('./services/standalone-operation-runner.js').ParagraphTargetDescriptor,
  markdown: string,
  options?: { author?: string; generateRedlines?: boolean; existingRevisions?: ExistingRevisionsPolicy }
): StructuredContentAnalysis & { operation: import('./services/standalone-operation-runner.js').RedlineDocumentOperation | null };

export function ingestWordOoxmlToPlainText(oxml: string): string;
export function ingestWordOoxmlToMarkdown(oxml: string): string;
export function ingestWordOoxmlToPlainTextResult(oxml: unknown): IngestionTextResult;
export function ingestWordOoxmlToMarkdownResult(oxml: unknown): IngestionTextResult;
export function extractCanonicalParagraphText(paragraph: Element | null | undefined, options?: { revisionView?: 'accepted' | 'rejected' | 'current' }): string;
export function extractParagraphRevisionSegments(paragraph: Element | null | undefined, options?: { mergeRuns?: boolean; [key: string]: unknown }): RevisionTextSegment[];
export function readCanonicalRunText(run: Element, options?: { revisionView?: 'accepted' | 'rejected' | 'current'; boundary?: Element | null }): string;
export function isNodeVisibleInRevisionView(node: Node, boundary?: Node | null, revisionView?: 'accepted' | 'rejected' | 'current'): boolean;

export interface InspectedParagraph {
  index: number; ref: string; paragraphId: string | null; fingerprint: string | null;
  text: string; exactText: string; excerpt: string; inTable: boolean;
  humanReference: string; styleId: string | null;
  table: { tableIndex: number; rowIndex: number; cellIndex: number } | null;
  structuralReferences: Array<{ type: 'footnote' | 'endnote' | 'comment'; id: string | null }>;
  headingLevel: number | null; nearestHeading: { level: number; text: string } | null;
  list: { numId: string; level: number; label: string | null; format: string | null } | null;
  hasRevisions: boolean; revisionAuthors: string[]; commentIds: string[];
  segments: RevisionTextSegment[];
}
export interface RevisionToken {
  algorithm: 'sha256';
  version: 1;
  scope: 'document-parts' | 'package';
  value: string;
  coveredParts?: string[];
}

export interface InspectedComment { id: string; author: string | null; date: string | null; text: string; paraId?: string | null; parentParaId?: string; parentCommentId?: string | null; done?: boolean; paragraphIndex?: number; targetRef?: string; anchoredText?: string; }
export interface DocumentInspectionOptions { revisionView?: 'accepted' | 'rejected' | 'current'; excerptLength?: number; revisedOnly?: boolean; inTable?: boolean; skipEmpty?: boolean; search?: string; indexes?: number[]; range?: { start: number; end: number } | [number, number]; digestFn?: (bytes: Uint8Array) => string; }
export interface DocumentInspectionResult {
  status: 'ok' | 'error'; paragraphs: InspectedParagraph[]; comments: InspectedComment[];
  revisionAuthors?: string[]; commentAuthors?: string[]; counts?: { paragraphs: number; comments: number; revisedParagraphs: number };
  revisionToken?: RevisionToken | null;
  coveredParts?: string[];
  warnings: string[]; error?: RedlineError;
}
export function inspectDocumentParts(parts: { documentXml: string; commentsXml?: string | null; commentsExtendedXml?: string | null; numberingXml?: string | null; stylesXml?: string | null; parts?: Map<string, unknown>; additionalParts?: Record<string, unknown> }, options?: DocumentInspectionOptions): DocumentInspectionResult;
export function buildRevisionTokenFraming(options: { scope: string; entries?: unknown }): { framing: Uint8Array; scope: string; version: number; coveredParts: string[] };
export function computeRevisionToken(options: { scope: string; entries?: unknown; digestFn?: (bytes: Uint8Array) => Promise<string> }): Promise<RevisionToken>;
export function computeRevisionTokenSync(options: { scope: string; entries?: unknown; digestFn: (bytes: Uint8Array) => string }): RevisionToken;
export function computeDocumentPartsRevisionToken(parts: unknown, options?: { digestFn?: (bytes: Uint8Array) => Promise<string> }): Promise<RevisionToken>;
export function validateRevisionToken(token: unknown): { valid: boolean; error?: { code: string; message: string } };
export function areRevisionTokensEqual(a: string, b: string): boolean;
export function normalizeOpcEntryName(name: string): string;
export function extractDocumentPartsEntries(parts: unknown): Array<{ name: string; payload: unknown }>;

export function ingestOoxml(oxml: string): unknown;
export function preprocessMarkdown(text: string): { cleanText: string; formatHints: unknown[] };
export function serializeToOoxml(runModel: unknown[], pPrXml?: string | null, formatHints?: unknown[], options?: Record<string, unknown>): string;
export function wrapInDocumentFragment(rawOoxml: string, options?: Record<string, unknown>): string;

export function injectCommentsIntoOoxml(oxml: string, comments: unknown[], options?: Record<string, unknown>): unknown;
export function injectCommentsIntoPackage(packageXml: string, comments: unknown[], options?: Record<string, unknown>): unknown;
export function applyCommentReplyToParts(options: { commentsXml: string; commentsExtendedXml?: string | null; parentCommentId: string | number; commentId: string | number; commentContent: string; author: string; date?: string }): { status: 'ok' | 'error'; hasChanges?: boolean; commentsXml?: string; commentsExtendedXml?: string; commentsXmlMode?: 'replace'; commentsExtendedXmlMode?: 'replace'; parentCommentId?: string; paraId?: string; parentParaId?: string; error?: RedlineError };
export function acceptTrackedChangesInOoxml(oxml: string, options?: RevisionFilterOptions): AcceptTrackedChangesResult;
export function rejectTrackedChangesInOoxml(oxml: string, options?: RevisionFilterOptions): RejectTrackedChangesResult;
export function deleteCommentsByAuthorInOoxml(oxml: string, options?: RevisionFilterOptions): DeleteCommentsResult;
export function containsTrackedChanges(xmlDoc: Document | Element): boolean;

export type RedlineValidationSeverity = 'error' | 'warning';

export interface RedlineValidationIssue {
  code:
    | 'PARSE_ERROR'
    | 'NESTED_REVISION'
    | 'DEL_CONTAINS_T'
    | 'MISSING_REVISION_METADATA'
    | 'DUPLICATE_REVISION_ID'
    | 'MISSING_SPACE_PRESERVE'
    | 'EMPTY_TEXT_ELEMENT'
    | 'EMPTY_REVISION_WRAPPER'
    | string;
  severity: RedlineValidationSeverity;
  message: string;
}

export interface RedlineValidationResult {
  valid: boolean;
  issues: RedlineValidationIssue[];
}

export function validateRedlineOoxml(oxml: string): RedlineValidationResult;

export function applyHighlightToOoxml(oxml: string, targetText: string, color: string, options?: Record<string, unknown>): string;
export function generateTableOoxml(headersOrData: unknown, rowsOrOptions?: unknown, options?: Record<string, unknown>): string;
export function extractReplacementNodesFromOoxml(oxml: string): unknown;
export function validateDocxPackage(zip: unknown): Promise<unknown> | unknown;
export function ensureNumberingArtifactsInZip(zip: unknown, numberingXml: string): Promise<unknown> | unknown;
export function ensureCommentsArtifactsInZip(zip: unknown, commentsXml: string): Promise<unknown> | unknown;
export function ensureCommentsExtendedArtifactsInZip(zip: unknown, commentsExtendedXml: string): Promise<unknown> | unknown;
export function createDynamicNumberingIdState(numberingXml?: string): unknown;

export function parseOoxml(ooxml: string): Document | null;
export function serializeOoxml(node: Node): string;
export function sanitizeAiResponse(text: string): string;

export class ReconciliationPipeline {
  constructor(options?: Record<string, unknown>);
  execute(oxml: string, modifiedText: string, options?: Record<string, unknown>): Promise<unknown>;
}

export class NumberingService {
  constructor(...args: unknown[]);
}

// Supporting public utilities. These declarations intentionally expose stable
// boundary shapes while leaving internal planning records extensible.
export const NS_W: string;
export const WORD_MAIN_NS: string;
export const DiffOp: Readonly<Record<'EQUAL' | 'DELETE' | 'INSERT', string>>;
export const RunKind: Readonly<Record<string, string>>;
export const ContainerKind: Readonly<Record<string, string>>;
export const ContentType: Readonly<Record<string, string>>;
export const RoutePlanKind: Readonly<Record<string, string>>;

export function escapeXml(value: unknown): string;
export function applyFormattingRemovalToOoxml(oxml: string, options?: Record<string, unknown>): string;
export function removeFormattingFromRPr(rPr: Element, options?: Record<string, unknown>): Element;

export function buildCommentElement(comment: Record<string, unknown>, id?: string | number): string;
export function buildCommentsPartXml(comments: unknown[]): string;

export function reserveNextNumberingId(state: unknown): number;
export function reserveNextNumberingIdPair(state: unknown): { abstractNumId: number; numId: number };
export function overwriteParagraphNumIds(oxml: string, numId: string | number): string;
export function extractFirstParagraphNumId(oxml: string): string | null;
export function buildExplicitDecimalMultilevelNumberingXml(options?: Record<string, unknown>): string;
export function remapNumberingPayloadForDocument(numberingXml: string, state: unknown): unknown;
export function mergeNumberingXmlBySchemaOrder(baseXml: string, incomingXml: string): string;

export function buildListMarkdown(items: unknown[], options?: Record<string, unknown>): string;
export function inferNumberingStyleFromMarker(marker: string): string;
export function normalizeListItemsWithLevels(items: unknown[]): unknown[];
export function parseMarkdownListContent(text: string): unknown;
export function hasListItems(parsed: unknown): boolean;

export function buildReconciliationPlan(params?: Record<string, unknown>): Record<string, unknown>;
export function normalizeContentEscapesForRouting(content: string): string;
export function buildSingleLineListStructuralFallbackPlan(options?: Record<string, unknown>): Record<string, unknown> | null;
export function executeSingleLineListStructuralFallback(plan: unknown, options?: Record<string, unknown>): Promise<RedlineResult>;
export function resolveSingleLineListFallbackNumberingAction(plan: unknown, sequenceState?: unknown): Record<string, unknown>;
export function recordSingleLineListFallbackExplicitSequence(sequenceState: unknown, numberingKey: string | null, numId: string | number | null, startAt: number | null): void;
export function clearSingleLineListFallbackExplicitSequence(sequenceState: unknown, numberingKey: string | null): void;
export function enforceListBindingOnParagraphNodes(nodes: Node[], options?: Record<string, unknown>): number;
export function stripSingleLineListMarkerPrefix(text: string): string;

export function getParagraphText(paragraph: Element | null | undefined): string;
export function getParagraphId(paragraph: Element | null | undefined): string | null;
export function createParagraphFingerprint(paragraph: Element | null | undefined): string | null;
export function getDocumentParagraphNodes(xmlDoc: Document | Element | null | undefined): Element[];
export function normalizeWhitespaceForTargeting(text: string): string;
export function isMarkdownTableText(text: string): boolean;
export function parseParagraphReference(reference: unknown): unknown;
export function stripLeadingParagraphMarker(text: string): string;
export function splitLeadingParagraphMarker(text: string): Record<string, unknown>;
export function findContainingWordElement(node: Node | null, localName: string): Element | null;
export function findParagraphByReference(xmlDoc: Document | Element, reference: unknown): Element | null;
export function findParagraphByStrictText(xmlDoc: Document | Element, text: string): Element | null;
export function findParagraphByBestTextMatch(xmlDoc: Document | Element, text: string): Element | null;
export function findStrictTargetCandidates(xmlDoc: Document | Element, text: string): ParagraphTargetCandidate[];
export function resolveTargetParagraph(xmlDoc: Document | Element, options?: Record<string, unknown>): ParagraphTargetResolution;
export function buildTargetReferenceSnapshot(xmlDoc: Document | Element): Map<number, {
  text: string;
  normalizedText: string;
  inTable: boolean;
}>;
export function resolveTargetParagraphWithSnapshot(
  xmlDoc: Document | Element,
  options?: Record<string, unknown>
): ParagraphTargetResolution;
export function resolveParagraphRangeByRefs(
  xmlDoc: Document | Element,
  startRef: string | number | null,
  endRef: string | number | null,
  options?: Record<string, unknown>
): Element[] | null;
export function extractParagraphIdFromOoxml(oxml: string): string | null;

export function getParagraphListInfo(paragraph: Element): Record<string, unknown> | null;
export function collectContiguousListParagraphBlock(paragraph: Element): Element[];
export function synthesizeExpandedListScopeEdit(options: Record<string, unknown>): Record<string, unknown> | null;
export function planListInsertionOnlyEdit(options: Record<string, unknown>): Record<string, unknown> | null;
export function stripRedundantLeadingListMarkers(text: string): string;

export function inferTableReplacementParagraphBlock(options: Record<string, unknown>): Record<string, unknown> | null;
export function isLikelyStructuredTableSourceParagraph(paragraph: Element): boolean;
export function synthesizeTableMarkdownFromMultilineCellEdit(options: Record<string, unknown>): string | null;

export function getBodyElementFromDocument(xmlDoc: Document): Element | null;
export function insertBodyElementBeforeSectPr(body: Element, element: Element): Element;
export function normalizeBodySectionOrderStandalone(documentXml: string): string;
export function sanitizeNestedParagraphsInTables(documentXml: string): string;
export function getPackagePartName(part: unknown): string | null;

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
  affectedTargets: unknown[];
  warnings: string[];
}

export class ReceiptCollector {
  constructor();
  beginOperation(operationIndex: number, operationId?: string | null, authorUsed?: string | null): void;
  recordRevision(id: string | number, kind?: string, partName?: string): void;
  recordComment(id: string | number, partName?: string): void;
  recordNumbering(id: string | number, partName?: string): void;
  recordRelationship(id: string | number, partName?: string): void;
  recordAffectedTarget(target: unknown): void;
  recordWarning(warning: unknown): void;
  commitOperation(disposition?: string): MutationReceipt | null;
  abortOperation(disposition?: string): MutationReceipt | null;
  createSavepoint(): unknown;
  restoreSavepoint(savepoint: unknown): void;
  clear(): void;
  markRolledBack(): void;
  getReceipts(): MutationReceipt[];
  getCurrentReceipt(): MutationReceipt | null;
}

export function createEmptyReceipt(
  operationIndex: number,
  operationId?: string | null,
  authorUsed?: string | null,
  disposition?: 'applied' | 'no_change' | 'refused' | 'not_attempted'
): MutationReceipt;

export function reconcileReceiptsAgainstOutput(
  parts: {
    documentXml: string;
    commentsXml?: string | null;
    numberingXml?: string | null;
    numberingXmlParts?: string[];
    relationshipsXml?: string | null;
  },
  receipts: MutationReceipt[]
): { valid: boolean; error?: { code: string; message: string } };
