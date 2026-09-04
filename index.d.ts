export type OoxmlSourceType = 'package' | 'document' | 'fragment';
export type RedlineStatus = 'ok' | 'no-op' | 'error';
export type ExistingRevisionsPolicy = 'reject-input' | 'accept-all-first' | 'accept-all-first-keep-normalized';

export type {
  BatchOperationItemResult,
  CommentDocumentOperation,
  DeleteDocumentOperation,
  DocumentOperation,
  DocumentOperationBatchResult,
  HighlightDocumentOperation,
  OperationConflict,
  OperationPreflightItemResult,
  OperationPreflightResult,
  ParagraphTargetDescriptor,
  RedlineDocumentOperation,
  ResolvedCommentAnchor,
  ResolvedDocumentTarget,
  StandaloneRunnerOptions
} from './services/standalone-operation-runner.js';

export interface RedlineError {
  code: 'PARSE_ERROR' | 'TARGET_NOT_FOUND' | 'EXISTING_REVISIONS' | 'DIFF_TOKEN_LIMIT' | string;
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

export function ingestWordOoxmlToPlainText(oxml: string): string;
export function ingestWordOoxmlToMarkdown(oxml: string): string;
export function ingestWordOoxmlToPlainTextResult(oxml: unknown): IngestionTextResult;
export function ingestWordOoxmlToMarkdownResult(oxml: unknown): IngestionTextResult;
export function extractCanonicalParagraphText(paragraph: Element | null | undefined, options?: { revisionView?: 'accepted' | 'rejected' | 'current' }): string;
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
}
export interface InspectedComment { id: string; author: string | null; date: string | null; text: string; paragraphIndex?: number; targetRef?: string; anchoredText?: string; }
export interface DocumentInspectionOptions { revisionView?: 'accepted' | 'rejected' | 'current'; excerptLength?: number; revisedOnly?: boolean; inTable?: boolean; skipEmpty?: boolean; search?: string; indexes?: number[]; range?: { start: number; end: number } | [number, number]; }
export interface DocumentInspectionResult {
  status: 'ok' | 'error'; paragraphs: InspectedParagraph[]; comments: InspectedComment[];
  revisionAuthors?: string[]; commentAuthors?: string[]; counts?: { paragraphs: number; comments: number; revisedParagraphs: number };
  warnings: string[]; error?: RedlineError;
}
export function inspectDocumentParts(parts: { documentXml: string; commentsXml?: string | null; numberingXml?: string | null; stylesXml?: string | null }, options?: DocumentInspectionOptions): DocumentInspectionResult;
export function ingestOoxml(oxml: string): unknown;
export function preprocessMarkdown(text: string): { cleanText: string; formatHints: unknown[] };
export function serializeToOoxml(runModel: unknown[], pPrXml?: string | null, formatHints?: unknown[], options?: Record<string, unknown>): string;
export function wrapInDocumentFragment(rawOoxml: string, options?: Record<string, unknown>): string;

export function injectCommentsIntoOoxml(oxml: string, comments: unknown[], options?: Record<string, unknown>): unknown;
export function injectCommentsIntoPackage(packageXml: string, comments: unknown[], options?: Record<string, unknown>): unknown;
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
