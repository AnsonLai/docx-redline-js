export type OoxmlSourceType = 'package' | 'document' | 'fragment';
export type RedlineStatus = 'ok' | 'no-op' | 'error';
export type ExistingRevisionsPolicy = 'reject-input' | 'accept-all-first';

export interface RedlineError {
  code: 'PARSE_ERROR' | 'TARGET_NOT_FOUND' | 'EXISTING_REVISIONS' | string;
  message: string;
}

export interface RedlineOptions {
  generateRedlines?: boolean;
  author?: string;
  targetParagraphId?: string | null;
  existingRevisions?: ExistingRevisionsPolicy;
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
}

export interface RejectTrackedChangesResult {
  oxml: string;
  hasChanges: boolean;
  rejectedCount: number;
  warnings: string[];
}

export interface DeleteCommentsResult {
  oxml: string;
  hasChanges: boolean;
  commentsRemoved: number;
  referencesRemoved: number;
  warnings: string[];
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

export function applyHighlightToOoxml(oxml: string, targetText: string, color: string, options?: Record<string, unknown>): string;
export function generateTableOoxml(headersOrData: unknown, rowsOrOptions?: unknown, options?: Record<string, unknown>): string;
export function extractReplacementNodesFromOoxml(oxml: string): unknown;
export function validateDocxPackage(zip: unknown): Promise<unknown> | unknown;
export function ensureNumberingArtifactsInZip(zip: unknown, numberingXml: string): Promise<unknown> | unknown;
export function ensureCommentsArtifactsInZip(zip: unknown, commentsXml: string): Promise<unknown> | unknown;
export function createDynamicNumberingIdState(numberingXml?: string): unknown;

export function parseOoxml(ooxml: string): Document;
export function serializeOoxml(node: Node): string;
export function sanitizeAiResponse(text: string): string;

export class ReconciliationPipeline {
  constructor(options?: Record<string, unknown>);
  execute(oxml: string, modifiedText: string, options?: Record<string, unknown>): Promise<unknown>;
}

export class NumberingService {
  constructor(...args: unknown[]);
}
