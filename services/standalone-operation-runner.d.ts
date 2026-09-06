import type { ExistingRevisionsPolicy, RedlineError, RedlineStatus } from '../index.js';

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
  revisionView?: 'accepted' | 'rejected';
}

export interface InsertionAffinity {
  formatting?: 'left' | 'right' | 'none';
  hyperlink?: 'inside' | 'outside' | 'preserve';
  revision?: 'coalesce_same_author' | 'separate';
  bookmark?: 'inside' | 'outside';
  comment?: 'inside' | 'outside';
}

export interface DocumentOperationBase {
  target?: string | ParagraphTargetDescriptor;
  targetRef?: number | string | null;
  author?: string;
  generateRedlines?: boolean;
  existingRevisions?: ExistingRevisionsPolicy;
  pairReplacements?: boolean;
  insertionAffinity?: InsertionAffinity;
}

export interface RedlineDocumentOperation extends DocumentOperationBase {
  type: 'redline' | 'replace' | 'format' | 'list-change' | 'table-reconciliation' | 'insert';
  modified: string;
  structuredContent?: boolean;
  targetEnd?: ParagraphTargetDescriptor;
  targetEndRef?: number | string | null;
}

export interface DeleteDocumentOperation extends DocumentOperationBase {
  type: 'delete';
  modified?: '';
}

export interface CommentDocumentOperation extends DocumentOperationBase {
  type: 'comment';
  textToComment?: string;
  commentContent: string;
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

export interface ParagraphFormatProperties {
  alignment?: 'left' | 'center' | 'right' | 'both';
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  style?: string | null;
}

export interface HighlightDocumentOperation extends DocumentOperationBase {
  type: 'highlight';
  textToHighlight: string;
  color?: string;
}

export interface CharacterFormatDocumentOperation extends DocumentOperationBase {
  type: 'format' | 'character-format';
  textToFormat: string;
  properties: CharacterFormatProperties;
  formattingRevisionPolicy?: 'always' | 'coalesce-own-insertion';
}

export interface ParagraphFormatDocumentOperation extends DocumentOperationBase {
  type: 'paragraph-format';
  properties: ParagraphFormatProperties;
}

export type DocumentOperation =
  | RedlineDocumentOperation
  | DeleteDocumentOperation
  | CommentDocumentOperation
  | HighlightDocumentOperation
  | CharacterFormatDocumentOperation
  | ParagraphFormatDocumentOperation;

export interface ResolvedDocumentTarget {
  index: number;
  paragraphId: string | null;
  text: string;
  fingerprint?: string;
  inTable?: boolean;
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
  pairReplacements?: boolean;
  insertionAffinity?: InsertionAffinity;
  onInfo?: (message: string) => void;
  onWarn?: (message: string) => void;
  [key: string]: unknown;
}

export interface DocumentOperationResult {
  documentXml: string;
  hasChanges: boolean;
  numberingXml?: string | null;
  commentsXml?: string | null;
  warnings?: string[];
  status?: RedlineStatus;
  error?: RedlineError;
  operationType: 'redline' | 'comment' | 'highlight' | 'format' | 'paragraph-format';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
}

export interface BatchOperationItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'highlight' | 'format' | 'paragraph-format';
  status: 'applied' | 'no_change' | 'error';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
  warnings?: string[];
  error?: RedlineError;
}

export interface DocumentOperationBatchResult {
  documentXml: string;
  hasChanges: boolean;
  commentsXml: string | null;
  numberingXmlParts: string[];
  results: BatchOperationItemResult[];
  executionOrder: number[];
  /** Authors attached to committed changes. Empty when an atomic batch rolls back. */
  authorsUsed: string[];
  rolledBack?: boolean;
  status?: RedlineStatus;
  error?: RedlineError;
  warnings?: string[];
}

export interface OperationPreflightItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'highlight' | 'format' | 'paragraph-format';
  status: 'ready' | 'error';
  authorUsed: string;
  resolvedBy?: string;
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
  error?: RedlineError & { candidates?: ResolvedDocumentTarget[] };
}

export interface OperationConflict {
  code: 'OVERLAPPING_TEXT_EDITS' | 'REVISION_ORDER_CONFLICT' | string;
  message: string;
  operationIndexes: number[];
  target: ResolvedDocumentTarget;
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
