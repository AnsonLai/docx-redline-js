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
  revisionView?: 'accepted' | 'rejected';
  captureRef?: string;
  select?: string;
}

export interface InsertionAffinity {
  formatting?: 'left' | 'right' | 'none';
  hyperlink?: 'inside' | 'outside' | 'preserve';
  revision?: 'coalesce_same_author' | 'separate';
  bookmark?: 'inside' | 'outside';
  comment?: 'inside' | 'outside';
}

export interface DocumentOperationBase {
  operationId?: string;
  captureKey?: string;
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

export interface CommentReplyDocumentOperation {
  type: 'comment_reply';
  operationId?: string;
  parentCommentId: number | string;
  commentContent: string;
  author?: string;
  date?: string;
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
  | CommentReplyDocumentOperation
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

export interface DocumentOperationResult {
  documentXml: string;
  hasChanges: boolean;
  numberingXml?: string | null;
  commentsXml?: string | null;
  commentsExtendedXml?: string | null;
  commentsXmlMode?: 'merge' | 'replace';
  commentsExtendedXmlMode?: 'merge' | 'replace';
  warnings?: string[];
  status?: RedlineStatus;
  error?: RedlineError;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'highlight' | 'format' | 'paragraph-format';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
  receipt?: MutationReceipt;
}

export interface BatchOperationItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'highlight' | 'format' | 'paragraph-format';
  status: 'applied' | 'no_change' | 'error';
  authorUsed: string;
  resolvedBy?: string;
  resolvedTarget?: ResolvedDocumentTarget;
  resolvedAnchor?: ResolvedCommentAnchor;
  warnings?: string[];
  error?: RedlineError;
  receipt?: MutationReceipt;
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
}

export interface OperationPreflightItemResult {
  index: number;
  type: string;
  operationType: 'redline' | 'comment' | 'comment_reply' | 'highlight' | 'format' | 'paragraph-format';
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
