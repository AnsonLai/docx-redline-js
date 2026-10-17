import type { RevisionToken } from '../index.js';

/** Canonical parts read from a Word body Flat OPC package. */
export interface WordFlatOpcParts {
  documentXml: string;
  commentsXml: string | null;
  commentsExtendedXml: string | null;
  commentsIdsXml: string | null;
  commentsExtensibleXml: string | null;
  numberingXml: string | null;
  stylesXml: string | null;
  additionalParts: Record<string, string>;
}

/** Inspect `body.getOoxml()` output (or a bare document/fragment) in the accepted view. */
export function inspectWordFlatOpc(flatOpc: string, options?: Record<string, unknown>): { parts: WordFlatOpcParts; inspection: any };

export interface WordFlatOpcBatchOptions {
  author?: string;
  generateRedlines?: boolean;
  existingRevisions?: string;
  sanitizeInput?: boolean;
  strictTargets?: boolean;
  /** `true` binds the batch to the source's own revision token. */
  expectedRevision?: boolean | RevisionToken | string;
  runner?: (...args: any[]) => Promise<any>;
  onInfo?: (...args: any[]) => void;
  onWarn?: (...args: any[]) => void;
}

export interface WordFlatOpcBatchResult {
  status: 'ready' | 'noop' | 'refused' | 'error';
  source: WordFlatOpcParts & { paragraphs: any[] };
  operations: any[];
  result: any;
  /** The single package to write with `insertOoxml(payload, 'Replace')`; present when status is 'ready'. */
  insertionPayload?: string;
  error?: Error;
}

/** Run one atomic canonical batch on a Word body Flat OPC source and build the insertion package. */
export function prepareWordFlatOpcBatch(
  flatOpc: string,
  operations: any[] | ((source: WordFlatOpcBatchResult['source']) => any[] | Promise<any[]>),
  options?: WordFlatOpcBatchOptions
): Promise<WordFlatOpcBatchResult>;

/** One-based index of Word's export-only empty trailing paragraph, or null unless the host count certifies it. */
export function detectWordExportTailIndex(parts: Partial<WordFlatOpcParts>, options?: { hostParagraphCount?: number }): number | null;
/** Revision token that ignores Word rsids and a certified export tail. */
export function computeWordSourceRevisionToken(parts: Partial<WordFlatOpcParts>, options?: { hostParagraphCount?: number }): Promise<RevisionToken>;
/** Token for comparing a restored body, ignoring ids Word regenerates on Replace. */
export function computeWordRestorationToken(parts: Partial<WordFlatOpcParts>, options?: { hostParagraphCount?: number }): Promise<RevisionToken & { commentIdentity: any[] }>;
