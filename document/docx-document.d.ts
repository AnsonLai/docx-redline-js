import type { DocumentOperation, DocumentOperationBatchResult, OperationPreflightResult, StandaloneRunnerOptions } from '../services/standalone-operation-runner.js';
import type { DocumentInspectionOptions, DocumentInspectionResult, RevisionToken } from '../index.js';

export interface DocxApplyOptions extends StandaloneRunnerOptions {
  author?: string;
  validate?: boolean;
}

export interface DocxApplyResult extends DocumentOperationBatchResult {
  written: boolean;
  uint8Array: Uint8Array;
  buffer: Uint8Array;
  inspection?: DocumentInspectionResult;
  artifactsChanged?: string[];
  validation?: { originalIssues: unknown[]; generatedIssues: unknown[] };
  issues?: unknown[];
  toUint8Array(): Uint8Array;
  toBuffer(): Uint8Array;
}

export class DocxDocument {
  constructor(input: Uint8Array | ArrayBuffer | string);
  readonly originalBytes: Uint8Array;
  readonly originalBuffer: Uint8Array;
  readonly entries: Map<string, Uint8Array>;
  inspect(options?: DocumentInspectionOptions): DocumentInspectionResult;
  getRevisionToken(): RevisionToken;
  readonly revisionToken: RevisionToken;
  preflight(operations: DocumentOperation[], author?: string, options?: StandaloneRunnerOptions): OperationPreflightResult;
  applyOperations(operations: DocumentOperation[], options?: DocxApplyOptions): Promise<DocxApplyResult>;
  resolveRevisions(action: 'accept' | 'reject', options?: { author?: string; allAuthors?: boolean; validate?: boolean }): Promise<DocxApplyResult>;
  deleteComments(options?: { author?: string; allAuthors?: boolean; validate?: boolean }): Promise<DocxApplyResult>;
  toUint8Array(): Uint8Array;
  toBuffer(): Uint8Array;
}

export function computePackageRevisionToken(input: unknown): RevisionToken;

export function openDocx(input: Uint8Array | ArrayBuffer | string): DocxDocument;
