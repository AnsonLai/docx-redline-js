import {
  applyRedlineToOxml,
  configureXmlProvider,
  ingestWordOoxmlToPlainTextResult,
  parseOoxmlSafe,
  reconcileMarkdownTableOoxml,
  validateRedlineOoxml,
  type RedlineOptions,
  type RedlineResult
} from '../../index.js';
import {
  applyOperationsToDocumentXml,
  preflightOperations,
  type DocumentOperation
} from '@ansonlai/docx-redline-js/standalone-runner';
import { openDocx } from '@ansonlai/docx-redline-js/node';

configureXmlProvider({ DOMParser, XMLSerializer });

const options: RedlineOptions = {
  author: 'Type Test',
  generateRedlines: true,
  existingRevisions: 'reject-input',
  sanitizeInput: false
};

async function exercisePublicTypes(oxml: string): Promise<RedlineResult> {
  const parsed = parseOoxmlSafe(oxml);
  if (parsed.error) throw new Error(parsed.error.message);

  const ingestion = ingestWordOoxmlToPlainTextResult(oxml);
  const result = await applyRedlineToOxml(oxml, ingestion.text, 'Replacement', options);
  const tableResult = await reconcileMarkdownTableOoxml(
    oxml,
    ingestion.text,
    '| A |\n| --- |\n| B |',
    options
  );
  const validation = validateRedlineOoxml(tableResult.oxml);
  if (!validation.valid) throw new Error(validation.issues[0]?.message);
  return result;
}

void exercisePublicTypes('<w:p/>');

const operations: DocumentOperation[] = [
  {
    type: 'replace',
    target: { exactText: 'Old text', index: 1 },
    modified: 'New text',
    author: 'Editor'
  },
  {
    type: 'comment',
    target: 'New text',
    commentContent: 'Review this paragraph',
    author: 'Reviewer'
  }
];

void applyOperationsToDocumentXml('<w:document/>', operations, 'Fallback Author');
void preflightOperations('<w:document/>', operations, 'Fallback Author');
const docx = openDocx(new Uint8Array());
void docx.inspect({ skipEmpty: true });
