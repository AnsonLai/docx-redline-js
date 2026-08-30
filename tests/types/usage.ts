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
