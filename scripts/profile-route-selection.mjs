import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider } from '../adapters/xml-adapter.js';
import { applyRedlineToOxml } from '../engine/oxml-engine.js';
import { createRouteFrequencyCollector } from '../engine/route-selection.js';

configureXmlProvider({ DOMParser, XMLSerializer });
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const paragraph = text => `<w:p xmlns:w="${W}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cases = [
    [paragraph('Alpha'), 'Alpha', 'Beta'],
    [paragraph('Alpha'), 'Alpha', '1. One\n2. Two'],
    [paragraph('Alpha'), 'Alpha', '| A |\n| --- |\n| B |'],
    [`<w:document xmlns:w="${W}"><w:body>${paragraph('1. One')}${paragraph('2. Two')}</w:body></w:document>`, '1. One\n2. Two', '1. One changed\n2. Two']
];
const collector = createRouteFrequencyCollector();
for (const [oxml, original, modified] of cases) {
    await applyRedlineToOxml(oxml, original, modified, { author:'Route profiler', _routeInstrumentation: collector });
}
process.stdout.write(`${JSON.stringify({ cases: cases.length, routes: collector.snapshot() }, null, 2)}\n`);
