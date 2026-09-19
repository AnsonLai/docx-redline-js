/**
 * XML adapter for parser/serializer portability.
 *
 * Transparently uses browser-provided DOMParser/XMLSerializer when available in
 * standard web runtimes, automatically falling back to pure-JS @xmldom/xmldom in
 * Node.js and sandbox environments. Consumers can also explicitly override these
 * constructors via configureXmlProvider.
 */

import { DOMParser as XmlDomParser, XMLSerializer as XmlDomSerializer } from '@xmldom/xmldom';
import { warn as logWarning, error as logError } from './logger.js';

let _DOMParser = null;
let _XMLSerializer = null;

/**
 * Configures XML provider constructors.
 *
 * @param {Object} [options={}] - Provider overrides
 * @param {typeof DOMParser} [options.DOMParser] - DOMParser constructor
 * @param {typeof XMLSerializer} [options.XMLSerializer] - XMLSerializer constructor
 */
export function configureXmlProvider(options = {}) {
    if ('DOMParser' in options) _DOMParser = options.DOMParser;
    if ('XMLSerializer' in options) _XMLSerializer = options.XMLSerializer;
}

/**
 * Resolves the active DOMParser constructor.
 *
 * Priority order:
 * 1. Explicitly configured constructor (via configureXmlProvider)
 * 2. Host environment native DOMParser (globalThis.DOMParser)
 * 3. Pure-JS @xmldom/xmldom fallback
 *
 * @returns {typeof DOMParser}
 */
export function resolveDomParserConstructor() {
    return _DOMParser || globalThis.DOMParser || XmlDomParser;
}

/**
 * Resolves the active XMLSerializer constructor.
 *
 * Priority order:
 * 1. Explicitly configured constructor (via configureXmlProvider)
 * 2. Host environment native XMLSerializer (globalThis.XMLSerializer)
 * 3. Pure-JS @xmldom/xmldom fallback
 *
 * @returns {typeof XMLSerializer}
 */
export function resolveXmlSerializerConstructor() {
    return _XMLSerializer || globalThis.XMLSerializer || XmlDomSerializer;
}

/**
 * Creates a parser instance.
 *
 * @param {Object} [options={}] - Parser options (e.g. { onError } for xmldom)
 * @returns {DOMParser}
 */
export function createParser(options = {}) {
    const ParserCtor = resolveDomParserConstructor();
    if (!ParserCtor) {
        throw new Error('DOMParser is not configured and no fallback XML parser is available.');
    }
    return new ParserCtor(options);
}

/**
 * Creates a serializer instance.
 *
 * @returns {XMLSerializer}
 */
export function createSerializer() {
    const SerializerCtor = resolveXmlSerializerConstructor();
    if (!SerializerCtor) {
        throw new Error('XMLSerializer is not configured and no fallback XML serializer is available.');
    }
    return new SerializerCtor();
}

/**
 * Parses XML text into a DOM document.
 *
 * @param {string} xmlString - XML string
 * @param {string} [contentType='text/xml'] - MIME type
 * @returns {Document}
 */
export function parseXml(xmlString, contentType = 'text/xml') {
    const result = parseOoxmlSafe(xmlString, contentType);
    if (result.error) {
        const parseError = new Error(result.error.message);
        parseError.code = result.error.code;
        throw parseError;
    }
    return result.doc;
}

function browserParseError(doc) {
    if (!doc?.documentElement) return null;
    if (String(doc.documentElement.localName || doc.documentElement.nodeName).toLowerCase() === 'parsererror') {
        return doc.documentElement;
    }
    return doc.getElementsByTagName?.('parsererror')?.[0] || null;
}

/**
 * Parses OOXML without allowing parser/provider exceptions to escape.
 *
 * `@xmldom/xmldom` reports recoverable diagnostics through `onError` and
 * throws for fatal errors. Browser DOMParser implementations instead return a
 * `<parsererror>` document. This helper normalizes both behaviors.
 *
 * @param {unknown} xmlString
 * @param {string} [contentType='application/xml']
 * @returns {{ doc: Document|null, error: {code:'PARSE_ERROR', message:string}|null, warnings: string[] }}
 */
export function parseOoxmlSafe(xmlString, contentType = 'application/xml') {
    const warnings = [];
    if (typeof xmlString !== 'string' || xmlString.trim() === '') {
        return {
            doc: null,
            error: { code: 'PARSE_ERROR', message: 'Input is not a non-empty XML string.' },
            warnings
        };
    }

    const onError = (level, message) => {
        const diagnostic = String(message || 'XML parser diagnostic.');
        if (level === 'fatalError') {
            logError('[XmlAdapter] XML fatal parse error:', diagnostic);
        } else {
            warnings.push(diagnostic);
            logWarning(`[XmlAdapter] XML ${level || 'warning'}:`, diagnostic);
        }
    };

    try {
        const parser = createParser({ onError });
        const doc = parser.parseFromString(xmlString, contentType);
        const parseError = browserParseError(doc);
        if (!doc?.documentElement || parseError) {
            const message = parseError?.textContent || 'Could not parse XML input.';
            logError('[XmlAdapter] XML parse error:', message);
            return { doc: null, error: { code: 'PARSE_ERROR', message }, warnings };
        }
        return { doc, error: null, warnings };
    } catch (caught) {
        const message = caught?.message || String(caught || 'Could not parse XML input.');
        logError('[XmlAdapter] XML parse error:', message);
        return { doc: null, error: { code: 'PARSE_ERROR', message }, warnings };
    }
}

/**
 * Serializes a node to XML text.
 *
 * @param {Node} node - Node to serialize
 * @returns {string}
 */
export function serializeXml(node) {
    const serializer = createSerializer();
    return serializer.serializeToString(node);
}
