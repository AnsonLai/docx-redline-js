/**
 * XML adapter for parser/serializer portability.
 *
 * Default behavior uses browser-provided DOMParser/XMLSerializer.
 * Consumers can override these constructors for non-browser runtimes.
 */

import { warn as logWarning, error as logError } from './logger.js';

let _DOMParser = globalThis.DOMParser;
let _XMLSerializer = globalThis.XMLSerializer;

/**
 * Configures XML provider constructors.
 *
 * @param {Object} [options={}] - Provider overrides
 * @param {typeof DOMParser} [options.DOMParser] - DOMParser constructor
 * @param {typeof XMLSerializer} [options.XMLSerializer] - XMLSerializer constructor
 */
export function configureXmlProvider(options = {}) {
    if (options.DOMParser) _DOMParser = options.DOMParser;
    if (options.XMLSerializer) _XMLSerializer = options.XMLSerializer;
}

/**
 * Creates a parser instance.
 *
 * @returns {DOMParser}
 */
export function createParser(options = {}) {
    if (!_DOMParser && globalThis.DOMParser) {
        _DOMParser = globalThis.DOMParser;
    }
    if (!_DOMParser) {
        throw new Error('DOMParser is not configured. Call configureXmlProvider({ DOMParser, XMLSerializer }) first.');
    }
    return new _DOMParser(options);
}

/**
 * Creates a serializer instance.
 *
 * @returns {XMLSerializer}
 */
export function createSerializer() {
    if (!_XMLSerializer && globalThis.XMLSerializer) {
        _XMLSerializer = globalThis.XMLSerializer;
    }
    if (!_XMLSerializer) {
        throw new Error('XMLSerializer is not configured. Call configureXmlProvider({ DOMParser, XMLSerializer }) first.');
    }
    return new _XMLSerializer();
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
