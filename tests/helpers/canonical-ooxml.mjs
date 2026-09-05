import { createHash } from 'node:crypto';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { configureXmlProvider, parseOoxmlSafe } from '../../adapters/xml-adapter.js';

configureXmlProvider({ DOMParser, XMLSerializer });

export const CANONICAL_PREFIX_MAP = new Map([
    ['http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w'],
    ['http://schemas.microsoft.com/office/word/2010/wordml', 'w14'],
    ['http://schemas.microsoft.com/office/word/2012/wordml', 'w15'],
    ['http://schemas.microsoft.com/office/word/2016/wordml16', 'w16'],
    ['http://schemas.microsoft.com/office/word/2016/wordml16cex', 'w16cex'],
    ['http://schemas.microsoft.com/office/word/2016/wordml16cid', 'w16cid'],
    ['http://schemas.microsoft.com/office/word/2016/wordml16se', 'w16se'],
    ['http://schemas.openxmlformats.org/markup-compatibility/2006', 'mc'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r'],
    ['http://schemas.openxmlformats.org/officeDocument/2006/math', 'm'],
    ['http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'wp'],
    ['http://schemas.openxmlformats.org/drawingml/2006/main', 'a'],
    ['http://schemas.openxmlformats.org/drawingml/2006/picture', 'pic'],
    ['http://schemas.openxmlformats.org/package/2006/relationships', 'rel'],
    ['http://schemas.openxmlformats.org/package/2006/content-types', 'ct'],
    ['http://www.w3.org/XML/1998/namespace', 'xml'],
    ['http://www.w3.org/2000/xmlns/', 'xmlns']
]);

const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

function escapeText(text) {
    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\r/g, '&#xD;');
}

function escapeAttributeValue(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/"/g, '&quot;')
        .replace(/\t/g, '&#x9;')
        .replace(/\n/g, '&#xA;')
        .replace(/\r/g, '&#xD;');
}

/**
 * Builds an in-scope namespace prefix resolver by walking ancestors.
 */
function getInScopeNamespaces(node) {
    const namespaces = new Map();
    let cursor = node;
    while (cursor) {
        if (cursor.nodeType === 1 && cursor.attributes) {
            for (let i = 0; i < cursor.attributes.length; i++) {
                const attr = cursor.attributes[i];
                if (attr.namespaceURI === XMLNS_NS || attr.name === 'xmlns' || attr.name.startsWith('xmlns:')) {
                    const prefix = attr.localName === 'xmlns' || attr.name === 'xmlns' ? '' : attr.localName;
                    if (!namespaces.has(prefix)) {
                        namespaces.set(prefix, attr.value);
                    }
                }
            }
        }
        cursor = cursor.parentNode;
    }
    return namespaces;
}

/**
 * Collects all namespace URIs actually utilized in an element subtree.
 */
function collectUsedNamespaceUris(node, usedUris) {
    if (node.nodeType === 1) {
        if (node.namespaceURI && node.namespaceURI !== XMLNS_NS) {
            usedUris.add(node.namespaceURI);
        }
        if (node.attributes) {
            for (let i = 0; i < node.attributes.length; i++) {
                const attr = node.attributes[i];
                if (attr.namespaceURI && attr.namespaceURI !== XMLNS_NS && attr.namespaceURI !== XML_NS) {
                    usedUris.add(attr.namespaceURI);
                }
                // If mc:Ignorable, collect the URIs referenced in its value
                if (attr.namespaceURI === MC_NS && attr.localName === 'Ignorable') {
                    const inScope = getInScopeNamespaces(node);
                    const prefixes = String(attr.value || '').trim().split(/\s+/).filter(Boolean);
                    for (const p of prefixes) {
                        const uri = inScope.get(p);
                        if (uri) usedUris.add(uri);
                    }
                }
            }
        }
        for (let child = node.firstChild; child; child = child.nextSibling) {
            collectUsedNamespaceUris(child, usedUris);
        }
    }
}

/**
 * Builds a deterministic mapping of namespace URI to canonical prefix for a subtree.
 */
function buildSubtreePrefixMap(rootNode) {
    const usedUris = new Set();
    collectUsedNamespaceUris(rootNode, usedUris);

    const prefixMap = new Map();
    const dynamicUris = [];

    for (const uri of Array.from(usedUris).sort()) {
        if (CANONICAL_PREFIX_MAP.has(uri)) {
            prefixMap.set(uri, CANONICAL_PREFIX_MAP.get(uri));
        } else {
            dynamicUris.push(uri);
        }
    }

    dynamicUris.sort();
    let dynamicIndex = 1;
    for (const uri of dynamicUris) {
        prefixMap.set(uri, `ns${dynamicIndex++}`);
    }

    return prefixMap;
}

/**
 * Canonicalizes a QName-list valued attribute such as mc:Ignorable.
 */
function canonicalizeQNameList(value, node, prefixMap) {
    const inScope = getInScopeNamespaces(node);
    const prefixes = String(value || '').trim().split(/\s+/).filter(Boolean);
    const canonicalPrefixes = [];

    for (const p of prefixes) {
        const uri = inScope.get(p);
        if (uri && prefixMap.has(uri)) {
            canonicalPrefixes.push(prefixMap.get(uri));
        } else if (p) {
            canonicalPrefixes.push(p);
        }
    }

    canonicalPrefixes.sort();
    return canonicalPrefixes.join(' ');
}

/**
 * Serializes a DOM element to canonical XML string.
 */
function serializeCanonicalElement(node, prefixMap, isRoot) {
    const uri = node.namespaceURI;
    const prefix = uri ? prefixMap.get(uri) || '' : '';
    const qName = prefix ? `${prefix}:${node.localName}` : node.localName;

    let out = `<${qName}`;

    // Root element emits all used namespace declarations sorted by prefix
    if (isRoot) {
        const sortedDecls = Array.from(prefixMap.entries())
            .filter(([nsUri]) => nsUri !== XML_NS && nsUri !== XMLNS_NS)
            .sort(([, prefA], [, prefB]) => prefA.localeCompare(prefB));

        for (const [nsUri, pref] of sortedDecls) {
            out += ` xmlns:${pref}="${escapeAttributeValue(nsUri)}"`;
        }
    }

    // Process attributes (excluding xmlns declarations)
    const attrs = [];
    if (node.attributes) {
        for (let i = 0; i < node.attributes.length; i++) {
            const attr = node.attributes[i];
            if (attr.namespaceURI === XMLNS_NS || attr.name === 'xmlns' || attr.name.startsWith('xmlns:')) {
                continue;
            }

            let attrPrefix = '';
            let attrVal = attr.value;

            if (attr.namespaceURI === XML_NS) {
                attrPrefix = 'xml';
            } else if (attr.namespaceURI === MC_NS && attr.localName === 'Ignorable') {
                attrPrefix = prefixMap.get(MC_NS) || 'mc';
                attrVal = canonicalizeQNameList(attr.value, node, prefixMap);
            } else if (attr.namespaceURI) {
                attrPrefix = prefixMap.get(attr.namespaceURI) || '';
            }

            const attrQName = attrPrefix ? `${attrPrefix}:${attr.localName}` : attr.localName;
            attrs.push({
                nsUri: attr.namespaceURI || '',
                localName: attr.localName,
                qName: attrQName,
                value: attrVal
            });
        }
    }

    // Sort attributes by namespaceURI, then by localName
    attrs.sort((a, b) => {
        if (a.nsUri !== b.nsUri) {
            // unprefixed/empty namespace first
            if (!a.nsUri) return -1;
            if (!b.nsUri) return 1;
            return a.nsUri.localeCompare(b.nsUri);
        }
        return a.localName.localeCompare(b.localName);
    });

    for (const attr of attrs) {
        out += ` ${attr.qName}="${escapeAttributeValue(attr.value)}"`;
    }

    // Children serialization
    const children = Array.from(node.childNodes || []).filter(child =>
        child.nodeType === 1 || child.nodeType === 3 || child.nodeType === 4
    );

    if (children.length === 0) {
        out += `></${qName}>`;
        return out;
    }

    out += '>';
    for (const child of children) {
        if (child.nodeType === 1) {
            out += serializeCanonicalElement(child, prefixMap, false);
        } else if (child.nodeType === 3 || child.nodeType === 4) {
            out += escapeText(child.nodeValue || '');
        }
    }
    out += `</${qName}>`;

    return out;
}

/**
 * Computes canonical XML representation, SHA-256 digest, and raw bytes for a DOM node or XML string.
 *
 * @param {string|Element|Document} input
 * @returns {{ canonicalXml: string, sha256: string, bytes: Buffer }}
 */
export function canonicalizeOoxml(input) {
    let rootNode;
    if (typeof input === 'string') {
        const parsed = parseOoxmlSafe(input, 'application/xml');
        if (parsed.error || !parsed.doc) {
            throw new Error(`Failed to parse XML for canonicalization: ${parsed.error?.message || 'invalid XML'}`);
        }
        rootNode = parsed.doc.documentElement;
    } else if (input?.nodeType === 9) { // Document
        rootNode = input.documentElement;
    } else if (input?.nodeType === 1) { // Element
        rootNode = input;
    } else {
        throw new TypeError('canonicalizeOoxml requires an XML string, Document, or Element.');
    }

    if (!rootNode) {
        throw new Error('canonicalizeOoxml: No documentElement or Element found.');
    }

    const prefixMap = buildSubtreePrefixMap(rootNode);
    const canonicalXml = serializeCanonicalElement(rootNode, prefixMap, true);
    const bytes = Buffer.from(canonicalXml, 'utf8');
    const sha256 = createHash('sha256').update(bytes).digest('hex');

    return {
        canonicalXml,
        sha256,
        bytes
    };
}
