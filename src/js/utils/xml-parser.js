/**
 * XML Parser - Compatible with tXml API
 * Uses browser native DOMParser when available
 */

let _order = 1;

/**
 * Main parsing function - compatible with tXml API
 * @param {string} xmlString - XML string to parse
 * @param {Object} options - Parsing options
 * @returns {Array|Object} Parsed XML structure with pos property
 */
export function tXml(xmlString, options = {}) {
    // Reset order for each parse (used in simplify)
    if (!options.attrValue) {
        _order = 1;
    }

    // Handle attribute value filtering (getElementById, getElementsByClassName)
    if (options.attrValue !== undefined) {
        const result = parseByAttributeValue(xmlString, options);
        result.pos = xmlString.length;
        return result;
    }

    // Try to use DOMParser (browser) or fast-xml-parser (Node.js)
    const doc = parseXmlString(xmlString);

    if (!doc) {
        // Fallback to simple character-based parsing (original tXml behavior)
        const fallbackResult = parseWithFallback(xmlString, options);
        if (options.simplify) {
            const simplified = simplify(fallbackResult);
            simplified.pos = xmlString.length;
            return simplified;
        }
        fallbackResult.pos = xmlString.length;
        return fallbackResult;
    }

    // Convert DOM to array format compatible with tXml
    let result = convertDomToArray(doc, options.parseNode);

    // Apply filter if provided
    if (options.filter) {
        result = filter(result, options.filter);
    }

    // Apply simplify if requested
    if (options.simplify) {
        console.log('tXml: calling simplify with result =', result);
        result = simplify(result);
        console.log('tXml: after simplify result =', result);
        // If result is still an array (should not happen), try to simplify again
        if (Array.isArray(result)) {
            result = simplify(result);
        }
    }

    // Fix XML declaration structure to match original tXml behavior
    if (options.simplify && result && typeof result === 'object' && result['?xml'] !== undefined) {
        console.log('Fixing XML declaration structure');
        // Original tXml nests root element under ?xml
        // If result has both ?xml and Types as top-level keys, we need to move Types under ?xml
        const xmlNode = result['?xml'];
        if (result.Types !== undefined && xmlNode && typeof xmlNode === 'object' && !xmlNode.Types) {
            xmlNode.Types = result.Types;
            delete result.Types;
            // Ensure xmlNode has attrs from attributes
            if (xmlNode.attributes && !xmlNode.attrs) {
                xmlNode.attrs = xmlNode.attributes;
                delete xmlNode.attributes;
            }
        }
    }

    // Set pos property as original tXml does (position where parsing stopped)
    result.pos = xmlString.length;
    return result;
}

/**
 * Parse XML string using native APIs
 * @param {string} xmlString - XML string to parse
 * @returns {Document|Object|null} Parsed DOM document or fast-xml-parser object
 */
function parseXmlString(xmlString) {
    // Extract XML declaration if present
    let xmlDeclaration = null;
    const xmlDeclMatch = xmlString.match(/^<\?xml\s+([^>]*)\?>/);
    if (xmlDeclMatch) {
        xmlDeclaration = xmlDeclMatch[1];
    }

    // Browser: use DOMParser
    if (typeof DOMParser !== 'undefined') {
        try {
            const parser = new DOMParser();
            // Handle XML declarations by removing them temporarily
            const cleanedXml = xmlString.replace(/^<\?xml[^>]*\?>/, '');
            const doc = parser.parseFromString(cleanedXml, 'application/xml');

            // Check for parsing errors
            const errorNode = doc.querySelector('parsererror');
            if (errorNode) {
                console.warn('XML parsing error:', errorNode.textContent);
                return null;
            }

            // Attach xmlDeclaration to the document for later use
            if (xmlDeclaration) {
                doc._xmlDeclaration = xmlDeclaration;
            }

            return doc;
        } catch (e) {
            console.warn('DOMParser failed:', e);
            return null;
        }
    }

    // Node.js: use fast-xml-parser if available
    try {
        if (typeof require === 'function') {
            const { XMLParser } = require('fast-xml-parser');
            const parser = new XMLParser({
                ignoreAttributes: false,
                attributeNamePrefix: '',
                textNodeName: '#text',
                allowBooleanAttributes: true
            });
            return parser.parse(xmlString);
        }
    } catch (e) {
        // fast-xml-parser not available, will use fallback
    }

    return null;
}

/**
 * Convert DOM structure to array format compatible with tXml
 * @param {Document|Object} doc - Parsed DOM document
 * @param {boolean} parseNodeOnly - Whether to parse a single node
 * @returns {Array|Object} Converted structure
 */
function convertDomToArray(doc, parseNodeOnly) {
    // Handle fast-xml-parser output format
    if (doc && typeof doc === 'object' && !doc.nodeType) {
        // fast-xml-parser returns an object directly
        if (parseNodeOnly) {
            return [convertFastXmlObjToNode(doc)];
        }
        return convertFastXmlObjToArray(doc);
    }

    // Handle DOMParser output format
    if (!doc || !doc.childNodes) {
        return parseNodeOnly ? {} : [];
    }

    console.log('convertDomToArray: doc._xmlDeclaration =', doc._xmlDeclaration, 'doc.childNodes.length =', doc.childNodes.length);
    const result = [];

    // Check if there's an XML declaration attached
    if (doc._xmlDeclaration) {
        const xmlNode = createXmlDeclarationNode(doc._xmlDeclaration);
        if (xmlNode) {
            // Make the root element a child of the XML declaration node
            // to match original tXml behavior
            const rootElement = doc.documentElement;
            if (rootElement) {
                const convertedRoot = convertDomNodeToArray(rootElement);
                if (convertedRoot) {
                    xmlNode.children = [convertedRoot];
                }
            }
            result.push(xmlNode);
        }
    } else {
        // No XML declaration, process all child nodes as before
        const rootElement = doc.documentElement || doc.childNodes[0];

        if (parseNodeOnly && rootElement) {
            return convertDomNodeToArray(rootElement);
        }

        for (let i = 0; i < doc.childNodes.length; i++) {
            const converted = convertDomNodeToArray(doc.childNodes[i]);
            if (converted !== null) {
                result.push(converted);
            }
        }
    }

    return result;
}

/**
 * Create XML declaration node from attributes string
 * @param {string} xmlDecl - XML declaration attributes string
 * @returns {Object} XML declaration node compatible with tXml
 */
function createXmlDeclarationNode(xmlDecl) {
    const attrs = {};
    const attrRegex = /(\w+)=["']([^"']*)["']/g;
    let match;

    while ((match = attrRegex.exec(xmlDecl)) !== null) {
        attrs[match[1]] = match[2];
    }

    return {
        tagName: '?xml',
        attributes: attrs
    };
}

/**
 * Convert DOM element to tXml node format
 * @param {Node} node - DOM node
 * @returns {Object|string|null} Converted node object, text string, or null for comments
 */
function convertDomNodeToArray(node) {
    if (!node) return null;

    // Handle text nodes
    if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent;
        // Return empty string? Original tXml includes non-empty text nodes
        return text.trim().length > 0 ? text : null;
    }

    // Handle comment nodes - original tXml skips them
    if (node.nodeType === Node.COMMENT_NODE) {
        return null;
    }

    const tagName = node.tagName || node.nodeName;
    const result = { tagName };

    // Parse attributes
    if (node.attributes && node.attributes.length > 0) {
        result.attributes = {};
        for (let i = 0; i < node.attributes.length; i++) {
            const attr = node.attributes[i];
            result.attributes[attr.name] = attr.value;
        }
    }

    // Special handling for script and style tags (preserve raw content)
    if (tagName.toLowerCase() === 'script' || tagName.toLowerCase() === 'style') {
        // Get raw text content as a single child string
        const textContent = node.textContent;
        if (textContent.trim().length > 0) {
            result.children = [textContent];
        }
        return result;
    }

    // Parse children recursively
    const children = [];
    for (let i = 0; i < node.childNodes.length; i++) {
        const child = convertDomNodeToArray(node.childNodes[i]);
        if (child !== null) {
            children.push(child);
        }
    }

    // Self-closing tags detection (img, br, input, meta, link)
    const selfClosingTags = ['img', 'br', 'input', 'meta', 'link'];
    const isSelfClosing = selfClosingTags.includes(tagName.toLowerCase()) ||
                         (node.childNodes.length === 0 && !node.hasChildNodes());

    if (!isSelfClosing && children.length > 0) {
        result.children = children;
    }

    return result;
}

/**
 * Convert fast-xml-parser object to array format
 */
function convertFastXmlObjToArray(obj) {
    const result = [];

    for (const key in obj) {
        if (key === '#text') {
            continue;
        }

        const value = obj[key];

        if (Array.isArray(value)) {
            value.forEach(item => {
                const node = convertFastXmlObjToNode(item, key);
                if (node) result.push(node);
            });
        } else {
            const node = convertFastXmlObjToNode(value, key);
            if (node) result.push(node);
        }
    }

    return result;
}

/**
 * Convert fast-xml-parser object to node format
 */
function convertFastXmlObjToNode(item, explicitTagName = null) {
    if (typeof item === 'string') {
        return item;
    }

    if (!item || typeof item !== 'object') {
        return null;
    }

    let tagName = explicitTagName;
    let attrs = {};
    let children = [];

    // If explicitTagName is provided, use it (for ?xml nodes)
    if (!tagName) {
        for (const key in item) {
            if (key.startsWith('@')) {
                // Attribute
                const attrName = key.substring(1);
                attrs[attrName] = item[key];
            } else if (key === '#text') {
                // Text content
                const text = item[key];
                if (text && text.trim()) {
                    children.push(text);
                }
            } else {
                // Child element
                tagName = key;
                if (typeof item[key] === 'object') {
                    children.push(convertFastXmlObjToNode(item[key]));
                } else if (typeof item[key] === 'string') {
                    children.push(item[key]);
                }
            }
        }
    } else {
        // Use explicitTagName and collect all attributes
        for (const key in item) {
            if (key.startsWith('@')) {
                // Attribute
                const attrName = key.substring(1);
                attrs[attrName] = item[key];
            } else if (key === '#text') {
                // Text content
                const text = item[key];
                if (text && text.trim()) {
                    children.push(text);
                }
            }
        }
    }

    const result = { tagName };

    if (Object.keys(attrs).length > 0) {
        result.attributes = attrs;
    }

    if (children.length > 0) {
        result.children = children;
    }

    return result;
}

/**
 * Fallback character-based parser (original tXml logic)
 */
function parseWithFallback(xmlString, options) {
    // This is a simplified version of the original tXml parser
    // We should try to match its behavior as closely as possible
    const result = [];
    let pos = options.pos || 0;
    const selfClosingTags = ['img', 'br', 'input', 'meta', 'link', 'Default', 'Override'];
    const whitespace = '\n\t>/= ';

    function parseTagName(startPos) {
        let endPos = startPos;
        while (whitespace.indexOf(xmlString[endPos]) === -1 && xmlString[endPos]) {
            endPos++;
        }
        return xmlString.slice(startPos, endPos);
    }

    function parseQuotedValue(startPos) {
        const quoteChar = xmlString[startPos];
        const valueStart = startPos + 1;
        const valueEnd = xmlString.indexOf(quoteChar, valueStart);
        return xmlString.slice(valueStart, valueEnd);
    }

    while (pos < xmlString.length) {
        if (xmlString[pos] === '<') {
            pos++;
            const tagName = parseTagName(pos);
            pos += tagName.length;

            const node = { tagName };

            // Parse attributes
            while (xmlString[pos] && xmlString[pos] !== '>') {
                if (/[A-Za-z]/.test(xmlString[pos])) {
                    const attrName = parseTagName(pos);
                    pos += attrName.length;

                    // Skip whitespace
                    while (xmlString[pos] && /\s/.test(xmlString[pos])) pos++;

                    if (xmlString[pos] === '=') {
                        pos++;
                        while (xmlString[pos] && /\s/.test(xmlString[pos])) pos++;

                        if (xmlString[pos] === '"' || xmlString[pos] === "'") {
                            const attrValue = parseQuotedValue(pos);
                            pos += attrValue.length + 2;
                            if (!node.attributes) node.attributes = {};
                            node.attributes[attrName] = attrValue;
                        }
                    }
                } else {
                    pos++;
                }
            }

            if (xmlString[pos] === '>') {
                pos++;
            }

            // Handle self-closing tags
            const isSelfClosing = selfClosingTags.includes(tagName) ||
                               tagName.startsWith('?') ||
                               xmlString[pos - 2] === '/' ||
                               (tagName.includes('xml') && xmlString[pos - 1] === '?');

            if (!isSelfClosing && xmlString[pos]) {
                // Note: original tXml would recursively parse children here
                // For fallback simplicity, we skip child parsing
            }

            result.push(node);
        } else {
            pos++;
        }
    }

    return result;
}

/**
 * Parse nodes filtered by attribute value
 */
function parseByAttributeValue(xmlString, options) {
    const doc = parseXmlString(xmlString);
    if (!doc) {
        return [];
    }

    const optionsAttrName = options.attrName || 'id';
    const result = [];

    function searchNodes(node) {
        if (node.attributes && node.attributes[optionsAttrName] === options.attrValue) {
            return node;
        }
        if (node.children) {
            for (const child of node.children) {
                const found = searchNodes(child);
                if (found) return found;
            }
        }
        return null;
    }

    const array = convertDomToArray(doc);
    for (const item of array) {
        const found = searchNodes(item);
        if (found) result.push(found);
    }

    return result;
}

/**
 * Simplify parsed XML structure (compatible with tXml.simplify)
 */
export function simplify(children) {
    const result = {};

    if (children === undefined) return {};
    if (children.length === 1 && typeof children[0] === 'string') return children[0];

    children.forEach((child) => {
        if (typeof child === 'object' && child.tagName) {
            if (!result[child.tagName]) {
                result[child.tagName] = [];
            }

            const simplifiedChild = simplify(child.children || []);
            result[child.tagName].push(simplifiedChild);
            if (typeof simplifiedChild === 'object') {
                if (child.attributes) {
                    simplifiedChild.attrs = child.attributes;
                }

                if (simplifiedChild.attrs === undefined) {
                    simplifiedChild.attrs = { order: _order };
                } else {
                    simplifiedChild.attrs.order = _order;
                }
            }
            _order++;
        }
    });

    // Unwrap single-element arrays
    for (const key in result) {
        if (result[key].length === 1) {
            result[key] = result[key][0];
        }
    }

    // Fix XML declaration structure to match original tXml behavior
    // If we have both ?xml and other top-level keys, move other keys under ?xml
    if (result['?xml'] !== undefined && Object.keys(result).length > 1) {
        const xmlNode = result['?xml'];
        if (xmlNode && typeof xmlNode === 'object') {
            // Move all other keys under xmlNode
            for (const key in result) {
                if (key !== '?xml' && key !== 'pos') {
                    xmlNode[key] = result[key];
                    delete result[key];
                }
            }
            // Ensure attributes are moved to attrs
            if (xmlNode.attributes && !xmlNode.attrs) {
                xmlNode.attrs = xmlNode.attributes;
                delete xmlNode.attributes;
            }
        }
    }

    return result;
}

/**
 * Filter nodes based on a predicate function
 */
export function filter(nodes, predicate) {
    const result = [];

    nodes.forEach((node) => {
        if (typeof node === 'object' && predicate(node)) {
            result.push(node);
        }

        if (node.children) {
            const filteredChildren = filter(node.children, predicate);
            result.push(...filteredChildren);
        }
    });

    return result;
}

/**
 * Convert parsed XML back to string
 */
export function stringify(nodes) {
    let output = '';

    function processChildren(children) {
        if (children) {
            for (let i = 0; i < children.length; i++) {
                if (typeof children[i] === 'string') {
                    output += children[i].trim();
                } else {
                    processNode(children[i]);
                }
            }
        }
    }

    function processNode(node) {
        if (!node || !node.tagName) return;

        output += `<${node.tagName}`;

        for (const attrName in node.attributes) {
            const attrValue = node.attributes[attrName];
            if (attrValue === null) {
                output += ` ${attrName}`;
            } else if (attrValue.indexOf('"') === -1) {
                output += ` ${attrName}="${attrValue.trim()}"`;
            } else {
                output += ` ${attrName}='${attrValue.trim()}'`;
            }
        }

        output += '>';
        processChildren(node.children);
        output += `</${node.tagName}>`;
    }

    processChildren(nodes);
    return output;
}

/**
 * Convert parsed XML to content string
 */
export function toContentString(node) {
    if (Array.isArray(node)) {
        let result = '';
        node.forEach((item) => {
            result += ' ' + toContentString(item);
            result = result.trim();
        });
        return result;
    }

    if (typeof node === 'object') {
        return toContentString(node.children);
    }

    return ' ' + node;
}

/**
 * Get element by ID attribute
 */
export function getElementById(xmlString, idValue, simplifyResult) {
    const result = tXml(xmlString, { attrValue: idValue, simplify: simplifyResult });
    return simplifyResult ? result : result[0];
}

/**
 * Get elements by class name
 */
export function getElementsByClassName(xmlString, className, simplifyResult) {
    const pattern = `[a-zA-Z0-9-s ]*${className}[a-zA-Z0-9-s ]*`;
    return tXml(xmlString, { attrName: 'class', attrValue: pattern, simplify: simplifyResult });
}

/**
 * Parse XML stream (Node.js only)
 */
export function parseStream(source, startOrCallback, callback) {
    if (typeof callback === 'function') {
        // Third argument is callback
    } else if (typeof startOrCallback === 'function') {
        // Second argument is callback
        callback = startOrCallback;
        startOrCallback = 0;
    }

    if (typeof source === 'string') {
        const fs = require('fs');
        source = fs.createReadStream(source, { start: startOrCallback });
        startOrCallback = 0;
    }

    let buffer = '';

    source.on('data', (chunk) => {
        buffer += chunk;

        // Parse complete tags from buffer
        const result = tXml(buffer);
        source.emit('xml', result);

        // Keep incomplete tag in buffer
        const lastOpenTag = buffer.lastIndexOf('<');
        if (lastOpenTag > -1) {
            buffer = buffer.substring(lastOpenTag);
        } else {
            buffer = '';
        }
    });

    source.on('end', () => {
        if (buffer.trim()) {
            const result = tXml(buffer);
            source.emit('xml', result);
        }
    });

    return source;
}

// Attach static methods to tXml function (for backward compatibility)
tXml.simplify = simplify;
tXml.filter = filter;
tXml.stringify = stringify;
tXml.toContentString = toContentString;
tXml.getElementById = getElementById;
tXml.getElementsByClassName = getElementsByClassName;
tXml.parseStream = parseStream;

// Export tXml as default
export default tXml;

// Expose to global scope for browser (non-module environments)
if (typeof window !== 'undefined' && !window.tXml) {
    window.tXml = tXml;
}