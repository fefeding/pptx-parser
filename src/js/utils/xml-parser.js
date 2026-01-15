/**
 * XML Parser - Uses browser native APIs when available
 */

let _order = 1;

/**
 * Main parsing function
 * @param {string} xmlString - XML string to parse
 * @param {Object} options - Parsing options
 * @returns {Object} Parsed XML structure
 */
export function parserXml(xmlString, options = {}) {
    // Reset order for each parse
    if (!options.attrValue) {
        _order = 1;
    }

    if (options.attrValue !== undefined) {
        return parseByAttributeValue(xmlString, options);
    }

    // Try to use DOMParser (browser) or XMLDocument (Node.js)
    const doc = parseXmlString(xmlString);

    if (!doc) {
        // Fallback to simple character-based parsing
        return parseWithFallback(xmlString, options);
    }

    const result = convertDomToArray(doc, options.parseNode);
    result.pos = xmlString.length;

    if (options.filter) {
        return filter(result, options.filter);
    }

    if (options.simplify) {
        return simplify(result);
    }

    return result;
}

/**
 * Parse XML string using native APIs
 * @param {string} xmlString - XML string to parse
 * @returns {Document|Object|null} Parsed DOM document, object with xmlDeclaration, or null
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

    // Node.js: use native XML (if available) or fast-xml-parser
    try {
        if (typeof require === 'function') {
            // Try to use fast-xml-parser
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
 * Convert DOM structure to array format (for backward compatibility)
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

    const result = [];

    // Check if there's an XML declaration attached
    if (doc._xmlDeclaration) {
        const xmlNode = createXmlDeclarationNode(doc._xmlDeclaration);
        if (xmlNode) {
            result.push(xmlNode);
        }
    }

    const rootElement = doc.documentElement || doc.childNodes[0];

    if (parseNodeOnly && rootElement) {
        return convertDomNodeToArray(rootElement);
    }

    for (let i = 0; i < doc.childNodes.length; i++) {
        const converted = convertDomNodeToArray(doc.childNodes[i]);
        if (converted) {
            result.push(converted);
        }
    }

    return result;
}

/**
 * Create XML declaration node from attributes string
 * @param {string} xmlDecl - XML declaration attributes string
 * @returns {Object} XML declaration node
 */
function createXmlDeclarationNode(xmlDecl) {
    // Parse attributes like 'version="1.0" encoding="UTF-8" standalone="yes"'
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
 * Convert DOM element to array format
 * @param {Node} node - DOM node
 * @returns {Object} Converted node object
 */
function convertDomNodeToArray(node) {
    if (!node) return null;

    // Handle text nodes
    if (node.nodeType === Node.TEXT_NODE) {
        return node.textContent;
    }

    // Handle comment nodes
    if (node.nodeType === Node.COMMENT_NODE) {
        return null;
    }

    const result = {
        tagName: node.tagName || node.nodeName
    };

    // Parse attributes
    if (node.attributes && node.attributes.length > 0) {
        result.attributes = {};
        for (let i = 0; i < node.attributes.length; i++) {
            const attr = node.attributes[i];
            result.attributes[attr.name] = attr.value;
        }
    }

    // Parse children
    const children = [];
    for (let i = 0; i < node.childNodes.length; i++) {
        const child = convertDomNodeToArray(node.childNodes[i]);
        if (child !== null) {
            children.push(child);
        }
    }

    if (children.length > 0) {
        result.children = children;
    }

    return result;
}

/**
 * Convert fast-xml-parser object to array format
 * @param {Object} obj - fast-xml-parser object
 * @returns {Array} Array format
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
 * @param {Object} item - fast-xml-parser item
 * @param {string} explicitTagName - Explicit tag name (for ?xml nodes)
 * @returns {Object} Node object
 */
function convertFastXmlObjToNode(item, explicitTagName = null) {
    if (typeof item === 'string') {
        return item;
    }

    if (!item || typeof item !== 'object') {
        return null;
    }

    // Find the actual tag name (key that's not '@' or '#')
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

    const result = {
        tagName: tagName
    };

    if (Object.keys(attrs).length > 0) {
        result.attributes = attrs;
    }

    if (children.length > 0) {
        result.children = children;
    }

    return result;
}

/**
 * Fallback character-based parser (for environments without native APIs)
 * @param {string} xmlString - XML string to parse
 * @param {Object} options - Parsing options
 * @returns {Object} Parsed structure
 */
function parseWithFallback(xmlString, options) {
    // Simple character-based parser (original logic)
    // This is a simplified version - in practice, this should rarely be used
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
                // Parse children recursively would go here
                // For simplicity, skip this part in fallback
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
 * Simplify parsed XML structure
 * @param {Array} children - Parsed children array
 * @returns {Object} Simplified object structure
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
            if(typeof simplifiedChild === 'object') {
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

    return result;
}

/**
 * Filter nodes based on a predicate function
 * @param {Array} nodes - Nodes to filter
 * @param {Function} predicate - Filter function
 * @returns {Array} Filtered nodes
 */
export function filter(nodes, predicate) {
    const result = [];

    nodes.forEach((node) => {
        if (typeof node === 'object' && predicate(node)) {
            result.push(node);
        }

        if (node.children) {
            const filteredChildren = filter(node.children, predicate);
            result = result.concat(filteredChildren);
        }
    });

    return result;
}

/**
 * Convert parsed XML back to string
 * @param {Array} nodes - Parsed XML nodes
 * @returns {string} XML string
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
 * @param {*} node - Parsed XML node or array
 * @returns {string} Text content
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
 * @param {string} xmlString - XML string
 * @param {string} idValue - ID value to find
 * @param {boolean} simplifyResult - Whether to simplify result
 * @returns {*} Found element
 */
export function getElementById(xmlString, idValue, simplifyResult) {
    const result = parserXml(xmlString, { attrValue: idValue, simplify: simplifyResult });
    return simplifyResult ? result : result[0];
}

/**
 * Get elements by class name
 * @param {string} xmlString - XML string
 * @param {string} className - Class name to find
 * @param {boolean} simplifyResult - Whether to simplify result
 * @returns {*} Found elements
 */
export function getElementsByClassName(xmlString, className, simplifyResult) {
    const pattern = `[a-zA-Z0-9-s ]*${className}[a-zA-Z0-9-s ]*`;
    return parserXml(xmlString, { attrName: 'class', attrValue: pattern, simplify: simplifyResult });
}

/**
 * Parse XML stream (Node.js only)
 * @param {string|ReadStream} source - Source file path or ReadStream
 * @param {number|Function} startOrCallback - Start position or callback
 * @param {Function} callback - Optional callback
 * @returns {ReadStream} The stream
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
        const result = parserXml(buffer);
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
            const result = parserXml(buffer);
            source.emit('xml', result);
        }
    });

    return source;
}

// Export as default
const parserXmlDefault = {
    parse: parserXml,
    simplify,
    filter,
    stringify,
    toContentString,
    getElementById,
    getElementsByClassName,
    parseStream
};

// Add static methods for backward compatibility
parserXml.simplify = simplify;
parserXml.filter = filter;
parserXml.stringify = stringify;
parserXml.toContentString = toContentString;
parserXml.getElementById = getElementById;
parserXml.getElementsByClassName = getElementsByClassName;
parserXml.parseStream = parseStream;

export default parserXmlDefault;
