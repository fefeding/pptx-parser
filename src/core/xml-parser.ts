/**
 * XML Parser - A lightweight XML parser for PPTX parsing
 * Uses native DOMParser API but maintains original parseXml output format
 */

// Module-level variable to track element order across multiple parseXml calls
let elementOrder = 1;

interface XmlNode {
    tagName: string;
    attributes?: Record<string, string>;
    children?: (XmlNode | string)[];
}

interface ParseOptions {
    pos?: number;
    attrValue?: string;
    attrName?: string;
    [key: string]: any;
}

/**
 * Convert DOM node to custom structure (matching parseXml format)
 * @param {Node} domNode - DOM node to convert
 * @returns {Object|string} Converted node or text
 */
function domNodeToCustom(domNode: Node): XmlNode | string | null {
    // Text node
    if (domNode.nodeType === Node.TEXT_NODE) {
        const text = domNode.textContent;
        return text;
    }

    // CDATA section
    if (domNode.nodeType === Node.CDATA_SECTION_NODE) {
        return domNode.textContent;
    }

    // Element node
    if (domNode.nodeType === Node.ELEMENT_NODE) {
        const element = domNode as Element;
        const node: any = {};

        // Keep original case tagName
        node.tagName = element.tagName;

        // Parse attributes
        if (element.attributes && element.attributes.length > 0) {
            node.attributes = {};
            for (let i = 0; i < element.attributes.length; i++) {
                const attr = element.attributes[i];
                node.attributes[attr.name] = attr.value;
            }
        }

        // Parse children
        let hasChildren = false;
        for (let j = 0; j < element.childNodes.length; j++) {
            const childNode = domNodeToCustom(element.childNodes[j]);
            if (childNode !== null && childNode !== "") {
                if (!hasChildren) {
                    node.children = [];
                    hasChildren = true;
                }
                node.children.push(childNode as XmlNode | string);
            }
        }

        return node;
    }

    // Skip other node types (comments, processing instructions, etc.)
    return null;
}

/**
 * Main XML parsing function using native DOMParser
 * @param {string} xmlString - The XML string to parse
 * @param {Object} options - Parsing options
 * @returns {Array|Object} Parsed XML structure
 */
function parseXml(xmlString: string, options: ParseOptions = {}): any {
    options = options || {};
    const position = options.pos || 0;

    let result: any = null;

    // Create DOMParser and parse XML
    let parser: any;
    try {
        if (typeof DOMParser !== 'undefined') {
            parser = new DOMParser();
        } else if (typeof require !== 'undefined') {
            // Node.js environment fallback
            const xmldom = require('xmldom');
            parser = new xmldom.DOMParser();
        } else {
            throw new Error('No XML parser available');
        }

        const xmlDoc = parser.parseFromString(xmlString, "text/xml");

        // Check for parsing errors
        const parserError = xmlDoc.getElementsByTagName("parsererror");
        if (parserError.length > 0) {
            return [];
        }

        // If filtering by attribute value
        if (options.attrValue !== undefined) {
            options.attrName = options.attrName || "id";
            result = [];

            // Find all matching elements using XPath or iteration
            const attrName = options.attrName;
            const attrValue = options.attrValue;
            const matchingElements: any[] = [];

            if (typeof xmlDoc.evaluate !== 'undefined') {
                const xpath = "//*[@" + attrName + "='" + attrValue + "']";
                const xpathResult = xmlDoc.evaluate(xpath, xmlDoc, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                for (let k = 0; k < xpathResult.snapshotLength; k++) {
                    matchingElements.push(xpathResult.snapshotItem(k));
                }
            } else {
                // Fallback: iterate through all elements
                const allElements = xmlDoc.getElementsByTagName("*");
                for (let m = 0; m < allElements.length; m++) {
                    const elem = allElements[m];
                    if (elem.hasAttribute(attrName) && elem.getAttribute(attrName) === attrValue) {
                        matchingElements.push(elem);
                    }
                }
            }

            // Convert matching elements to custom structure
            for (let n = 0; n < matchingElements.length; n++) {
                const customNode = domNodeToCustom(matchingElements[n]);
                result.push(customNode);
            }
        } else {
            // Parse all nodes from document element
            const documentElement = xmlDoc.documentElement;

            if (documentElement) {
                // Create a virtual "?xml" node wrapper to match original parseXml behavior
                const xmlDeclarationNode = { tagName: "?xml", children: [domNodeToCustom(documentElement)] };

                // Parse document element itself (for parseNode option)
                if (options.parseNode) {
                    result = xmlDeclarationNode;
                } else {
                    // Simplify returns the structure with "?xml" key
                    result = [xmlDeclarationNode];
                }
            } else {
                result = [];
            }
        }

        // Apply filter if specified
        if (options.filter) {
            result = parseXml.filter(result, options.filter);
        }

        // Apply simplify if specified
        if (options.simplify) {
            result = parseXml.simplify(result);
        }

        result.pos = position;
        return result;
    } catch (e) {
        return [];
    }
}

/**
 * Simplify parsed XML structure to a more user-friendly format
 * @param {Array} nodes - Array of parsed nodes
 * @returns {Object} Simplified structure
 */
parseXml.simplify = function(nodes: any): any {
    const simplified: any = {};

    if (nodes === undefined) {
        return {};
    }

    // If single text node, return text
    if (nodes.length === 1 && typeof nodes[0] === "string") {
        return nodes[0];
    }

    // Process each node
    nodes.forEach(function(node: any) {
        if (typeof node === "object") {
            // Create array for this tag name
            if (!simplified[node.tagName]) {
                simplified[node.tagName] = [];
            }

            // Recursively simplify children
            const simplifiedNode = parseXml.simplify(node.children || []);
            simplified[node.tagName].push(simplifiedNode);

            // Add attributes (only if simplifiedNode is an object, not a string)
            if (node.attributes && typeof simplifiedNode === "object") {
                simplifiedNode.attrs = node.attributes;

                // Add order attribute
                simplifiedNode.attrs.order = elementOrder;
            } else if (typeof simplifiedNode === "object") {
                // Create attrs object if node has no attributes
                simplifiedNode.attrs = {
                    order: elementOrder
                };
            }
            elementOrder++;
        }
    });

    // Unwrap single-element arrays
    for (const tagName in simplified) {
        if (simplified[tagName].length === 1) {
            simplified[tagName] = simplified[tagName][0];
        }
    }

    return simplified;
};

/**
 * Filter nodes by predicate function
 * @param {Array} nodes - Array of nodes to filter
 * @param {Function} predicate - Filter predicate function
 * @returns {Array} Filtered nodes
 */
parseXml.filter = function(nodes: any, predicate: (node: any) => boolean): any[] {
    const filtered: any[] = [];

    nodes.forEach(function(node: any) {
        if (typeof node === "object" && predicate(node)) {
            filtered.push(node);
        }
        if (node.children) {
            const childFiltered = parseXml.filter(node.children, predicate);
            Array.prototype.push.apply(filtered, childFiltered);
        }
    });

    return filtered;
};

/**
 * Convert parsed XML structure back to XML string
 * @param {Array} nodes - Array of parsed nodes
 * @returns {string} XML string
 */
parseXml.stringify = function(nodes: any): string {
    let xmlOutput = "";

    function processChildren(children: any) {
        if (children) {
            for (let i = 0; i < children.length; i++) {
                if (typeof children[i] === "string") {
                    xmlOutput += children[i].trim();
                } else {
                    processNode(children[i]);
                }
            }
        }
    }

    function processNode(node: any) {
        xmlOutput += "<" + node.tagName;

        for (const attrName in node.attributes) {
            let attrValue = node.attributes[attrName];
            if (attrValue === null) {
                xmlOutput += " " + attrName;
            } else if (attrValue.indexOf('"') === -1) {
                xmlOutput += " " + attrName + '="' + attrValue.trim() + '"';
            } else {
                xmlOutput += " " + attrName + "='" + attrValue.trim() + "'";
            }
        }

        xmlOutput += ">";
        processChildren(node.children);
        xmlOutput += "</" + node.tagName + ">";
    }

    processChildren(nodes);
    return xmlOutput;
};

/**
 * Extract text content from parsed XML
 * @param {Array|Object|string} nodes - Parsed XML structure
 * @returns {string} Text content
 */
parseXml.toContentString = function(nodes: any): string {
    if (Array.isArray(nodes)) {
        let result = "";
        nodes.forEach(function(node: any) {
            result += " " + parseXml.toContentString(node);
            result = result.trim();
        });
        return result;
    }
    if (typeof nodes === "object") {
        return parseXml.toContentString(nodes.children);
    }
    return " " + nodes;
};

/**
 * Get element by id attribute
 * @param {string} xmlString - XML string
 * @param {string} idValue - ID value to find
 * @param {boolean} simplify - Whether to simplify result
 * @returns {Object|Array} Found element
 */
parseXml.getElementById = function(xmlString: string, idValue: string, simplify?: boolean): any {
    const result = parseXml(xmlString, {
        attrValue: idValue,
        simplify: simplify
    });
    return simplify ? result : result[0];
};

/**
 * Get elements by class name
 * @param {string} xmlString - XML string
 * @param {string} className - Class name to find
 * @param {boolean} simplify - Whether to simplify result
 * @returns {Array} Found elements
 */
parseXml.getElementsByClassName = function(xmlString: string, className: string, simplify?: boolean): any {
    return parseXml(xmlString, {
        attrName: "class",
        attrValue: "[a-zA-Z0-9-s ]*" + className + "[a-zA-Z0-9-s ]*",
        simplify: simplify
    });
};

/**
 * Parse XML stream (Node.js only)
 * @param {Stream} stream - Readable stream
 * @param {number|string} position - Start position
 * @returns {Stream} The stream for chaining
 */
parseXml.parseStream = function(stream: any, position?: number | string): any {
    if (typeof position === "function") {
        position = 0;
    }

    if (typeof position === "string") {
        position = (position as string).length + 2;
    }

    if (typeof stream === "string") {
        const fs = require("fs");
        stream = fs.createReadStream(stream, {
            start: position as number
        });
        position = 0;
    }

    let currentPosition = position as number;
    let buffer = "";
    let eventPos = 0;

    stream.on("data", function(data: any) {
        eventPos++;
        buffer += data;

        for (;;) {
            currentPosition = buffer.indexOf("<", currentPosition) + 1;
            const node = parseXml(buffer, {
                pos: currentPosition,
                parseNode: true
            });

            currentPosition = node.pos;

            if (currentPosition > buffer.length - 1 || eventPos > currentPosition) {
                if (eventPos) {
                    buffer = buffer.slice(eventPos);
                    currentPosition = 0;
                    eventPos = 0;
                }
                return;
            }

            stream.emit("xml", node);
            eventPos = currentPosition;
        }
    });

    stream.on("end", function() {
    });

    return stream;
};

/**
 * Reset element order counter
 * Useful for starting a fresh document
 */
function resetOrder(): void {
    elementOrder = 1;
}

// Also attach to parseXml object for compatibility
parseXml.resetOrder = resetOrder;

// Export for Node.js environment
if (typeof module !== "undefined") {
    (module as any).exports = parseXml;
}

export default parseXml;
export { parseXml };
export { resetOrder };

