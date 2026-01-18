/**
 * XML Parser - A lightweight XML parser for PPTX parsing
 * Uses native DOMParser API but maintains original parseXml output format
 */

// Module-level variable to track element order across multiple parseXml calls
var elementOrder = 1;

/**
 * Convert DOM node to custom structure (matching parseXml format)
 * @param {Node} domNode - DOM node to convert
 * @returns {Object|string} Converted node or text
 */
function domNodeToCustom(domNode) {
    // Text node
    if (domNode.nodeType === Node.TEXT_NODE) {
        var text = domNode.textContent;
        return text;
    }

    // CDATA section
    if (domNode.nodeType === Node.CDATA_SECTION_NODE) {
        return domNode.textContent;
    }

    // Element node
    if (domNode.nodeType === Node.ELEMENT_NODE) {
        var node = {};

        // Keep original case tagName
        node.tagName = domNode.tagName;

        // Parse attributes
        for (var i = 0; i < domNode.attributes.length; i++) {
            if (!node.attributes) {
                node.attributes = {};
            }
            var attr = domNode.attributes[i];
            node.attributes[attr.name] = attr.value;
        }

        // Parse children
        var hasChildren = false;
        for (var j = 0; j < domNode.childNodes.length; j++) {
            var childNode = domNodeToCustom(domNode.childNodes[j]);
            if (childNode !== null && childNode !== "") {
                if (!hasChildren) {
                    node.children = [];
                    hasChildren = true;
                }
                node.children.push(childNode);
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
function parseXml(xmlString, options) {
    "use strict";

    options = options || {};
    var position = options.pos || 0;

    var result = null;

    // Create DOMParser and parse XML
    var parser;
    try {
        if (typeof DOMParser !== 'undefined') {
            parser = new DOMParser();
        } else if (typeof require !== 'undefined') {
            // Node.js environment fallback
            var xmldom = require('xmldom');
            parser = new xmldom.DOMParser();
        } else {
            throw new Error('No XML parser available');
        }

        var xmlDoc = parser.parseFromString(xmlString, "text/xml");

        // Check for parsing errors
        var parserError = xmlDoc.getElementsByTagName("parsererror");
        if (parserError.length > 0) {
            console.error("XML Parse Error:", parserError[0].textContent);
            return [];
        }

        // If filtering by attribute value
        if (options.attrValue !== undefined) {
            options.attrName = options.attrName || "id";
            result = [];

            // Find all matching elements using XPath or iteration
            var attrName = options.attrName;
            var attrValue = options.attrValue;
            var matchingElements = [];

            if (typeof xmlDoc.evaluate !== 'undefined') {
                var xpath = "//*[@" + attrName + "='" + attrValue + "']";
                var xpathResult = xmlDoc.evaluate(xpath, xmlDoc, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
                for (var k = 0; k < xpathResult.snapshotLength; k++) {
                    matchingElements.push(xpathResult.snapshotItem(k));
                }
            } else {
                // Fallback: iterate through all elements
                var allElements = xmlDoc.getElementsByTagName("*");
                for (var m = 0; m < allElements.length; m++) {
                    var elem = allElements[m];
                    if (elem.hasAttribute(attrName) && elem.getAttribute(attrName) === attrValue) {
                        matchingElements.push(elem);
                    }
                }
            }

            // Convert matching elements to custom structure
            for (var n = 0; n < matchingElements.length; n++) {
                var customNode = domNodeToCustom(matchingElements[n]);
                result.push(customNode);
            }
        } else {
            // Parse all nodes from document element
            var documentElement = xmlDoc.documentElement;

            if (documentElement) {
                // Create a virtual "?xml" node wrapper to match original parseXml behavior
                var xmlDeclarationNode = { tagName: "?xml", children: [domNodeToCustom(documentElement)] };

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
        console.error("Error parsing XML:", e);
        return [];
    }
}

/**
 * Simplify parsed XML structure to a more user-friendly format
 * @param {Array} nodes - Array of parsed nodes
 * @returns {Object} Simplified structure
 */
parseXml.simplify = function(nodes) {
    var simplified = {};

    if (nodes === undefined) {
        return {};
    }

    // If single text node, return text
    if (nodes.length === 1 && typeof nodes[0] === "string") {
        return nodes[0];
    }

    // Process each node
    nodes.forEach(function(node) {
        if (typeof node === "object") {
            // Create array for this tag name
            if (!simplified[node.tagName]) {
                simplified[node.tagName] = [];
            }

            // Recursively simplify children
            var simplifiedNode = parseXml.simplify(node.children || []);
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
    for (var tagName in simplified) {
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
parseXml.filter = function(nodes, predicate) {
    var filtered = [];

    nodes.forEach(function(node) {
        if (typeof node === "object" && predicate(node)) {
            filtered.push(node);
        }
        if (node.children) {
            var childFiltered = parseXml.filter(node.children, predicate);
            filtered = filtered.concat(childFiltered);
        }
    });

    return filtered;
};

/**
 * Convert parsed XML structure back to XML string
 * @param {Array} nodes - Array of parsed nodes
 * @returns {string} XML string
 */
parseXml.stringify = function(nodes) {
    function processChildren(children) {
        if (children) {
            for (var i = 0; i < children.length; i++) {
                if (typeof children[i] === "string") {
                    xmlOutput += children[i].trim();
                } else {
                    processNode(children[i]);
                }
            }
        }
    }

    function processNode(node) {
        xmlOutput += "<" + node.tagName;

        for (var attrName in node.attributes) {
attrValue = node.attributes[attrName];
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

    var xmlOutput = "";
    processChildren(nodes);
    return xmlOutput;
};

/**
 * Extract text content from parsed XML
 * @param {Array|Object|string} nodes - Parsed XML structure
 * @returns {string} Text content
 */
parseXml.toContentString = function(nodes) {
    if (Array.isArray(nodes)) {
result = "";
        nodes.forEach(function(node) {
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
parseXml.getElementById = function(xmlString, idValue, simplify) {
result = parseXml(xmlString, {
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
parseXml.getElementsByClassName = function(xmlString, className, simplify) {
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
parseXml.parseStream = function(stream, position) {
    if (typeof position === "function") {
        position = 0;
    }

    if (typeof position === "string") {
        position = position.length + 2;
    }

    if (typeof stream === "string") {
        var fs = require("fs");
        stream = fs.createReadStream(stream, {
            start: position
        });
        position = 0;
    }

    var currentPosition = position;
    var buffer = "";
    var eventPos = 0;

    stream.on("data", function(data) {
        eventPos++;
        buffer += data;

        for (;;) {
            currentPosition = buffer.indexOf("<", currentPosition) + 1;
node = parseXml(buffer, {
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
        console.log("end");
    });

    return stream;
};

/**
 * Reset element order counter
 * Useful for starting a fresh document
 */
function resetOrder() {
    elementOrder = 1;
}

// Also attach to parseXml object for compatibility
parseXml.resetOrder = resetOrder;

// Export for Node.js environment
if (typeof module !== "undefined") {
    module.exports = parseXml;
}

// ES Module exports
export default parseXml;
export { parseXml };
export { resetOrder };


