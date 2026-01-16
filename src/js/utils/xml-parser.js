/**
 * XML Parser - A lightweight XML parser for PPTX parsing
 * Based on parseXml library
 */

var elementOrder = 1;

/**
 * Main XML parsing function
 * @param {string} xmlString - The XML string to parse
 * @param {Object} options - Parsing options
 * @returns {Array|Object} Parsed XML structure
 */
function parseXml(xmlString, options) {
    "use strict";

    /**
     * Parse children nodes from current position
     */
    function parseChildren() {
        var children = [];
        while (xmlString[position]) {
            if (xmlString.charCodeAt(position) === lessThanCode) {
                // Check for closing tag
                if (xmlString.charCodeAt(position + 1) === slashCode) {
                    position = xmlString.indexOf(greaterThanChar, position);
                    if (position + 1) {
                        position += 1;
                    }
                    return children;
                }

                // Check for special tags (comments, CDATA, DOCTYPE)
                if (xmlString.charCodeAt(position + 1) === exclamationCode) {
                    if (xmlString.charCodeAt(position + 2) === dashCode) {
                        // Parse comment: <!-- ... -->
                        while (-1 !== position && 
                               !(xmlString.charCodeAt(position) === greaterThanCode &&
                                 xmlString.charCodeAt(position - 1) === dashCode &&
                                 xmlString.charCodeAt(position - 2) === dashCode) &&
                                 position !== -1) {
                            position = xmlString.indexOf(greaterThanChar, position + 1);
                        }
                        if (position === -1) {
                            position = xmlString.length;
                        }
                    } else {
                        // Parse special content like <![CDATA[...]]> or <!DOCTYPE...>
                        position += 2;
                        while (xmlString.charCodeAt(position) !== greaterThanCode && xmlString[position]) {
                            position++;
                        }
                    }
                    position++;
                    continue;
                }

                // Parse regular node
                var node = parseNode();
                children.push(node);
            } else {
                // Parse text content
                var text = parseTextContent();
                if (text.trim().length > 0) {
                    children.push(text);
                }
                position++;
            }
        }
        return children;
    }

    /**
     * Parse text content between tags
     */
    function parseTextContent() {
        var start = position;
        position = xmlString.indexOf(lessThanChar, position) - 1;
        if (position === -2) {
            position = xmlString.length;
        }
        return xmlString.slice(start, position + 1);
    }

    /**
     * Parse tag name from current position
     */
    function parseTagName() {
        var start = position;
        while (delimiterChars.indexOf(xmlString[position]) === -1 && xmlString[position]) {
            position++;
        }
        return xmlString.slice(start, position);
    }

    /**
     * Parse a single XML node (tag with attributes and children)
     */
    function parseNode() {
        var node = {};
        position++;
        node.tagName = parseTagName();
        var hasAttributes = false;

        // Parse attributes
        while (xmlString.charCodeAt(position) !== greaterThanCode && xmlString[position]) {
            var charCode = xmlString.charCodeAt(position);
            
            // Check if this is an attribute name (letter)
            if ((charCode > 64 && charCode < 91) || (charCode > 96 && charCode < 123)) {
                var attrName = parseTagName();
                var attrValue = null;
                
                // Skip to attribute value
                var nextChar = xmlString.charCodeAt(position);
                while (nextChar && 
                       nextChar !== singleQuoteCode && 
                       nextChar !== doubleQuoteCode && 
                       !((nextChar > 64 && nextChar < 91) || (nextChar > 96 && nextChar < 123)) && 
                       nextChar !== greaterThanCode) {
                    position++;
                    nextChar = xmlString.charCodeAt(position);
                }

                if (!hasAttributes) {
                    node.attributes = {};
                    hasAttributes = true;
                }

                // Parse quoted attribute value
                if (nextChar === singleQuoteCode || nextChar === doubleQuoteCode) {
                    attrValue = parseQuotedString();
                    if (position === -1) {
                        return node;
                    }
                } else {
                    attrValue = null;
                    position--;
                }

                node.attributes[attrName] = attrValue;
            }
            position++;
        }

        // Parse children or self-closing tag
        if (xmlString.charCodeAt(position - 1) !== slashCode) {
            // Check for special tags with raw content
            if (node.tagName === "script") {
                var contentStart = position + 1;
                position = xmlString.indexOf("</script>", position);
                node.children = [xmlString.slice(contentStart, position - 1)];
                position += 8;
            } else if (node.tagName === "style") {
                var contentStart = position + 1;
                position = xmlString.indexOf("</style>", position);
                node.children = [xmlString.slice(contentStart, position - 1)];
                position += 7;
            } else if (voidTags.indexOf(node.tagName) === -1) {
                // Regular tag with children
                position++;
                node.children = parseChildren();
            }
        } else {
            // Self-closing tag
            position++;
        }

        return node;
    }

    /**
     * Parse a quoted string (attribute value)
     */
    function parseQuotedString() {
        var quoteChar = xmlString[position];
        var startPosition = ++position;
        position = xmlString.indexOf(quoteChar, startPosition);
        return xmlString.slice(startPosition, position);
    }

    /**
     * Find attribute position in XML string (for attribute filtering)
     */
    function findAttributePosition() {
        var pattern = new RegExp("\\s" + options.attrName + "\\s*=['\"]" + options.attrValue + "['\"]");
        var match = pattern.exec(xmlString);
        return match ? match.index : -1;
    }

    // Initialize variables
    options = options || {};
    var position = options.pos || 0;
    
    // Character codes for common XML characters
    var lessThanChar = "<";
    var lessThanCode = "<".charCodeAt(0);
    var greaterThanChar = ">";
    var greaterThanCode = ">".charCodeAt(0);
    var dashCode = "-".charCodeAt(0);
    var slashCode = "/".charCodeAt(0);
    var exclamationCode = "!".charCodeAt(0);
    var singleQuoteCode = "'".charCodeAt(0);
    var doubleQuoteCode = '"'.charCodeAt(0);
    var delimiterChars = "\n\t>/= ";
    
    // Tags that don't have children (void elements)
    var voidTags = ["img", "br", "input", "meta", "link"];
    
    var result = null;

    // If filtering by attribute value
    if (options.attrValue !== undefined) {
        options.attrName = options.attrName || "id";
        result = [];
        while (-1 !== (position = findAttributePosition())) {
            position = xmlString.lastIndexOf("<", position);
            if (position !== -1) {
                result.push(parseNode());
                xmlString = xmlString.substring(position);
                position = 0;
            }
        }
    } else {
        // Parse all nodes or a single node
        result = options.parseNode ? parseNode() : parseChildren();
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
    
    // If single text node, return the text
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

            // Add attributes
            if (node.attributes) {
                simplifiedNode.attrs = node.attributes;
            }

            // Add order attribute
            if (simplifiedNode.attrs === undefined) {
                simplifiedNode.attrs = {
                    order: elementOrder
                };
            } else {
                simplifiedNode.attrs.order = elementOrder;
            }
            elementOrder++;
            console.log(elementOrder);
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
            var attrValue = node.attributes[attrName];
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
        var result = "";
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
    var result = parseXml(xmlString, {
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
        // Callback function is not used in this implementation
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
            var node = parseXml(buffer, {
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
 * Useful for testing when you need to start fresh
 */
parseXml.resetOrder = function() {
    elementOrder = 1;
};

// Export for Node.js environment
if (typeof module !== "undefined") {
    module.exports = parseXml;
}

// Support for ES modules - export as default and named
// export default parseXml;
// export { parseXml };
