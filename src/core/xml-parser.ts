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
            const matchingElements = xmlDoc.querySelectorAll(`[${options.attrName}="${options.attrValue}"]`);
            for (let i = 0; i < matchingElements.length; i++) {
                result.push(domNodeToCustom(matchingElements[i]));
            }
            return result;
        }

        // Simplify the DOM to match expected structure
        const rootElement = xmlDoc.documentElement;
        if (!rootElement) {
            return [];
        }

        // Convert the DOM to custom format
        const simplifiedResult = domNodeToCustom(rootElement);

        // Add simplify method to maintain compatibility
        (simplifiedResult as any).simplify = function(children: any[] = []): any {
            if (!this.children) return this;
            
            const result: any = {};
            result.tagName = this.tagName;
            if (this.attributes) {
                result.attributes = this.attributes;
            }
            
            if (children.length > 0) {
                // Convert children array to object if they have unique tagNames
                const childObj: any = {};
                for (const child of children) {
                    if (typeof child === 'object') {
                        childObj[child.tagName] = child;
                    }
                }
                result.children = childObj;
            } else if (this.children) {
                // Simplify own children
                const childObj: any = {};
                for (const child of this.children) {
                    if (typeof child === 'object') {
                        const node = child as XmlNode;
                        if (!childObj[node.tagName]) {
                            childObj[node.tagName] = node;
                        } else if (Array.isArray(childObj[node.tagName])) {
                            (childObj[node.tagName] as any[]).push(node);
                        } else {
                            childObj[node.tagName] = [childObj[node.tagName], node];
                        }
                    }
                }
                result.children = childObj;
            }
            
            return result;
        };

        return simplifiedResult;

    } catch (e) {
        console.error("XML parsing error:", e);
        return [];
    }
}

export default parseXml;
