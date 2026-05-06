/**
 * Mapper Module
 * Maps Markdown AST tokens to native docx.js objects.
 */

window.TokenMapper = {
    /**
     * Converts an array of marked.js tokens into an array of docx.js elements.
     * @param {Array} tokens - The AST tokens from marked.js.
     * @returns {Array} Array of docx components (e.g., Paragraphs).
     */
    mapTokensToDocx: function(tokens) {
        if (!window.docx) {
            console.error("docx library is not loaded.");
            return [];
        }

        const docElements = [];

        tokens.forEach(token => {
            const element = this.processToken(token);
            if (element) {
                // If it returns an array of elements (like a list), spread them
                if (Array.isArray(element)) {
                    docElements.push(...element);
                } else {
                    docElements.push(element);
                }
            }
        });

        return docElements;
    },

    /**
     * Processes a single token.
     * @param {Object} token - A single marked.js token.
     * @param {Object} options - Formatting options (e.g., list depth).
     */
    processToken: function(token, options = {}) {
        const { Paragraph, TextRun, HeadingLevel, AlignmentType } = window.docx;

        switch (token.type) {
            case 'heading':
                const headingLevel = `HEADING_${token.depth}`;
                return new Paragraph({
                    children: this.processInlineTokens(token.tokens || []),
                    heading: HeadingLevel[headingLevel] || HeadingLevel.HEADING_1,
                });

            case 'paragraph':
                return new Paragraph({
                    children: this.processInlineTokens(token.tokens || [])
                });

            case 'space':
                // Could return a blank paragraph or ignore
                return new Paragraph({ children: [] });

            case 'blockquote':
                return new Paragraph({
                    children: this.processInlineTokens(token.tokens || [{ type: 'text', text: token.text || "" }]),
                    style: "Quote"
                });

            case 'list':
                const listElements = [];
                // Process each list item
                token.items.forEach(item => {
                    // For simplicity, handle as bullet points
                    // A more robust implementation would handle nested lists and ordered lists
                    const itemParagraph = new Paragraph({
                        children: this.processInlineTokens(item.tokens || []),
                        bullet: { level: options.depth || 0 }
                    });
                    listElements.push(itemParagraph);
                });
                return listElements;

            case 'hr':
                // Simple horizontal rule implementation (can be styled better in template)
                return new Paragraph({
                    text: "---",
                    alignment: AlignmentType.CENTER
                });

            default:
                // Graceful fallback for unrecognized tokens
                console.warn(`Unrecognized token type: ${token.type}. Falling back to plain text paragraph.`);
                return new Paragraph({
                    children: [new TextRun({ text: token.raw || token.text || "" })]
                });
        }
    },

    /**
     * Processes inline tokens (text, strong, em, link) into TextRuns.
     * @param {Array} tokens - Inline tokens array.
     */
    processInlineTokens: function(tokens) {
        if (!tokens || tokens.length === 0) return [];

        const { TextRun } = window.docx;

        const toRuns = (inlineTokens, inherited = {}) => {
            const runs = [];
            if (!inlineTokens || inlineTokens.length === 0) return runs;

            inlineTokens.forEach(token => {
                switch (token.type) {
                    case 'text':
                    case 'escape':
                    case 'html': {
                        if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                            runs.push(...toRuns(token.tokens, inherited));
                            return;
                        }
                        const text = token.text ?? token.raw ?? "";
                        if (text) runs.push(new TextRun({ text, ...inherited }));
                        return;
                    }
                    case 'strong': {
                        const next = { ...inherited, bold: true };
                        if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                            runs.push(...toRuns(token.tokens, next));
                        } else {
                            const text = token.text ?? "";
                            if (text) runs.push(new TextRun({ text, ...next }));
                        }
                        return;
                    }
                    case 'em': {
                        const next = { ...inherited, italics: true };
                        if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                            runs.push(...toRuns(token.tokens, next));
                        } else {
                            const text = token.text ?? "";
                            if (text) runs.push(new TextRun({ text, ...next }));
                        }
                        return;
                    }
                    case 'del': {
                        const next = { ...inherited, strike: true };
                        if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                            runs.push(...toRuns(token.tokens, next));
                        } else {
                            const text = token.text ?? "";
                            if (text) runs.push(new TextRun({ text, ...next }));
                        }
                        return;
                    }
                    case 'codespan': {
                        const text = token.text ?? "";
                        if (text) runs.push(new TextRun({ text, ...inherited, font: "Courier New", color: "d63384" }));
                        return;
                    }
                    case 'link': {
                        const text = token.text ?? token.href ?? "";
                        const href = token.href ? ` (${token.href})` : "";
                        if (text) {
                            runs.push(new TextRun({ text: `${text}${href}`, ...inherited, color: "0000EE", underline: {} }));
                        }
                        return;
                    }
                    case 'br': {
                        runs.push(new TextRun({ text: "", break: 1 }));
                        return;
                    }
                    default: {
                        if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                            runs.push(...toRuns(token.tokens, inherited));
                            return;
                        }
                        if (token.type === 'paragraph' || token.type === 'heading') {
                            if (Array.isArray(token.tokens) && token.tokens.length > 0) {
                                runs.push(...toRuns(token.tokens, inherited));
                                return;
                            }
                        }
                        const text = token.text ?? token.raw ?? "";
                        if (text) runs.push(new TextRun({ text, ...inherited }));
                        return;
                    }
                }
            });

            return runs;
        };

        return toRuns(tokens, {});
    }
};
