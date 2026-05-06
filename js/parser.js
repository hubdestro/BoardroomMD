/**
 * Parser Module
 * Uses marked.js solely for generating an Abstract Syntax Tree (AST).
 * Does NOT generate HTML.
 */

window.MarkdownParser = {
    /**
     * Parses raw markdown text into a flat/nested array of tokens.
     * @param {string} rawMarkdown - The input markdown string.
     * @returns {Array} Array of token objects.
     */
    parse: function(rawMarkdown) {
        if (!window.marked) {
            console.error("marked.js is not loaded.");
            return [];
        }

        try {
            // Use marked.lexer to get the AST tokens instead of HTML
            const tokens = window.marked.lexer(rawMarkdown);
            console.log("AST Tokens generated:", tokens);
            return tokens;
        } catch (error) {
            console.error("Error parsing markdown:", error);
            return [];
        }
    }
};
