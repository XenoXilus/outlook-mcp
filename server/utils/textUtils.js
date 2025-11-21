import DOMPurify from 'isomorphic-dompurify';

/**
 * Strips HTML tags from a string and decodes common entities.
 * @param {string} html - The HTML string to strip.
 * @returns {string} The plain text content.
 */
export function stripHtml(html) {
    if (!html) return '';

    // First use DOMPurify to sanitize (good practice)
    const cleanHtml = DOMPurify.sanitize(html);

    // Then strip tags using a temporary DOM element approach (simulated here for Node.js)
    // Since we're in Node, we can use a regex approach or a library. 
    // Given we have isomorphic-dompurify, we can use it to get text content if we had a window,
    // but in Node, regex is often faster for simple stripping after sanitization.
    // However, a better approach for "text content" is to replace block tags with newlines first.

    let text = cleanHtml
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>/gi, '\n\n')
        .replace(/<\/div>/gi, '\n')
        .replace(/<\/tr>/gi, '\n')
        .replace(/<\/li>/gi, '\n')
        .replace(/<[^>]+>/g, ''); // Strip remaining tags

    // Decode common entities
    text = text
        .replace(/&nbsp;/g, ' ')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'");

    // Collapse multiple newlines and trim
    return text.replace(/\n\s*\n/g, '\n\n').trim();
}

/**
 * Truncates text to a specified length.
 * @param {string} text - The text to truncate.
 * @param {number} maxLength - The maximum length.
 * @returns {string} The truncated text.
 */
export function truncateText(text, maxLength = 1000) {
    if (!text) return '';
    if (text.length <= maxLength) return text;

    return text.substring(0, maxLength) + `\n... [truncated, original length: ${text.length} chars]`;
}
