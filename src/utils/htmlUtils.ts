// Utility functions for HTML processing and image extraction

// Check if text contains HTML tags that need rendering
export const containsHtml = (text: string): boolean => {
    return /<(img|br|sub|sup|em|strong|b|i|u|table|div|span|p|math|svg)[^>]*>/i.test(text);
};

// Strip HTML tags from text to get plain text (for AI operations)
// Strip HTML tags from text to get plain text (for AI operations)
export const stripHtmlForAI = (text: string): string => {
    // Protect math, svg, and img tags from stripping
    const protectedText = text.replace(/<(math|svg|img|table|tr|td|th)([\s\S]*?)(\/?>|<\/\1>)/gi, (match) => {
        return match.replace(/</g, '___LT___').replace(/>/g, '___GT___');
    });

    const stripped = protectedText
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/\s+/g, ' ')
        .trim();

    // Restore protected tags
    return stripped.replace(/___LT___/g, '<').replace(/___GT___/g, '>');
};

// Extremely robust regex to catch src regardless of other attributes
// Matches: <img ... src="data:..." ...> or <img src='...'> etc.
export const IMAGE_SRC_REGEX = /<img[^>]+?src\s*=\s*["']?([^"'>\s]+?)["']?[^>]*?>/gi;

// Extract image and SVG HTML tags from text
export const extractImageHtml = (text: string): string[] => {
    const imgMatches = text.match(/<img[^>]+?src\s*=\s*["']?([^"'>\s]+?)["']?[^>]*?>/gi) || [];
    const svgMatches = text.match(/<svg[\s\S]*?<\/svg>/gi) || [];
    return [...imgMatches, ...svgMatches];
};

// Re-attach image and SVG HTML to text after AI processing
export const reattachImages = (newText: string, originalHtml?: string): string => {
    if (!originalHtml) return newText;
    const media = extractImageHtml(originalHtml);
    if (media.length === 0) return newText;
    // Append media at the end of the new text
    return newText + '<br>' + media.join('<br>');
};
