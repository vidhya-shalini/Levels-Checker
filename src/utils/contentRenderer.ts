// Enhanced content rendering utilities
// Handles MathML, complex LaTeX, and embedded content

// MathML to LaTeX conversion map (basic)
const MATHML_TO_LATEX: Record<string, string> = {
  '<mi>': '',
  '</mi>': '',
  '<mo>': '',
  '</mo>': '',
  '<mn>': '',
  '</mn>': '',
  '<msup>': '^{',
  '</msup>': '}',
  '<msub>': '_{',
  '</msub>': '}',
  '<mfrac>': '\\frac{',
  '</mfrac>': '}',
  '<mrow>': '{',
  '</mrow>': '}',
  '<mtext>': '\\text{',
  '</mtext>': '}',
  '<msqrt>': '\\sqrt{',
  '</msqrt>': '}',
};

/**
 * Extract MathML from HTML content
 */
export const extractMathML = (html: string): string[] => {
  const mathmlRegex = /<math[^>]*>([\s\S]*?)<\/math>/gi;
  const results: string[] = [];
  let match;

  while ((match = mathmlRegex.exec(html)) !== null) {
    results.push(match[0]);
  }

  return results;
};

/**
 * Convert simple MathML to text representation
 */
export const convertMathMLToText = (mathml: string): string => {
  try {
    // Create a temporary div and parse MathML
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = mathml;

    // Extract text content
    let text = tempDiv.textContent || '';

    // Basic replacements for common MathML elements
    text = text.replace(/<\/?mi>/gi, '');
    text = text.replace(/<\/?mo>/gi, '');
    text = text.replace(/<\/?mn>/gi, '');
    text = text.replace(/<\/?mtext>/gi, '');
    text = text.replace(/<\/msup>/gi, '}');
    text = text.replace(/<msup>/gi, '^{');
    text = text.replace(/<\/msub>/gi, '}');
    text = text.replace(/<msub>/gi, '_{');

    // Clean whitespace
    text = text.replace(/\s+/g, ' ').trim();

    return text;
  } catch (e) {
    console.error('[MathML Converter] Error:', e);
    return mathml;
  }
};

/**
 * Detect if content contains complex mathematical content
 */
export const containsComplexMath = (text: string): boolean => {
  return (
    /<math/i.test(text) ||
    /\$\$[\s\S]*?\$\$/i.test(text) ||
    /\\\[[\s\S]*?\\\]/i.test(text) ||
    /\\frac\{/i.test(text) ||
    /\\sqrt\{/i.test(text) ||
    /\\int|\\sum|\\prod|\\lim/i.test(text)
  );
};

/**
 * Render MathML content to image
 */
export const renderMathMLToImage = async (mathml: string, fontSize: number = 24): Promise<string | null> => {
  try {
    // Try using MathJax if available, otherwise fall back to text rendering
    if (typeof (window as any).MathJax !== 'undefined') {
      // This would require MathJax to be loaded in the application
      console.warn('[MathML Renderer] MathJax rendering not yet implemented');
    }

    // Fallback: render as text
    const text = convertMathMLToText(mathml);
    return renderTextAsImage(text, fontSize);
  } catch (e) {
    console.error('[MathML Renderer] Error:', e);
    return null;
  }
};

/**
 * Render complex LaTeX expression to canvas image
 */
export const renderComplexLatex = async (
  latex: string,
  fontSize: number = 24
): Promise<string | null> => {
  try {
    // Check if LaTeX contains complex structures
    const isComplex =
      /\\int|\\sum|\\prod|\\frac|\\sqrt|\\begin/i.test(latex) ||
      /\^{[^}]{3,}}/i.test(latex); // complex superscripts

    if (isComplex && typeof (window as any).MathJax !== 'undefined') {
      // Could use MathJax for complex expressions
      console.log('[LaTeX Renderer] Complex expression detected:', latex.substring(0, 50));
    }

    // Fallback: simplified rendering
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    // Enhanced cleanup for complex LaTeX
    const displayText = simplifyLataxForDisplay(latex);

    // Measure text
    ctx.font = `${fontSize}px 'Cambria Math', 'Times New Roman', serif`;
    const metrics = ctx.measureText(displayText);
    const textWidth = metrics.width + 30;
    const textHeight = fontSize * 1.8 + 30;

    canvas.width = Math.max(textWidth, 150);
    canvas.height = Math.max(textHeight, 80);

    // Draw background and text
    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = 'black';
    ctx.font = `${fontSize}px 'Cambria Math', 'Times New Roman', serif`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    ctx.fillText(displayText, canvas.width / 2, canvas.height / 2);

    return canvas.toDataURL('image/png');
  } catch (e) {
    console.error('[LaTeX Renderer] Error:', e);
    return null;
  }
};

/**
 * Simplify LaTeX for display without MathJax
 */
const simplifyLataxForDisplay = (latex: string): string => {
  const result = latex
    // Fractions
    .replace(/\\frac\{([^}]*)\}\{([^}]*)\}/g, '($1)/($2)')
    // Roots
    .replace(/\\sqrt\[([^\]]*)\]\{([^}]*)\}/g, '$1√($2)')
    .replace(/\\sqrt\{([^}]*)\}/g, '√($1)')
    // Complex operators
    .replace(/\\int/gi, '∫')
    .replace(/\\sum/gi, 'Σ')
    .replace(/\\prod/gi, 'Π')
    .replace(/\\lim/gi, 'lim')
    // Superscripts and subscripts
    .replace(/\^\{([^}]+)\}/g, (_, exp) => '^' + exp)
    .replace(/\^(\d+)/g, (_, d) => '^' + d)
    .replace(/_\{([^}]+)\}/g, (_, exp) => '_' + exp)
    .replace(/_(\d+)/g, (_, d) => '_' + d)
    // Greek letters
    .replace(/\\alpha/gi, 'α')
    .replace(/\\beta/gi, 'β')
    .replace(/\\gamma/gi, 'γ')
    .replace(/\\delta/gi, 'δ')
    .replace(/\\epsilon/gi, 'ε')
    .replace(/\\pi/gi, 'π')
    .replace(/\\theta/gi, 'θ')
    .replace(/\\lambda/gi, 'λ')
    .replace(/\\mu/gi, 'μ')
    .replace(/\\sigma/gi, 'σ')
    .replace(/\\omega/gi, 'ω')
    // Symbols
    .replace(/\\infty/gi, '∞')
    .replace(/\\partial/gi, '∂')
    .replace(/\\nabla/gi, '∇')
    .replace(/\\leq/gi, '≤')
    .replace(/\\geq/gi, '≥')
    .replace(/\\neq/gi, '≠')
    .replace(/\\approx/gi, '≈')
    .replace(/\\times/gi, '×')
    .replace(/\\cdot/gi, '·')
    .replace(/\\pm/gi, '±')
    .replace(/\\rightarrow/gi, '→')
    .replace(/\\leftarrow/gi, '←')
    // Bracket cleanup
    .replace(/[{}]/g, '')
    // Remaining backslashes
    .replace(/\\\\/g, ' ')
    .replace(/\\/g, '');

  return result.replace(/\s+/g, ' ').trim();
};

/**
 * Render plain text as image
 */
const renderTextAsImage = async (text: string, fontSize: number = 20): Promise<string | null> => {
  try {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    ctx.font = `${fontSize}px 'Arial', sans-serif`;
    const metrics = ctx.measureText(text);
    const width = metrics.width + 20;
    const height = fontSize * 1.5 + 20;

    canvas.width = Math.max(width, 100);
    canvas.height = Math.max(height, 50);

    ctx.fillStyle = 'white';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    ctx.fillStyle = 'black';
    ctx.font = `${fontSize}px 'Arial', sans-serif`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';
    ctx.fillText(text, canvas.width / 2, canvas.height / 2);

    return canvas.toDataURL('image/png');
  } catch (e) {
    console.error('[TextRenderer] Error:', e);
    return null;
  }
};

/**
 * Process all mathematical content in HTML
 */
export const processMathematicalContent = async (html: string): Promise<{ processedHtml: string; mathCount: number }> => {
  let processedHtml = html;
  let mathCount = 0;

  // Extract and process MathML
  const mathmlElements = extractMathML(html);
  for (const mathml of mathmlElements) {
    const imageDataUrl = await renderMathMLToImage(mathml);
    if (imageDataUrl) {
      const altText = convertMathMLToText(mathml);
      const imgTag = `<img src="${imageDataUrl}" alt="Math: ${altText}" class="math-render" style="vertical-align: middle; max-height: 20px; margin: 0 2px;">`;
      processedHtml = processedHtml.replace(mathml, imgTag);
      mathCount++;
    }
  }

  return { processedHtml, mathCount };
};

/**
 * Extract textual descriptions of complex diagrams
 */
export const extractDiagramDescriptions = (html: string): string[] => {
  const descriptions: string[] = [];

  // Look for figure captions
  const figureRegex = /<figure[^>]*>([\s\S]*?)<\/figure>/gi;
  const captionRegex = /<figcaption[^>]*>([\s\S]*?)<\/figcaption>/gi;

  let match;
  while ((match = figureRegex.exec(html)) !== null) {
    const captionMatch = captionRegex.exec(match[1]);
    if (captionMatch) {
      const caption = captionMatch[1].replace(/<[^>]*>/g, '').trim();
      if (caption.length > 5) {
        descriptions.push(caption);
      }
    }
    captionRegex.lastIndex = 0;
  }

  // Also look for alt text in images
  const imgRegex = /<img[^>]*alt\s*=\s*["']?([^"'>\s]+)["']?[^>]*>/gi;
  while ((match = imgRegex.exec(html)) !== null) {
    const alt = match[1];
    if (alt && alt.length > 5 && !descriptions.includes(alt)) {
      descriptions.push(alt);
    }
  }

  return descriptions;
};
