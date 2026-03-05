
import mammoth from 'mammoth';
import * as pdfjsLib from 'pdfjs-dist';

// Configure PDF.js worker
// Use unpkg as it reliably mirrors npm versions
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

export interface ParseFileResult {
  html: string;
  pdfPageImages?: (string | null)[];
}

export const parseFile = async (file: File): Promise<ParseFileResult> => {
  const fileType = file.name.split('.').pop()?.toLowerCase();

  if (fileType === 'html' || fileType === 'htm') {
    const html = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => resolve(e.target?.result as string);
      reader.onerror = (e) => reject(e);
      reader.readAsText(file);
    });
    return { html };
  } else if (fileType === 'docx') {
    const html = await parseDocx(file);
    return { html };
  } else if (fileType === 'pdf') {
    return parsePdf(file);
  } else {
    throw new Error('Unsupported file type');
  }
};

const parseDocx = async (file: File): Promise<string> => {
  const arrayBuffer = await file.arrayBuffer();

  // We need TWO passes:
  // 1. convertToHtml for the full HTML with images
  // 2. extractRawText for reliable metadata extraction (mammoth's HTML often splits table cell text)
  const [htmlResult, textResult] = await Promise.all([
    mammoth.convertToHtml(
      { arrayBuffer },
      {
        convertImage: mammoth.images.imgElement(function (image: any) {
          return image.read("base64").then(function (imageBuffer: string) {
            const contentType = image.contentType || 'image/png';
            console.log(`[DOCX Parser] Found embedded image: type=${contentType}, size=${imageBuffer.length}`);
            return {
              src: `data:${contentType};base64,${imageBuffer}`
            };
          });
        })
      }
    ),
    mammoth.extractRawText({ arrayBuffer })
  ]);

  let htmlOutput = htmlResult.value;

  // Fix common Mammoth truncation artifacts in headers for better preview presentation
  htmlOutput = htmlOutput.replace(/>\s*bject\s*/gi, '>Subject ');
  htmlOutput = htmlOutput.replace(/>\s*anch\s*</gi, '>Branch<');
  htmlOutput = htmlOutput.replace(/>\s*ranch\s*</gi, '>Branch<');

  // Count images for debugging
  const imgCount = (htmlOutput.match(/<img/gi) || []).length;
  console.log(`[DOCX Parser] Converted DOCX to HTML: ${htmlOutput.length} chars, ${imgCount} images found`);
  if (htmlResult.messages.length > 0) {
    console.log('[DOCX Parser] Mammoth messages:', htmlResult.messages);
  }

  // ============================================================
  // DOCX Metadata Extraction from raw text
  // ============================================================
  // Mammoth's HTML output often garbles table header cells by splitting
  // labels across different <td> elements (e.g., "Date:" → "Da"|"te:").
  // We extract metadata from the RAW TEXT which preserves cell content
  // more reliably, then inject it into the HTML with IDs so that
  // extractMetadata() in htmlParser.ts can find them.
  const rawText = textResult.value;
  console.log('[DOCX Parser] Raw text length:', rawText.length);
  console.log('[DOCX Parser] Raw text first 1200 chars:', rawText.substring(0, 1200));

  let subjectCode = '';
  let subjectName = '';
  let date = '';
  let branch = '';
  let yearSem = '';
  let maxMarks = '';
  let time = '';

  // Split raw text into lines for line-by-line analysis
  // Mammoth extractRawText often puts each table cell on its own line
  const rawLines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const headerLines = rawLines.slice(0, 50); // Search header region (wider search)

  // Log all header lines for debugging
  console.log('[DOCX Parser] Header lines:');
  for (let d = 0; d < Math.min(30, headerLines.length); d++) {
    console.log(`[DOCX Parser]   Line ${d}: "${headerLines[d]}"`);
  }

  for (let i = 0; i < headerLines.length; i++) {
    const line = headerLines[i];
    const nextLine = (i + 1 < headerLines.length) ? headerLines[i + 1] : '';
    const nextLine2 = (i + 2 < headerLines.length) ? headerLines[i + 2] : '';

    // ========== Subject Code / Name ==========
    // Pattern: "EC 5102 / CIRCUIT THEORY" or "EC5102 / CIRCUIT THEORY"
    // Note: subject code may have a SPACE between letters and digits (e.g. "EC 5102")
    if (!subjectCode && /Subject\s*Code/i.test(line)) {
      // Same line: "Subject Code / Name EC 5102 / CIRCUIT THEORY"
      const sameLine = line.match(/Subject\s*Code\s*[/\\\\]?\s*Name\s*[:.\\-]?\s*([A-Z]{2,}\s*\d{3,})\s*[/\\\\]\s*(.+?)(?:\s+Time|\s+Max|$)/i);
      if (sameLine) {
        subjectCode = sameLine[1].replace(/\s+/g, ' ').trim();
        subjectName = sameLine[2].trim();
      } else {
        // Same line but code and name may not be separated by slash clearly
        const codeLine = line.match(/Subject\s*Code\s*[/\\\\]?\s*Name\s*[:.\\-]?\s*([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]?\s*(.*)/i);
        if (codeLine) {
          subjectCode = codeLine[1].replace(/\s+/g, ' ').trim();
          subjectName = codeLine[2].trim().replace(/\s+Time\s+.*$/i, '').replace(/\s+Max\s+.*$/i, '').trim();
        }
      }
      // If not found on same line, check next line — value might be on a separate line
      if (!subjectCode && nextLine) {
        // Next line might be "EC 5102 / CIRCUIT THEORY" or "EC5102 / CIRCUIT THEORY"
        const nextMatch = nextLine.match(/^([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]\s*(.+)/i);
        if (nextMatch) {
          subjectCode = nextMatch[1].replace(/\s+/g, ' ').trim();
          subjectName = nextMatch[2].trim().replace(/\s+Time\s+.*$/i, '').replace(/\s+Max\s+.*$/i, '').trim();
        }
      }
      // Label-only line: "Subject Code / Name" and value is entirely on next line
      if (!subjectCode && /^\s*Subject\s*Code\s*[/\\\\]?\s*Name\s*[:.\\-]?\s*$/i.test(line) && nextLine) {
        const nextMatch = nextLine.match(/^([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]\s*(.+)/i);
        if (nextMatch) {
          subjectCode = nextMatch[1].replace(/\s+/g, ' ').trim();
          subjectName = nextMatch[2].trim();
        } else if (nextLine.length > 3 && !/^(Date|Branch|Time|Max|Year)/i.test(nextLine)) {
          // Next line is just the raw value without clear code/name separation
          subjectName = nextLine;
        }
      }
    }

    // ========== Date ==========
    if (!date && /Date/i.test(line)) {
      const dateMatch = line.match(/Date\s*[:.\\-]?\s*([\d./-]+[\d])/i);
      if (dateMatch) {
        date = dateMatch[1].trim();
      } else if (nextLine && /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(nextLine)) {
        date = nextLine;
      }
      // Label-only "Date" or "Date:" line
      if (!date && /^\s*Date\s*[:.\\-]?\s*$/i.test(line) && nextLine) {
        const dm = nextLine.match(/^([\d./-]+)$/);
        if (dm) date = dm[1].trim();
      }
    }

    // ========== Max Marks ==========
    if (!maxMarks && /Max/i.test(line)) {
      const mMatch = line.match(/Max\.?\s*Marks?\s*[:.\\-]?\s*(\d+(?:\s*Marks?)?)/i);
      if (mMatch) maxMarks = mMatch[1].trim();
      // Label on one line, value on next
      if (!maxMarks && /^\s*Max\.?\s*Marks?\s*[:.\\-]?\s*$/i.test(line) && nextLine) {
        const nm = nextLine.match(/^(\d+(?:\s*Marks?)?)/i);
        if (nm) maxMarks = nm[1].trim();
      }
    }

    // ========== Time ==========
    if (!time && /Time/i.test(line) && !/Time\s*(?:Variant|Invariant|domain|signal|Shifting)/i.test(line)) {
      const tMatch = line.match(/Time\s*[:.\\-]?\s*(\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?))/i);
      if (tMatch) {
        time = tMatch[1].trim();
      } else if (nextLine && /^\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?)/i.test(nextLine)) {
        time = nextLine.match(/^\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?)/i)![0];
      }
      // Label-only "Time" line
      if (!time && /^\s*Time\s*[:.\\-]?\s*$/i.test(line) && nextLine) {
        const tm = nextLine.match(/^(\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?))/i);
        if (tm) time = tm[1].trim();
      }
    }

    // ========== Branch ==========
    // e.g. "Branch COMMON TO ECE, VLSI, ACT" or "Branch: B.E - EC(ACT)"
    // or label "Branch" on one line and "COMMON TO ECE, VLSI, ACT" on the next
    if (!branch && /Branch/i.test(line)) {
      // Same line with value
      const brMatch = line.match(/Branch\s*[:.\\-]?\s*(.+?)(?:\s+Year|\s+Date|$)/i);
      if (brMatch && brMatch[1].trim().length > 1) {
        branch = brMatch[1].trim();
      }
      // Label-only "Branch" or "Branch:" line — grab next line as value
      if (!branch && /^\s*Branch\s*[:.\\-]?\s*$/i.test(line) && nextLine) {
        if (!/Year|Sem|Date|Subject|Q\.?\s*No|Time|Max/i.test(nextLine) && nextLine.length > 1) {
          branch = nextLine.replace(/\s+Year\s*[/]?\s*Sem.*$/i, '').trim();
        }
      }
      // If brMatch found something but it's really short (e.g. just "B"), also check next line
      if (!branch && nextLine && !/Year|Sem|Date|Subject|Q\.?\s*No|Time|Max/i.test(nextLine) && nextLine.length > 1) {
        branch = nextLine.replace(/\s+Year\s*[/]?\s*Sem.*$/i, '').trim();
      }
    }

    // ========== Year / Sem ==========
    if (!yearSem && /Year\s*[/\\\\]?\s*Sem/i.test(line)) {
      const ysMatch = line.match(/Year\s*[/\\\\]?\s*Sem(?:ester)?\s*[:.\\-]?\s*(.+?)$/i);
      if (ysMatch && ysMatch[1].trim().length > 0) {
        const val = ysMatch[1].trim();
        if (val !== ':' && val !== '-') {
          yearSem = val;
        }
      }
      // Label-only "Year / Sem" or "Year/Sem:" line — grab next line
      if (!yearSem && nextLine && nextLine.length > 0) {
        let valueCandidate = nextLine.trim();
        // If the next line is just a colon/dash, we skip it and look at the line after that
        if ((valueCandidate === ':' || valueCandidate === '-') && nextLine2) {
          valueCandidate = nextLine2.trim();
        }
        // Accept almost anything as year/sem value (Roman numerals, digits, slashes)
        if (valueCandidate && valueCandidate !== ':' && valueCandidate !== '-' && !/Subject|Branch|Date|Time|Max|Q\.?\s*No|Course/i.test(valueCandidate)) {
          yearSem = valueCandidate;
        }
      }
    }
  }

  // ========== Standalone value scanning ==========
  // In mammoth raw text, table cells often become separate lines.
  // Sometimes the label "Subject Code / Name" is on one line and the value
  // like "EC 5102 / CIRCUIT THEORY" appears as a standalone line without any label.
  // Scan for standalone subject code patterns anywhere in header lines.
  if (!subjectCode) {
    for (let i = 0; i < headerLines.length; i++) {
      const line = headerLines[i];
      // Match standalone "EC 5102 / CIRCUIT THEORY" or "EC5102 / CIRCUIT THEORY"
      const codeMatch = line.match(/^([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]\s*([A-Z][A-Za-z\s&(),]+)/i);
      if (codeMatch) {
        subjectCode = codeMatch[1].replace(/\s+/g, ' ').trim();
        subjectName = codeMatch[2].trim();
        console.log('[DOCX Parser] Found subject code via standalone scan:', subjectCode, '/', subjectName);
        break;
      }
    }
  }

  // Fallback: sweep the entire header block as one string
  const fullHeaderBlock = headerLines.join(' ');
  if (!subjectCode) {
    // Allow optional space in subject code (e.g. "EC 5102" or "EC5102")
    const m = fullHeaderBlock.match(/([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]\s*([A-Z][A-Za-z\s&]+?)(?:\s+Time|\s+Max|\s+Branch|$)/i);
    if (m) { subjectCode = m[1].replace(/\s+/g, ' ').trim(); if (!subjectName) subjectName = m[2].trim(); }
  }
  if (!date) {
    const m = fullHeaderBlock.match(/Date\s*[:.\\-]?\s*([\d./-]+)/i);
    if (m) date = m[1].trim();
  }
  if (!maxMarks) {
    const m = fullHeaderBlock.match(/Max\.?\s*Marks?\s*[:.\\-]?\s*(\d+(?:\s*Marks?)?)/i);
    if (m) maxMarks = m[1].trim();
  }
  if (!branch) {
    // Allow commas and more special chars in branch names
    const m = fullHeaderBlock.match(/Branch\s*[:.\\-]?\s*([A-Za-z.\s&(),/\\-]+?)(?:\s+Year|\s+Date|$)/i);
    if (m) branch = m[1].trim();
  }
  if (!yearSem) {
    const m = fullHeaderBlock.match(/Year\s*[/\\\\]?\s*Sem(?:ester)?\s*[:.\\-]?\s*([A-Za-z0-9\s/IVX]+?)(?:\s|$)/i);
    if (m) yearSem = m[1].trim();
  }
  if (!time) {
    const m = fullHeaderBlock.match(/Time\s*[:.\\-]?\s*(\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?))/i);
    if (m) time = m[1].trim();
  }

  // ========== ALWAYS scan mammoth HTML tables (row-by-row) ==========
  // Mammoth often truncates label text in merged cells (e.g. "Branch" → "anch",
  // "Date:" → "te:"). Instead of relying on labels, we scan for VALUE PATTERNS
  // in each table row and also match partial/truncated labels.
  {
    const tempParser = new DOMParser();
    const tempDoc = tempParser.parseFromString(htmlOutput, 'text/html');
    const allTables = tempDoc.querySelectorAll('table');

    // Helper: clean cell text
    const cleanCell = (el: Element): string =>
      (el.textContent || '').replace(/[\xA0\u200B\n\r\t]/g, ' ').replace(/\s+/g, ' ').trim();

    allTables.forEach(table => {
      const rows = table.querySelectorAll('tr');
      rows.forEach(row => {
        const cells = row.querySelectorAll('td, th');
        if (cells.length < 2) return; // Skip single-cell rows

        // Collect all cell texts for this row
        const cellTexts: string[] = [];
        cells.forEach(cell => cellTexts.push(cleanCell(cell)));
        const rowText = cellTexts.join(' ');

        console.log('[DOCX HTML Scan] Row cells:', cellTexts);

        // ---- DATE ----
        // Look for any cell containing a date pattern (dd-mm-yyyy, dd/mm/yyyy, dd.mm.yyyy)
        if (!date) {
          for (const ct of cellTexts) {
            const dm = ct.match(/(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/);
            if (dm) {
              date = dm[1];
              console.log('[DOCX HTML Scan] Found date:', date);
              break;
            }
          }
        }

        // ---- SUBJECT CODE / NAME ----
        // Look for cell containing "XX 1234 / SOME NAME" or "XX1234 / SOME NAME"
        if (!subjectCode) {
          for (const ct of cellTexts) {
            const cm = ct.match(/([A-Z]{2,}\s*\d{3,})\s*[/\\\\—–\-]\s*([A-Z][A-Za-z\s&(),]+)/);
            if (cm) {
              subjectCode = cm[1].replace(/\s+/g, ' ').trim();
              subjectName = cm[2].trim();
              console.log('[DOCX HTML Scan] Found subject:', subjectCode, '/', subjectName);
              break;
            }
          }
        }

        // ---- MAX MARKS ----
        if (!maxMarks) {
          // Check if any cell contains "Max" (even partial "ax. Marks" or "Marks")
          for (let c = 0; c < cellTexts.length; c++) {
            if (/Max|Marks/i.test(cellTexts[c])) {
              // Look for a number in this cell or adjacent cells
              const mm = cellTexts[c].match(/(\d{2,3})\s*(?:Marks?)?/i);
              if (mm) { maxMarks = mm[1].trim(); break; }
              // Check next cell for the number
              if (c + 1 < cellTexts.length) {
                const nm = cellTexts[c + 1].match(/^(\d{2,3})\s*(?:Marks?)?$/i);
                if (nm) { maxMarks = nm[1].trim(); break; }
              }
            }
          }
        }

        // ---- TIME ----
        if (!time) {
          for (const ct of cellTexts) {
            const tm = ct.match(/(\d+(?:\.\d+)?\s*(?:hrs?|hours?|mins?))/i);
            if (tm) {
              time = tm[1].trim();
              console.log('[DOCX HTML Scan] Found time:', time);
              break;
            }
          }
        }

        // ---- BRANCH ----
        // Mammoth truncates "Branch" to "ranch" or "anch". Look for partial label OR
        // look for the branch value pattern directly in the row.
        if (!branch) {
          // Strategy 1: Find a cell with partial "Branch" label (ranch, anch, Branch)
          for (let c = 0; c < cellTexts.length; c++) {
            if (/(?:^|B)?r?anch\b/i.test(cellTexts[c]) || /^Branch/i.test(cellTexts[c])) {
              // The value is in the NEXT cell in the same row
              if (c + 1 < cellTexts.length) {
                const val = cellTexts[c + 1];
                if (val.length > 1 && !/Year|Sem|Time|Max|Date|Subject/i.test(val)) {
                  branch = val;
                  console.log('[DOCX HTML Scan] Found branch (via label):', branch);
                }
              }
              // Or inline: "Branch: COMMON TO ECE..."
              const inline = cellTexts[c].match(/(?:B?ranch)\s*[:.\\-]?\s*(.+)/i);
              if (!branch && inline && inline[1].trim().length > 1) {
                branch = inline[1].trim();
                console.log('[DOCX HTML Scan] Found branch (inline):', branch);
              }
              break;
            }
          }
          // Strategy 2: If the row also contains "Year" or "Sem", there might be
          // a branch value that isn't near a label. Find non-label, non-date cells.
          if (!branch && /Year|Sem/i.test(rowText)) {
            for (let c = 0; c < cellTexts.length; c++) {
              const ct = cellTexts[c];
              // Skip cells that are labels or known other values
              if (/Year|Sem|Time|Max|Date|Subject|Code|Name|Marks|hrs|^\d/i.test(ct)) continue;
              if (ct.length < 3) continue;
              // Skip cells that look like year/sem values (Roman numerals / digits with /)
              if (/^[IVX\d]+\s*[/\\\\]\s*[IVX\d]+$/i.test(ct)) continue;
              // This could be a branch value
              if (ct.length > 2) {
                branch = ct;
                console.log('[DOCX HTML Scan] Found branch (by elimination):', branch);
                break;
              }
            }
          }
        }

        // ---- YEAR / SEM ----
        if (!yearSem) {
          // Strategy 1: Find a cell with "Year" or "Sem" label
          for (let c = 0; c < cellTexts.length; c++) {
            if (/Year\s*[/\\\\]?\s*Sem/i.test(cellTexts[c]) || /(?:^|\s)Sem(?:ester)?/i.test(cellTexts[c])) {
              // Check if value is inline
              const ysInline = cellTexts[c].match(/(?:Year\s*[/\\\\]?\s*)?Sem(?:ester)?\s*[:.\\-]?\s*(.+)/i);
              if (ysInline && ysInline[1].trim().length > 0) {
                let parsed = ysInline[1].trim();
                if (parsed !== ':' && parsed !== '-') yearSem = parsed;
              }
              // Or the value is in subsequent cells (skip empty/colon cells)
              if (!yearSem) {
                for (let nextC = c + 1; nextC < cellTexts.length; nextC++) {
                  const parsed = cellTexts[nextC].trim();
                  if (parsed.length > 0 && parsed !== ':' && parsed !== '-') {
                    yearSem = parsed;
                    break;
                  }
                }
              }
              if (yearSem) {
                console.log('[DOCX HTML Scan] Found yearSem (via label):', yearSem);
              }
              break;
            }
          }
          // Strategy 2: Find a cell that looks like year/sem value (Roman numerals / digits with /)
          if (!yearSem) {
            for (const ct of cellTexts) {
              if (/^[IVX]+\s*[/\\\\]\s*[IVX]+$/i.test(ct) || /^\d\s*[/\\\\]\s*\d$/.test(ct)) {
                yearSem = ct;
                console.log('[DOCX HTML Scan] Found yearSem (value pattern):', yearSem);
                break;
              }
            }
          }
        }
      });
    });
  }

  console.log('[DOCX Parser] Extracted metadata:', { subjectCode, subjectName, date, branch, yearSem, maxMarks, time });

  // Inject metadata into the HTML output with proper IDs so extractMetadata can find them.
  // We inject hidden divs with IDs that extractMetadata looks for (#code, #branchCell, #yearSemesterCell).
  // We also inject visible metadata if mammoth failed to render them correctly.
  let metadataInjection = '';

  if (subjectCode || subjectName) {
    metadataInjection += `<div id="code" style="display:none">Subject Code / Name: ${subjectCode}${subjectName ? ' / ' + subjectName : ''}</div>`;
  }
  if (branch) {
    metadataInjection += `<div id="branchCell" style="display:none">Branch: ${branch}</div>`;
  }
  if (yearSem) {
    metadataInjection += `<div id="yearSemesterCell" style="display:none">Year / Sem: ${yearSem}</div>`;
  }
  if (date) {
    metadataInjection += `<div id="dateCell" style="display:none">Date: ${date}</div>`;
  }
  if (maxMarks) {
    metadataInjection += `<div id="maxMarksCell" style="display:none">Max. Marks: ${maxMarks}</div>`;
  }
  if (time) {
    metadataInjection += `<div id="timeCell" style="display:none">Time: ${time}</div>`;
  }

  // Prepend metadata divs to the HTML output
  return metadataInjection + htmlOutput;
};

// ============================================================
// PDF Parsing: Text-flattening + Sequential Question Detection
// ============================================================

interface ExtractedPdfImage {
  dataUrl: string;
  x: number;      // X position on page (PDF coords)
  y: number;      // Y position on page (PDF coords, bottom-up)
  width: number;   // Width in PDF points
  height: number;  // Height in PDF points
}

const extractImagesFromPdfPage = async (page: any): Promise<ExtractedPdfImage[]> => {
  try {
    const images: ExtractedPdfImage[] = [];
    const operatorList = await page.getOperatorList();

    if (!operatorList) return images;

    // pdf.js OPS constants for image operations
    // paintImageXObject = 85, paintJpegXObject = 84, paintInlineImageXObject = 87
    const OPS_paintImageXObject = 85;
    const OPS_paintJpegXObject = 84;
    const OPS_paintInlineImageXObject = 87;
    const OPS_transform = 12; // setTransform / transform
    const OPS_save = 10;      // save graphics state
    const OPS_restore = 11;   // restore graphics state

    // Track current transform matrix to get image positions
    // PDF transform matrix: [a, b, c, d, e, f] where e=x, f=y, a=scaleX, d=scaleY
    const transformStack: number[][] = [[1, 0, 0, 1, 0, 0]];
    let currentTransform = [1, 0, 0, 1, 0, 0];

    for (let i = 0; i < operatorList.fnArray.length; i++) {
      const fn = operatorList.fnArray[i];
      const args = operatorList.argsArray[i];

      // Track transform matrix for image positioning
      if (fn === OPS_save) {
        transformStack.push([...currentTransform]);
      } else if (fn === OPS_restore) {
        if (transformStack.length > 1) {
          currentTransform = transformStack.pop()!;
        }
      } else if (fn === OPS_transform && args && args.length >= 6) {
        // Multiply current transform with new transform
        const [a, b, c, d, e, f] = args;
        const ct = currentTransform;
        currentTransform = [
          ct[0] * a + ct[2] * b,
          ct[1] * a + ct[3] * b,
          ct[0] * c + ct[2] * d,
          ct[1] * c + ct[3] * d,
          ct[0] * e + ct[2] * f + ct[4],
          ct[1] * e + ct[3] * f + ct[5]
        ];
      }

      // Check for image paint operations
      if (fn === OPS_paintImageXObject || fn === OPS_paintJpegXObject) {
        const imageName = args[0];
        if (!imageName) continue;

        try {
          // Get the image object from the page
          const imgData = await new Promise<any>((resolve) => {
            page.objs.get(imageName, (data: any) => {
              resolve(data);
            });
          });

          if (!imgData) continue;

          // Get image dimensions from the transform matrix
          const imgWidth = Math.abs(currentTransform[0]);  // Scale X = width in PDF points
          const imgHeight = Math.abs(currentTransform[3]); // Scale Y = height in PDF points
          const imgX = currentTransform[4];                // X position
          const imgY = currentTransform[5];                // Y position (PDF bottom-up)

          // Skip tiny images (logos, icons, decorations) — only keep diagrams
          // Typical diagrams are at least 50x50 PDF points
          if (imgWidth < 40 || imgHeight < 40) {
            console.log(`[PDF-IMG] Skipping small image "${imageName}": ${imgWidth.toFixed(0)}x${imgHeight.toFixed(0)}pts`);
            continue;
          }

          // Render the image to a canvas to get a data URL
          const canvas = document.createElement('canvas');

          if (imgData.bitmap) {
            // ImageBitmap — draw directly
            canvas.width = imgData.bitmap.width;
            canvas.height = imgData.bitmap.height;
            const ctx = canvas.getContext('2d');
            if (ctx) {
              ctx.drawImage(imgData.bitmap, 0, 0);
            }
          } else if (imgData.data && imgData.width && imgData.height) {
            // Raw pixel data (RGBA or RGB)
            canvas.width = imgData.width;
            canvas.height = imgData.height;
            const ctx = canvas.getContext('2d');
            if (ctx) {
              const imageData = ctx.createImageData(imgData.width, imgData.height);

              if (imgData.data.length === imgData.width * imgData.height * 4) {
                // RGBA data
                imageData.data.set(imgData.data);
              } else if (imgData.data.length === imgData.width * imgData.height * 3) {
                // RGB data — convert to RGBA
                for (let px = 0; px < imgData.width * imgData.height; px++) {
                  imageData.data[px * 4] = imgData.data[px * 3];
                  imageData.data[px * 4 + 1] = imgData.data[px * 3 + 1];
                  imageData.data[px * 4 + 2] = imgData.data[px * 3 + 2];
                  imageData.data[px * 4 + 3] = 255;
                }
              } else {
                // Unknown format, try treating as RGBA
                imageData.data.set(imgData.data.slice(0, imageData.data.length));
              }

              ctx.putImageData(imageData, 0, 0);
            }
          } else if (imgData.src) {
            // JPEG/PNG source URL
            await new Promise<void>((resolve) => {
              const tempImg = new Image();
              tempImg.onload = () => {
                canvas.width = tempImg.width;
                canvas.height = tempImg.height;
                const ctx = canvas.getContext('2d');
                ctx?.drawImage(tempImg, 0, 0);
                resolve();
              };
              tempImg.onerror = () => resolve();
              tempImg.src = imgData.src;
            });
          } else {
            console.log(`[PDF-IMG] Unknown image data format for "${imageName}":`, Object.keys(imgData));
            continue;
          }

          if (canvas.width > 10 && canvas.height > 10) {
            const dataUrl = canvas.toDataURL('image/png');
            images.push({
              dataUrl,
              x: imgX,
              y: imgY,
              width: imgWidth,
              height: imgHeight
            });
            console.log(`[PDF-IMG] Extracted image "${imageName}": ${canvas.width}x${canvas.height}px, pos=(${imgX.toFixed(1)}, ${imgY.toFixed(1)}), size=${imgWidth.toFixed(0)}x${imgHeight.toFixed(0)}pts`);
          }
        } catch (err) {
          console.warn(`[PDF-IMG] Failed to extract image "${imageName}":`, err);
        }
      }
    }

    console.log(`[PDF-IMG] Total extracted images from page: ${images.length}`);
    return images;
  } catch (e) {
    console.warn('[PDF-IMG] Error extracting embedded images:', e);
    return [];
  }
};

const renderPdfPageToImage = async (page: any, scale: number = 2.0): Promise<string | null> => {
  try {
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;

    await page.render({ canvasContext: ctx, viewport }).promise;
    const dataUrl = canvas.toDataURL('image/png');
    console.log(`[PDF] Rendered page to image: ${canvas.width}x${canvas.height}, data length: ${dataUrl.length}`);
    return dataUrl;
  } catch (err) {
    console.error('[PDF] Failed to render page to image:', err);
    return null;
  }
};

const parsePdf = async (file: File): Promise<ParseFileResult> => {
  try {
    console.log('Starting PDF parsing for:', file.name);
    const arrayBuffer = await file.arrayBuffer();

    // Ensure worker is set
    if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;
    }

    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    console.log('PDF Loaded, pages:', pdf.numPages);

    // Track metadata for every parsed line across all pages
    const globalLineMetadata: { page: number; y: number; height: number; }[] = [];

    // Extract text AND render images from all pages
    const pageTexts: string[] = [];
    const pageImages: (string | null)[] = [];
    const allExtractedImages: ExtractedPdfImage[] = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);

      // 1. Extract text content (existing logic)
      const textContent = await page.getTextContent();

      // Get items with position
      const items = textContent.items.map((item: any) => ({
        str: item.str as string,
        x: item.transform[4] as number,
        y: item.transform[5] as number,
      }));

      // ============================================================
      // FIX: Column-aware Y-snapping for table Q.No items
      // ============================================================
      // Problem: In PDF tables, Q.No column items ("11", "A", "12", etc.) are often
      // vertically centered in their cell, while question text starts from the top.
      // This means Q.No items have LOWER Y than the first line of question text.
      // When sorted by Y desc, the first question text line appears BEFORE the Q.No,
      // causing it to be wrongly attributed to the previous question.
      //
      // Fix: Detect Q.No column items, group numbers with their suffixes, and snap
      // each group's Y to the highest Y of text in the same table row.
      // ============================================================

      // Step 1: Detect the Q.No column X-range using standalone numbers
      // These numbers ("1", "2", "11", "12", "13") reliably appear in the Q.No column
      const ROW_HEIGHT_THRESHOLD = 30; // Max vertical distance within one table row

      const numberItems: typeof items = [];
      for (const item of items) {
        const s = item.str.trim();
        if (/^\d{1,2}$/.test(s)) {
          numberItems.push(item);
        }
      }

      // Find Q.No column X boundary from standalone number positions
      let qnoColumnMaxX = 0;
      if (numberItems.length >= 3) {
        // Sort by X, find the leftmost cluster
        const sortedByX = [...numberItems].sort((a, b) => a.x - b.x);
        // Use the 25th percentile X + margin to find the Q.No column boundary
        const q25Idx = Math.floor(sortedByX.length * 0.25);
        const q25X = sortedByX[q25Idx].x;
        qnoColumnMaxX = q25X + 50;
        console.log(`[PDF] Q.No column detected: maxX=${qnoColumnMaxX.toFixed(0)} (based on ${numberItems.length} number items, q25X=${q25X.toFixed(0)})`);
      }

      // Detect OR separator Y positions EARLY — used as upper-bound clamps in Y-snapping.
      // Or separators in a Part B table appear as a standalone "OR" text between question rows.
      // A Q.No label that appears BELOW an OR separator must not snap above that OR's Y,
      // otherwise question text from above the OR (previous question) would be stolen.
      const orItemYs: number[] = [];
      for (const item of items) {
        if (/^\s*OR\s*$/i.test(item.str)) {
          orItemYs.push(item.y);
          console.log(`[PDF] OR separator detected early at y=${item.y.toFixed(1)}, x=${item.x.toFixed(1)}`);
        }
      }

      // Step 2: Group Q.No number items with their nearby suffix letters
      // e.g., "11" at (30, 355) and "B" at (45, 354) form one Q.No group
      if (qnoColumnMaxX > 0) {
        // Find Q.No numbers that are actually in the Q.No column
        const qnoNumbers = numberItems.filter(it => it.x <= qnoColumnMaxX);

        // Find suffix letters in Q.No column
        const suffixItems = items.filter(it => {
          const s = it.str.trim();
          return it.x <= qnoColumnMaxX && /^[ABab]$/.test(s);
        });

        // Build groups: each number plus any suffix within 8px Y and 50px X
        interface QNoGroup {
          members: typeof items;
          baseY: number;
        }
        const groups: QNoGroup[] = [];
        const usedSuffixes = new Set<any>();

        for (const numItem of qnoNumbers) {
          const members = [numItem];
          for (const suf of suffixItems) {
            if (usedSuffixes.has(suf)) continue;
            if (Math.abs(suf.y - numItem.y) < 8 && Math.abs(suf.x - numItem.x) < 50) {
              members.push(suf);
              usedSuffixes.add(suf);
            }
          }
          const baseY = Math.max(...members.map(m => m.y));
          groups.push({ members, baseY });
        }

        // Sort groups by baseY desc (top of page first)
        groups.sort((a, b) => b.baseY - a.baseY);

        // Store original baseY values before snapping begins, so we can use them
        // as upper bounds (the pre-snap Y of the group directly above defines the
        // ceiling for each group's text search range).
        const originalBaseYs = groups.map(g2 => g2.baseY);

        // Step 3: For each group, snap Y to the highest Y of text in the same row.
        // The upper bound is the ORIGINAL Y of the DIRECTLY PRECEDING Q.No group (groups[g-1]),
        // i.e., the Q.No row that is immediately above this one in the document.
        // This ensures that even very tall rows (e.g. 16-mark questions where the Q.No
        // label is centered far below the first text line) get correctly snapped.
        // Without this fix, the first line(s) of a question would appear ABOVE the
        // Q.No label's Y-snapped position and be wrongly appended to the previous question.
        for (let g = 0; g < groups.length; g++) {
          const group = groups[g];

          // Upper bound = original Y of the group directly above this one (groups[g-1]).
          // groups[] is sorted desc (highest Y first = topmost on page first),
          // so groups[g-1] is immediately above group[g] in the document.
          // We use the ORIGINAL baseY (pre-snap) to avoid cross-row text stealing.
          // If this is the first group (no group above), use a large offset to allow
          // reaching the page top.
          let upperBound: number;
          if (g === 0) {
            // Topmost Q.No group: allow searching up to 300px above (covers header rows)
            upperBound = group.baseY + 300;
          } else {
            // Use the ORIGINAL pre-snap Y of the group directly above as the ceiling.
            // This covers the full inter-row gap: from this group's centered Q.No label
            // up to (but not including) the row above's Q.No cell center.
            upperBound = originalBaseYs[g - 1];
          }

          // CRITICAL OR-BOUNDARY CLAMP:
          // If an OR separator lies between this group's baseY and the upper bound,
          // clamp the upper bound to just below the OR separator's Y.
          // This prevents Q.No groups below an OR divider from snapping their label
          // above the OR line, which would steal question text from the row above.
          // Example: Q12A's label is below the OR between Q11A/Q11B. Without this clamp,
          // Q12A would snap up into Q11B's text territory and absorb "Compute the even..."
          // into what looks like Q11B content in the flattened text.
          if (orItemYs.length > 0) {
            // Find any OR separator that is between group.baseY and upperBound
            for (const orY of orItemYs) {
              if (orY > group.baseY && orY < upperBound) {
                // This OR separator divides this group from the question above.
                // Clamp upper bound to just below the OR Y so we don't snap above it.
                const clampedBound = orY - 1;
                if (clampedBound < upperBound) {
                  console.log(`[PDF] Y-snap upper bound clamped to OR separator: orY=${orY.toFixed(1)}, old bound=${upperBound.toFixed(1)}, new bound=${clampedBound.toFixed(1)} for group baseY=${group.baseY.toFixed(1)}`);
                  upperBound = clampedBound;
                }
              }
            }
          }

          // Find the highest Y among text column items above Q.No but within bounds
          let snapY = group.baseY;
          for (const other of items) {
            if (group.members.includes(other)) continue;
            if (other.x <= qnoColumnMaxX) continue;
            if (other.y <= group.baseY) continue;
            if (other.y >= upperBound) continue;
            if (!other.str.trim()) continue;
            if (other.y > snapY) {
              snapY = other.y;
            }
          }

          // Apply snap to all members of the group
          if (snapY > group.baseY) {
            for (const member of group.members) {
              console.log(`[PDF] Y-snap: "${member.str.trim()}" at (x=${member.x.toFixed(0)}, y=${member.y.toFixed(0)}) → y=${snapY.toFixed(0)}`);
              member.y = snapY;
            }
            // Update baseY to the snapped value for reference (qnoRowYs uses member.y directly)
            group.baseY = snapY;
          }
        }
      }

      // ============================================================
      // ROW ASSIGNMENT: Tag each text item with the Q.No row it belongs to.
      // After Y-snapping, each Q.No group sits at the Y of the FIRST (topmost)
      // line of its question row. A text item belongs to a Q.No row if its Y
      // is at or below that Q.No's Y but above the next Q.No's Y below it.
      // During line grouping, items from different Q.No rows are forced onto
      // separate lines even if their Y-coordinates are within 5px.
      // This is critical for preventing text from adjacent question rows
      // (e.g., Q11A last line and Q11B first line) from merging.
      //
      // ADDITIONAL FIX: OR separator items act as row boundaries.
      // In an OR-paired question table, text between OR and the next Q.No label
      // physically belongs to the next question (e.g. Q12A not Q11B).
      // We detect OR item Y-positions and insert them into qnoRowYs as virtual
      // boundaries so that getRowId correctly assigns such text to the next row.
      // ============================================================

      // Collect Q.No group Y positions after snapping (sorted desc = top to bottom)
      const qnoRowYs: number[] = [];
      if (qnoColumnMaxX > 0) {
        const qnoItems = items.filter(it => {
          const s = it.str.trim();
          return it.x <= qnoColumnMaxX && /^\d{1,2}$/.test(s);
        });
        // Use unique Y values (multiple Q.No items may share the same Y after snapping)
        const ySet = new Set<number>();
        for (const qi of qnoItems) {
          ySet.add(qi.y);
        }
        // Also insert OR separator Ys as virtual row boundaries.
        // Text that appears in the PDF between an OR divider and the next Q.No label
        // belongs to the NEXT question. By inserting OR Ys as extra boundary rows,
        // getRowId will assign such text to a distinct rowId, breaking line grouping
        // at the OR boundary and preventing cross-question text bleeding.
        for (const orY of orItemYs) {
          // Only insert if this OR Y is not already near a Q.No Y (avoid duplicates)
          const alreadyNear = Array.from(ySet).some(qy => Math.abs(qy - orY) < 5);
          if (!alreadyNear) {
            ySet.add(orY);
            console.log(`[PDF] Inserted OR boundary y=${orY.toFixed(1)} into qnoRowYs`);
          }
        }
        qnoRowYs.push(...Array.from(ySet).sort((a, b) => b - a)); // Sorted desc (top first)
        console.log(`[PDF] Q.No row Y positions (${qnoRowYs.length}): ${qnoRowYs.map(y => y.toFixed(1)).join(', ')}`);
      }

      // Assign each item a rowId based on the Q.No group it belongs to.
      // rowId = index in qnoRowYs of the nearest Q.No at or above the item's Y.
      // Items above all Q.No groups get rowId = -1 (header area).
      // Items below all Q.No groups get the last rowId.
      const itemRowIds = new Map<any, number>();

      const getRowId = (itemY: number): number => {
        if (qnoRowYs.length === 0) return -1;
        // If item is above all Q.No rows, it's in the header area
        if (itemY > qnoRowYs[0] + 2) return -1;

        // Find the Q.No row this item belongs to.
        // Item belongs to row r if: qnoRowYs[r] >= itemY > qnoRowYs[r+1]
        // (at or below row r's Y, but above row r+1's Y)
        for (let r = 0; r < qnoRowYs.length; r++) {
          const rowY = qnoRowYs[r];
          const nextRowY = (r + 1 < qnoRowYs.length) ? qnoRowYs[r + 1] : -Infinity;

          // Item is at or below this row's top, and above the next row's top
          if (itemY <= rowY + 2 && itemY > nextRowY) {
            return r;
          }
        }
        // Fallback: assign to last row
        return qnoRowYs.length - 1;
      };

      // Tag all items with their row IDs
      for (const item of items) {
        itemRowIds.set(item, getRowId(item.y));
      }

      // Sort: top to bottom (Y desc), then left to right (X asc)
      items.sort((a: any, b: any) => {
        if (Math.abs(a.y - b.y) < 5) return a.x - b.x;
        return b.y - a.y;
      });

      // Group into lines by Y-coordinate
      // CRITICAL: Force line breaks when items belong to different Q.No rows,
      // even if their Y-coordinates are within 5px of each other.
      const lines: string[] = [];
      const lineYs: number[] = [];
      const lineHeights: number[] = [];
      let currentLine: typeof items = [];

      if (items.length > 0) {
        currentLine.push(items[0]);
        for (let j = 1; j < items.length; j++) {
          const yDiff = Math.abs(items[j].y - currentLine[0].y);
          const sameRow = itemRowIds.get(items[j]) === itemRowIds.get(currentLine[0]);
          const sameLine = yDiff < 5 && sameRow;

          if (sameLine) {
            currentLine.push(items[j]);
          } else {
            currentLine.sort((a: any, b: any) => a.x - b.x);
            lines.push(currentLine.map((it: any) => it.str).join(' '));
            lineYs.push(currentLine[0].y);
            lineHeights.push(Math.max(...currentLine.map((it: any) => it.height || 10)));
            currentLine = [items[j]];
          }
        }
        if (currentLine.length > 0) {
          currentLine.sort((a: any, b: any) => a.x - b.x);
          lines.push(currentLine.map((it: any) => it.str).join(' '));
          lineYs.push(currentLine[0].y);
          lineHeights.push(Math.max(...currentLine.map((it: any) => it.height || 10)));
        }
      }

      for (let l = 0; l < lines.length; l++) {
        globalLineMetadata.push({ page: i - 1, y: lineYs[l], height: lineHeights[l] });
      }

      pageTexts.push(lines.join('\n'));

      // 2. Render page to image (captures diagrams, figures, math, everything visual)
      // This is used as a fallback for vector-drawn diagrams that can't be extracted as images
      const pageImage = await renderPdfPageToImage(page, 2.0);
      pageImages.push(pageImage);

      // 3. Extract actual embedded image objects from the PDF page
      // These are the clean diagram images (not screenshots of the page)
      const extractedImages = await extractImagesFromPdfPage(page);
      // Tag each image with its page number
      extractedImages.forEach(img => (img as any).pageNum = i - 1);
      allExtractedImages.push(...extractedImages);

      console.log(`[PDF] Page ${i}: ${lines.length} text lines, rendered image: ${pageImage ? 'yes' : 'no'}, embedded images: ${extractedImages.length}`);
    }

    const fullText = pageTexts.join('\n');
    console.log('Full extracted text length:', fullText.length);
    console.log(`[PDF] Total page images: ${pageImages.filter(Boolean).length}/${pageImages.length}`);
    console.log(`[PDF] Total extracted embedded images: ${allExtractedImages.length}`);

    const html = await convertTextToHtml(fullText, pageImages, pageTexts, globalLineMetadata, allExtractedImages);
    return { html, pdfPageImages: pageImages };
  } catch (error) {
    console.error('PDF Parsing Error:', error);
    throw error;
  }
};

const cropFromPdfDataUrl = async (dataUrl: string, pdfYTop: number, pdfYBottom: number): Promise<string> => {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const scale = 2.0;
      const pdfHeight = img.height / scale;

      // PDF Y is bottom-up (0 is bottom). Canvas Y is top-down (0 is top).
      // Math: canvasY = (pdfHeight - pdfY) * scale
      let startY = (pdfHeight - pdfYTop) * scale;
      let endY = (pdfHeight - pdfYBottom) * scale;

      // Add a small padding (10 pixels)
      startY = Math.max(0, startY - 10);
      endY = Math.min(img.height, endY + 10);

      const cropHeight = Math.max(1, endY - startY);

      // Compute horizontal limits to cleanly slice out adjacent table columns/borders
      // The diagram column safely lives between ~12% to 82% of the A4 page width
      const hMarginLeft = img.width * 0.12;
      const hMarginRight = img.width * 0.18;
      const cropWidth = img.width - hMarginLeft - hMarginRight;

      const canvas = document.createElement('canvas');
      canvas.width = cropWidth;
      canvas.height = cropHeight;
      const ctx = canvas.getContext('2d');
      // drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh)
      ctx?.drawImage(img, hMarginLeft, startY, cropWidth, cropHeight, 0, 0, cropWidth, cropHeight);
      resolve(canvas.toDataURL('image/png'));
    };
    img.onerror = () => resolve('');
    img.src = dataUrl;
  });
};

const convertTextToHtml = async (
  fullText: string,
  pageImages?: (string | null)[],
  pageTexts?: string[],
  globalLineMetadata?: { page: number; y: number; height: number; }[],
  extractedImages?: ExtractedPdfImage[]
): Promise<string> => {
  let html = '<div class="pdf-content">';

  // Extract metadata from header area using line-by-line search
  // PDF table cells may be on the same or different Y-coordinates,
  // so labels and values could be on the same line or adjacent lines
  let subject = '';
  let subjectCode = '';
  let date = '';
  let maxMarks = '';
  let time = '';
  let branch = '';
  let yearSem = '';

  const headerLines = fullText.split('\n').slice(0, 30); // Only search header area

  for (let i = 0; i < headerLines.length; i++) {
    const line = headerLines[i].trim();
    const nextLine = (i + 1 < headerLines.length) ? headerLines[i + 1].trim() : '';

    // Subject Code / Name
    if (!subjectCode && /Subject\s*Code/i.test(line)) {
      // Same line: "Subject Code / Name EC4411 / SIGNAL PROCESSING TECHNIQUES Time..."
      const sameLine = line.match(/Subject\s*Code\s*[/\\\\]\s*Name\s*[:.]?\s*([A-Z]+\d+)\s*[/\\\\]\s*(.+?)(?:\s+Time|\s+Max|$)/i);
      if (sameLine) {
        subjectCode = sameLine[1].trim();
        subject = sameLine[2].trim();
      } else {
        // Next line: "EC4411 / SIGNAL PROCESSING TECHNIQUES"
        const nextMatch = nextLine.match(/^([A-Z]+\d+)\s*[/\\\\]\s*(.+)/i);
        if (nextMatch) {
          subjectCode = nextMatch[1].trim();
          subject = nextMatch[2].trim().replace(/\s+Time\s+.*$/i, '').replace(/\s+Max\s+.*$/i, '').trim();
        }
      }
    }

    // Date
    if (!date && /Date/i.test(line)) {
      const dateMatch = line.match(/Date\s*[:.]?\s*([\d./-]+)/i);
      if (dateMatch) {
        date = dateMatch[1].trim();
      } else if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(nextLine)) {
        date = nextLine;
      }
    }

    // Max Marks
    if (!maxMarks && /Max/i.test(line)) {
      const mMatch = line.match(/Max\.?\s*Marks?\s*[:.]?\s*(\d+(?:\s*Marks?)?)/i);
      if (mMatch) maxMarks = mMatch[1].trim();
    }

    // Time
    if (!time && /Time/i.test(line) && !/Time\s*(?:Variant|Invariant|domain|signal|Shifting)/i.test(line)) {
      const tMatch = line.match(/Time\s*[:.]?\s*(\d+\s*(?:hrs?|hours?|mins?))/i);
      if (tMatch) {
        time = tMatch[1].trim();
      } else if (nextLine && /^\d+\s*(?:hrs?|hours?|mins?)/i.test(nextLine)) {
        time = nextLine.match(/^\d+\s*(?:hrs?|hours?|mins?)/i)![0];
      }
    }

    // Branch
    if (!branch && /Branch/i.test(line)) {
      // Branch: B.E - EC(ACT) & EC(VLSI)
      const brMatch = line.match(/Branch\s*[:.]?\s*([A-Za-z.\s&()-]+?)(?:\s+Year|$)/i);
      if (brMatch && brMatch[1].trim().length > 1) {
        branch = brMatch[1].trim();
      } else if (nextLine && !/Year|Sem|Date|Subject|Q\.?\s*No/i.test(nextLine) && nextLine.length > 1) {
        branch = nextLine.replace(/\s+Year\s*[/]?\s*Sem.*$/i, '').trim();
      }
    }

    // Year / Sem
    if (!yearSem && /Year\s*[/\\\\]?\s*Sem/i.test(line)) {
      const ysMatch = line.match(/Year\s*[/\\\\]?\s*Sem(?:ester)?\s*[:.]?\s*(.+?)$/i);
      if (ysMatch && ysMatch[1].trim().length > 0) {
        yearSem = ysMatch[1].trim();
      } else if (nextLine && /^[IVX1-4IVX]+\s*[/\\\\]\s*[IVX1-8IVX]+/i.test(nextLine)) {
        yearSem = nextLine.trim();
      }
    }
  }

  // Fallback sweeping extraction across the whole header text block
  const fullHeaderBlock = headerLines.join(' ');
  if (!subjectCode) {
    const m = fullHeaderBlock.match(/(?:Subject\s*Code\s*[/\\\\]\s*Name\s*[:.]?\s*)?([A-Z]{2,}\d{3,})\s*[/\\\\]\s*([A-Z][A-Za-z\s&]+?)(?:\s+Time|\s+Max|\s+Branch|$)/i);
    if (m) { subjectCode = m[1].trim(); if (!subject) subject = m[2].trim(); }
  }
  if (!date) {
    const m = fullHeaderBlock.match(/Date\s*[:.]?\s*([\d./-]+)/i);
    if (m) date = m[1].trim();
  }
  if (!maxMarks) {
    const m = fullHeaderBlock.match(/Max\.?\s*Marks?\s*[:.]?\s*(\d+(?:\s*Marks?)?)/i);
    if (m) maxMarks = m[1].trim();
  }
  if (!branch) {
    const m = fullHeaderBlock.match(/Branch\s*[:.]?\s*([A-Za-z.\s&()-]+?)(?:\s+Year|\s+Date|$)/i);
    if (m) branch = m[1].trim();
  }
  if (!yearSem) {
    const m = fullHeaderBlock.match(/Year\s*[/\\\\]?\s*Sem(?:ester)?\s*[:.]?\s*([A-Za-z0-9\s/IVX]+?)(?:\s|$)/i);
    if (m) yearSem = m[1].trim();
  }
  if (!time) {
    const m = fullHeaderBlock.match(/Time\s*[:.]?\s*(\d+\s*(?:hrs?|hours?|mins?))/i);
    if (m) time = m[1].trim();
  }



  // Detect IA type from text using robust roman numeral and word spacing detection
  let iaTypeText = 'Internal Assessment';

  // Collapse whitespace so the regex can match across newlines in the PDF parse output
  const cleanText = fullText.replace(/\s+/g, ' ');

  if (/Internal.{1,30}?\b(?:3|III|I\s*I\s*I|l\s*l\s*l)\b/i.test(cleanText) ||
    /\bIA\s*[-–—]?\s*(?:3|III|I\s*I\s*I|l\s*l\s*l)\b/i.test(cleanText) ||
    /Model\s+Exam(?:ination)?/i.test(cleanText)) {
    iaTypeText = 'Internal Assessment 3';
  } else if (/Internal.{1,30}?\b(?:2|II|I\s*I|l\s*l)\b(?!\s*(?:I|l))/i.test(cleanText) ||
    /\bIA\s*[-–—]?\s*(?:2|II|I\s*I|l\s*l)\b(?!\s*(?:I|l))/i.test(cleanText)) {
    iaTypeText = 'Internal Assessment 2';
  } else if (/Internal.{1,30}?\b(?:1|I|l)\b(?!\s*(?:[IV]|I|l))/i.test(cleanText) ||
    /\bIA\s*[-–—]?\s*(?:1|I|l)\b(?!\s*(?:[IV]|I|l))/i.test(cleanText) ||
    /\bEPC\b/i.test(cleanText)) {
    iaTypeText = 'Internal Assessment 1';
  }

  // Debug: dump first 15 header lines to console
  console.log('[PDF] --- Header lines (0-14) ---');
  for (let d = 0; d < Math.min(15, headerLines.length); d++) {
    console.log(`[PDF] Header ${d}: "${headerLines[d].trim().substring(0, 120)}"`);
  }
  console.log('[PDF] Extracted metadata:', { subject, subjectCode, date, maxMarks, time, branch, yearSem, iaTypeText });

  // Header table with proper IDs for metadata extraction
  html += `
    <table style="width: 100%; border-collapse: collapse; border: 1px solid black; margin-bottom: 20px;">
      <tr>
        <td colspan="4" style="text-align: center; border: 1px solid black; padding: 10px;">
          <h2 style="margin: 0;">CHENNAI INSTITUTE OF TECHNOLOGY</h2>
          <p style="margin: 5px 0;">Autonomous</p>
          <p style="margin: 2px 0;">Sarathy Nagar, Kundrathur, Chennai – 600 069.</p>
          <p style="margin: 5px 0;" id="assTypeCell"><strong>${iaTypeText}</strong></p>
        </td>
      </tr>
      <tr>
        <td style="border: 1px solid black; padding: 5px;">Date: ${date}</td>
        <td style="border: 1px solid black; padding: 5px;" id="code">Subject Code / Name: ${subjectCode}${subject ? ' / ' + subject : ''}</td>
        <td style="border: 1px solid black; padding: 5px;">Max. Marks: ${maxMarks || '50 Marks'}</td>
        <td style="border: 1px solid black; padding: 5px;">Time: ${time || '1.30 hrs'}</td>
      </tr>
      <tr>
        <td style="border: 1px solid black; padding: 5px;" colspan="2" id="branchCell">Branch: ${branch}</td>
        <td style="border: 1px solid black; padding: 5px;" colspan="2" id="yearSemesterCell">Year / Sem: ${yearSem}</td>
      </tr>
    </table>
  `;

  // ---- Extract Course Objectives ----
  const courseObjectives: string[] = [];
  const objSectionMatch = fullText.match(/Course\s+Objective[s]?\s*[:\-–]?\s*([\s\S]*?)(?=Course\s+Outcome|On\s+[Cc]ompletion|CO\s*\.?\s*NO|PART|Q\.?\s*No|$)/i);
  if (objSectionMatch) {
    const objText = objSectionMatch[1].trim();

    // Strategy 1: Try to find numbered items like "1 To introduce..." or "1. To study..."
    // This handles PDF text where numbers and text are on separate lines or same line
    const numberedItemRegex = /(\d+)\s*[.)\s]\s*(To\s+[A-Za-z][\s\S]*?)(?=\d+\s*[.)\s]\s*To\s+|$)/gi;
    let objMatch;
    while ((objMatch = numberedItemRegex.exec(objText)) !== null) {
      const cleaned = objMatch[2].trim().replace(/\s+/g, ' ').trim();
      // Remove trailing RBT level patterns like "L2" or "L3" 
      const withoutLevel = cleaned.replace(/\s+L\d\s*$/i, '').trim();
      if (withoutLevel.length > 10 && !courseObjectives.includes(withoutLevel)) {
        courseObjectives.push(withoutLevel);
      }
    }

    // Strategy 2: If numbered items didn't work, try splitting on line breaks and filtering
    if (courseObjectives.length < 2) {
      courseObjectives.length = 0;
      const lines = objText.split(/\n/);
      lines.forEach(line => {
        let cleaned = line.trim().replace(/\s+/g, ' ').trim();
        // Remove leading numbers like "1 ", "1. ", "1) "
        cleaned = cleaned.replace(/^\d+\s*[.)]\s*/, '').trim();
        // Remove CO.NO header, column headers
        if (/^(CO\.?\s*NO|Questions?|RBT|Marks?|Q\.?\s*No)/i.test(cleaned)) return;
        // Remove trailing RBT levels
        cleaned = cleaned.replace(/\s+L\d\s*$/i, '').trim();
        if (cleaned.length > 15 && !courseObjectives.includes(cleaned)) {
          courseObjectives.push(cleaned);
        }
      });
    }

    // Strategy 3: Try splitting on "N. " or "N " followed by "To " pattern within continuous text
    if (courseObjectives.length < 2) {
      courseObjectives.length = 0;
      const splitItems = objText.split(/(?=\d+\s*[.)\s]\s*To\s)/i);
      splitItems.forEach(item => {
        let cleaned = item.trim().replace(/\s+/g, ' ').trim();
        cleaned = cleaned.replace(/^\d+\s*[.)]\s*/, '').trim();
        cleaned = cleaned.replace(/\s+L\d\s*$/i, '').trim();
        if (cleaned.length > 15 && !courseObjectives.includes(cleaned)) {
          courseObjectives.push(cleaned);
        }
      });
    }
  }

  // ---- Extract Course Outcomes ----
  const courseOutcomes: { co: string; description: string; level?: string }[] = [];

  // Strategy 1: Look for CO section and extract numbered items with RBT levels
  const coSectionMatch = fullText.match(/(?:Course\s+Outcome|CO\s*Statements?|On\s+[Cc]ompletion\s+of\s+the\s+course)[:\s]*([\s\S]*?)(?=PART|Q\.?\s*No|Part\s*[-–—]?\s*A|$)/i);
  const coSearchText = coSectionMatch ? coSectionMatch[1] : '';

  if (coSearchText) {
    // Try numbered format: "1 Summarize the classification... L2" or "CO1 description L3"
    // Common in PDF tables where columns get flattened
    const numberedCORegex = /(\d+)\s+([A-Z][a-zA-Z\s,.'()\-–&]+?)(?:\s+(L[1-6]))?(?=\s*\d+\s+[A-Z]|\s*$)/g;
    let coMatch;
    while ((coMatch = numberedCORegex.exec(coSearchText)) !== null) {
      const coNum = coMatch[1];
      const desc = coMatch[2].trim().replace(/\s+/g, ' ').trim();
      const level = coMatch[3]?.toUpperCase() || '';
      if (desc.length > 10) {
        const coKey = `CO${coNum}`;
        if (!courseOutcomes.find(c => c.co === coKey)) {
          courseOutcomes.push({ co: coKey, description: desc, level });
        }
      }
    }

    // Strategy 2: Line-by-line extraction from CO section
    if (courseOutcomes.length < 2) {
      courseOutcomes.length = 0;
      const coLines = coSearchText.split(/\n/);
      let coCounter = 1;
      coLines.forEach(line => {
        const cleaned = line.trim().replace(/\s+/g, ' ').trim();
        // Skip header-like lines
        if (/^(CO\.?\s*NO|Course\s+Outcome|RBT|Marks?|Q\.?\s*No)/i.test(cleaned)) return;
        if (cleaned.length < 10) return;

        // Extract CO number and level from the line
        const coIdMatch = cleaned.match(/^(?:CO\s*\.?\s*)?(\d+)\s+/);
        const levelMatch = cleaned.match(/\b(L[1-6])\b\s*$/i);

        if (coIdMatch) {
          const coNum = coIdMatch[1];
          let desc = cleaned.substring(coIdMatch[0].length).trim();
          const level = levelMatch ? levelMatch[1].toUpperCase() : '';
          if (level) desc = desc.replace(/\s*L[1-6]\s*$/i, '').trim();

          if (desc.length > 10) {
            const coKey = `CO${coNum}`;
            if (!courseOutcomes.find(c => c.co === coKey)) {
              courseOutcomes.push({ co: coKey, description: desc, level });
            }
          }
        } else if (cleaned.length > 15) {
          // No CO number found, use sequential numbering
          const level = levelMatch ? levelMatch[1].toUpperCase() : '';
          let desc = cleaned;
          if (level) desc = desc.replace(/\s*L[1-6]\s*$/i, '').trim();

          const coKey = `CO${coCounter}`;
          if (!courseOutcomes.find(c => c.co === coKey)) {
            courseOutcomes.push({ co: coKey, description: desc, level });
            coCounter++;
          }
        }
      });
    }
  }

  // Strategy 3: Fallback - use CO regex for "CO1: description" format anywhere in text  
  if (courseOutcomes.length === 0) {
    const coRegex = /CO\s*\.?\s*(\d+)\s*[:\-–.]\s*(.+?)(?=CO\s*\.?\s*\d+\s*[:\-–.]|\n\n|PART|$)/gi;
    let coMatch;
    while ((coMatch = coRegex.exec(fullText)) !== null) {
      const coNum = coMatch[1];
      const desc = coMatch[2].trim().replace(/\s+/g, ' ').trim();
      if (desc.length > 5) {
        courseOutcomes.push({ co: `CO${coNum}`, description: desc });
      }
    }
  }

  console.log('PDF Extracted Course Objectives:', courseObjectives);
  console.log('PDF Extracted Course Outcomes:', courseOutcomes);

  // Embed COs as hidden data for metadata extraction
  if (courseObjectives.length > 0) {
    html += `<div id="courseObjectives" style="display:none;">${JSON.stringify(courseObjectives)}</div>`;
  }
  if (courseOutcomes.length > 0) {
    html += `<div id="courseOutcomes" style="display:none;">${JSON.stringify(courseOutcomes)}</div>`;
  }

  // ---- Question Detection ----
  // Strategy: Find the ACTUAL question section first (after "Q.No" / "Part A" headers),
  // then scan for sequential numbered questions only within that section.
  // This prevents Course Objectives/Outcomes numbered items from being picked up.

  interface QuestionData {
    num: number;
    suffix: string; // Captured suffix (e.g. 'A', 'B')
    text: string;
    co: string;
    rbt: string;
    marks: string;
    pageIndex?: number; // Track which PDF page this question came from
    startLine: number; // The line index where this question started
    endLine: number; // The line index where this question ended
  }

  // ---- Step 1: Find where actual questions begin ----
  // Look for markers that indicate the start of the question section:
  // - "Q.No" header (often in table headers)
  // - "Part A" followed by marks info like "(5 x 2 = 10 marks)" or "(Answer all"
  // - "PART – A" or similar
  // We want to skip everything before: metadata, Course Objectives, Course Outcomes

  // Split the full text into lines for line-by-line parsing
  const allLines = fullText.split('\n');
  console.log(`[PDF] Total lines extracted: ${allLines.length}, globalLineMetadata size: ${globalLineMetadata.length}`);

  // Find the line index where questions actually start
  let questionStartLineIdx = 0;

  // Priority 1: Find "Q.No" header line
  for (let i = 0; i < allLines.length; i++) {
    if (/Q\.?\s*No\b/i.test(allLines[i])) {
      questionStartLineIdx = i;
      console.log(`[PDF] Found Q.No header at line ${i}: "${allLines[i].substring(0, 60)}..."`);
      break;
    }
  }

  // Priority 2: Find "Part A" with marks info if no Q.No found
  if (questionStartLineIdx === 0) {
    for (let i = 0; i < allLines.length; i++) {
      if (/Part\s*[-–—]?\s*A/i.test(allLines[i]) &&
        /(?:\d+\s*(?:×|x|X|\*)\s*\d+|marks?|Answer\s+all)/i.test(allLines[i])) {
        questionStartLineIdx = i;
        console.log(`[PDF] Found Part A with marks at line ${i}: "${allLines[i].substring(0, 60)}..."`);
        break;
      }
    }
  }

  // Priority 3: Find end of Course Outcomes section
  let coEndLineIdx = 0;
  for (let i = 0; i < allLines.length; i++) {
    if (/Course\s+Outcome/i.test(allLines[i]) || /On\s+[Cc]ompletion/i.test(allLines[i])) {
      // Scan forward until we find a non-CO line (until Part A or Q.No)
      for (let j = i + 1; j < allLines.length; j++) {
        if (/Part\s*[-–—]?\s*A/i.test(allLines[j]) || /Q\.?\s*No\b/i.test(allLines[j])) {
          coEndLineIdx = j;
          break;
        }
      }
      if (coEndLineIdx === 0) coEndLineIdx = i + 6; // Skip ~5 CO entries
      break;
    }
  }

  // Use the later of the two start positions
  const effectiveStartLine = Math.max(questionStartLineIdx, coEndLineIdx);
  console.log(`[PDF] Effective question parsing starts at line ${effectiveStartLine} (qHeader: ${questionStartLineIdx}, coEnd: ${coEndLineIdx})`);

  // Debug: dump lines around the question section so we can see exact PDF text structure
  console.log(`[PDF] --- Lines ${effectiveStartLine} to ${Math.min(effectiveStartLine + 40, allLines.length - 1)} ---`);
  for (let d = effectiveStartLine; d < Math.min(effectiveStartLine + 40, allLines.length); d++) {
    console.log(`[PDF] Line ${d}: "${allLines[d].trim().substring(0, 100)}"`);
  }

  // ---- Step 2: Parse questions line-by-line ----
  // Strategy: 
  // - Each line starting with a 1-2 digit number (possibly followed by A/B suffix) 
  //   followed by English text is a question header
  // - Lines that DON'T start with a question number are continuation of the previous question
  // - Strip metadata columns (CO, RBT, Marks) from the right side of each line

  /**
   * Strip PDF table metadata columns (CO, RBT, Marks) from the end of a line.
   * PDF tables linearize: "Define Nyquist rate. CO1 L2 2"
   * We want: { text: "Define Nyquist rate.", co: "CO1", rbt: "L2", marks: "2" }
   */
  const stripLineMetadata = (line: string): { text: string; co: string; rbt: string; marks: string } => {
    let text = line.trim();
    let co = '', rbt = '', marks = '';

    // Pattern 1: "CO1 L2 2" or "CO1 L3 12" or "CO1 L5 6" at end
    const endMeta = text.match(/(?:\s+|^)(CO\s*\d+)\s+(L\d(?:-L\d)?)\s+(\d{1,2})\s*$/i);
    if (endMeta) {
      co = endMeta[1].replace(/\s/g, '');
      rbt = endMeta[2];
      marks = endMeta[3];
      text = text.substring(0, text.length - endMeta[0].length).trim();
      return { text, co, rbt, marks };
    }

    // Pattern 2: "CO1 L2" at end (no marks)
    const endMetaNoMarks = text.match(/(?:\s+|^)(CO\s*\d+)\s+(L\d(?:-L\d)?)\s*$/i);
    if (endMetaNoMarks) {
      co = endMetaNoMarks[1].replace(/\s/g, '');
      rbt = endMetaNoMarks[2];
      text = text.substring(0, text.length - endMetaNoMarks[0].length).trim();
      return { text, co, rbt, marks };
    }

    // Pattern 3: Just "L5 6" at end (CO on a different line)
    const endLevelMarks = text.match(/(?:\s+|^)(L\d(?:-L\d)?)\s+(\d{1,2})\s*$/i);
    if (endLevelMarks) {
      rbt = endLevelMarks[1];
      marks = endLevelMarks[2];
      text = text.substring(0, text.length - endLevelMarks[0].length).trim();
      // Check if text now ends with CO
      const coEnd = text.match(/(?:\s+|^)(CO\s*\d+)\s*$/i);
      if (coEnd) {
        co = coEnd[1].replace(/\s/g, '');
        text = text.substring(0, text.length - coEnd[0].length).trim();
      }
      return { text, co, rbt, marks };
    }

    // Pattern 4: Just a standalone marks number at end: "...signal. 2" or "...function. 12"
    // Only strip if preceded by text ending with period/closing paren
    const endMarksOnly = text.match(/([.)\]])\s+(\d{1,2})\s*$/);
    if (endMarksOnly) {
      marks = endMarksOnly[2];
      text = text.substring(0, text.length - endMarksOnly[0].length + endMarksOnly[1].length).trim();
      return { text, co, rbt, marks };
    }

    return { text, co, rbt, marks };
  };

  /**
   * Check if a line starts with a question number pattern.
   * Returns the parsed number, suffix, and rest of the text, or null if not a question line.
   *
   * Must distinguish between:
   *   QUESTION: "1 Give the relation between..." or "11 A Evaluate whether..."
   *   NOT A QUESTION: "1. from -2 to 2" (sub-item), "3. x(-n)" (sub-item), "5. x(2n)"
   */
  const parseQuestionLine = (line: string, currentQuestionState: any): { num: number; suffix: string; rest: string } | null => {
    const trimmed = line.trim();

    // Strip metadata first before pattern matching so that standalone numbers with metadata aren't missed
    const preClean = stripLineMetadata(trimmed);
    const preCleanLeadingCO = stripLeadingCO(preClean.text);
    const cleanForMatch = preCleanLeadingCO.text.trim();

    // REJECT: Lines where number precedes a Part header
    // e.g., "1 Part A (5 x 2 = 10 marks)" — this is NOT a question
    if (/^\d{1,2}\s*[.)\s]\s*Part\s*[-–—]?\s*[A-C]/i.test(trimmed)) {
      return null;
    }

    // Pattern 1a: "11 A text..." or "11A text..." (number + suffix + text)
    const mSuffix = cleanForMatch.match(/^(\d{1,2})\s*([ABab])\s+(.+)/);
    if (mSuffix) {
      const num = parseInt(mSuffix[1]);
      if (num >= 1 && num <= 20) {
        const rest = mSuffix[3];
        if (/[A-Za-z]{2,}/.test(rest)) {
          return { num, suffix: mSuffix[2].toUpperCase(), rest };
        }
      }
    }

    // Pattern 1b: "11 A" or "11A" alone (number + suffix, text on next continuation line)
    const mSuffixOnly = cleanForMatch.match(/^(\d{1,2})\s*([ABab])\s*$/);
    if (mSuffixOnly) {
      const num = parseInt(mSuffixOnly[1]);
      if (num >= 1 && num <= 20) {
        return { num, suffix: mSuffixOnly[2].toUpperCase(), rest: '' };
      }
    }

    // Pattern 2: "1 Give the relation..." or "2. Define..." (number + question text)
    // CRITICAL: The text MUST start with a capitalized word of 3+ letters.
    // This prevents sub-items like "1. from -2 to 2", "3. x(-n)" from matching.
    const mNum = cleanForMatch.match(/^(\d{1,2})\s*[.)\s]\s*([A-Z][a-zA-Z\u03b4]{2,}.*)/);
    if (mNum) {
      const num = parseInt(mNum[1]);
      if (num >= 1 && num <= 20) {
        // Reject if text starts with Part/Answer/Q.No header words
        if (/^(?:Part|Answer|Questions?)\b/i.test(mNum[2])) return null;
        return { num, suffix: '', rest: mNum[2] };
      }
    }

    // Pattern 3: Standalone number "1" or "2" alone on a line
    // In PDF tables, the Q.No cell may be at a different Y-coordinate than the question text
    // So "1" appears on its own line, with "Give the relation..." on the next line
    const mStandalone = trimmed.match(/^(\d{1,2})\s*[.):]?\s*$/);
    if (mStandalone) {
      const num = parseInt(mStandalone[1]);
      if (num >= 1 && num <= 20) {
        return { num, suffix: '', rest: '' };
      }
    }

    // Pattern 4: Number alone, but followed by metadata that got stripped
    const mCleanedStandalone = cleanForMatch.match(/^(\d{1,2})\s*[.):]?\s*$/);
    if (mCleanedStandalone) {
      const num = parseInt(mCleanedStandalone[1]);
      if (num >= 1 && num <= 20) {
        return { num, suffix: '', rest: '' };
      }
    }

    // Pattern 5: Number followed by random diagram text, IF it's exactly the expected next question number
    const mNextExpected = cleanForMatch.match(/^(\d{1,2})\s*[.)\s]\s*(.*)/);
    if (mNextExpected) {
      const num = parseInt(mNextExpected[1]);
      if (num >= 1 && num <= 20) {
        const isNextExpected = !currentQuestionState || num === currentQuestionState.num + 1 || num === 1;
        if (isNextExpected) {
          if (!/^(?:Part|Answer|Questions?)\b/i.test(mNextExpected[2])) {
            return { num, suffix: '', rest: mNextExpected[2] };
          }
        }
      }
    }

    return null;
  };

  // Track which lines are Part B / Part C headers
  let partBLineIdx = -1;
  let partCLineIdx = -1;
  for (let i = effectiveStartLine; i < allLines.length; i++) {
    if (/Part\s*[-–—]?\s*B/i.test(allLines[i]) && /(?:\d+|marks?|Answer)/i.test(allLines[i])) {
      partBLineIdx = i;
    }
    if (/Part\s*[-–—]?\s*C/i.test(allLines[i]) && /(?:\d+|marks?|Answer)/i.test(allLines[i])) {
      partCLineIdx = i;
    }
  }

  // Build questions by scanning lines
  const questions: QuestionData[] = [];
  let currentQuestion: { num: number; suffix: string; textParts: string[]; co: string; rbt: string; marks: string; startLine: number; endLine: number } | null = null;

  // Track Part B/C header lines so the standalone number guard can accept
  // the first question number after these headers
  let lastPartBHeaderLine = -1;
  let lastPartCHeaderLine = -1;

  // CRITICAL FIX: Buffer for orphaned text lines that appear BEFORE their Q.No line.
  // In PDF tables, Q.No is vertically centered in the cell while question text starts
  // from the top. Even with Y-snapping, some first lines may still appear before the
  // Q.No line in the flattened text. We buffer these orphaned lines and prepend them
  // to the next question that's found.
  let orphanedTextBuffer: { text: string; lineIdx: number }[] = [];

  /**
   * Strip leading CO codes from text.
   * In PDF tables, the CO column value sometimes appears at the START of the next line
   * e.g., "CO1 waveform has a triangular shape" → "waveform has a triangular shape"
   */
  const stripLeadingCO = (text: string): { text: string; co: string } => {
    // Match "CO1 " or "CO1, " at the very start of text
    const m = text.match(/^(CO\s*\d+)\s*,?\s+/i);
    if (m) {
      return { text: text.substring(m[0].length).trim(), co: m[1].replace(/\s/g, '') };
    }
    return { text, co: '' };
  };

  const finalizeQuestion = () => {
    if (!currentQuestion) return;

    console.log(`[PDF-DEBUG] Finalizing Q${currentQuestion.num}${currentQuestion.suffix}: textParts count=${currentQuestion.textParts.length}`);
    currentQuestion.textParts.forEach((tp, idx) => {
      console.log(`[PDF-DEBUG]   textPart[${idx}]: "${tp.substring(0, 80)}${tp.length > 80 ? '...' : ''}"`);
    });

    let qText = currentQuestion.textParts.join(' ').trim();
    // Clean up
    qText = qText.replace(/\s{2,}/g, ' ').trim();
    // Remove trailing "OR" separator
    qText = qText.replace(/\s*\bOR\b\s*$/i, '').trim();
    // Remove "All the Best" and similar
    qText = qText.replace(/\s*~?\*?\s*All\s+the\s+Best\s*\*?~?\s*$/i, '').trim();
    qText = qText.replace(/~\*/g, '').replace(/\*~/g, '').trim();
    // Remove Part headers embedded in text
    qText = qText.replace(/Part\s*[-–—]?\s*[A-C]\s*[-–—]?\s*(?:\([^)]*\)|\d+\s*[×x*]\s*\d+[^)]*)?/gi, '').trim();
    // Remove "Answer all the questions"
    qText = qText.replace(/\(?Answer\s+all\s+(?:the\s+)?questions\)?/gi, '').trim();
    // Remove "Q.No" headers in text
    qText = qText.replace(/Q\.?\s*No\b[^A-Za-z]*/gi, '').trim();
    // Remove CO/RBT/Marks patterns: "CO1 L3 12", "CO1 L5"
    qText = qText.replace(/\bCO\s*\d+\s+L\d(?:-L\d)?\s+\d{1,2}\b/gi, '').trim();
    qText = qText.replace(/\bCO\s*\d+\s+L\d(?:-L\d)?\b/gi, '').trim();
    // Remove standalone CO codes at start or inline: "CO1 waveform..." → "waveform..."
    qText = qText.replace(/^CO\s*\d+\s*,?\s+/i, '').trim();
    // Remove standalone CO codes that appear inline (from column bleed)
    qText = qText.replace(/\bCO\d+\b(?!\s*[&,]\s*CO)/gi, '').trim();
    // Remove standalone RBT levels like "L5" that appear inline
    qText = qText.replace(/(\s)L[1-6](\s)/g, '$1$2').trim();
    // Collapse multiple spaces and commas
    qText = qText.replace(/\s{2,}/g, ' ').trim();
    qText = qText.replace(/^[,;\s]+/, '').trim();

    if (qText.length >= 5) {
      questions.push({
        num: currentQuestion.num,
        suffix: currentQuestion.suffix,
        text: qText,
        co: currentQuestion.co,
        rbt: currentQuestion.rbt,
        marks: currentQuestion.marks,
        pageIndex: globalLineMetadata[currentQuestion.startLine]?.page || 0,
        startLine: currentQuestion.startLine,
        endLine: currentQuestion.endLine
      });
    } else {
      console.log(`[PDF] Discarded Q${currentQuestion.num}${currentQuestion.suffix} — text too short: "${qText}"`);
    }
    currentQuestion = null;
  };

  for (let i = effectiveStartLine; i < allLines.length; i++) {
    const line = allLines[i].trim();
    if (!line) continue;

    console.log(`[PDF-DEBUG] Line ${i}: "${line.substring(0, 120)}${line.length > 120 ? '...' : ''}"`);

    // CRITICAL: When we hit Part headers or Q.No headers, finalize the current question
    // BEFORE skipping. This prevents Q5's text from bleeding into Part B content.
    // Also handles "1 Part A (...)" where q.no column merges with part header
    if (/^(?:\d{1,2}\s*)?Part\s*[-–—]?\s*[A-C]/i.test(line) ||
      /^Q\.?\s*No\b/i.test(line)) {
      console.log(`[PDF-DEBUG]   → HEADER/PART line, skipping`);
      finalizeQuestion();
      // Clear orphaned text buffer on Part/section boundaries
      // to prevent cross-section text contamination
      orphanedTextBuffer = [];

      // Track that we just entered a Part B or Part C section.
      // This helps the standalone number guard below to accept the first Q.No
      // even when currentQuestion is null.
      if (/Part\s*[-–—]?\s*B/i.test(line)) {
        lastPartBHeaderLine = i;
        console.log(`[PDF-DEBUG]   → Marked Part B header at line ${i}`);
      }
      if (/Part\s*[-–—]?\s*C/i.test(line)) {
        lastPartCHeaderLine = i;
        console.log(`[PDF-DEBUG]   → Marked Part C header at line ${i}`);
      }

      // IMPORTANT: If a Q.No number was merged with the Part header (e.g., "1 Part A (5 x 2 = 10 marks)"),
      // the number represents the first question of that part. Initialize a new question for it
      // so that subsequent continuation lines are captured as this question's text.
      // HOWEVER: For Part B/C headers like "11 Part B (5 OR choice pairs...)" the number is NOT
      // question text — it's just Y-snapping putting the Q.No in the Part header line.
      // We should NOT create a question from these; let the real Q11 be found on subsequent lines.
      const mergedQNo = line.match(/^(\d{1,2})\s*Part\s*[-–—]?\s*([A-C])/i);
      if (mergedQNo) {
        const qNum = parseInt(mergedQNo[1]);
        const partLetter = mergedQNo[2].toUpperCase();
        // Only create merged Q.No for Part A (Q1-Q5), never for Part B/C
        // because Part B/C first Q.No (Q6/Q11/Q16) will appear as a standalone number
        // on the next line and should be picked up normally.
        if (qNum >= 1 && qNum <= 5 && partLetter === 'A') {
          console.log(`[PDF-DEBUG]   → Merged Q${qNum} with Part A header`);
          currentQuestion = {
            num: qNum,
            suffix: '',
            textParts: [],
            co: '',
            rbt: '',
            marks: '',
            startLine: i,
            endLine: i
          };
        } else {
          console.log(`[PDF-DEBUG]   → Ignoring merged Q${qNum} with Part ${partLetter} header (not Part A, will be parsed as standalone)`);
        }
      }

      // FIX: Q.No header line may absorb the first question number due to Y-snapping proximity.
      // e.g., "Q.No 1 Give the relation between continuous time..." or "Q.No 1 Part A..."
      // Extract the question number AND any question text from within the header line.
      if (/^Q\.?\s*No\b/i.test(line) && !mergedQNo) {
        // Strategy: Search the header line for a pattern like "<number> <suffix?> <question text>"
        // anywhere within the line (not just right after Q.No).
        // The number might appear after various header content: "Q.No Part A (5 x 2 = 10 marks) CO RBT Marks 1 Give..."
        // or directly after: "Q.No 1 Give the relation..."

        // Try to find a question number + text pattern anywhere in the line
        // Look for: number + optional suffix + capitalized text (at least 10 chars)
        // Use a broad search that skips over header content
        let qnoEmbedded: RegExpMatchArray | null = null;

        // Strategy 1: Direct match after Q.No (handles "Q.No 1 Give the relation...")
        qnoEmbedded = line.match(/Q\.?\s*No\s+(\d{1,2})\s*([ABab])?\s+([A-Z][a-zA-Z\u03b4].{10,})/i);

        // Strategy 2: Match number + text anywhere in the line after header content
        // Handles "Q.No Part A ... CO RBT Marks 1 Give the relation..."
        if (!qnoEmbedded) {
          // Split by common header words and look for number + text in the latter part
          const marksIdx = line.search(/\bMarks?\b/i);
          if (marksIdx > 0) {
            const afterMarks = line.substring(marksIdx + 5).trim();
            const m = afterMarks.match(/^\s*(\d{1,2})\s*([ABab])?\s+([A-Z][a-zA-Z\u03b4].{10,})/i);
            if (m) {
              qnoEmbedded = m;
            }
          }
        }

        if (qnoEmbedded) {
          const qNum = parseInt(qnoEmbedded[1]);
          const suffix = (qnoEmbedded[2] || '').toUpperCase();
          const rest = qnoEmbedded[3];
          // Reject if the "question text" is actually a Part header or column header
          if (/^(?:Part|Answer|Questions?|CO\s|RBT|Marks?)\b/i.test(rest)) {
            console.log(`[PDF-DEBUG]   → Rejected Q${qNum}${suffix} embedded in Q.No header — text is a header: "${rest.substring(0, 40)}"`);
          } else if (qNum >= 1 && qNum <= 20) {
            const cleaned = stripLineMetadata(rest);
            let initialText = cleaned.text;
            const leadingCOInit = stripLeadingCO(initialText);
            initialText = leadingCOInit.text;
            console.log(`[PDF-DEBUG]   → Q${qNum}${suffix} embedded in Q.No header line: "${initialText.substring(0, 80)}"`);
            currentQuestion = {
              num: qNum,
              suffix: suffix,
              textParts: initialText ? [initialText] : [],
              co: cleaned.co || leadingCOInit.co,
              rbt: cleaned.rbt,
              marks: cleaned.marks,
              startLine: i,
              endLine: i
            };
          }
        } else {
          // Check for just a standalone number in the header line (no question text)
          // e.g., "Q.No 1 Part A (4 x 2 = 8 marks) CO RBT Marks"
          // Search for the number anywhere in the line
          const headerWithNum = line.match(/Q\.?\s*No\s+(\d{1,2})(?:\s|$)/i) ||
            line.match(/\bMarks?\b\s+(\d{1,2})(?:\s|$)/i);
          if (headerWithNum) {
            const qNum = parseInt(headerWithNum[1]);
            // Only accept if it's a plausible first question number (1-5)
            // and not likely a marks value or column data
            if (qNum >= 1 && qNum <= 5) {
              console.log(`[PDF-DEBUG]   → Q${qNum} number embedded in Q.No header (no text yet)`);
              currentQuestion = {
                num: qNum,
                suffix: '',
                textParts: [],
                co: '',
                rbt: '',
                marks: '',
                startLine: i,
                endLine: i
              };
            }
          }
        }
      }
      continue;
    }
    // "OR" separator lines: finalize current question and reset.
    // CRITICAL: After OR, lines that appear before the next Q.No label belong to
    // the NEXT question (e.g. 11B), not the previous one (e.g. 11A).
    // We finalize the current question and set currentQuestion=null so that
    // subsequent text lines go into orphanedTextBuffer, ready to be prepended
    // to the next question found.
    if (/^\s*OR\s*$/i.test(line)) {
      console.log(`[PDF-DEBUG]   → OR separator: finalizing Q${currentQuestion?.num}${currentQuestion?.suffix} and resetting for next OR pair`);
      finalizeQuestion(); // sets currentQuestion = null
      orphanedTextBuffer = []; // clear any stale orphans from previous row
      continue;
    }
    // Skip "Answer all" lines
    if (/^\(?Answer\s+all/i.test(line)) { console.log(`[PDF-DEBUG]   → Answer all, skipping`); continue; }
    // Skip lines that are just column headers (CO, RBT, Marks on their own)
    if (/^(?:CO|RBT|Marks?|Questions?)$/i.test(line)) { console.log(`[PDF-DEBUG]   → Column header, skipping`); continue; }
    // Skip standalone CO column values like "CO1", "CO2" 
    if (/^CO\s*\d+$/i.test(line)) { console.log(`[PDF-DEBUG]   → Standalone CO value, skipping`); continue; }
    // Skip "Course Objective/Outcome" header lines that may appear after Q.No
    if (/^Course\s+(Objective|Outcome)/i.test(line)) { console.log(`[PDF-DEBUG]   → Course header, skipping`); continue; }
    // Skip "CO.NO" or "On Completion" lines
    if (/^(?:CO\.?\s*NO|On\s+[Cc]ompletion)/i.test(line)) { console.log(`[PDF-DEBUG]   → CO.NO/Completion, skipping`); continue; }
    // Skip ~* All the Best *~ lines
    if (/All\s+the\s+Best/i.test(line)) { console.log(`[PDF-DEBUG]   → All the Best, skipping`); continue; }

    const parsed = parseQuestionLine(line, currentQuestion);

    if (parsed) {
      console.log(`[PDF-DEBUG]   → parseQuestionLine matched: num=${parsed.num}, suffix="${parsed.suffix}", rest="${parsed.rest.substring(0, 80)}${parsed.rest.length > 80 ? '...' : ''}"`);
      // Guard for standalone numbers (no text, no suffix):
      // These could be marks values ("2", "6", "8", "12", "16") on their own line
      // Only treat as a new question if it's a plausible NEXT question number
      if (!parsed.rest && !parsed.suffix) {
        // Accept the number if:
        // 1. No current question (e.g., at start or after Part header)
        // 2. It's the next sequential number after current question
        // 3. It's Q1 (always valid as first question)
        // 4. It's the first question after Part B header (typically Q6 or Q11)
        // 5. It's the first question after Part C header (typically Q14 or Q16)
        const isAfterPartBHeader = lastPartBHeaderLine >= 0 && i > lastPartBHeaderLine && i <= lastPartBHeaderLine + 5;
        const isAfterPartCHeader = lastPartCHeaderLine >= 0 && i > lastPartCHeaderLine && i <= lastPartCHeaderLine + 5;
        const isNextExpected = !currentQuestion ||
          parsed.num === (currentQuestion.num + 1) ||
          parsed.num === 1 || // Q1 is always valid as first question
          isAfterPartBHeader ||
          isAfterPartCHeader;

        if (!isNextExpected) {
          console.log(`[PDF-DEBUG]   → Standalone ${parsed.num} not expected, skipping (current Q${currentQuestion?.num})`);
          // Likely a marks value, not a question number — skip silently
          continue;
        }
      }

      // This line starts a new question — finalize previous question first
      finalizeQuestion();

      // Strip metadata from the rest of the text on this line
      const cleaned = stripLineMetadata(parsed.rest);
      let initialText = cleaned.text;
      // Also strip leading CO codes
      const leadingCOInit = stripLeadingCO(initialText);
      initialText = leadingCOInit.text;

      console.log(`[PDF-DEBUG]   → Starting Q${parsed.num}${parsed.suffix}: initialText="${initialText.substring(0, 80)}", co=${cleaned.co || leadingCOInit.co}, rbt=${cleaned.rbt}, marks=${cleaned.marks}`);

      // CRITICAL FIX: Prepend any orphaned text that appeared BEFORE this Q.No line.
      // In PDF tables, the first line(s) of question text may appear before the Q.No
      // because Q.No is vertically centered in the cell. These orphaned lines were
      // buffered and should be prepended to this question's text.
      //
      // HOWEVER: If this is Q2 (or later) and NO Q1 has been parsed yet, the orphaned
      // text likely belongs to Q1 (whose number was lost). In that case, synthesize Q1
      // from the orphaned text rather than incorrectly prepending it to Q2.
      const initialParts: string[] = [];
      if (orphanedTextBuffer.length > 0) {
        // Check if we should create Q1 from orphaned text instead of prepending to current.
        // This only applies to Part A Q1 (the very first question, num=1 not yet seen).
        const noQ1Yet = questions.length === 0 && !currentQuestion;
        const isQ2OrLater = parsed.num >= 2;
        const isPartAOrphan = noQ1Yet && isQ2OrLater && !parsed.suffix;

        if (isPartAOrphan) {
          // These orphaned lines are Q1's text — create Q1 from them
          console.log(`[PDF-DEBUG]   → Orphaned text belongs to Q1 (synthesizing Q1 from ${orphanedTextBuffer.length} lines)`);
          const q1Text = orphanedTextBuffer.map(o => o.text).join(' ').trim();
          if (q1Text.length >= 5) {
            // Extract metadata from q1Text
            const q1Cleaned = stripLineMetadata(q1Text);
            let q1CleanedText = q1Cleaned.text;
            const q1LeadingCO = stripLeadingCO(q1CleanedText);
            q1CleanedText = q1LeadingCO.text;

            const q1StartLine = orphanedTextBuffer[0].lineIdx;
            const q1EndLine = orphanedTextBuffer[orphanedTextBuffer.length - 1].lineIdx;

            questions.push({
              num: 1,
              suffix: '',
              text: q1CleanedText,
              co: q1Cleaned.co || q1LeadingCO.co,
              rbt: q1Cleaned.rbt,
              marks: q1Cleaned.marks,
              pageIndex: globalLineMetadata[q1StartLine]?.page || 0,
              startLine: q1StartLine,
              endLine: q1EndLine
            });
            console.log(`[PDF-DEBUG]   → Synthesized Q1: "${q1CleanedText.substring(0, 80)}"`);
          }
          orphanedTextBuffer = [];
        } else {
          // Prepend orphaned lines to this question (e.g. 11B's text that appeared before "11 B" label)
          console.log(`[PDF-DEBUG]   → Prepending ${orphanedTextBuffer.length} orphaned text line(s) to Q${parsed.num}${parsed.suffix}`);
          orphanedTextBuffer.forEach((orphanLine, idx) => {
            console.log(`[PDF-DEBUG]     orphan[${idx}]: "${orphanLine.text.substring(0, 80)}"`);
          });
          initialParts.push(...orphanedTextBuffer.map(o => o.text));
          orphanedTextBuffer = [];
        }
      }
      if (initialText) {
        initialParts.push(initialText);
      }

      currentQuestion = {
        num: parsed.num,
        suffix: parsed.suffix,
        textParts: initialParts,
        co: cleaned.co || leadingCOInit.co,
        rbt: cleaned.rbt,
        marks: cleaned.marks,
        startLine: i,
        endLine: i
      };
    } else if (currentQuestion) {
      // Continuation line — add to current question's text
      const cleaned = stripLineMetadata(line);
      let cleanedText = cleaned.text;

      // Also strip leading CO codes that bleed from the CO column
      // e.g., "CO1 waveform has a triangular shape" → "waveform has a triangular shape"
      const leadingCO = stripLeadingCO(cleanedText);
      cleanedText = leadingCO.text;
      if (leadingCO.co && !currentQuestion.co) currentQuestion.co = leadingCO.co;

      if (cleanedText.length > 0) {
        console.log(`[PDF-DEBUG]   → Continuation for Q${currentQuestion.num}${currentQuestion.suffix}: "${cleanedText.substring(0, 80)}${cleanedText.length > 80 ? '...' : ''}"`);
        currentQuestion.textParts.push(cleanedText);
        currentQuestion.endLine = i;
      } else {
        console.log(`[PDF-DEBUG]   → Continuation line empty after strip (original: "${line.substring(0, 60)}")`);
      }
      // Pick up CO/RBT/Marks if not already set
      if (!currentQuestion.co && cleaned.co) currentQuestion.co = cleaned.co;
      if (!currentQuestion.rbt && cleaned.rbt) currentQuestion.rbt = cleaned.rbt;
      if (!currentQuestion.marks && cleaned.marks) currentQuestion.marks = cleaned.marks;
    } else {
      // CRITICAL FIX: No currentQuestion and no question match — this line is orphaned.
      // In PDF tables, question text may appear BEFORE the Q.No line due to vertical centering.
      // Buffer this text so it can be prepended to the next question that's found.
      const cleaned = stripLineMetadata(line);
      let cleanedText = cleaned.text;
      const leadingCO = stripLeadingCO(cleanedText);
      cleanedText = leadingCO.text;

      // Only buffer meaningful text (not just numbers, marks, metadata)
      if (cleanedText.length > 3 && /[a-zA-Z]{2,}/.test(cleanedText)) {
        console.log(`[PDF-DEBUG]   → Buffering orphaned text: "${cleanedText.substring(0, 80)}"`);
        orphanedTextBuffer.push({ text: cleanedText, lineIdx: i });
      } else {
        console.log(`[PDF-DEBUG]   → No currentQuestion and no match, line ignored: "${line.substring(0, 60)}"`);
      }
    }
  }

  // Finalize the last question
  finalizeQuestion();

  // Sort questions by (num, suffix) to ensure correct ordering
  questions.sort((a, b) => {
    if (a.num !== b.num) return a.num - b.num;
    return (a.suffix || '').localeCompare(b.suffix || '');
  });

  // ============================================================
  // CROP QUESTION IMAGES: For each question, crop its FULL visual
  // region from the rendered PDF page. This captures the question
  // text AND any diagrams/images/equations — producing the same
  // quality as an HTML file upload.
  // ============================================================
  if (globalLineMetadata && globalLineMetadata.length > 0 && pageImages && pageImages.length > 0) {
    for (let qi = 0; qi < questions.length; qi++) {
      const q = questions[qi];

      const startMeta = globalLineMetadata[q.startLine];
      const endMeta = globalLineMetadata[q.endLine];
      if (!endMeta || !startMeta) continue;

      // Determine the bottom boundary of this question's region.
      // It's the top of the NEXT question/OR separator, or the page bottom.
      let bottomBoundaryY: number | null = null;
      let bottomBoundaryPage = startMeta.page;

      // Search for the next boundary: next question start or OR separator
      const searchEndLine = (qi + 1 < questions.length) ? questions[qi + 1].startLine : allLines.length;
      for (let idx = q.endLine + 1; idx < searchEndLine; idx++) {
        if (/^\s*OR\s*$/i.test(allLines[idx].trim())) {
          const orMeta = globalLineMetadata[idx];
          if (orMeta && orMeta.page === startMeta.page) {
            bottomBoundaryY = orMeta.y + 5; // Just above the OR line
          }
          break;
        }
        const meta = globalLineMetadata[idx];
        if (meta && meta.page !== startMeta.page) {
          break; // Page boundary
        }
      }

      // If no OR found, use the next question's start line
      if (bottomBoundaryY === null && qi + 1 < questions.length) {
        const nextQ = questions[qi + 1];
        const nextStartMeta = globalLineMetadata[nextQ.startLine];
        if (nextStartMeta && nextStartMeta.page === startMeta.page) {
          bottomBoundaryY = nextStartMeta.y + Math.max(nextStartMeta.height, 10) + 5;
        }
      }

      // Fallback: use the end of the question's last text line with padding
      if (bottomBoundaryY === null) {
        bottomBoundaryY = endMeta.y - Math.max(endMeta.height, 10) - 10;
      }

      // The top boundary is the question's first text line
      const topBoundaryY = startMeta.y + Math.max(startMeta.height, 10) + 5;

      // Only crop if the region is on the same page and has enough height
      if (startMeta.page === (bottomBoundaryPage) || true) {
        const cropHeight = topBoundaryY - bottomBoundaryY;
        if (cropHeight > 30) {
          const pgImg = pageImages[startMeta.page];
          if (pgImg) {
            const croppedDataUrl = await cropFromPdfDataUrl(pgImg, topBoundaryY, bottomBoundaryY);
            if (croppedDataUrl) {
              // Store the cropped image as a separate property for the question
              (q as any).croppedImage = croppedDataUrl;
              console.log(`[PDF-CROP] Cropped Q${q.num}${q.suffix}: topY=${topBoundaryY.toFixed(1)}, bottomY=${bottomBoundaryY.toFixed(1)}, height=${cropHeight.toFixed(1)}pts`);
            }
          }
        }
      }
    }
  }

  // ============================================================
  // POST-PROCESSING: Fix text bleeding between Part B OR pairs.
  // In PDF tables, text from Q(n+1)A's cell can bleed into Q(n)B
  // because Q(n+1)A's text starts at the TOP of its cell while
  // its Q.No label is vertically centered LOWER. The flattened
  // text puts Q(n+1)A's first lines as continuations of Q(n)B.
  //
  // Detection strategies:
  // 1. If Q(n+1)A starts with a lowercase word, it's definitely
  //    a sentence fragment continued from Q(n)B's trailing text.
  // 2. If Q(n)B's text and Q(n+1)A's text, when joined, form a
  //    coherent sentence that was split, rejoin and redistribute.
  // ============================================================
  for (let qi = 0; qi < questions.length - 1; qi++) {
    const qB = questions[qi];
    const qNextA = questions[qi + 1];

    // Only process Part B cross-pair boundaries:
    // Q(n)B followed by Q(n+1)A (or Q(n) no-suffix followed by Q(n+1)A)
    const isCrossPair =
      (qB.suffix === 'B' || qB.suffix === '') &&
      qNextA.suffix === 'A' &&
      qNextA.num === qB.num + 1 &&
      qB.num >= 6; // Only Part B questions (num >= 6)

    if (!isCrossPair) continue;

    const bText = qB.text;
    const aText = qNextA.text;

    // Strategy 1: Q(n+1)A starts with lowercase → definitely a fragment.
    // Find the start of the sentence in Q(n)B that continues into Q(n+1)A.
    // Join them, then split: everything from the sentence start goes to Q(n+1)A.
    if (/^[a-z]/.test(aText)) {
      // The fragment in qNextA continues a sentence from the end of qB.
      // Try to find the sentence boundary in qB's text by looking for a
      // capital letter that starts the trailing sentence.
      // Walk backwards from the end of bText to find the last sentence start.
      // A "sentence start" is a capital letter preceded by ". " or start-of-text.
      let splitIdx = -1;
      for (let ci = bText.length - 1; ci >= 0; ci--) {
        const ch = bText[ci];
        if (/[A-Z]/.test(ch)) {
          // Check if preceded by ". " or ") " or start of string
          const before = ci > 0 ? bText.substring(Math.max(0, ci - 2), ci) : '';
          if (ci === 0 || /[.)]\s*$/.test(before) || /\)\s*$/.test(before) || /,\s*$/.test(before)) {
            splitIdx = ci;
            // Don't break — keep searching further back for a better split point
            // Actually, we want the LAST (rightmost) sentence start that makes the
            // fragment coherent. The first capital letter we find walking backwards
            // is the best candidate. But we need to be careful not to split too early.
            break;
          }
        }
      }

      if (splitIdx > 0) {
        const movedText = bText.substring(splitIdx).trim();
        const remainingBText = bText.substring(0, splitIdx).trim();
        const newAText = movedText + ' ' + aText;

        console.log(`[PDF-POSTFIX] Text bleeding fix: Q${qB.num}${qB.suffix} → Q${qNextA.num}${qNextA.suffix}`);
        console.log(`[PDF-POSTFIX]   Moved: "${movedText.substring(0, 80)}..."`);
        console.log(`[PDF-POSTFIX]   Q${qB.num}${qB.suffix} now: "${remainingBText.substring(0, 80)}..."`);
        console.log(`[PDF-POSTFIX]   Q${qNextA.num}${qNextA.suffix} now: "${newAText.substring(0, 80)}..."`);

        qB.text = remainingBText;
        qNextA.text = newAText;
      }
      continue;
    }

    // Strategy 2: Trailing numbered sub-items after sentence-ending punctuation.
    // In PDF tables, numbered sub-items from Q(n+1)A's cell can appear at the
    // end of Q(n)B's text because they were in the flattened text before Q(n+1)A's
    // Q.No label. Detect pattern: "sentence. N. text..." at the end of Q(n)B.
    // Example: Q11B ends with "...signal. 1. from -2 to 2" — move "1. from..." to Q12A.
    const trailingSplit = bText.match(/^(.*[.!?;])\s+(\d+\.\s+.*)$/s);
    if (trailingSplit) {
      const remainingBText = trailingSplit[1].trim();
      const movedText = trailingSplit[2].trim();

      // Sanity: remaining text should be substantial
      if (remainingBText.length >= 20 && movedText.length >= 3) {
        const newAText = movedText + ' ' + aText;

        console.log(`[PDF-POSTFIX] Text bleeding fix (trailing sub-item): Q${qB.num}${qB.suffix} → Q${qNextA.num}${qNextA.suffix}`);
        console.log(`[PDF-POSTFIX]   Moved: "${movedText.substring(0, 80)}..."`);
        console.log(`[PDF-POSTFIX]   Q${qB.num}${qB.suffix} now: "${remainingBText.substring(0, 80)}..."`);
        console.log(`[PDF-POSTFIX]   Q${qNextA.num}${qNextA.suffix} now: "${newAText.substring(0, 80)}..."`);

        qB.text = remainingBText;
        qNextA.text = newAText;
        continue;
      }
    }

    // Strategy 3: Q(n+1)A starts with uppercase but Q(n)B has trailing text
    // from Q(n+1)A's cell. Detect by looking for question-start verbs at
    // sentence boundaries in Q(n)B. If multiple verbs found, the LAST one
    // likely starts Q(n+1)A's text.
    const questionStartVerbs = /(?:^|[.)]\s+)(Compute|Determine|Find|Design|Derive|Draw|Sketch|State|Explain|Define|Apply|From\s+image|Calculate|Evaluate|Obtain|Prove|Show|Verify|Check|Write|List|Classify|Compare|Discuss|Illustrate|Implement|Analyze|Describe)\b/g;

    const verbMatches: { index: number; verb: string }[] = [];
    let vm: RegExpExecArray | null;
    while ((vm = questionStartVerbs.exec(bText)) !== null) {
      const verbStart = vm.index + vm[0].indexOf(vm[1]);
      verbMatches.push({ index: verbStart, verb: vm[1] });
    }

    if (verbMatches.length >= 2) {
      const lastMatch = verbMatches[verbMatches.length - 1];
      const splitIdx = lastMatch.index;
      const movedText = bText.substring(splitIdx).trim();
      const remainingBText = bText.substring(0, splitIdx).trim();

      if (remainingBText.length >= 20 && movedText.length >= 10) {
        const newAText = movedText + ' ' + aText;

        console.log(`[PDF-POSTFIX] Text bleeding fix (verb heuristic): Q${qB.num}${qB.suffix} → Q${qNextA.num}${qNextA.suffix}`);
        console.log(`[PDF-POSTFIX]   Split verb: "${lastMatch.verb}" at index ${splitIdx}`);
        console.log(`[PDF-POSTFIX]   Moved: "${movedText.substring(0, 80)}..."`);

        qB.text = remainingBText;
        qNextA.text = newAText;
      }
    }
  }

  console.log(`[PDF] Questions parsed: ${questions.length}`);
  if (questions.length > 0) {
    console.log(`[PDF] Question list: ${questions.map(q => `Q${q.num}${q.suffix}`).join(', ')}`);
    questions.forEach(q => {
      console.log(`  Q${q.num}${q.suffix}: "${q.text.substring(0, 60)}..." CO=${q.co} RBT=${q.rbt} Marks=${q.marks}`);
    });
  }

  // Set page indices based on cumulative character positions
  const pageCharOffsets: number[] = [];
  if (pageImages && pageImages.length > 0 && pageTexts && pageTexts.length > 0) {
    let cumLen = 0;
    for (let p = 0; p < pageTexts.length; p++) {
      pageCharOffsets.push(cumLen);
      cumLen += pageTexts[p].length + 1;
    }
  }

  // Assign page indices to questions based on which line they started on
  if (pageTexts && pageTexts.length > 0) {
    // Build line-to-page mapping
    const lineToPage: number[] = [];
    const lineCount = 0;
    for (let p = 0; p < pageTexts.length; p++) {
      const pageLineCount = pageTexts[p].split('\n').length;
      for (let l = 0; l < pageLineCount; l++) {
        lineToPage.push(p);
      }
    }
    // Note: questions have startLine relative to allLines, which is built from all pages
    // Just assign page 0 for now since most IA papers are single-page
    questions.forEach(q => { q.pageIndex = 0; });
  }

  // ---- Detect Part boundaries from question numbers ----
  // Use a simple, reliable approach based on question numbering conventions:
  // IA papers: Part A = Q1-Q5/Q10, Part B = Q11-Q13/Q15, Part C = Q16+
  let partBStartQ = -1;
  let partCStartQ = -1;

  // If we found Part B header line, use the first question that appears after it
  if (partBLineIdx >= 0) {
    const partBQ = questions.find(q => q.startLine > partBLineIdx);
    if (partBQ) {
      partBStartQ = partBQ.num;
    } else {
      const fallback = questions.find(q => q.num >= 6 && q.num !== 10);
      if (fallback) partBStartQ = fallback.num;
    }
  }

  // If we found Part C header line, Part C typically starts at Q16
  if (partCLineIdx >= 0) {
    const partCQ = questions.find(q => q.startLine > partCLineIdx);
    if (partCQ) {
      partCStartQ = partCQ.num;
    } else {
      const fallbackCQ = questions.find(q => q.num >= 16);
      if (fallbackCQ) partCStartQ = fallbackCQ.num;
    }
  }

  // Fallback: if Part B header not found but we have questions with large numbers,
  // infer Part B boundary from the gap in question numbering
  if (partBStartQ < 0 && questions.length > 0) {
    // Look for a gap: e.g., Q5 -> Q11 (gap of 6+)
    for (let i = 1; i < questions.length; i++) {
      if (!questions[i].suffix && !questions[i - 1].suffix) {
        const gap = questions[i].num - questions[i - 1].num;
        if (gap >= 3) {
          partBStartQ = questions[i].num;
          break;
        }
      }
    }
  }

  console.log(`[PDF] Part boundaries: B starts at Q${partBStartQ}, C starts at Q${partCStartQ}`);

  // ---- Generate HTML Tables ----

  const generateTable = (id: string, title: string, qs: typeof questions) => {
    let tableHtml = `
      <h3 style="text-align: center; margin-top: 20px; font-weight: bold;">${title}</h3>
      <table id="${id}" style="width: 100%; border-collapse: collapse; border: 1px solid black;">
        <thead>
          <tr style="background-color: #f0f0f0;">
            <th style="border: 1px solid black; padding: 8px; width: 50px;">Q.No</th>
            <th style="border: 1px solid black; padding: 8px;">Questions</th>
            <th style="border: 1px solid black; padding: 8px; width: 60px;">CO</th>
            <th style="border: 1px solid black; padding: 8px; width: 80px;">RBT</th>
            <th style="border: 1px solid black; padding: 8px; width: 60px;">Marks</th>
          </tr>
        </thead>
        <tbody>
    `;

    let prevNum = -1;

    for (const q of qs) {
      // Insert OR separator row between A/B variants of the same question number
      if (q.num === prevNum && q.suffix === 'B') {
        tableHtml += `
          <tr>
            <td colspan="5" style="border: 1px solid black; padding: 4px; text-align: center; font-weight: bold; background-color: #f9f9f9;">OR</td>
          </tr>
        `;
      }
      prevNum = q.num;

      // Build question cell content:
      // - Always include extracted text (for parsing/search)
      // - If we have a cropped image from the PDF page render, embed it
      //   This produces the same quality as HTML uploads with inline images
      let questionCellContent = q.text;
      const croppedImg = (q as any).croppedImage;
      if (croppedImg) {
        questionCellContent += `<br><img src="${croppedImg}" alt="Question ${q.num}${q.suffix}" style="max-width: 100%; display: block; margin: 8px 0; border-radius: 4px;" />`;
      }

      // Question row — includes text and cropped PDF image
      tableHtml += `
        <tr>
          <td style="border: 1px solid black; padding: 8px; text-align: center; vertical-align: top;">${q.num}${q.suffix ? ' ' + q.suffix : ''}</td>
          <td style="border: 1px solid black; padding: 8px; vertical-align: top;">${questionCellContent}</td>
          <td style="border: 1px solid black; padding: 8px; text-align: center; vertical-align: top;">${q.co}</td>
          <td style="border: 1px solid black; padding: 8px; text-align: center; vertical-align: top;">${q.rbt}</td>
          <td style="border: 1px solid black; padding: 8px; text-align: center; vertical-align: top;">${q.marks}</td>
        </tr>
      `;
    }

    tableHtml += '</tbody></table>';
    return tableHtml;
  };

  if (questions.length > 0) {
    if (partBStartQ > 0) {
      const partAQs = questions.filter(q => q.num < partBStartQ);
      let partBQs: QuestionData[];
      let partCQs: QuestionData[] = [];

      if (partCStartQ > 0) {
        // Has Part C - split Part B and Part C
        partBQs = questions.filter(q => q.num >= partBStartQ && q.num < partCStartQ);
        partCQs = questions.filter(q => q.num >= partCStartQ);
      } else {
        partBQs = questions.filter(q => q.num >= partBStartQ);
      }

      if (partAQs.length > 0) html += generateTable('parta', 'Part A', partAQs);
      if (partBQs.length > 0) html += generateTable('partb', 'Part B', partBQs);
      if (partCQs.length > 0) html += generateTable('partc', 'Part C', partCQs);
    } else {
      html += generateTable('parta', questionStartLineIdx > 0 ? 'Part A' : 'Questions', questions);
    }
  } else {
    html += '<p style="color: red; text-align: center; padding: 20px; font-weight: bold;">No numbered questions (1., 2., 3...) detected in this document.</p>';
  }

  html += '</div>';
  return html;
};
