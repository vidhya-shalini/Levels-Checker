import { ParseResult } from './parser/htmlParser';
import { Question } from '@/components/QuestionDisplay';
import { CIT_LOGO_BASE64 } from '@/assets/citLogoBase64';

// CIT logo is directly available as a Base64 data URL
// No fetch needed — baked into the JS bundle at build time
// This guarantees the logo works in downloaded HTML files (offline, no server needed)
const fetchLogoAsDataUrl = async (): Promise<string> => {
    // Validate the base64 string is actually present and valid
    if (CIT_LOGO_BASE64 && CIT_LOGO_BASE64.startsWith('data:image/')) {
        return CIT_LOGO_BASE64;
    }
    // Fallback: try to dynamically import the PNG file and convert to base64
    console.warn('[HtmlGen] CIT_LOGO_BASE64 is missing or invalid, attempting fallback...');
    try {
        // Vite resolves this to a URL at build time; fetch it and convert to data URL
        const logoModule = await import('@/assets/cit-logo.png');
        const logoUrl = typeof logoModule === 'string' ? logoModule : logoModule.default;
        if (logoUrl) {
            const resp = await fetch(logoUrl);
            const blob = await resp.blob();
            return await new Promise<string>((resolve) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result as string);
                reader.readAsDataURL(blob);
            });
        }
    } catch (err) {
        console.warn('[HtmlGen] Fallback logo fetch also failed:', err);
    }
    return ''; // No logo available
};

/**
 * Compute Part B marks distribution string from the actual questions.
 * Groups questions by marks and outputs e.g. "(2 x 12 = 24 marks)(1 x 16 = 16 marks)"
 */
const computePartBDistribution = (partB: Question[]): string => {
    // Only count "A" variants (the primary question, not OR alternatives)
    const primaryQuestions = partB.filter(q => !q.questionNumber.toUpperCase().includes('B'));
    if (primaryQuestions.length === 0) return '';

    // Group by marks
    const markGroups = new Map<number, number>();
    primaryQuestions.forEach(q => {
        const m = q.marks || 0;
        markGroups.set(m, (markGroups.get(m) || 0) + 1);
    });

    // Sort by marks value
    const entries = Array.from(markGroups.entries()).sort((a, b) => a[0] - b[0]);

    if (entries.length === 1) {
        const [marks, count] = entries[0];
        return `${count} x ${marks} = ${count * marks} marks`;
    }

    // Multiple mark groups: (2 x 12 = 24 marks)(1 x 16 = 16 marks)
    return entries.map(([marks, count]) => `(${count} x ${marks} = ${count * marks} marks)`).join('');
};

const createQuestionsTableHtml = async (questions: Question[], partLabel: string, pdfPageImages?: (string | null)[]): Promise<string> => {
    console.log(`[HtmlGen] Creating table for ${questions.length} questions`);

    let html = '<table style="width: 100%; border-collapse: collapse; border: 1px solid black; table-layout: fixed;">';

    // Single header row with part label inside the table (no separate heading outside)
    html += `<tr style="background-color: #f0f0f0;">
        <th style="border: 1px solid black; padding: 3px; text-align: center; width: 6%; font-size: 10pt;">Q.No</th>
        <th style="border: 1px solid black; padding: 3px; text-align: center; width: 70%; font-size: 10pt;"><strong>${partLabel}</strong><br><span style="font-size: 9pt; font-weight: normal;">(Answer all the questions)</span></th>
        <th style="border: 1px solid black; padding: 3px; text-align: center; width: 7%; font-size: 10pt;">CO</th>
        <th style="border: 1px solid black; padding: 3px; text-align: center; width: 7%; font-size: 10pt;">RBT</th>
        <th style="border: 1px solid black; padding: 3px; text-align: center; width: 7%; font-size: 10pt;">Marks</th>
    </tr>`;

    let prevNum = -1;
    for (const q of questions) {
        const m = q.questionNumber.match(/(\d+)/);
        const num = m ? parseInt(m[1]) : -1;
        if (q.questionNumber.toUpperCase().includes('B') && num === prevNum) {
            html += `<tr>
                <td colspan="5" style="border: 1px solid black; text-align: center; padding: 2px; font-weight: bold; font-size: 10pt;">OR</td>
            </tr>`;
        }
        prevNum = num;

        // For enhanced questions, use the enhanced text (q.text) which has been updated
        // For non-enhanced questions, use originalHtml for formatting fidelity
        let questionContent = '';

        if (q.isEnhanced) {
            questionContent = (q.text || '').replace(/\s+/g, ' ').trim();
            console.log(`[HtmlGen] Q${q.questionNumber}: Using ENHANCED text`);
        } else if (q.originalHtml) {
            questionContent = q.originalHtml;
            console.log(`[HtmlGen] Q${q.questionNumber}: Using originalHtml`);
        } else {
            questionContent = (q.text || '').replace(/\s+/g, ' ').trim();
            console.log(`[HtmlGen] Q${q.questionNumber}: Using plain text fallback`);
        }

        // Handle PDF fallback logic
        if (!questionContent && pdfPageImages?.length) {
            const pageMatch = ((q.text || '') + (q.originalHtml || '')).match(/data-page="(\d+)"/);
            if (pageMatch) {
                const pageIdx = parseInt(pageMatch[1]);
                const url = pdfPageImages[pageIdx];
                if (url) {
                    questionContent = `<div style="margin: 4px 0;"><img src="${url}" style="max-width: 400px; height: auto; display: block;" alt="PDF page"/></div>`;
                }
            }
        }

        const wrappedContent = `<div class="question-content" style="line-height: 1.3; word-wrap: break-word;">${questionContent}</div>`;

        html += `<tr style="vertical-align: top;">
            <td style="border: 1px solid black; padding: 3px 4px; text-align: center; font-size: 10pt;" contenteditable="true">${q.questionNumber}</td>
            <td style="border: 1px solid black; padding: 3px 4px; word-wrap: break-word; overflow-wrap: break-word; font-size: 10pt;" contenteditable="true">${wrappedContent}</td>
            <td style="border: 1px solid black; padding: 3px 4px; text-align: center; font-size: 10pt;" contenteditable="true">${q.co}</td>
            <td style="border: 1px solid black; padding: 3px 4px; text-align: center; font-size: 10pt;" contenteditable="true">${q.detectedLevel}</td>
            <td style="border: 1px solid black; padding: 3px 4px; text-align: center; font-size: 10pt;" contenteditable="true">${q.marks}</td>
        </tr>`;
    }

    html += '</table>';
    return html;
};

export const generateHtmlDocument = async (parseResult: ParseResult, iaType: 'IA1' | 'IA2' | 'IA3'): Promise<Blob> => {
    const logoUrl = await fetchLogoAsDataUrl();
    const { metadata, partA, partB, partC, pdfPageImages } = parseResult;

    // ---- Compute marks distribution dynamically from actual questions ----
    const partACount = partA.length;
    const partAPerQ = partA.length > 0 ? partA[0].marks : 2;
    const partATotal = partACount * partAPerQ;
    const partALabel = `Part A (${partACount} x ${partAPerQ} = ${partATotal} marks)`;

    const partBDistribution = computePartBDistribution(partB);
    const partBLabel = `Part B - ${partBDistribution}`;

    let partCLabel = '';
    if (partC && partC.length > 0) {
        const partCPrimary = partC.filter(q => !q.questionNumber.toUpperCase().includes('B'));
        const partCPerQ = partCPrimary.length > 0 ? partCPrimary[0].marks : 15;
        partCLabel = `Part C (${partCPrimary.length} x ${partCPerQ} = ${partCPrimary.length * partCPerQ} marks)`;
    }

    let subjectDisplay = '________________';
    if (metadata.subjectCode && metadata.subjectName) {
        subjectDisplay = `${metadata.subjectCode} / ${metadata.subjectName}`;
    } else if (metadata.subjectCode) {
        subjectDisplay = metadata.subjectCode;
    } else if (metadata.subjectName) {
        subjectDisplay = metadata.subjectName;
    }

    let html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Internal Assessment ${iaType.replace('IA', '')}</title>
    <style>
        @page {
            size: A4;
            margin: 15mm 12mm 15mm 12mm;
        }
        * { box-sizing: border-box; }
        body {
            font-family: 'Times New Roman', 'Calibri', serif;
            margin: 0;
            padding: 0;
            line-height: 1.3;
            color: #000;
            font-size: 10pt;
            background: #d0d0d0;
        }

        /* A4 page container */
        .page-container {
            width: 210mm;
            min-height: 297mm;
            margin: 10mm auto;
            padding: 10mm 12mm;
            background: #fff;
            box-shadow: 0 2px 12px rgba(0,0,0,0.3);
            position: relative;
        }

        /* Editable cell highlight */
        [contenteditable="true"] {
            outline: none;
            transition: background-color 0.2s;
        }
        [contenteditable="true"]:hover {
            background-color: #f8f8ff;
        }
        [contenteditable="true"]:focus {
            background-color: #eef0ff;
            outline: 2px solid #4a90d9;
            outline-offset: -1px;
        }

        h1, h2, h3, h4 { margin: 3px 0; font-weight: bold; }
        h1 { font-size: 13pt; }
        strong, b { font-weight: bold; }
        em, i { font-style: italic; }
        u { text-decoration: underline; }

        ul, ol { margin: 2px 0; padding-left: 16px; }
        li { margin: 1px 0; }
        p { margin: 2px 0; line-height: 1.3; }

        img {
            max-width: 100%;
            height: auto;
            display: block;
            margin: 3px 0;
        }

        .question-content table { border-collapse: collapse; margin: 3px 0; }
        .question-content td, .question-content th { border: 1px solid #000; padding: 2px 4px; font-size: 10pt; }

        svg { max-width: 100%; height: auto; }

        /* Main header table */
        .header-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 0;
            border: 1px solid black;
        }
        .header-table td {
            padding: 3px 6px;
            border: 1px solid black;
            vertical-align: middle;
            font-size: 10pt;
        }

        /* Metadata table */
        .metadata-table {
            width: 100%;
            border-collapse: collapse;
            border: 1px solid black;
        }
        .metadata-table td {
            padding: 3px 6px;
            border: 1px solid black;
            font-size: 10pt;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 2px;
        }
        th, td {
            border: 1px solid black;
            padding: 3px 4px;
            text-align: left;
            font-size: 10pt;
        }
        th {
            background-color: #f0f0f0;
            font-weight: bold;
            text-align: center;
        }

        .question-content {
            line-height: 1.3;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }

        .footer {
            text-align: center;
            margin-top: 10px;
            font-weight: bold;
            font-size: 10pt;
        }

        .section-note {
            text-align: center;
            font-size: 9pt;
            margin: 3px 0 2px 0;
        }

        /* ===== Floating Toolbar ===== */
        .toolbar {
            position: fixed;
            top: 10px;
            right: 15px;
            display: flex;
            gap: 8px;
            z-index: 10000;
            background: #ffffff;
            padding: 8px 14px;
            border-radius: 10px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            border: 1px solid #ddd;
            align-items: center;
        }
        .toolbar-label {
            font-size: 11px;
            color: #666;
            margin-right: 4px;
            font-family: 'Segoe UI', Arial, sans-serif;
        }
        .toolbar button {
            padding: 7px 16px;
            border: 1px solid #ccc;
            border-radius: 6px;
            cursor: pointer;
            font-size: 12px;
            font-family: 'Segoe UI', Arial, sans-serif;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 5px;
            transition: all 0.2s;
        }
        .toolbar button:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 8px rgba(0,0,0,0.12);
        }
        .btn-save {
            background: #2563eb;
            color: white;
            border-color: #1d4ed8 !important;
        }
        .btn-save:hover { background: #1d4ed8; }
        .btn-print {
            background: #7c3aed;
            color: white;
            border-color: #6d28d9 !important;
        }
        .btn-print:hover { background: #6d28d9; }

        .edit-hint {
            position: fixed;
            bottom: 15px;
            left: 50%;
            transform: translateX(-50%);
            background: #1e293b;
            color: #fff;
            padding: 8px 20px;
            border-radius: 8px;
            font-size: 12px;
            font-family: 'Segoe UI', Arial, sans-serif;
            z-index: 10000;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
            opacity: 1;
            transition: opacity 1s;
        }

        /* Print styles */
        @media print {
            body { background: white; margin: 0; padding: 0; }
            .page-container {
                width: 100%;
                min-height: auto;
                margin: 0;
                padding: 5mm 8mm;
                box-shadow: none;
            }
            .toolbar, .edit-hint { display: none !important; }
            [contenteditable="true"]:focus { outline: none; background-color: transparent; }
            [contenteditable="true"]:hover { background-color: transparent; }
        }
    </style>
</head>
<body>

<!-- Floating Toolbar -->
<div class="toolbar" id="toolbar">
    <span class="toolbar-label">\u270f\ufe0f Edit Mode</span>
    <button class="btn-save" onclick="saveAsHtml()" title="Save your edits as HTML">
        <span>\ud83d\udcbe</span> Save
    </button>
    <button class="btn-print" onclick="printDocument()" title="Print the document">
        <span>\ud83d\udda8\ufe0f</span> Print
    </button>
</div>

<!-- Edit Hint -->
<div class="edit-hint" id="editHint">
    \u270f\ufe0f Click on any cell to edit it directly. Use the toolbar at the top-right to Save or Print.
</div>

<div class="page-container" id="paperContent">`;

    // ---- HEADER SECTION ----
    // Two-column layout: logo + institute name (NO empty right box)
    html += `<table class="header-table">
        <tr>
            <td colspan="2" style="text-align: right; padding-right: 15px; border-bottom: 1px solid black;" contenteditable="true">
                <strong>Reg. No. ____________________</strong>
            </td>
        </tr>
        <tr>
            <td style="width: 15%; text-align: center; padding: 8px; border-right: 1px solid black;">
                ${logoUrl ? `<img src="${logoUrl}" style="width: 70px; height: 70px;" alt="CIT Logo"/>` : ''}
            </td>
            <td style="width: 85%; text-align: center; padding: 5px;" contenteditable="true">
                <h1 style="margin: 2px 0; font-size: 12pt;">CHENNAI INSTITUTE OF TECHNOLOGY</h1>
                <p style="margin: 1px 0; font-size: 10pt;"><strong>Autonomous</strong></p>
                <p style="margin: 1px 0; font-size: 9pt;">Sarathy Nagar, Pudupedu, Chennai – 600 069.</p>
                <p style="margin: 5px 0 0 0; font-size: 11pt; font-weight: bold;">Internal Assessment ${iaType.replace('IA', '')}</p>
            </td>
        </tr>
    </table>`;

    // ---- METADATA TABLE ----
    html += `<table class="metadata-table">
        <tr>
            <td style="width: 65%;" contenteditable="true">Date: ${metadata.date || '________________'}</td>
            <td style="width: 15%; text-align: center;">Max. Marks</td>
            <td style="width: 20%; text-align: center;" contenteditable="true">${metadata.maxMarks}</td>
        </tr>
        <tr>
            <td contenteditable="true">Subject Code / Name: ${subjectDisplay}</td>
            <td style="text-align: center;">Time</td>
            <td style="text-align: center;" contenteditable="true">${metadata.time}</td>
        </tr>
        <tr>
            <td contenteditable="true">Branch: ${metadata.branch || '________________'}</td>
            <td style="text-align: center;">Year / Sem</td>
            <td style="text-align: center;" contenteditable="true">${metadata.yearSem || '_____/_____'}</td>
        </tr>
    </table>`;

    // RBT verb mapping for detecting level from CO description
    const rbtVerbsForCO: Record<string, string[]> = {
        'L1': ['define', 'list', 'state', 'identify', 'name', 'recall', 'recognize', 'label', 'memorize', 'enumerate'],
        'L2': ['explain', 'describe', 'discuss', 'summarize', 'interpret', 'classify', 'compare', 'contrast', 'illustrate', 'understand'],
        'L3': ['apply', 'demonstrate', 'calculate', 'solve', 'implement', 'use', 'execute', 'compute', 'derive'],
        'L4': ['analyze', 'examine', 'investigate', 'categorize', 'distinguish', 'organize', 'break down', 'outline', 'differentiate'],
        'L5': ['evaluate', 'justify', 'assess', 'critique', 'judge', 'defend', 'validate', 'appraise'],
        'L6': ['design', 'create', 'develop', 'formulate', 'propose', 'plan', 'invent', 'compose', 'build']
    };

    const detectRBTFromText = (text: string): string => {
        const lower = text.toLowerCase().trim();
        // Check from highest to lowest
        for (const level of ['L6', 'L5', 'L4', 'L3', 'L2', 'L1']) {
            for (const verb of rbtVerbsForCO[level]) {
                // Match word boundary to avoid partial matches (e.g. "analyze" vs "analyzer")
                const regex = new RegExp('\\b' + verb + '\\b', 'i');
                if (regex.test(lower)) {
                    return level;
                }
            }
        }
        return '';
    };

    // ---- COURSE OBJECTIVES ----
    // If courseObjectives is empty, use courseOutcomes descriptions as objectives
    const objectives = (metadata.courseObjectives && metadata.courseObjectives.length > 0)
        ? metadata.courseObjectives
        : (metadata.courseOutcomes && metadata.courseOutcomes.length > 0)
            ? metadata.courseOutcomes.map(co => co.description)
            : [];

    if (objectives.length > 0) {
        html += `<p style="text-align: center; font-weight: bold; font-size: 10pt; margin: 10px 0 4px 0;">Course Objectives:</p>
        <table>
            <tr style="background-color: #f0f0f0;">
                <th style="width: 12%; font-size: 10pt;">CO.NO</th>
                <th style="width: 88%; font-size: 10pt;">Course Objectives</th>
            </tr>`;
        objectives.forEach((obj: string, idx: number) => {
            html += `<tr>
                <td style="text-align: center;" contenteditable="true">${idx + 1}</td>
                <td contenteditable="true">${obj}</td>
            </tr>`;
        });
        html += `</table>`;
    }

    // ---- COURSE OUTCOMES (with RBT Level column) ----
    if (metadata.courseOutcomes && metadata.courseOutcomes.length > 0) {
        html += `<p class="section-note" style="margin-top: 8px;">On Completion of the course the students will be able to</p>
        <table>
            <tr style="background-color: #f0f0f0;">
                <th style="width: 12%; font-size: 10pt;">CO.NO</th>
                <th style="width: 74%; font-size: 10pt;">Course Outcomes</th>
                <th style="width: 14%; font-size: 10pt;">RBT Level</th>
            </tr>`;
        metadata.courseOutcomes.forEach((co: any) => {
            // Use provided level, or detect from description text
            let rbtLevel = co.level || '';
            if (!rbtLevel) {
                // Detect from action verbs in the CO description
                rbtLevel = detectRBTFromText(co.description || '');
            }
            html += `<tr>
                <td style="text-align: center;" contenteditable="true">${co.co}</td>
                <td contenteditable="true">${co.description}</td>
                <td style="text-align: center;" contenteditable="true">${rbtLevel}</td>
            </tr>`;
        });
        html += `</table>`;
    }

    // ---- PART A (with spacing from previous section) ----
    html += `<div style="margin-top: 10px;"></div>`;
    html += await createQuestionsTableHtml(partA, partALabel, pdfPageImages);

    // ---- PART B (with spacing from Part A) ----
    html += `<div style="margin-top: 10px;"></div>`;
    html += await createQuestionsTableHtml(partB, partBLabel, pdfPageImages);

    // ---- PART C (with spacing from Part B) ----
    if (partC && partC.length > 0) {
        html += `<div style="margin-top: 10px;"></div>`;
        html += await createQuestionsTableHtml(partC, partCLabel, pdfPageImages);
    }

    // Footer
    html += `<div class="footer">-* All the Best *-</div>`;

    // Close page container
    html += `</div>`;

    // JavaScript for toolbar actions
    html += `
<script>
    // Auto-hide edit hint after 5 seconds
    setTimeout(function() {
        var hint = document.getElementById('editHint');
        if (hint) {
            hint.style.opacity = '0';
            setTimeout(function() { hint.style.display = 'none'; }, 1000);
        }
    }, 5000);

    // Save as HTML (preserves edits)
    function saveAsHtml() {
        var toolbar = document.getElementById('toolbar');
        var hint = document.getElementById('editHint');
        toolbar.style.display = 'none';
        if (hint) hint.style.display = 'none';

        var content = '<!DOCTYPE html>' + document.documentElement.outerHTML;

        toolbar.style.display = 'flex';

        var blob = new Blob([content], { type: 'text/html' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = document.title.replace(/\\s+/g, '_') + '_Edited.html';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast('\\u2705 Saved successfully!');
    }

    // Print
    function printDocument() {
        window.print();
    }

    // Toast notification
    function showToast(msg) {
        var existing = document.getElementById('saveToast');
        if (existing) existing.remove();

        var toast = document.createElement('div');
        toast.id = 'saveToast';
        toast.textContent = msg;
        toast.style.cssText = 'position:fixed;bottom:20px;right:20px;background:#1e293b;color:#fff;padding:10px 24px;border-radius:8px;font-size:13px;font-family:Segoe UI,Arial,sans-serif;z-index:99999;box-shadow:0 4px 15px rgba(0,0,0,0.2);transition:opacity 0.5s;';
        document.body.appendChild(toast);
        setTimeout(function() { toast.style.opacity = '0'; }, 2000);
        setTimeout(function() { toast.remove(); }, 2500);
    }
</script>

</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html' });
    return blob;
};
