import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, ImageRun, BorderStyle } from 'docx';
import { ParseResult } from './parser/htmlParser';
import { Question } from '@/components/QuestionDisplay';
import { IMAGE_SRC_REGEX, stripHtmlForAI } from '@/utils/htmlUtils';
import { containsMathExpression, renderMathToImage, extractMathExpressions } from '@/utils/mathRenderer';
import { extractAllImages, base64ToUint8Array, fetchImageAsBuffer, scaleDimensions, convertSvgToPng } from '@/utils/imageExtractor';
import { CIT_LOGO_BASE64 } from '@/assets/citLogoBase64';


// Helper: get actual image dimensions from data URI or blob URL
const getImageDimensions = (dataUrl: string): Promise<{ width: number; height: number }> => {
    return new Promise((resolve) => {
        const img = new Image();
        img.onload = () => {
            // Ensure we never return 0 for dimensions to avoid invisible 'boxes'
            resolve({
                width: Math.max(img.naturalWidth, 20),
                height: Math.max(img.naturalHeight, 20)
            });
        };
        img.onerror = () => {
            resolve({ width: 450, height: 300 }); // fallback
        };
        img.src = dataUrl;
    });
};

// Helper: scale dimensions to fit within maxWidth while preserving aspect ratio
const scaleToFit = (width: number, height: number, maxWidth: number = 450): { width: number; height: number } => {
    return scaleDimensions(width, height, maxWidth, 600);
};

// CIT logo as an ArrayBuffer from the pre-encoded Base64
// No fetch needed — baked into the JS bundle at build time
const fetchLogo = async (): Promise<ArrayBuffer> => {
    try {
        const base64Data = CIT_LOGO_BASE64.split(',')[1];
        if (!base64Data) return new ArrayBuffer(0);
        const binaryString = atob(base64Data);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes.buffer;
    } catch (e) {
        console.warn('Could not decode CIT logo, using empty buffer');
        return new ArrayBuffer(0);
    }
};

export const generateWordDocument = async (parseResult: ParseResult, iaType: 'IA1' | 'IA2' | 'IA3'): Promise<Blob> => {
    const logoBuffer = await fetchLogo();
    const { metadata, partA, partB, partC, pdfPageImages } = parseResult;

    let subjectDisplay = '________________';
    if (metadata.subjectCode && metadata.subjectName) {
        subjectDisplay = `${metadata.subjectCode} / ${metadata.subjectName}`;
    } else if (metadata.subjectCode) {
        subjectDisplay = metadata.subjectCode;
    } else if (metadata.subjectName) {
        subjectDisplay = metadata.subjectName;
    }

    const doc = new Document({
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: 720, // 0.5 inch
                            right: 720,
                            bottom: 720,
                            left: 720,
                        },
                    },
                },
                children: [
                    // Header Table
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        columnSpan: 3,
                                        children: [
                                            new Paragraph({
                                                alignment: AlignmentType.RIGHT,
                                                children: [
                                                    new TextRun({ text: "Reg. No. ", bold: true }),
                                                    new TextRun({ text: "____________________" }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: { size: 15, type: WidthType.PERCENTAGE },
                                        verticalAlign: 'center',
                                        children: [
                                            logoBuffer.byteLength > 0 ? new Paragraph({
                                                alignment: AlignmentType.CENTER,
                                                children: [
                                                    new ImageRun({
                                                        data: logoBuffer,
                                                        transformation: { width: 80, height: 80 },
                                                        type: "png",
                                                    }),
                                                ],
                                            }) : new Paragraph({ text: "" }),
                                        ],
                                    }),
                                    new TableCell({
                                        width: { size: 70, type: WidthType.PERCENTAGE },
                                        verticalAlign: 'center',
                                        children: [
                                            new Paragraph({
                                                alignment: AlignmentType.CENTER,
                                                children: [new TextRun({ text: "CHENNAI INSTITUTE OF TECHNOLOGY", bold: true, size: 24 })],
                                            }),
                                            new Paragraph({
                                                alignment: AlignmentType.CENTER,
                                                children: [new TextRun({ text: "Autonomous", bold: true, size: 20 })],
                                            }),
                                            new Paragraph({
                                                alignment: AlignmentType.CENTER,
                                                children: [new TextRun({ text: "Sarathy Nagar, Kundrathur, Chennai – 600 069.", size: 18 })],
                                            }),
                                            new Paragraph({
                                                alignment: AlignmentType.CENTER,
                                                spacing: { before: 100 },
                                                children: [new TextRun({ text: `Internal Assessment ${iaType.replace('IA', '')}`, bold: true, size: 22 })],
                                            }),
                                        ],
                                    }),
                                    new TableCell({
                                        width: { size: 15, type: WidthType.PERCENTAGE },
                                        children: [new Paragraph({ text: "" })],
                                    }),
                                ],
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        columnSpan: 3,
                                        children: [
                                            new Table({
                                                width: { size: 100, type: WidthType.PERCENTAGE },
                                                borders: {
                                                    top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                    bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                    left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                    right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                    insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                    insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                                },
                                                rows: [
                                                    new TableRow({
                                                        children: [
                                                            new TableCell({
                                                                children: [new Paragraph({ children: [new TextRun({ text: "Date: " + (metadata.date || "________________"), bold: true })] })],
                                                            }),
                                                            new TableCell({
                                                                children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Max. Marks: " + metadata.maxMarks, bold: true })] })],
                                                            }),
                                                        ],
                                                    }),
                                                    new TableRow({
                                                        children: [
                                                            new TableCell({
                                                                children: [new Paragraph({ children: [new TextRun({ text: "Subject Code/Name: " + subjectDisplay, bold: true })] })],
                                                            }),
                                                            new TableCell({
                                                                children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Time: " + metadata.time, bold: true })] })],
                                                            }),
                                                        ],
                                                    }),
                                                    new TableRow({
                                                        children: [
                                                            new TableCell({
                                                                children: [new Paragraph({ children: [new TextRun({ text: "Branch: " + (metadata.branch || "________________"), bold: true })] })],
                                                            }),
                                                            new TableCell({
                                                                children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Year/Sem: " + (metadata.yearSem || "_____/_____"), bold: true })] })],
                                                            }),
                                                        ],
                                                    }),
                                                ],
                                            }),
                                        ],
                                    }),
                                ],
                            }),
                        ],
                    }),

                    new Paragraph({ text: "" }),

                    // Course Objectives
                    ...(metadata.courseObjectives && metadata.courseObjectives.length > 0 ? [
                        new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 200, after: 100 },
                            children: [new TextRun({ text: "Course Objectives:", bold: true, size: 22 })],
                        }),
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            borders: {
                                insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                                insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                            },
                            rows: [
                                new TableRow({
                                    tableHeader: true,
                                    children: [
                                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "CO.NO", bold: true })], alignment: AlignmentType.CENTER })] }),
                                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Course Objectives", bold: true })], alignment: AlignmentType.CENTER })] }),
                                    ],
                                }),
                                ...metadata.courseObjectives.map((obj, idx) =>
                                    new TableRow({
                                        children: [
                                            new TableCell({ children: [new Paragraph({ text: String(idx + 1), alignment: AlignmentType.CENTER })] }),
                                            new TableCell({ children: [new Paragraph({ text: obj })] }),
                                        ],
                                    })
                                ),
                            ],
                        }),
                    ] : []),

                    // Course Outcomes
                    ...(metadata.courseOutcomes && metadata.courseOutcomes.length > 0 ? [
                        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 100 }, children: [new TextRun({ text: "Course Outcomes:", bold: true, size: 22 })] }),
                        new Table({
                            width: { size: 100, type: WidthType.PERCENTAGE },
                            borders: { insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" }, insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" } },
                            rows: [
                                new TableRow({
                                    tableHeader: true,
                                    children: [
                                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "CO.NO", bold: true })], alignment: AlignmentType.CENTER })] }),
                                        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Course Outcomes", bold: true })], alignment: AlignmentType.CENTER })] }),
                                    ],
                                }),
                                ...metadata.courseOutcomes.map((co: any) =>
                                    new TableRow({
                                        children: [
                                            new TableCell({ children: [new Paragraph({ text: co.co, alignment: AlignmentType.CENTER })] }),
                                            new TableCell({ children: [new Paragraph({ text: co.description })] }),
                                        ],
                                    })
                                ),
                            ],
                        }),
                    ] : []),

                    // Part A
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 }, children: [new TextRun({ text: "PART – A (5 x 2 = 10 Marks)", bold: true })] }),
                    await createQuestionsTable(partA, pdfPageImages),

                    // Part B
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 }, children: [new TextRun({ text: "PART – B Marks", bold: true })] }),
                    await createQuestionsTable(partB, pdfPageImages),

                    // Part C
                    ...(partC && partC.length > 0 ? [
                        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 200, after: 200 }, children: [new TextRun({ text: "PART – C (1 x 15 = 15 Marks)", bold: true })] }),
                        await createQuestionsTable(partC, pdfPageImages)
                    ] : []),

                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 400 }, children: [new TextRun({ text: "* All the Best *", bold: true })] }),
                ],
            },
        ],
    });

    return await Packer.toBlob(doc);
};

const extractImagesFromHtml = async (html: string, questionNumber: string) => {
    const images: { src: string; isBase64: boolean; data?: string; type?: string; buffer?: Uint8Array }[] = [];
    if (!html) return images;

    try {
        const extractedImages = extractAllImages(html);
        console.log(`[DocGen] Q${questionNumber}: Found ${extractedImages.length} images via extractor`);

        for (const extracted of extractedImages) {
            try {
                if (extracted.isBase64 && extracted.base64Data) {
                    const type = extracted.type || 'png';
                    // Always try to normalize SVG or unsupported types to PNG for Word compatibility
                    if (type === 'svg' || type === 'webp' || extracted.src.includes('svg') || extracted.src.includes('webp')) {
                        console.log(`[DocGen] Q${questionNumber}: Normalizing ${type} to PNG for Word...`);
                        try {
                            const svgData = extracted.base64Data.includes('<svg')
                                ? extracted.base64Data
                                : decodeURIComponent(escape(atob(extracted.base64Data)));
                            const png = await convertSvgToPng(svgData, 600, 450); // Higher res for better quality
                            if (png) {
                                images.push({ src: extracted.src, isBase64: true, type: 'png', data: png.base64 });
                                continue;
                            }
                        } catch (svgErr) {
                            console.error(`[DocGen] Failed normalizing ${type} for Q${questionNumber}:`, svgErr);
                        }
                    }
                    // Map jpeg to jpg for docx library
                    const finalType = (type === 'jpeg' || type === 'jpg') ? 'jpg' : type;
                    images.push({ src: extracted.src, isBase64: true, type: finalType as any, data: extracted.base64Data });
                } else if (!extracted.isBase64 && extracted.src) {
                    console.log(`[DocGen] Q${questionNumber}: Fetching remote/local image ${extracted.src.substring(0, 50)}...`);
                    const res = await fetchImageAsBuffer(extracted.src);
                    if (res && res.buffer.byteLength > 0) {
                        // Map jpeg to jpg for docx library
                        const finalType = (res.type === 'jpeg') ? 'jpg' : res.type;
                        images.push({ src: extracted.src, isBase64: false, buffer: res.buffer, type: finalType });
                    }
                }
            } catch (err) {
                console.error(`[DocGen] Failed processing image in Q${questionNumber}:`, err);
            }
        }
    } catch (e) {
        console.error(`[DocGen] Extraction loop error Q${questionNumber}:`, e);
    }
    return images;
};

const createQuestionsTable = async (questions: Question[], pdfPageImages?: (string | null)[]): Promise<Table> => {
    console.log(`[DocGen] Creating table for ${questions.length} questions`);

    const questionRowsData = await Promise.all(questions.map(async (q) => {
        // Source selection: prefer originalHtml if available as it's the most raw data including images
        const sources = [q.originalHtml, q.text].filter(Boolean) as string[];

        // Extract images from ALL sources to ensure none are missed (e.g. from subdivision rows)
        const images: any[] = [];
        const seenSrcs = new Set<string>();
        for (const source of sources) {
            const extracted = await extractImagesFromHtml(source, q.questionNumber);
            for (const img of extracted) {
                if (!seenSrcs.has(img.src)) {
                    images.push(img);
                    seenSrcs.add(img.src);
                }
            }
        }

        // Clean text for the doc paragraph, stripping out images and math
        const textToClean = (q.isFixed || q.isEnhanced || !q.originalHtml) ? (q.text || '') : q.originalHtml;
        const cleanText = stripHtmlForAI(textToClean)
            .replace(/<img[^>]*>/gi, '')
            .replace(/&lt;math[\s\S]*?&gt;/gi, '')
            .replace(/<math[\s\S]*?<\/math>/gi, '')
            .replace(/\s+/g, ' ')
            .trim();

        // Math handling
        const mathImages: any[] = [];
        const mathSource = q.originalHtml || q.text || '';
        if (containsMathExpression(mathSource)) {
            try {
                const exprs = extractMathExpressions(mathSource);
                for (const e of exprs) {
                    const url = await renderMathToImage(e.expression, 22);
                    if (url) {
                        const dims = await getImageDimensions(url);
                        const scaled = scaleToFit(dims.width, dims.height, 400);
                        const b64 = url.split(',')[1];
                        if (b64) {
                            const bin = atob(b64);
                            const buf = new Uint8Array(bin.length);
                            for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
                            mathImages.push({ buffer: buf, width: scaled.width, height: scaled.height });
                        }
                    }
                }
            } catch (e) { console.warn(`[DocGen] Math error Q${q.questionNumber}:`, e); }
        }

        return { q, cleanText, images, mathImages };
    }));

    const finalRows = await Promise.all(questionRowsData.map(async ({ q, cleanText, images, mathImages }) => {
        const children: Paragraph[] = [new Paragraph({ text: cleanText })];

        // Add math images
        for (const m of mathImages) {
            try {
                children.push(new Paragraph({
                    children: [new ImageRun({
                        data: m.buffer,
                        transformation: { width: m.width, height: m.height },
                        type: "png"
                    })]
                }));
            } catch (e) { console.error(`[DocGen] Math img embed fail Q${q.questionNumber}:`, e); }
        }

        // Add question images
        for (const img of images) {
            try {
                let data: Uint8Array | undefined;
                let width = 400, height = 300;
                let finalType: any = 'png';

                if (img.isBase64 && img.data) {
                    // Extract exactly the base64 data regardless of any complex header before it
                    const b64Parts = img.data.replace(/\s/g, '').split(/base64,/i);
                    const cleanB64 = b64Parts.length > 1 ? b64Parts[1] : b64Parts[0];
                    const bin = atob(cleanB64);
                    data = new Uint8Array(bin.length);
                    for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);

                    const url = `data:image/${img.type || 'png'};base64,${cleanB64}`;
                    const dims = await getImageDimensions(url).catch(() => ({ width: 450, height: 300 }));
                    const scaled = scaleToFit(dims.width, dims.height, 450);
                    width = scaled.width; height = scaled.height;

                    // Map common types to docx-supported strings
                    const supportedTypes = ['jpg', 'png', 'gif', 'bmp', 'wmf', 'emf'];
                    const mappedType = img.type === 'jpeg' ? 'jpg' : img.type;
                    finalType = supportedTypes.includes(mappedType || '') ? mappedType : 'png';
                } else if (img.buffer) {
                    data = img.buffer;
                    const supportedTypes = ['jpg', 'png', 'gif', 'bmp', 'wmf', 'emf'];
                    const mappedType = img.type === 'jpeg' ? 'jpg' : img.type;
                    finalType = supportedTypes.includes(mappedType || '') ? mappedType : 'png';
                    const scaled = scaleToFit(450, 300, 450);
                    width = scaled.width; height = scaled.height;
                }

                if (data && data.length > 10) {
                    children.push(new Paragraph({
                        children: [new ImageRun({
                            data,
                            transformation: { width: Math.round(width), height: Math.round(height) },
                            type: finalType as any
                        })]
                    }));
                    console.log(`[DocGen] Q${q.questionNumber}: Successfully embedded image (${finalType}, ${data.length} bytes)`);
                }
            } catch (e) {
                console.error(`[DocGen] Img embed fail Q${q.questionNumber}:`, e);
            }
        }

        // PDF fallback logic
        if (images.length === 0 && mathImages.length === 0 && pdfPageImages?.length) {
            const pageMatch = ((q.text || '') + (q.originalHtml || '')).match(/data-page="(\d+)"/);
            if (pageMatch) {
                const pageIdx = parseInt(pageMatch[1]);
                const url = pdfPageImages[pageIdx];
                if (url) {
                    console.log(`[DocGen] Q${q.questionNumber}: Using PDF fallback page ${pageIdx}`);
                    try {
                        const b64 = url.split(',')[1];
                        const bin = atob(b64);
                        const buf = new Uint8Array(bin.length);
                        for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
                        const dims = await getImageDimensions(url);
                        const scaled = scaleToFit(dims.width, dims.height, 450);
                        children.push(new Paragraph({
                            children: [new ImageRun({
                                data: buf,
                                transformation: { width: scaled.width, height: scaled.height },
                                type: "png"
                            })]
                        }));
                    } catch (e) { console.error(`[DocGen] PDF fallback fail Q${q.questionNumber}:`, e); }
                }
            }
        }

        return { q, children };
    }));

    const tableRows: TableRow[] = [];
    let prevNum = -1;
    for (const { q, children } of finalRows) {
        const m = q.questionNumber.match(/(\d+)/);
        const num = m ? parseInt(m[1]) : -1;
        if (q.questionNumber.toUpperCase().includes('B') && num === prevNum) {
            tableRows.push(new TableRow({ children: [new TableCell({ columnSpan: 5, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "OR", bold: true })] })] })] }));
        }
        prevNum = num;
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ text: q.questionNumber, alignment: AlignmentType.CENTER })] }),
                new TableCell({ children }),
                new TableCell({ children: [new Paragraph({ text: q.co, alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ text: q.detectedLevel, alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ text: q.marks.toString(), alignment: AlignmentType.CENTER })] }),
            ]
        }));
    }

    return new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
        },
        rows: [
            new TableRow({
                tableHeader: true,
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Q.No", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 8, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Questions", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 62, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "CO", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "RBT Level", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Marks", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 10, type: WidthType.PERCENTAGE } }),
                ],
            }),
            ...tableRows,
        ],
    });
};
