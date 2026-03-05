// Local question fixer — works without Supabase edge functions
// Uses rule-based verb swapping to match target RBT levels

const RBT_VERBS: Record<string, string[]> = {
    L1: ['Identify', 'Label', 'List', 'State', 'Define', 'Name', 'Recognize'],
    L2: ['Describe', 'Compare', 'Restate', 'Illustrate', 'Explain', 'Summarize', 'Interpret'],
    L3: ['Calculate', 'Apply', 'Solve', 'Demonstrate', 'Compute', 'Use', 'Determine'],
    L4: ['Break down', 'Categorize', 'Outline', 'Simplify', 'Analyze', 'Differentiate', 'Examine'],
    L5: ['Select', 'Rate', 'Defend', 'Verify', 'Evaluate', 'Justify', 'Critique'],
    L6: ['Plan', 'Modify', 'Propose', 'Develop', 'Design', 'Create', 'Construct']
};

// All known verbs flattened for detection
const ALL_VERBS = Object.values(RBT_VERBS).flat();

function pickRandomVerb(level: string): string {
    const verbs = RBT_VERBS[level] || RBT_VERBS.L3;
    return verbs[Math.floor(Math.random() * verbs.length)];
}

function stripHtmlTags(html: string): string {
    return html.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Fixes a question locally by swapping/prepending the action verb to match the target level.
 */
export function fixQuestionLocally(
    questionText: string,
    _currentLevel: string,
    targetLevel: string,
    convertToSingle: boolean = false
): string {
    const plain = stripHtmlTags(questionText);
    const targetVerb = pickRandomVerb(targetLevel);

    if (convertToSingle) {
        // Remove subdivision markers and merge into a single question
        const cleaned = plain
            .replace(/\b(i+\)|[ivx]+\)|[a-d]\)|[a-d]\.)\s*/gi, '')
            .replace(/\s+/g, ' ')
            .trim();

        // Remove any existing leading verb
        let body = cleaned;
        for (const verb of ALL_VERBS) {
            const regex = new RegExp(`^${verb}\\b\\s*`, 'i');
            if (regex.test(body)) {
                body = body.replace(regex, '').trim();
                break;
            }
        }
        return `${targetVerb} ${body}`;
    }

    // Standard level change: replace the leading action verb
    let body = plain;
    let verbReplaced = false;

    for (const verb of ALL_VERBS) {
        const regex = new RegExp(`^${verb}\\b\\s*`, 'i');
        if (regex.test(body)) {
            body = body.replace(regex, '').trim();
            verbReplaced = true;
            break;
        }
    }

    if (!verbReplaced) {
        // No known verb found at start — just prepend the target verb
        // Try to lowercase the first character to flow naturally
        if (body.length > 0) {
            body = body.charAt(0).toLowerCase() + body.slice(1);
        }
    }

    return `${targetVerb} ${body}`;
}

/**
 * Fixes a question by using the fast local fixer.
 * Optionally tries the Supabase edge function with a short timeout for a better result.
 * Local fix is ALWAYS used as the primary — edge function is a "nice to have" upgrade.
 */
export async function fixQuestionWithFallback(
    supabase: any,
    questionText: string,
    currentLevel: string,
    targetLevel: string,
    convertToSingle: boolean,
    subdivisionLevels?: string[],
    originalHtml?: string
): Promise<{ fixedQuestion: string; usedFallback: boolean }> {
    // Always produce a local fix immediately (instant, no network)
    const localFixed = fixQuestionLocally(questionText, currentLevel, targetLevel, convertToSingle);

    // Try the edge function with a short timeout (2 seconds max)
    // If it succeeds quickly, use the better AI result; otherwise use local fix
    try {
        const edgePromise = supabase.functions.invoke('fix-question', {
            body: {
                questionText,
                currentLevel,
                targetLevel,
                convertToSingle,
                subdivisionLevels
            }
        });

        // Race: edge function vs 2-second timeout
        const timeoutPromise = new Promise<null>((resolve) => setTimeout(() => resolve(null), 2000));
        const result = await Promise.race([edgePromise, timeoutPromise]);

        if (result && !result.error && result.data?.fixedQuestion) {
            return { fixedQuestion: result.data.fixedQuestion, usedFallback: false };
        }
    } catch {
        // Edge function failed — use local fix (already computed)
    }

    return { fixedQuestion: localFixed, usedFallback: true };
}
