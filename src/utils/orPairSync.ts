import { Question } from '@/components/QuestionDisplay';

/**
 * Synchronize OR pair levels in Part B (and optionally Part C) questions.
 * When one question in an OR pair (e.g., Q13A/Q13B) has a different level
 * from its partner, both are set to the higher level.
 * This ensures OR choice questions always display matching RBT levels
 * in the UI, preview, and downloaded documents.
 */
export function syncOrPairLevels(questions: Question[]): Question[] {
    // Build a map: base question number → highest level among pair members
    const orPairHighest = new Map<number, string>();

    questions.forEach(q => {
        const m = q.questionNumber.match(/(\d+)\s*([ABab])/);
        if (m && m[2]) {
            const base = parseInt(m[1]);
            const existing = orPairHighest.get(base);
            if (existing) {
                const existingNum = parseInt(existing.replace('L', ''));
                const currentNum = parseInt(q.detectedLevel.replace('L', ''));
                if (currentNum > existingNum) {
                    orPairHighest.set(base, q.detectedLevel);
                }
            } else {
                orPairHighest.set(base, q.detectedLevel);
            }
        }
    });

    // Apply: set both OR pair members to the higher level
    return questions.map(q => {
        const m = q.questionNumber.match(/(\d+)\s*([ABab])/);
        if (m && m[2]) {
            const syncLevel = orPairHighest.get(parseInt(m[1]));
            if (syncLevel && syncLevel !== q.detectedLevel) {
                return { ...q, detectedLevel: syncLevel, expectedLevel: syncLevel };
            }
        }
        return q;
    });
}
