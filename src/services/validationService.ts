import { Question } from '@/components/QuestionDisplay';
import { stripHtmlForAI } from '@/utils/htmlUtils';
import { getLevelNumber, getHigherLevel } from './parser/htmlParser';

export interface PartAnalysis {
  part: string;
  totalQuestions: number;
  levelCounts: Record<string, number>;
  levelPercentages: Record<string, number>;
  errors: ValidationError[];
  isValid: boolean;
}

export interface ValidationResult {
  status: 'accepted' | 'rejected';
  errors: ValidationError[];
  levelDistribution: {
    l1l2: number;
    l3: number;
    l4l5l6: number;
  };
  partAnalysis: {
    partA: PartAnalysis;
    partB: PartAnalysis;
    partC?: PartAnalysis;
  };
  allErrorsFixed: boolean;
}

export interface ValidationError {
  part: string;
  questionNumber: string;
  issue: string;
  suggestion: string;
}

// Count questions by level
export const countLevels = (questions: Question[]): Record<string, number> => {
  const counts: Record<string, number> = { L1: 0, L2: 0, L3: 0, L4: 0, L5: 0, L6: 0 };
  questions.forEach(q => {
    if (counts[q.detectedLevel] !== undefined) {
      counts[q.detectedLevel]++;
    }
  });
  return counts;
};

// Calculate percentage for each level
export const calculateLevelPercentages = (counts: Record<string, number>, total: number): Record<string, number> => {
  const percentages: Record<string, number> = {};
  for (const level in counts) {
    percentages[level] = total > 0 ? Math.round((counts[level] / total) * 100) : 0;
  }
  return percentages;
};

// Calculate distribution summary
const calculateDistribution = (questions: Question[]): { l1l2: number; l3: number; l4l5l6: number } => {
  const counts = countLevels(questions);
  const total = questions.length;

  if (total === 0) return { l1l2: 0, l3: 0, l4l5l6: 0 };

  return {
    l1l2: Math.round(((counts.L1 + counts.L2) / total) * 100),
    l3: Math.round((counts.L3 / total) * 100),
    l4l5l6: Math.round(((counts.L4 + counts.L5 + counts.L6) / total) * 100),
  };
};

// Analyze a single part and return analysis object
const analyzePartSimple = (
  questions: Question[],
  part: string
): PartAnalysis => {
  const levelCounts = countLevels(questions);
  const total = questions.length;
  const levelPercentages = calculateLevelPercentages(levelCounts, total);
  const errors = questions.filter(q => q.hasError).map(q => ({
    part,
    questionNumber: q.questionNumber,
    issue: q.errorMessage || 'Unknown error',
    suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
  }));

  // if (part === 'B' && errors.length > 0) console.error(`Part B errors found: ${JSON.stringify(errors)}`);

  // if (part === 'B' && errors.length > 0) console.log(`Part B errors found: ${JSON.stringify(errors)}`);

  return {
    part,
    totalQuestions: total,
    levelCounts,
    levelPercentages,
    errors,
    isValid: errors.length === 0
  };
};

// Validate OR pairs and handle structure/level mismatches
const validateOrPairsWithUpgrade = (questions: Question[], part: string): void => {
  // Group by base question number
  const pairs: Map<number, Question[]> = new Map();

  questions.forEach(q => {
    const match = q.questionNumber.match(/(\d+)\s*([ABab])/);
    if (match && match[2]) {
      const base = parseInt(match[1]);
      if (!pairs.has(base)) {
        pairs.set(base, []);
      }
      pairs.get(base)!.push(q);
    }
  });

  // Check each pair
  pairs.forEach((pair, baseNum) => {
    if (pair.length === 2) {
      const [q1, q2] = pair;

      // SKIP if either question is already fixed - don't re-validate fixed questions
      if (q1.isFixed || q2.isFixed) {
        return;
      }

      // Check subdivision structure mismatch
      const q1HasSubs = q1.hasSubdivisions || false;
      const q2HasSubs = q2.hasSubdivisions || false;

      if (q1HasSubs !== q2HasSubs) {
        // One has subdivisions, one doesn't - CONVERT subdivisions TO single question
        const questionWithSubs = q1HasSubs ? q1 : q2;
        const questionWithoutSubs = q1HasSubs ? q2 : q1;

        if (!questionWithSubs.isFixed) {
          questionWithSubs.hasError = true;
          questionWithSubs.expectedLevel = questionWithoutSubs.detectedLevel;
          questionWithSubs.expectedCo = questionWithoutSubs.co;
          questionWithSubs.errorMessage = `OR pair ${baseNum}: Structure mismatch. Convert to single ${questionWithoutSubs.detectedLevel} question to match OR pair.`;
        }
      }

      // Check level mismatch — OR pair level matching OVERRIDES mark-based errors
      // Use the maximum of their expected OR detected levels to ensure they sync up
      if (!q1.isFixed && !q2.isFixed) {
        const l1Str = q1.expectedLevel || q1.detectedLevel;
        const l2Str = q2.expectedLevel || q2.detectedLevel;

        const level1 = getLevelNumber(l1Str);
        const level2 = getLevelNumber(l2Str);

        if (level1 !== level2) {
          const higherLevel = getHigherLevel(l1Str, l2Str);

          if (q1.detectedLevel !== higherLevel) {
            q1.hasError = true;
            q1.expectedLevel = higherLevel;
            q1.errorMessage = `OR pair ${baseNum}: Level mismatch. Change to ${higherLevel} to match OR pair.`;
          }
          if (q2.detectedLevel !== higherLevel) {
            q2.hasError = true;
            q2.expectedLevel = higherLevel;
            q2.errorMessage = `OR pair ${baseNum}: Level mismatch. Change to ${higherLevel} to match OR pair.`;
          }
        }
      }

      // Check CO matching for Part B OR pairs
      if (part === 'B' && q1.co !== q2.co) {
        // Don't flag as error for 16-mark questions in IA2 (which should have different COs)
        const is16Mark = q1.marks === 16 || q2.marks === 16;
        if (!is16Mark && !q1.hasError && !q2.hasError && !q1.isFixed && !q2.isFixed) {
          q2.hasError = true;
          q2.expectedLevel = q2.detectedLevel;
          q2.expectedCo = q1.co; // Track expected CO for fix
          q2.errorMessage = `OR pair ${baseNum}: CO mismatch. Should be ${q1.co} to match OR pair.`;
        }
      }
    }
  });
};

// Check and fix Part A L1/L2 distribution
const checkPartADistribution = (questions: Question[], targetL1Percent: number = 60): void => {
  if (questions.length === 0) return;

  const l1Count = questions.filter(q => q.detectedLevel === 'L1').length;
  const total = questions.length;
  const l1Percent = (l1Count / total) * 100;

  // Relaxed margin: +/- 20% (e.g. 40-80% for a 60% target)
  const margin = 20;
  const minL1 = Math.max(0, targetL1Percent - margin);
  const maxL1 = Math.min(100, targetL1Percent + margin);

  if (l1Percent < minL1 || l1Percent > maxL1) {
    // Only flag a non-fixed question with this warning
    const candidate = questions.find(q => !q.isFixed && !q.hasError);
    if (candidate) {
      candidate.hasError = true;
      candidate.errorMessage = `Part A Level Distribution mismatch: found ${l1Percent.toFixed(0)}% L1 (target ${targetL1Percent}% ±${margin}%).`;
    }
  }
};

// Validate marks-based level restrictions
const validateMarkBasedLevels = (questions: Question[], allowedLevels: string[], requiredLevels: string[], markThreshold: number, isAboveThreshold: boolean): void => {
  questions.forEach(q => {
    if (q.hasError || q.isFixed) return;

    const meetsThreshold = isAboveThreshold ? q.marks >= markThreshold : q.marks < markThreshold;

    if (meetsThreshold) {
      if (requiredLevels.length > 0 && !requiredLevels.includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = requiredLevels[0];
        q.errorMessage = `${q.marks}-mark questions must be ${requiredLevels.join('/')} only. Found ${q.detectedLevel}.`;
      }
    } else if (!allowedLevels.includes(q.detectedLevel)) {
      q.hasError = true;
      q.expectedLevel = allowedLevels[0];
      q.errorMessage = `This question allows only ${allowedLevels.join('/')} levels. Found ${q.detectedLevel}.`;
    }
  });
};

// Validate IA3 Part A: Each CO must have one L1 and one L2 question
const validateIA3PartACOLevels = (questions: Question[]): void => {
  // Group by CO
  const coGroups: Map<string, Question[]> = new Map();

  questions.forEach(q => {
    if (!coGroups.has(q.co)) {
      coGroups.set(q.co, []);
    }
    coGroups.get(q.co)!.push(q);
  });

  // Check each CO group - should have 2 questions with different levels (L1 and L2)
  coGroups.forEach((group, co) => {
    if (group.length >= 2) {
      const l1Questions = group.filter(q => q.detectedLevel === 'L1');
      const l2Questions = group.filter(q => q.detectedLevel === 'L2');

      // If both are L1, mark one to change to L2
      if (l1Questions.length >= 2 && l2Questions.length === 0) {
        const q = l1Questions[1];
        if (!q.hasError && !q.isFixed) {
          q.hasError = true;
          q.expectedLevel = 'L2';
          q.errorMessage = `${co} has two L1 questions. One must be L2.`;
        }
      }

      // If both are L2, mark one to change to L1
      if (l2Questions.length >= 2 && l1Questions.length === 0) {
        const q = l2Questions[1];
        if (!q.hasError && !q.isFixed) {
          q.hasError = true;
          q.expectedLevel = 'L1';
          q.errorMessage = `${co} has two L2 questions. One must be L1.`;
        }
      }
    }
  });
};

// Validate IA2 Part A: CO2 and CO3 each must have L1 and L2 questions
const validateIA2PartACOLevels = (questions: Question[]): void => {
  const co2Questions = questions.filter(q => q.co === 'CO2');
  const co3Questions = questions.filter(q => q.co === 'CO3');

  // Check CO2
  if (co2Questions.length >= 2) {
    const l1 = co2Questions.filter(q => q.detectedLevel === 'L1');
    const l2 = co2Questions.filter(q => q.detectedLevel === 'L2');

    if (l1.length === 0 && l2.length >= 2) {
      const q = l2[0];
      if (!q.hasError && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L1';
        q.errorMessage = 'CO2 needs one L1 question.';
      }
    }
    if (l2.length === 0 && l1.length >= 2) {
      const q = l1[0];
      if (!q.hasError && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L2';
        q.errorMessage = 'CO2 needs one L2 question.';
      }
    }
  }

  // Check CO3
  if (co3Questions.length >= 2) {
    const l1 = co3Questions.filter(q => q.detectedLevel === 'L1');
    const l2 = co3Questions.filter(q => q.detectedLevel === 'L2');

    if (l1.length === 0 && l2.length >= 2) {
      const q = l2[0];
      if (!q.hasError && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L1';
        q.errorMessage = 'CO3 needs one L1 question.';
      }
    }
    if (l2.length === 0 && l1.length >= 2) {
      const q = l1[0];
      if (!q.hasError && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L2';
        q.errorMessage = 'CO3 needs one L2 question.';
      }
    }
  }
};

// Validate IA2 Part B 16-mark CO requirement (CO2 for one, CO3 for other)
const validateIA2PartB16MarkCOs = (questions: Question[]): void => {
  const sixteenMarkQuestions = questions.filter(q => q.marks === 16);

  if (sixteenMarkQuestions.length >= 2) {
    const cos = sixteenMarkQuestions.map(q => q.co);
    const hasCO2 = cos.includes('CO2');
    const hasCO3 = cos.includes('CO3');

    if (!hasCO2 || !hasCO3) {
      // Check if both are same CO
      if (cos[0] === cos[1]) {
        const q = sixteenMarkQuestions[1];
        const targetCO = cos[0] === 'CO2' ? 'CO3' : 'CO2';
        if (!q.hasError && !q.isFixed) {
          q.hasError = true;
          q.expectedLevel = q.detectedLevel;
          q.errorMessage = `16-mark questions should have one CO2 and one CO3. Change this to ${targetCO}.`;
        }
      }
    }
  }
};

// Validate IA1 paper
export const validateIA1 = (questions: { partA: Question[]; partB: Question[] }): ValidationResult => {
  const allQuestions = [...questions.partA, ...questions.partB];

  // Check if ALL questions have auto-detected levels (no explicit CO/RBT in document)
  // This happens with aptitude papers, general papers, etc.
  const allAutoDetected = allQuestions.length > 0 && allQuestions.every(q => q.levelAutoDetected === true);

  if (allAutoDetected) {
    // Skip strict level validation — paper has no explicit CO/RBT info
    allQuestions.forEach(q => {
      q.hasError = false;
      q.errorMessage = undefined;
      q.expectedLevel = q.detectedLevel;
    });

    const partAAnalysis = analyzePartSimple(questions.partA, 'A');
    const partBAnalysis = analyzePartSimple(questions.partB, 'B');
    partAAnalysis.errors = [];
    partAAnalysis.isValid = true;
    partBAnalysis.errors = [];
    partBAnalysis.isValid = true;

    const dist = calculateDistribution(allQuestions);

    return {
      status: 'accepted',
      errors: [],
      levelDistribution: dist,
      partAnalysis: {
        partA: partAAnalysis,
        partB: partBAnalysis
      },
      allErrorsFixed: true
    };
  }

  // Reset errors only on non-fixed questions — never re-flag fixed questions
  allQuestions.forEach(q => {
    if (q.isFixed) return;
    q.hasError = false;
    q.errorMessage = undefined;
    q.expectedLevel = q.detectedLevel;
  });

  // Part A validation: L1 and L2 only, with allowance for one L3 (5-10%)
  const partAL3Questions = questions.partA.filter(q => q.detectedLevel === 'L3');
  const allowedL3Count = Math.max(1, Math.floor(questions.partA.length * 0.1));

  questions.partA.forEach((q, index) => {
    if (q.detectedLevel === 'L3') {
      // Allow up to 10% L3 questions (at least 1)
      const l3Index = partAL3Questions.indexOf(q);
      if (l3Index >= allowedL3Count && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L2';
        q.errorMessage = `Too many L3 questions in Part A. Only ${allowedL3Count} L3 question(s) allowed.`;
      }
    } else if (!['L1', 'L2'].includes(q.detectedLevel) && !q.isFixed) {
      q.hasError = true;
      q.expectedLevel = 'L1';
      q.errorMessage = `Invalid level ${q.detectedLevel}. Part A allows only L1, L2, and limited L3.`;
    }
  });

  // Check Part A distribution (60% L1, 40% L2)
  checkPartADistribution(questions.partA, 60);

  // Part B validation based on marks:
  // 8-mark: L2, L3, L4 allowed
  // 16-mark: L4, L5, L6 strictly required
  questions.partB.forEach(q => {
    // console.log(`Checking Q${q.questionNumber}: marks=${q.marks}, level=${q.detectedLevel}, fixed=${q.isFixed}, hasError=${q.hasError}`);
    if (q.hasError || q.isFixed) {
      // console.log('Skipping fixed/error question:', q.questionNumber);
      return;
    }

    if (q.marks === 16) {
      // 16-mark must be L4, L5, or L6
      if (!['L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L4';
        q.errorMessage = `16-mark questions must be L4/L5/L6 only. Found ${q.detectedLevel}.`;
      }
    } else if (q.marks === 8) {
      // 8-mark can be L2, L3, or L4
      if (!['L2', 'L3', 'L4'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `8-mark questions must be L2/L3/L4 only. Found ${q.detectedLevel}.`;
      }
    } else {
      // Default Part B validation for other marks
      if (!['L2', 'L3', 'L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `Part B requires L2-L6, found ${q.detectedLevel}`;
      }
    }
  });

  // Validate OR pairs - upgrade lower to higher
  validateOrPairsWithUpgrade(questions.partB, 'B');

  // Create part analyses
  // if (questions.partB.length > 0) console.log(`Before analysis: Q11 hasError=${questions.partB[0]?.hasError}`);
  const partAAnalysis = analyzePartSimple(questions.partA, 'A');
  const partBAnalysis = analyzePartSimple(questions.partB, 'B');

  // Collect all errors
  const errors = [...partAAnalysis.errors, ...partBAnalysis.errors];

  // Calculate distribution
  const dist = calculateDistribution(allQuestions);

  // Check if all errors are fixed
  const allErrorsFixed = allQuestions.every(q => !q.hasError);

  // Check for Missing Parts
  if (questions.partA.length === 0) {
    // Create a dummy error to block acceptance
    errors.push({
      part: 'A',
      questionNumber: 'All',
      issue: 'Missing Part A questions',
      suggestion: 'Ensure the document has a table for Part A or Questions 1-10'
    });
  }
  if (questions.partB.length === 0) {
    errors.push({
      part: 'B',
      questionNumber: 'All',
      issue: 'Missing Part B questions',
      suggestion: 'Ensure the document has a table for Part B or Questions 11+'
    });
  }

  return {
    status: errors.length > 0 ? 'rejected' : 'accepted',
    errors,
    levelDistribution: dist,
    partAnalysis: {
      partA: partAAnalysis,
      partB: partBAnalysis
    },
    allErrorsFixed
  };
};

// Validate IA2 paper
export const validateIA2 = (questions: { partA: Question[]; partB: Question[]; partC?: Question[] }): ValidationResult => {
  const allQuestions = [...questions.partA, ...questions.partB, ...(questions.partC || [])];

  // Auto-accept papers with no explicit CO/RBT levels
  const allAutoDetected = allQuestions.length > 0 && allQuestions.every(q => q.levelAutoDetected === true);
  if (allAutoDetected) {
    allQuestions.forEach(q => { q.hasError = false; q.errorMessage = undefined; q.expectedLevel = q.detectedLevel; });
    const partAAnalysis = analyzePartSimple(questions.partA, 'A');
    const partBAnalysis = analyzePartSimple(questions.partB, 'B');
    const partCAnalysis = questions.partC ? analyzePartSimple(questions.partC, 'C') : undefined;
    partAAnalysis.errors = []; partAAnalysis.isValid = true;
    partBAnalysis.errors = []; partBAnalysis.isValid = true;
    if (partCAnalysis) { partCAnalysis.errors = []; partCAnalysis.isValid = true; }
    return {
      status: 'accepted', errors: [], levelDistribution: calculateDistribution(allQuestions),
      partAnalysis: { partA: partAAnalysis, partB: partBAnalysis, partC: partCAnalysis }, allErrorsFixed: true
    };
  }

  // Reset errors only on non-fixed questions — never re-flag fixed questions
  allQuestions.forEach(q => {
    if (q.isFixed) return;
    q.hasError = false;
    q.errorMessage = undefined;
    q.expectedLevel = q.detectedLevel;
  });

  // Part A validation: L1 and L2 only (allow one L3), CO2/CO3
  const partAL3Questions = questions.partA.filter(q => q.detectedLevel === 'L3');
  const allowedL3Count = Math.max(1, Math.floor(questions.partA.length * 0.1));

  questions.partA.forEach(q => {
    if (q.isFixed) return;

    if (!['CO2', 'CO3'].includes(q.co)) {
      q.hasError = true;
      q.errorMessage = `Wrong CO (${q.co}). IA2 Part A should be CO2 or CO3.`;
    }

    if (q.detectedLevel === 'L3') {
      // Allow up to 10% L3 questions (at least 1)
      const l3Index = partAL3Questions.indexOf(q);
      if (l3Index >= allowedL3Count && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L2';
        q.errorMessage = (q.errorMessage ? q.errorMessage + ' ' : '') + `Too many L3 questions. Only ${allowedL3Count} L3 allowed.`;
      }
    } else if (!['L1', 'L2'].includes(q.detectedLevel) && !q.isFixed) {
      q.hasError = true;
      q.expectedLevel = 'L1';
      q.errorMessage = (q.errorMessage ? q.errorMessage + ' ' : '') + `Invalid level ${q.detectedLevel}. Part A allows only L1, L2, and limited L3.`;
    }
  });

  // Validate IA2 Part A CO2/CO3 level requirements
  validateIA2PartACOLevels(questions.partA);

  // Check Part A distribution (60% L1, 40% L2)
  checkPartADistribution(questions.partA, 60);

  // Part B validation based on marks
  questions.partB.forEach(q => {
    if (q.hasError || q.isFixed) return;

    if (q.marks === 16) {
      // 16-mark must be L4, L5, or L6
      if (!['L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L4';
        q.errorMessage = `16-mark questions must be L4/L5/L6 only. Found ${q.detectedLevel}.`;
      }
    } else if (q.marks === 12) {
      // 12-mark can be L2, L3, L4, or L5
      if (!['L2', 'L3', 'L4', 'L5'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `12-mark questions must be L2/L3/L4/L5 only. Found ${q.detectedLevel}.`;
      }
    } else {
      // Default Part B validation
      if (!['L2', 'L3', 'L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `Part B requires L2-L6, found ${q.detectedLevel}`;
      }
    }
  });

  // Part C validation (if exists) - similar to IA3
  let partCAnalysis: PartAnalysis | undefined;
  if (questions.partC && questions.partC.length > 0) {
    questions.partC.forEach(q => {
      if (q.hasError || q.isFixed) return;
      if (!['L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L5';
        q.errorMessage = `Part C (15-mark) allows only L4-L6, found ${q.detectedLevel}`;
      }
    });
    validateOrPairsWithUpgrade(questions.partC, 'C');
    partCAnalysis = analyzePartSimple(questions.partC, 'C');
  }

  // Validate 16-mark CO requirements (CO2 and CO3)
  validateIA2PartB16MarkCOs(questions.partB);

  // Validate OR pairs in Part B
  validateOrPairsWithUpgrade(questions.partB, 'B');

  // Create part analyses
  const partAAnalysis = analyzePartSimple(questions.partA, 'A');
  const partBAnalysis = analyzePartSimple(questions.partB, 'B');

  const errors = [
    ...partAAnalysis.errors,
    ...partBAnalysis.errors,
    ...(partCAnalysis?.errors || [])
  ];
  const dist = calculateDistribution(allQuestions);
  const allErrorsFixed = allQuestions.every(q => !q.hasError);

  return {
    status: errors.length > 0 ? 'rejected' : 'accepted',
    errors,
    levelDistribution: dist,
    partAnalysis: {
      partA: partAAnalysis,
      partB: partBAnalysis,
      partC: partCAnalysis
    },
    allErrorsFixed
  };
};

// Validate IA3 paper
export const validateIA3 = (questions: { partA: Question[]; partB: Question[]; partC?: Question[] }): ValidationResult => {
  const allQuestions = [...questions.partA, ...questions.partB, ...(questions.partC || [])];

  // Auto-accept papers with no explicit CO/RBT levels
  const allAutoDetected = allQuestions.length > 0 && allQuestions.every(q => q.levelAutoDetected === true);
  if (allAutoDetected) {
    allQuestions.forEach(q => { q.hasError = false; q.errorMessage = undefined; q.expectedLevel = q.detectedLevel; });
    const partAAnalysis = analyzePartSimple(questions.partA, 'A');
    const partBAnalysis = analyzePartSimple(questions.partB, 'B');
    partAAnalysis.errors = []; partAAnalysis.isValid = true;
    partBAnalysis.errors = []; partBAnalysis.isValid = true;
    const partCAnalysis = questions.partC ? analyzePartSimple(questions.partC, 'C') : undefined;
    if (partCAnalysis) { partCAnalysis.errors = []; partCAnalysis.isValid = true; }
    return {
      status: 'accepted', errors: [], levelDistribution: calculateDistribution(allQuestions),
      partAnalysis: { partA: partAAnalysis, partB: partBAnalysis, partC: partCAnalysis }, allErrorsFixed: true
    };
  }

  // Reset errors only on non-fixed questions — never re-flag fixed questions
  allQuestions.forEach(q => {
    if (q.isFixed) return;
    q.hasError = false;
    q.errorMessage = undefined;
    q.expectedLevel = q.detectedLevel;
  });

  // Part A validation: L1 and L2 only (allow one L3)
  const partAL3Questions = questions.partA.filter(q => q.detectedLevel === 'L3');
  const allowedL3Count = Math.max(1, Math.floor(questions.partA.length * 0.1));

  questions.partA.forEach(q => {
    if (q.detectedLevel === 'L3') {
      // Allow up to 10% L3 questions (at least 1)
      const l3Index = partAL3Questions.indexOf(q);
      if (l3Index >= allowedL3Count && !q.isFixed) {
        q.hasError = true;
        q.expectedLevel = 'L2';
        q.errorMessage = `Too many L3 questions in Part A. Only ${allowedL3Count} L3 question(s) allowed.`;
      }
    } else if (!['L1', 'L2'].includes(q.detectedLevel) && !q.isFixed) {
      q.hasError = true;
      q.expectedLevel = 'L1';
      q.errorMessage = `Invalid level ${q.detectedLevel}. Part A allows only L1, L2, and limited L3.`;
    }
  });

  // Validate IA3 Part A: Each CO must have one L1 and one L2
  validateIA3PartACOLevels(questions.partA);

  // Check Part A distribution (50% L1, 50% L2 for IA3)
  checkPartADistribution(questions.partA, 50);

  // Determine IA3 structure
  const hasPartC = questions.partC && questions.partC.length > 0;
  const partBMaxMarks = Math.max(...questions.partB.map(q => q.marks), 0);

  // Part B validation depends on structure
  if (hasPartC) {
    // Has Part C: Part B is 13-mark questions (L2-L5)
    questions.partB.forEach(q => {
      if (q.hasError || q.isFixed) return;

      if (!['L2', 'L3', 'L4', 'L5'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `Part B (13-mark) allows only L2-L5, found ${q.detectedLevel}`;
      }
    });
  } else {
    // No Part C: Part B has 16-mark questions (L2-L6)
    questions.partB.forEach(q => {
      if (q.hasError || q.isFixed) return;

      if (!['L2', 'L3', 'L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L3';
        q.errorMessage = `Part B (16-mark) allows L2-L6, found ${q.detectedLevel}`;
      }
    });
  }

  // Validate OR pairs in Part B
  validateOrPairsWithUpgrade(questions.partB, 'B');

  // Part C validation: L4-L6 only (if exists)
  let partCAnalysis: PartAnalysis | undefined;
  if (questions.partC && questions.partC.length > 0) {
    questions.partC.forEach(q => {
      if (q.hasError || q.isFixed) return;

      if (!['L4', 'L5', 'L6'].includes(q.detectedLevel)) {
        q.hasError = true;
        q.expectedLevel = 'L5';
        q.errorMessage = `Part C (15-mark) allows only L4-L6, found ${q.detectedLevel}`;
      }
    });

    // Validate OR pairs in Part C
    validateOrPairsWithUpgrade(questions.partC, 'C');

    partCAnalysis = analyzePartSimple(questions.partC, 'C');
  }

  // Create part analyses
  const partAAnalysis = analyzePartSimple(questions.partA, 'A');
  const partBAnalysis = analyzePartSimple(questions.partB, 'B');

  const errors = [
    ...partAAnalysis.errors,
    ...partBAnalysis.errors,
    ...(partCAnalysis?.errors || [])
  ];

  const dist = calculateDistribution(allQuestions);
  const allErrorsFixed = allQuestions.every(q => !q.hasError);

  return {
    status: errors.length > 0 ? 'rejected' : 'accepted',
    errors,
    levelDistribution: dist,
    partAnalysis: {
      partA: partAAnalysis,
      partB: partBAnalysis,
      partC: partCAnalysis
    },
    allErrorsFixed
  };
};
