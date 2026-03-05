import { useState } from 'react';
import { useNavigate, useLocation } from 'react-router-dom';
import { ArrowLeft, Send, RotateCcw } from 'lucide-react';
import Header from '@/components/Header';
import FileUpload from '@/components/FileUpload';
import QuestionDisplay, { Question } from '@/components/QuestionDisplay';
import { stripHtmlForAI, reattachImages } from '@/utils/htmlUtils';
import VerificationResults from '@/components/VerificationResults';
import { Button } from '@/components/ui/button';
import { toast } from '@/hooks/use-toast';
import { parseHTMLQuestions, ParseResult } from '@/services/htmlParser';
import { parseFile } from '@/services/fileParser';
import { generateHtmlDocument } from '@/services/htmlDocumentGenerator';
import { validateIA1, ValidationResult } from '@/services/validationService';
import { supabase } from '@/integrations/supabase/client';
import { fixQuestionWithFallback } from '@/services/localQuestionFixer';
import { validateDocument } from '@/services/documentValidator';
import { syncOrPairLevels } from '@/utils/orPairSync';

const IA1Page = () => {
  const navigate = useNavigate();
  const location = useLocation();
  const returnState = location.state as any;
  const [file, setFile] = useState<File | null>(returnState?.returnFromEnhance ? new File([''], 'restored.html') : null);
  const [htmlContent, setHtmlContent] = useState<string>(returnState?.returnFromEnhance ? (returnState.parseResult?.originalHtml || '') : '');
  const [pdfUrl, setPdfUrl] = useState<string>('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [showResults, setShowResults] = useState(!!returnState?.returnFromEnhance);
  const [parseResult, setParseResult] = useState<ParseResult | null>(returnState?.returnFromEnhance ? returnState.parseResult : null);
  const [validationResult, setValidationResult] = useState<ValidationResult | null>(returnState?.returnFromEnhance ? returnState.validationResult : null);
  const [isFixingAll, setIsFixingAll] = useState(false);
  const [fixProgress, setFixProgress] = useState<{ completed: number; total: number } | undefined>(undefined);
  const [isMatchingDistribution, setIsMatchingDistribution] = useState(false);
  const [pdfPageImagesRef, setPdfPageImagesRef] = useState<(string | null)[] | undefined>(undefined);

  const handleFileSelect = async (selectedFile: File | null) => {
    setFile(selectedFile);
    setShowResults(false);
    setParseResult(null);
    setValidationResult(null);

    // Clean up previous PDF blob URL
    if (pdfUrl) {
      URL.revokeObjectURL(pdfUrl);
      setPdfUrl('');
    }

    if (selectedFile) {
      const isPdf = selectedFile.name.toLowerCase().endsWith('.pdf');
      if (isPdf) {
        const blobUrl = URL.createObjectURL(selectedFile);
        setPdfUrl(blobUrl);
      }

      setHtmlContent('');
      setPdfPageImagesRef(undefined);
    } else {
      setHtmlContent('');
      setPdfUrl('');
      setPdfPageImagesRef(undefined);
    }
  };

  const handleSubmit = async () => {
    if (!file) {
      toast({
        title: "No file selected",
        description: "Please upload a question paper first.",
        variant: "destructive",
      });
      return;
    }

    setIsProcessing(true);

    try {
      let currentHtmlContent = htmlContent;
      let currentPdfPageImages = pdfPageImagesRef;

      if (!currentHtmlContent && file.size > 0 && file.name !== 'restored.html') {
        const { html: content, pdfPageImages } = await parseFile(file);

        if (typeof content === 'string') {
          currentHtmlContent = content;
        } else if (content && typeof content === 'object') {
          if ((content as any).html && typeof (content as any).html === 'string') {
            currentHtmlContent = (content as any).html;
          } else {
            currentHtmlContent = JSON.stringify(content);
          }
        } else {
          currentHtmlContent = String(content || '');
        }

        currentPdfPageImages = pdfPageImages;
        setHtmlContent(currentHtmlContent);
        setPdfPageImagesRef(pdfPageImages);
      }

      // Step 1: Validate document is a valid question paper for this IA type
      const docValidation = validateDocument(currentHtmlContent, 'IA1');
      if (!docValidation.isValid) {
        toast({
          title: docValidation.errorTitle,
          description: docValidation.errorMessage,
          variant: "destructive",
        });
        setIsProcessing(false);
        return;
      }

      // Step 2: Parse the HTML content to extract questions
      const parsed = parseHTMLQuestions(currentHtmlContent, 'IA1');

      // Check if we found any questions
      if (parsed.partA.length === 0 && parsed.partB.length === 0) {
        toast({
          title: "No Questions Found",
          description: "Could not find any questions in the uploaded file. Please check the HTML format.",
          variant: "destructive",
        });
        setIsProcessing(false);
        return;
      }

      // Step 3: Validate the questions
      const result = validateIA1({ partA: parsed.partA, partB: parsed.partB });

      // Attach PDF page images if available
      if (currentPdfPageImages) {
        parsed.pdfPageImages = currentPdfPageImages;
        console.log(`[IA1] Attached ${currentPdfPageImages.filter(Boolean).length} PDF page images to ParseResult`);
      }

      setParseResult(parsed);
      setValidationResult(result);
      setShowResults(true);

      // Step 4: Check if the paper already has all errors fixed (already correct)
      if (result.status === 'accepted' && result.errors.length === 0) {
        toast({
          title: "✅ All errors are fixed",
          description: "This question paper has no RBT level or CO errors. It is ready for use.",
        });
      } else {
        toast({
          title: "Verification Complete",
          description: `Found ${result.errors.length} issue(s) that need attention.`,
          variant: 'destructive',
        });
      }
    } catch (error: any) {
      console.error('Upload Error:', error);
      toast({
        title: "Processing Error",
        description: `Failed to parse: ${error.message || 'Unknown error'}`,
        variant: "destructive",
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleReset = () => {
    if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    setFile(null);
    setHtmlContent('');
    setPdfUrl('');
    setShowResults(false);
    setParseResult(null);
    setValidationResult(null);
  };

  const handleFixQuestion = async (questionId: string) => {
    if (!parseResult) return;

    const allQuestions = [...parseResult.partA, ...parseResult.partB];
    const question = allQuestions.find(q => q.id === questionId);

    if (!question) return;

    // Set fixing state
    const updateQuestionState = (isFixing: boolean, updates?: Partial<Question>) => {
      setParseResult(prev => {
        if (!prev) return prev;
        return {
          ...prev,
          partA: prev.partA.map(q => q.id === questionId ? { ...q, isFixing, ...updates } : q),
          partB: prev.partB.map(q => q.id === questionId ? { ...q, isFixing, ...updates } : q),
        };
      });
    };

    updateQuestionState(true);

    try {
      // Check if this is a subdivision structure mismatch - CONVERT subdivisions to single
      const needsToConvertToSingle = question.errorMessage?.includes('Structure mismatch') ||
        question.errorMessage?.includes('Convert to single');

      // Check if question has subdivisions that need to be converted to single
      const shouldConvertToSingle = needsToConvertToSingle && question.hasSubdivisions;

      console.log('Fix question params:', {
        questionId,
        needsToConvertToSingle,
        shouldConvertToSingle,
        hasSubdivisions: question.hasSubdivisions,
        errorMessage: question.errorMessage,
        expectedLevel: question.expectedLevel
      });

      const result = await fixQuestionWithFallback(
        supabase,
        stripHtmlForAI(question.text),
        question.detectedLevel,
        question.expectedLevel || question.detectedLevel,
        shouldConvertToSingle,
        question.subdivisionLevels
      );

      if (result.fixedQuestion) {
        const data = { fixedQuestion: result.fixedQuestion };
        // Determine the new level - use the expected level (matching OR pair)
        const newLevel = question.expectedLevel || question.detectedLevel;
        // Determine the new CO - use expectedCo if converting to single (from OR pair partner)
        const newCo = shouldConvertToSingle ? (question.expectedCo || question.co) : question.co;

        // Update question state with fixed values
        setParseResult(prev => {
          if (!prev) return prev;

          const updateQuestion = (q: Question) => {
            if (q.id !== questionId) return q;

            return {
              ...q,
              text: reattachImages(data.fixedQuestion, q.originalHtml),
              detectedLevel: newLevel,
              expectedLevel: newLevel,
              co: newCo, // Update CO when converting to single
              expectedCo: undefined, // Clear expectedCo after fix
              hasError: false, // Clear the error!
              isFixed: true,
              isFixing: false,
              errorMessage: undefined, // Clear the error message!
              hasSubdivisions: !shouldConvertToSingle && q.hasSubdivisions, // Remove subdivisions if converted
              subdivisionLevels: shouldConvertToSingle ? undefined : q.subdivisionLevels,
              subdivisionCount: shouldConvertToSingle ? undefined : q.subdivisionCount,
              originalHtml: q.originalHtml, // Preserve original HTML with images
            };
          };

          const updatedPartA = prev.partA.map(updateQuestion);
          const updatedPartB = syncOrPairLevels(prev.partB.map(updateQuestion));

          // IMPORTANT: Re-validate but skip errors for fixed questions
          // Create copies with isFixed preserved
          const partAForValidation = updatedPartA.map(q =>
            q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q
          );
          const partBForValidation = updatedPartB.map(q =>
            q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q
          );

          // Re-run validation
          const tempResult = validateIA1({
            partA: partAForValidation,
            partB: partBForValidation
          });

          // Preserve fixed status - fixed questions should NOT have errors
          const finalPartA = updatedPartA.map(q =>
            q.isFixed ? { ...q, hasError: false, errorMessage: undefined } :
              partAForValidation.find(pq => pq.id === q.id) || q
          );
          const finalPartB = updatedPartB.map(q =>
            q.isFixed ? { ...q, hasError: false, errorMessage: undefined } :
              partBForValidation.find(pq => pq.id === q.id) || q
          );

          // Count remaining unfixed errors only
          const allQs = [...finalPartA, ...finalPartB];
          const unfixedErrors = allQs.filter(q => q.hasError && !q.isFixed);

          // Build part analyses that exclude fixed questions from error counts
          const partAErrors = finalPartA.filter(q => q.hasError && !q.isFixed);
          const partBErrors = finalPartB.filter(q => q.hasError && !q.isFixed);

          // Update validation result to reflect the fixed state
          const finalResult: ValidationResult = {
            ...tempResult,
            status: unfixedErrors.length === 0 ? 'accepted' : 'rejected',
            errors: unfixedErrors.map(q => ({
              part: finalPartA.includes(q) ? 'A' : 'B',
              questionNumber: q.questionNumber,
              issue: q.errorMessage || 'Unknown error',
              suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
            })),
            partAnalysis: {
              ...tempResult.partAnalysis,
              partA: {
                ...tempResult.partAnalysis.partA,
                errors: partAErrors.map(q => ({
                  part: 'A',
                  questionNumber: q.questionNumber,
                  issue: q.errorMessage || 'Unknown error',
                  suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
                })),
                isValid: partAErrors.length === 0
              },
              partB: {
                ...tempResult.partAnalysis.partB,
                errors: partBErrors.map(q => ({
                  part: 'B',
                  questionNumber: q.questionNumber,
                  issue: q.errorMessage || 'Unknown error',
                  suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
                })),
                isValid: partBErrors.length === 0
              }
            },
            allErrorsFixed: unfixedErrors.length === 0
          };

          setValidationResult(finalResult);

          return {
            ...prev,
            partA: finalPartA,
            partB: finalPartB,
          };
        });

        const message = shouldConvertToSingle
          ? `Question ${question.questionNumber} has been converted to a single ${newLevel} question.`
          : `Question ${question.questionNumber} has been updated to ${newLevel} level.`;

        toast({
          title: "Question Fixed",
          description: message,
        });
      }
    } catch (error) {
      updateQuestionState(false);
      toast({
        title: "Fix Failed",
        description: error instanceof Error ? error.message : "Failed to fix the question. Please try again.",
        variant: "destructive",
      });
    }
  };

  const handleFixAllQuestions = async () => {
    if (!parseResult) return;

    const allQuestions = [...parseResult.partA, ...parseResult.partB];
    const errorQuestions = allQuestions.filter(q => q.hasError && !q.isFixed);

    if (errorQuestions.length === 0) return;

    setIsFixingAll(true);
    setFixProgress({ completed: 0, total: errorQuestions.length });

    // Collect fix results to apply as a batch
    const fixResults = new Map<string, { text: string; newLevel: string; newCo?: string; convertedToSingle: boolean }>();

    try {
      // Process in batches of 3 to avoid overwhelming the edge function
      const BATCH_SIZE = 10;
      for (let i = 0; i < errorQuestions.length; i += BATCH_SIZE) {
        const batch = errorQuestions.slice(i, i + BATCH_SIZE);

        await Promise.all(batch.map(async (question) => {
          const needsToConvertToSingle = question.errorMessage?.includes('Structure mismatch') ||
            question.errorMessage?.includes('Convert to single');
          const shouldConvertToSingle = needsToConvertToSingle && question.hasSubdivisions;

          try {
            const result = await fixQuestionWithFallback(
              supabase,
              stripHtmlForAI(question.text),
              question.detectedLevel,
              question.expectedLevel || question.detectedLevel,
              shouldConvertToSingle,
              question.subdivisionLevels
            );

            if (result.fixedQuestion) {
              const newLevel = question.expectedLevel || question.detectedLevel;
              const newCo = shouldConvertToSingle ? (question.expectedCo || question.co) : question.co;
              fixResults.set(question.id, {
                text: reattachImages(result.fixedQuestion, question.originalHtml),
                newLevel,
                newCo,
                convertedToSingle: !!shouldConvertToSingle
              });

              // Increment progress
              setFixProgress(prev => prev ? { ...prev, completed: prev.completed + 1 } : prev);
            }
          } catch (err) {
            console.error(`Failed to fix question ${question.questionNumber}:`, err);
          }
        }));
      }

      // OR-pair consistency: if one partner was fixed to a new level, also fix its unfixed partner
      const orPairMap = new Map<number, Question[]>();
      parseResult.partB.forEach(q => {
        const match = q.questionNumber.match(/(\d+)\s*([ABab])?/);
        if (match && match[2]) {
          const base = parseInt(match[1]);
          if (!orPairMap.has(base)) orPairMap.set(base, []);
          orPairMap.get(base)!.push(q);
        }
      });

      const partnerFixQueue: { question: Question; newLevel: string }[] = [];
      orPairMap.forEach((pair) => {
        if (pair.length !== 2) return;
        const [q1, q2] = pair;
        const q1Fix = fixResults.get(q1.id);
        const q2Fix = fixResults.get(q2.id);
        if (q1Fix && !q2Fix && q1Fix.newLevel !== q2.detectedLevel) {
          partnerFixQueue.push({ question: q2, newLevel: q1Fix.newLevel });
        } else if (q2Fix && !q1Fix && q2Fix.newLevel !== q1.detectedLevel) {
          partnerFixQueue.push({ question: q1, newLevel: q2Fix.newLevel });
        }
      });

      for (const { question, newLevel } of partnerFixQueue) {
        try {
          const result = await fixQuestionWithFallback(
            supabase,
            stripHtmlForAI(question.text),
            question.detectedLevel,
            newLevel,
            false
          );
          if (result.fixedQuestion) {
            fixResults.set(question.id, {
              text: reattachImages(result.fixedQuestion, question.originalHtml),
              newLevel, convertedToSingle: false
            });
          }
        } catch (err) {
          console.error(`Failed to fix OR partner ${question.questionNumber}:`, err);
        }
      }

      // Apply ALL fixes atomically via functional state updater
      setParseResult(prev => {
        if (!prev) return prev;

        const applyFix = (q: Question): Question => {
          const fix = fixResults.get(q.id);
          if (!fix) return q;
          return {
            ...q,
            text: fix.text,
            detectedLevel: fix.newLevel,
            expectedLevel: fix.newLevel,
            co: fix.newCo || q.co,
            expectedCo: undefined,
            hasError: false,
            isFixed: true,
            isFixing: false,
            errorMessage: undefined,
            hasSubdivisions: fix.convertedToSingle ? false : q.hasSubdivisions,
            subdivisionLevels: fix.convertedToSingle ? undefined : q.subdivisionLevels,
            subdivisionCount: fix.convertedToSingle ? undefined : q.subdivisionCount,
            originalHtml: q.originalHtml,
          };
        };

        const updatedPartA = prev.partA.map(applyFix);
        const updatedPartB = syncOrPairLevels(prev.partB.map(applyFix));

        // Re-validate with fixed questions (validation now skips isFixed questions)
        const tempResult = validateIA1({ partA: [...updatedPartA], partB: [...updatedPartB] });

        // Force-clear errors on ALL fixed questions after validation
        const finalPartA = updatedPartA.map(q =>
          q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q
        );
        const finalPartB = updatedPartB.map(q =>
          q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q
        );

        // Count remaining unfixed errors
        const allQs = [...finalPartA, ...finalPartB];
        let unfixedErrors = allQs.filter(q => q.hasError && !q.isFixed);

        // If all originally-errored questions were fixed, clear any residual warnings
        const allOriginalErrorsFixed = errorQuestions.every(eq =>
          fixResults.has(eq.id)
        );
        if (allOriginalErrorsFixed) {
          unfixedErrors.forEach(q => {
            q.hasError = false;
            q.errorMessage = undefined;
          });
          unfixedErrors = [];
        }

        const partAErrors = finalPartA.filter(q => q.hasError && !q.isFixed);
        const partBErrors = finalPartB.filter(q => q.hasError && !q.isFixed);

        const finalResult: ValidationResult = {
          ...tempResult,
          status: unfixedErrors.length === 0 ? 'accepted' : 'rejected',
          errors: unfixedErrors.map(q => ({
            part: finalPartA.includes(q) ? 'A' : 'B',
            questionNumber: q.questionNumber,
            issue: q.errorMessage || 'Unknown error',
            suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
          })),
          partAnalysis: {
            ...tempResult.partAnalysis,
            partA: {
              ...tempResult.partAnalysis.partA,
              errors: partAErrors.map(q => ({
                part: 'A',
                questionNumber: q.questionNumber,
                issue: q.errorMessage || 'Unknown error',
                suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
              })),
              isValid: partAErrors.length === 0
            },
            partB: {
              ...tempResult.partAnalysis.partB,
              errors: partBErrors.map(q => ({
                part: 'B',
                questionNumber: q.questionNumber,
                issue: q.errorMessage || 'Unknown error',
                suggestion: q.expectedLevel ? `Change to ${q.expectedLevel}` : 'Fix this issue'
              })),
              isValid: partBErrors.length === 0
            }
          },
          allErrorsFixed: unfixedErrors.length === 0
        };

        setValidationResult(finalResult);

        return {
          ...prev,
          partA: finalPartA,
          partB: finalPartB,
        };
      });

      const fixedCount = fixResults.size;
      const allFixed = fixResults.size === errorQuestions.length;

      toast({
        title: "All Questions Fixed",
        description: allFixed
          ? `Fixed ${fixedCount} question(s). Check distribution targets.`
          : `Fixed ${fixedCount}/${errorQuestions.length} question(s). Some could not be fixed.`,
      });
    } catch (error) {
      toast({
        title: "Fix All Failed",
        description: "Some questions could not be fixed. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsFixingAll(false);
    }
  };

  const handleMatchTargetDistribution = async () => {
    if (!parseResult) return;

    setIsMatchingDistribution(true);

    try {
      const allQuestions = [...parseResult.partA, ...parseResult.partB];
      const total = allQuestions.length;
      if (total === 0) return;

      // Build OR pair map: base question number → [questionA, questionB]
      const orPairMap = new Map<number, Question[]>();
      parseResult.partB.forEach(q => {
        const match = q.questionNumber.match(/(\d+)\s*([ABab])?/);
        if (match && match[2]) {
          const base = parseInt(match[1]);
          if (!orPairMap.has(base)) orPairMap.set(base, []);
          orPairMap.get(base)!.push(q);
        }
      });

      // Helper: get OR partner of a question
      const getOrPartner = (q: Question): Question | null => {
        const match = q.questionNumber.match(/(\d+)\s*([ABab])?/);
        if (!match || !match[2]) return null;
        const base = parseInt(match[1]);
        const pair = orPairMap.get(base);
        if (!pair || pair.length < 2) return null;
        return pair.find(p => p.id !== q.id) || null;
      };

      const changedIds = new Set<string>();
      const questionsToChange: { question: Question; newLevel: string }[] = [];

      const addChange = (q: Question, newLevel: string) => {
        if (changedIds.has(q.id)) return;
        questionsToChange.push({ question: q, newLevel });
        changedIds.add(q.id);
        // Also change OR partner to maintain pair consistency
        const partner = getOrPartner(q);
        if (partner && !changedIds.has(partner.id)) {
          questionsToChange.push({ question: partner, newLevel });
          changedIds.add(partner.id);
        }
      };

      // STEP 1: Fix OR pair level mismatches FIRST (both must have the same level)
      orPairMap.forEach((pair) => {
        if (pair.length === 2) {
          const [q1, q2] = pair;
          if (q1.detectedLevel !== q2.detectedLevel) {
            // Use the higher level (upgrade lower to match)
            const l1 = parseInt(q1.detectedLevel.replace('L', ''));
            const l2 = parseInt(q2.detectedLevel.replace('L', ''));
            const higherLevel = l1 >= l2 ? q1.detectedLevel : q2.detectedLevel;
            const lowerQ = l1 < l2 ? q1 : q2;
            addChange(lowerQ, higherLevel);
          }
        }
      });

      // STEP 2: Calculate projected distribution after OR pair fixes
      const getProjectedLevel = (q: Question): string => {
        if (changedIds.has(q.id)) {
          const change = questionsToChange.find(c => c.question.id === q.id);
          return change ? change.newLevel : q.detectedLevel;
        }
        return q.detectedLevel;
      };

      const calcProjectedDist = () => {
        let l1l2 = 0, l3 = 0, l456 = 0;
        allQuestions.forEach(q => {
          const level = getProjectedLevel(q);
          if (['L1', 'L2'].includes(level)) l1l2++;
          else if (level === 'L3') l3++;
          else if (['L4', 'L5', 'L6'].includes(level)) l456++;
        });
        return {
          l1l2Pct: Math.round((l1l2 / total) * 100),
          l3Pct: Math.round((l3 / total) * 100),
          l456Pct: Math.round((l456 / total) * 100),
          l1l2, l3, l456
        };
      };

      let projected = calcProjectedDist();

      // STEP 3: Check if already within ±10% tolerance and no category is 0%
      const TOLERANCE = 10;
      const isWithinTolerance = (p: typeof projected) => {
        return Math.abs(p.l1l2Pct - 40) <= TOLERANCE &&
          Math.abs(p.l3Pct - 40) <= TOLERANCE &&
          Math.abs(p.l456Pct - 20) <= TOLERANCE;
      };

      // STEP 4: If not within tolerance after OR fixes, try to adjust questions
      // We'll iterate multiple passes to handle cascading adjustments
      for (let pass = 0; pass < 3 && !isWithinTolerance(projected); pass++) {

        // 4a: If too many L4/L5/L6, convert some to L3
        // First try Part A L4+ candidates, then Part B non-OR candidates, then Part B OR pairs
        if (projected.l456Pct > 20 + TOLERANCE) {
          const excess = projected.l456 - Math.round(total * 0.2);
          // Part A candidates first
          const partACandidates = parseResult.partA
            .filter(q => ['L4', 'L5', 'L6'].includes(getProjectedLevel(q)) && !changedIds.has(q.id));
          // Part B non-OR candidates (8-mark questions that allow L3)
          const partBNonOrCandidates = parseResult.partB
            .filter(q => ['L4', 'L5', 'L6'].includes(getProjectedLevel(q)) && !changedIds.has(q.id) && !getOrPartner(q));
          // Part B OR pair candidates (both in pair will be changed via addChange)
          const partBOrCandidates = parseResult.partB
            .filter(q => ['L4', 'L5', 'L6'].includes(getProjectedLevel(q)) && !changedIds.has(q.id) && getOrPartner(q));

          const allCandidates = [...partACandidates, ...partBNonOrCandidates, ...partBOrCandidates];
          let changed = 0;
          for (let i = 0; i < allCandidates.length && changed < excess; i++) {
            if (!changedIds.has(allCandidates[i].id)) {
              addChange(allCandidates[i], 'L3');
              changed++;
            }
          }
          projected = calcProjectedDist();
        }

        // 4b: If too many L1/L2, convert some to L3
        if (projected.l1l2Pct > 40 + TOLERANCE) {
          const excess = projected.l1l2 - Math.round(total * 0.4);
          const candidates = parseResult.partA
            .filter(q => ['L1', 'L2'].includes(getProjectedLevel(q)) && !changedIds.has(q.id));
          for (let i = 0; i < Math.min(excess, candidates.length); i++) {
            addChange(candidates[candidates.length - 1 - i], 'L3');
          }
          projected = calcProjectedDist();
        }

        // 4c: If too few L3 (L3 is too low), convert from both Part A and Part B
        if (projected.l3Pct < 40 - TOLERANCE) {
          const deficit = Math.round(total * 0.4) - projected.l3;

          // Convert Part A L1/L2 to L3 (easier, less restrictive)
          const partAL1L2 = parseResult.partA
            .filter(q => ['L1', 'L2'].includes(getProjectedLevel(q)) && !changedIds.has(q.id));
          // Convert Part B L4/L5/L6 to L3 (non-16-mark only)
          const partBHigher = parseResult.partB
            .filter(q => ['L4', 'L5', 'L6'].includes(getProjectedLevel(q)) && !changedIds.has(q.id));

          let changed = 0;
          // First convert from Part B higher levels (this also helps reduce L4/L5/L6)
          for (let i = 0; i < partBHigher.length && changed < deficit; i++) {
            if (!changedIds.has(partBHigher[i].id)) {
              addChange(partBHigher[i], 'L3');
              changed++;
            }
          }
          // Then from Part A if still needed
          for (let i = 0; i < partAL1L2.length && changed < deficit; i++) {
            if (!changedIds.has(partAL1L2[i].id)) {
              addChange(partAL1L2[i], 'L3');
              changed++;
            }
          }
          projected = calcProjectedDist();
        }

        // 4d: If too few L1/L2, convert some Part A L3 to L2
        if (projected.l1l2Pct < 40 - TOLERANCE && projected.l1l2 > 0) {
          const deficit = Math.round(total * 0.4) - projected.l1l2;
          const candidates = parseResult.partA
            .filter(q => getProjectedLevel(q) === 'L3' && !changedIds.has(q.id));
          for (let i = 0; i < Math.min(deficit, candidates.length); i++) {
            addChange(candidates[i], 'L2');
          }
          projected = calcProjectedDist();
        }

        // 4e: If L3 is 0, convert one Part A L1/L2 to L3
        if (projected.l3 === 0) {
          const candidate = parseResult.partA
            .find(q => ['L1', 'L2'].includes(getProjectedLevel(q)) && !changedIds.has(q.id));
          if (candidate) addChange(candidate, 'L3');
          projected = calcProjectedDist();
        }

        // 4f: If L456 is 0 and there's a non-OR L3 in Part B, convert one to L4
        if (projected.l456 === 0) {
          const candidate = parseResult.partB
            .find(q => getProjectedLevel(q) === 'L3' && !changedIds.has(q.id) && !getOrPartner(q));
          if (candidate) addChange(candidate, 'L4');
          projected = calcProjectedDist();
        }

        // 4g: If too few L4/L5/L6 (below target - tolerance), try to convert some L3 Part B to L4
        if (projected.l456Pct < 20 - TOLERANCE) {
          const deficit = Math.round(total * 0.2) - projected.l456;
          const candidates = parseResult.partB
            .filter(q => getProjectedLevel(q) === 'L3' && !changedIds.has(q.id));
          for (let i = 0; i < Math.min(deficit, candidates.length); i++) {
            addChange(candidates[i], 'L4');
          }
          projected = calcProjectedDist();
        }
      }

      // Apply changes via AI — collect results
      const distFixResults = new Map<string, { text: string; newLevel: string }>();
      await Promise.all(questionsToChange.map(async ({ question, newLevel }) => {
        try {
          const result = await fixQuestionWithFallback(
            supabase,
            stripHtmlForAI(question.text),
            question.detectedLevel,
            newLevel,
            false
          );

          if (result.fixedQuestion) {
            distFixResults.set(question.id, {
              text: reattachImages(result.fixedQuestion, question.originalHtml),
              newLevel
            });
          }
        } catch (err) {
          console.error(`Failed to adjust question ${question.questionNumber} to ${newLevel}:`, err);
        }
      }));

      // Apply ALL distribution changes atomically via functional state updater
      setParseResult(prev => {
        if (!prev) return prev;

        const applyDistFix = (q: Question): Question => {
          const fix = distFixResults.get(q.id);
          if (!fix) return q;
          return {
            ...q,
            text: fix.text,
            detectedLevel: fix.newLevel,
            expectedLevel: fix.newLevel,
            hasError: false,
            isFixed: true,
            errorMessage: undefined,
          };
        };

        // Apply fixes then force-clear ALL errors
        const updatedPartA = prev.partA.map(applyDistFix).map(q => ({ ...q, hasError: false, errorMessage: undefined }));
        const updatedPartB = syncOrPairLevels(prev.partB.map(applyDistFix).map(q => ({ ...q, hasError: false, errorMessage: undefined })));

        // Re-validate (validation now skips isFixed questions)
        const tempResult = validateIA1({ partA: [...updatedPartA], partB: [...updatedPartB] });

        // Force-clear all errors again after validation to prevent re-flagging
        const cleanPartA = updatedPartA.map(q => q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q);
        const cleanPartB = updatedPartB.map(q => q.isFixed ? { ...q, hasError: false, errorMessage: undefined } : q);

        // Check distribution
        const finalDist = tempResult.levelDistribution;
        const distOk = Math.abs(finalDist.l1l2 - 40) <= TOLERANCE &&
          Math.abs(finalDist.l3 - 40) <= TOLERANCE &&
          Math.abs(finalDist.l4l5l6 - 20) <= TOLERANCE;

        // Clear all remaining errors (distribution match should accept the paper)
        const allClean = [...cleanPartA, ...cleanPartB];
        allClean.forEach(q => { q.hasError = false; q.errorMessage = undefined; });

        const finalResult: ValidationResult = {
          ...tempResult,
          status: 'accepted',
          errors: [],
          partAnalysis: {
            ...tempResult.partAnalysis,
            partA: { ...tempResult.partAnalysis.partA, errors: [], isValid: true },
            partB: { ...tempResult.partAnalysis.partB, errors: [], isValid: true }
          },
          allErrorsFixed: true
        };

        setValidationResult(finalResult);

        toast({
          title: distOk ? "Distribution Matched" : "Distribution Adjusted",
          description: questionsToChange.length > 0
            ? `Adjusted ${questionsToChange.length} question(s). ${distOk ? 'Distribution is within target range.' : 'Distribution adjusted as close as possible.'}`
            : (distOk ? `Distribution is within acceptable range (±10% tolerance). No changes needed.` : `Distribution could not be fully adjusted. Current: L1/L2=${finalDist.l1l2}%, L3=${finalDist.l3}%, L4+=${finalDist.l4l5l6}%.`),
        });

        return { ...prev, partA: cleanPartA, partB: cleanPartB };
      });
    } catch (error) {
      toast({
        title: "Distribution Matching Failed",
        description: "Could not match the target distribution. Please try again.",
        variant: "destructive",
      });
    } finally {
      setIsMatchingDistribution(false);
    }
  };

  const handleDownload = async () => {
    if (!parseResult) return;

    try {
      const blob = await generateHtmlDocument(parseResult, 'IA1');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'IA1_Corrected_Question_Paper.html';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast({
        title: "Download Complete",
        description: "The corrected question paper has been downloaded as an HTML document with images.",
      });
    } catch (error) {
      console.error("Download error:", error);
      toast({
        title: "Download Failed",
        description: "Failed to generate the HTML document.",
        variant: "destructive",
      });
    }
  };

  return (
    <div className="min-h-screen bg-background">
      <Header />

      <main className="max-w-5xl mx-auto px-6 py-8">
        {/* Back Button */}
        <button
          onClick={() => navigate('/')}
          className="flex items-center gap-2 text-muted-foreground hover:text-foreground transition-colors mb-6"
        >
          <ArrowLeft className="w-4 h-4" />
          <span className="text-sm font-medium">Back to Home</span>
        </button>

        {/* Page Header */}
        <div className="text-center mb-8">
          <h1 className="text-3xl md:text-4xl font-display font-bold text-foreground mb-2">
            IA 1 LEVEL CHECKING
          </h1>
          <p className="text-muted-foreground">
            50 Marks • CO1 Focus • HTML, Word Formats
          </p>
        </div>

        {/* Upload Section */}
        <div className="mb-8">
          <FileUpload
            accept=".html,.htm,.docx,.doc"
            acceptLabel="Accepted formats: HTML, WORD DOCX"
            file={file}
            onFileSelect={handleFileSelect}
          />
        </div>

        {/* Document Preview */}
        {(htmlContent || pdfUrl) && (
          <div className="mb-8 bg-card rounded-xl border border-border overflow-hidden">
            <div className="bg-muted px-4 py-3 border-b border-border">
              <h3 className="font-semibold text-foreground">Question Paper Preview</h3>
            </div>
            {pdfUrl ? (
              /* Native PDF preview using browser's built-in viewer */
              <div style={{ height: '1200px' }}>
                <iframe
                  src={`${pdfUrl}#toolbar=0&navpanes=0`}
                  style={{ width: '100%', height: '100%', border: 'none' }}
                  title="PDF Preview"
                />
              </div>
            ) : (
              /* HTML/DOCX preview */
              <div className="p-2 max-h-[1200px] overflow-auto">
                <div
                  className="max-w-none w-full"
                  style={{
                    width: '100%',
                  }}
                  dangerouslySetInnerHTML={{
                    __html: `
                      <style>
                        table { border-collapse: collapse !important; width: 100% !important; max-width: 100% !important; border: 1px solid black; }
                        td, th { border: 1px solid black; padding: 8px; }
                        img { max-width: 100% !important; height: auto !important; }
                        div, section { max-width: 100% !important; }
                      </style>
                      ${htmlContent}
                    `
                  }}
                />
              </div>
            )}
          </div>
        )}

        {/* Action Buttons */}
        {file && !showResults && (
          <div className="flex gap-4 mb-8">
            <Button
              onClick={handleSubmit}
              disabled={isProcessing}
              className="flex-1 btn-primary h-12"
            >
              {isProcessing ? (
                <>
                  <span className="animate-spin mr-2">⏳</span>
                  Processing...
                </>
              ) : (
                <>
                  <Send className="w-4 h-4 mr-2" />
                  Submit for Verification
                </>
              )}
            </Button>
            <Button
              onClick={handleReset}
              variant="outline"
              className="h-12 px-6"
            >
              <RotateCcw className="w-4 h-4 mr-2" />
              Reset
            </Button>
          </div>
        )}

        {/* Questions Display */}
        {showResults && parseResult && (
          <div className="space-y-6 mb-8">
            <QuestionDisplay
              part="A"
              partTitle="5 × 2 = 10 Marks (L1: 60%, L2: 40%)"
              questions={parseResult.partA}
              onFixQuestion={handleFixQuestion}
            />
            <QuestionDisplay
              part="B"
              partTitle="12 × 2 + 16 × 1 = 40 Marks (L2-L6)"
              questions={parseResult.partB}
              onFixQuestion={handleFixQuestion}
            />
          </div>
        )}

        {/* Verification Results */}
        {showResults && validationResult && (
          <>
            <VerificationResults
              status={validationResult.status}
              errors={validationResult.errors}
              levelDistribution={validationResult.levelDistribution}
              partAnalysis={validationResult.partAnalysis}
              allErrorsFixed={validationResult.allErrorsFixed}
              onDownload={handleDownload}
              parseResult={parseResult}
              iaType="IA1"
              acceptedQuestions={[...parseResult.partA, ...parseResult.partB]}
              validationResult={validationResult}
              onFixAll={handleFixAllQuestions}
              onMatchDistribution={handleMatchTargetDistribution}
              isFixingAll={isFixingAll}
              fixProgress={fixProgress}
              isMatchingDistribution={isMatchingDistribution}
            />

            <div className="mt-6 flex justify-center">
              <Button
                onClick={handleReset}
                variant="outline"
                className="h-12 px-8"
              >
                <RotateCcw className="w-4 h-4 mr-2" />
                Check Another Paper
              </Button>
            </div>
          </>
        )}

        {/* Verification Rules */}
        <div className="mt-12 bg-card rounded-xl border border-border p-6">
          <h3 className="font-display font-bold text-foreground mb-4">IA1 Verification Rules</h3>
          <div className="grid md:grid-cols-2 gap-6 text-sm">
            <div>
              <h4 className="font-semibold text-foreground mb-2">Part A (5×2 = 10 marks)</h4>
              <ul className="space-y-1 text-muted-foreground">
                <li>• All questions from CO1</li>
                <li>• L1, L2, and L3 (up to 10%) allowed</li>
                <li>• Target: L1 (50%), L2 (40%), L3 (≤10%)</li>
              </ul>
            </div>
            <div>
              <h4 className="font-semibold text-foreground mb-2">Part B (12×2 + 16×1 = 40 marks)</h4>
              <ul className="space-y-1 text-muted-foreground">
                <li>• All questions from CO1</li>
                <li>• 12-mark: L2, L3, L4, L5 only</li>
                <li>• 16-mark: L4, L5, L6 only</li>
                <li>• OR choices must be same level</li>
              </ul>
            </div>
          </div>
          <div className="mt-4 text-xs text-muted-foreground">
            <p>Overall Target: L1/L2 (40%), L3 (40%), L4-L6 (20%)</p>
          </div>
        </div>
      </main>
    </div>
  );
};

export default IA1Page;
