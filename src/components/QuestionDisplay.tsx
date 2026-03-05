import { Sparkles, CheckCircle } from 'lucide-react';
import { Button } from '@/components/ui/button';

export interface Question {
  id: string;
  questionNumber: string;
  text: string;
  marks: number;
  detectedLevel: string;
  expectedLevel: string;
  co: string;
  expectedCo?: string; // Expected CO when converting OR-pair structure mismatch
  hasError: boolean;
  errorMessage?: string;
  isFixed?: boolean;
  isFixing?: boolean;
  hasSubdivisions?: boolean;
  subdivisionCount?: number;
  subdivisionLevels?: string[];
  isEnhanced?: boolean;
  levelAutoDetected?: boolean;
  originalHtml?: string; // Preserves original HTML with images/formulas
  pageImageDataUrl?: string; // Base64 data URL of rendered PDF page (for circuit diagrams etc.)
}

interface QuestionDisplayProps {
  part: string;
  partTitle: string;
  questions: Question[];
  onFixQuestion: (questionId: string) => void;
}

const getLevelBadgeClass = (level: string) => {
  const levelMap: Record<string, string> = {
    L1: 'level-l1',
    L2: 'level-l2',
    L3: 'level-l3',
    L4: 'level-l4',
    L5: 'level-l5',
    L6: 'level-l6',
  };
  return levelMap[level] || 'level-l1';
};

import { containsHtml, stripHtmlForAI, extractImageHtml, reattachImages } from '@/utils/htmlUtils';

// Parse subdivisions from question text (e.g., "i) question1 ii) question2")
const parseSubdivisionsFromText = (text: string): { subdivisions: string[]; hasSubdivisions: boolean } => {
  // Use plain text for subdivision parsing
  const plainText = containsHtml(text) ? stripHtmlForAI(text) : text;
  const subdivisions: string[] = [];

  // Strategy 1: Split approach - split on Roman numeral patterns that start a subdivision
  // Match standalone Roman numeral markers like "i)", "ii)", "iii)", "iv)", "(i)", "(ii)"
  // The marker must be preceded by start-of-string or a space (not in the middle of a word)
  const splitPattern = /(?:^|\s)((?:i{1,3}|iv|v|vi{0,3})[)\.]|(?:\((?:i{1,3}|iv|v|vi{0,3})\)))\s*/gi;

  // Find all marker positions
  const markers: { index: number; marker: string }[] = [];
  let match;
  splitPattern.lastIndex = 0;
  while ((match = splitPattern.exec(plainText)) !== null) {
    markers.push({
      index: match.index + (match[0].startsWith(' ') ? 1 : 0),
      marker: match[1].trim()
    });
  }

  if (markers.length >= 2) {
    for (let i = 0; i < markers.length; i++) {
      const start = markers[i].index;
      const end = i + 1 < markers.length ? markers[i + 1].index : plainText.length;
      const subText = plainText.substring(start, end).trim();
      if (subText.length > 3) {
        subdivisions.push(subText);
      }
    }
  }

  // Strategy 2: If split approach didn't work, try simpler split
  if (subdivisions.length < 2) {
    subdivisions.length = 0;
    const simplePattern = /([ivxIVX]+\))\s*/gi;
    const parts = plainText.split(simplePattern).filter(p => p.trim().length > 0);

    if (parts.length >= 3) {
      for (let i = 0; i < parts.length - 1; i += 2) {
        if (/^[ivxIVX]+\)$/i.test(parts[i])) {
          const subText = parts[i] + ' ' + (parts[i + 1] || '').trim();
          if (subText.length > 5) {
            subdivisions.push(subText);
          }
        }
      }
    }
  }

  return {
    subdivisions,
    hasSubdivisions: subdivisions.length >= 2
  };
};

const QuestionDisplay = ({ part, partTitle, questions, onFixQuestion }: QuestionDisplayProps) => {
  return (
    <div className="bg-card rounded-xl border border-border overflow-hidden">
      <div className="gradient-primary px-4 py-3">
        <h3 className="text-lg font-display font-bold text-primary-foreground">
          Part {part}: {partTitle}
        </h3>
      </div>

      <div className="divide-y divide-border">
        {questions.map((question) => {
          // Parse subdivisions from text if the question has them
          const { subdivisions, hasSubdivisions } = parseSubdivisionsFromText(question.text);
          const showSubdivisions = hasSubdivisions || (question.hasSubdivisions && question.subdivisionLevels && question.subdivisionLevels.length > 1);

          return (
            <div
              key={question.id}
              className={`p-4 ${question.hasError ? 'bg-destructive/5' : ''} ${question.isFixed ? 'bg-success/10' : ''}`}
            >
              <div className="flex items-start justify-between gap-4">
                <div className="flex-1">
                  <div className="flex items-center gap-2 mb-2 flex-wrap">
                    <span className="font-semibold text-foreground">
                      Q{question.questionNumber}
                    </span>

                    {/* Show subdivision levels if present, otherwise show single level */}
                    {showSubdivisions && question.subdivisionLevels && question.subdivisionLevels.length > 1 ? (
                      <div className="flex items-center gap-1">
                        {question.subdivisionLevels.map((level, idx) => (
                          <span key={idx} className={`level-badge ${getLevelBadgeClass(level)}`}>
                            {level}
                          </span>
                        ))}
                      </div>
                    ) : (
                      <span className={`level-badge ${getLevelBadgeClass(question.detectedLevel)}`}>
                        {question.detectedLevel}
                      </span>
                    )}

                    {question.hasError && question.expectedLevel !== question.detectedLevel && (
                      <>
                        <span className="text-xs text-muted-foreground">→</span>
                        {/* Show target subdivision levels or single level */}
                        {question.expectedLevel.includes(',') ? (
                          <div className="flex items-center gap-1">
                            {question.expectedLevel.split(',').map((level, idx) => (
                              <span key={idx} className={`level-badge ${getLevelBadgeClass(level.trim())} opacity-70`}>
                                {level.trim()}
                              </span>
                            ))}
                          </div>
                        ) : (
                          <span className={`level-badge ${getLevelBadgeClass(question.expectedLevel)} opacity-70`}>
                            {question.expectedLevel}
                          </span>
                        )}
                      </>
                    )}
                    <span className="text-xs text-muted-foreground">
                      {question.co}
                    </span>
                    {/* Show total marks for the question */}
                    <span className="text-xs text-muted-foreground">
                      [{question.marks} marks]
                    </span>
                    {question.isFixed && (
                      <span className="flex items-center gap-1 text-xs text-success">
                        <CheckCircle className="w-3 h-3" />
                        Fixed
                      </span>
                    )}
                  </div>

                  {/* Display question with subdivisions formatted properly - VERTICAL layout */}
                  {showSubdivisions && subdivisions.length >= 2 ? (
                    <div className="space-y-3 mt-2">
                      {subdivisions.map((sub, idx) => {
                        // Calculate marks per subdivision (split equally from total)
                        const marksPerSub = Math.floor(question.marks / subdivisions.length);
                        const level = question.subdivisionLevels?.[idx] || question.detectedLevel;
                        return (
                          <div key={idx} className="flex items-start gap-3 p-3 bg-muted/30 rounded-lg border border-border/50">
                            <div className="flex-1">
                              <p className="text-sm text-foreground leading-relaxed">{sub}</p>
                            </div>
                            <div className="flex items-center gap-2 shrink-0">
                              <span className={`level-badge ${getLevelBadgeClass(level)}`}>
                                {level}
                              </span>
                              <span className="text-xs font-medium text-muted-foreground bg-background px-2 py-0.5 rounded">
                                {marksPerSub} marks
                              </span>
                            </div>
                          </div>
                        );
                      })}
                      {/* Show images from originalHtml after subdivisions */}
                      {question.originalHtml && containsHtml(question.originalHtml) && (
                        <div
                          className="text-sm text-foreground leading-relaxed question-html-content mt-2"
                          dangerouslySetInnerHTML={{ __html: extractImageHtml(question.originalHtml).join('') }}
                        />
                      )}
                    </div>
                  ) : (
                    /* Render as HTML if contains images/formulas, otherwise plain text */
                    containsHtml(question.originalHtml || question.text) ? (
                      <div
                        className="text-sm text-foreground leading-relaxed question-html-content"
                        dangerouslySetInnerHTML={{ __html: question.originalHtml || question.text }}
                      />
                    ) : (
                      <p className="text-sm text-foreground leading-relaxed">
                        {question.text}
                      </p>
                    )
                  )}

                  {question.hasError && question.errorMessage && !question.isFixed && (
                    <p className="mt-2 text-sm error-text">
                      ⚠️ {question.errorMessage}
                    </p>
                  )}
                </div>


              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default QuestionDisplay;
