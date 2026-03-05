import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
};

// RBT Level action verbs mapping - updated with user's specific verbs
const RBT_VERBS: Record<string, string[]> = {
  L1: ['Identify', 'Label', 'List', 'State', 'Define', 'Name', 'Recognize'],
  L2: ['Describe', 'Compare', 'Restate', 'Illustrate', 'Explain', 'Summarize', 'Interpret'],
  L3: ['Calculate', 'Apply', 'Solve', 'Demonstrate', 'Compute', 'Use', 'Determine'],
  L4: ['Break down', 'Categorize', 'Outline', 'Simplify', 'Analyze', 'Differentiate', 'Examine'],
  L5: ['Select', 'Rate', 'Defend', 'Verify', 'Evaluate', 'Justify', 'Critique'],
  L6: ['Plan', 'Modify', 'Propose', 'Develop', 'Design', 'Create', 'Construct']
};

// Helper function to delay execution
const delay = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

// Helper function to make API call with retry logic
async function callAIWithRetry(
  apiKey: string,
  systemPrompt: string,
  userPrompt: string,
  maxRetries = 5
): Promise<{ success: boolean; data?: unknown; error?: string; retryAfter?: number }> {
  let lastError = "";
  
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    if (attempt > 0) {
      const waitTime = Math.pow(2, attempt) * 1500;
      console.log(`Retry attempt ${attempt + 1}, waiting ${waitTime}ms`);
      await delay(waitTime);
    }

    try {
      const response = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "google/gemini-2.5-flash-lite",
          messages: [
            { role: "system", content: systemPrompt },
            { role: "user", content: userPrompt },
          ],
          stream: false,
        }),
      });

      if (response.ok) {
        const data = await response.json();
        return { success: true, data };
      }

      if (response.status === 429) {
        const retryAfter = parseInt(response.headers.get("Retry-After") || "5", 10);
        lastError = "Rate limit exceeded";
        console.log(`Rate limited, retry-after: ${retryAfter}s`);
        
        if (attempt < maxRetries - 1) {
          await delay(retryAfter * 1000);
          continue;
        }
        return { success: false, error: "Rate limit exceeded. Please wait a moment and try again.", retryAfter };
      }

      if (response.status === 402) {
        return { success: false, error: "Payment required. Please add credits to your workspace." };
      }

      const errorText = await response.text();
      console.error("AI gateway error:", response.status, errorText);
      lastError = `AI gateway error: ${response.status}`;
    } catch (err) {
      console.error("Fetch error:", err);
      lastError = err instanceof Error ? err.message : "Network error";
    }
  }

  return { success: false, error: lastError || "Failed after multiple retries" };
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const body = await req.json();
    const { questionText, level, subjectContext, useAlternativeVerb } = body;

    // Support batch requests: body.questions = [{ questionText, level, questionNumber, useAlternativeVerb? }]
    const batch = Array.isArray(body.questions) ? body.questions : null;

    if (!batch && (!questionText || !level)) {
      return new Response(
        JSON.stringify({ error: "Missing required fields: questionText, level" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const LOVABLE_API_KEY = Deno.env.get("LOVABLE_API_KEY");
    if (!LOVABLE_API_KEY) {
      throw new Error("LOVABLE_API_KEY is not configured");
    }

    // If batch provided, send one combined prompt asking for JSON output of enhanced questions
    if (batch) {
      // Build a prompt listing questions with their levels
      const questionListText = batch.map((b: any, i: number) => {
        const reTag = b.useAlternativeVerb ? ' (RE-ENHANCE - use a different action verb)' : '';
        const qNum = b.questionNumber ? `#${b.questionNumber} ` : '';
        return `${i + 1}. ${qNum}[${b.level}] ${b.questionText}${reTag}`;
      }).join('\n');

      const systemPrompt = `You are an expert educational content enhancer. Enhance the following questions according to their specified RBT levels.

    Rules:
    - Keep each question at the same RBT level.
    - Each enhanced question MUST start with an appropriate action verb for its level.
    - If a question includes the note '(RE-ENHANCE - use a different action verb)', use a different appropriate verb than the one currently used.
    - Maintain the same topic and meaning.
    - Return a JSON array of objects with keys: questionNumber (if present), questionIndex (1-based index in the list), and enhancedQuestion. Return ONLY the JSON array.`;

      const userPrompt = `Questions:\n${questionListText}\n\nReturn a JSON array like [{"questionIndex":1,"questionNumber":"Q.1","enhancedQuestion":"..."}, ...]`;

      const result = await callAIWithRetry(LOVABLE_API_KEY, systemPrompt, userPrompt);
      if (!result.success) {
        const status = result.error?.includes("Rate limit") ? 429 : result.error?.includes("Payment") ? 402 : 500;
        return new Response(
          JSON.stringify({ error: result.error, retryAfter: result.retryAfter }),
          { status, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }

      const data = result.data as { choices?: Array<{ message?: { content?: string } }> };
      const content = data.choices?.[0]?.message?.content?.trim() || '';
      let parsed: any = null;
      try { parsed = JSON.parse(content); } catch (e) {
        // If response not strict JSON, try to extract JSON substring
        const jsonMatch = content.match(/\[\s*\{[\s\S]*\}\s*\]/);
        if (jsonMatch) {
          try { parsed = JSON.parse(jsonMatch[0]); } catch (e2) { parsed = null; }
        }
      }

      if (!parsed) {
        // Fallback: return each original unchanged
        const fallback = batch.map((b: any, i: number) => ({ questionIndex: i + 1, enhancedQuestion: b.questionText }));
        return new Response(JSON.stringify({ success: true, enhancedQuestions: fallback }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
      }

      return new Response(JSON.stringify({ success: true, enhancedQuestions: parsed }), { headers: { ...corsHeaders, "Content-Type": "application/json" } });
    }

    // Single-question flow (backwards compatibility)
    const levelVerbs = RBT_VERBS[level] || RBT_VERBS.L2;
    
    // Detect current verb used in question and exclude it if re-enhancing
    const excludedVerbs: string[] = [];
    if (useAlternativeVerb) {
      for (const verb of levelVerbs) {
        if (questionText.toLowerCase().startsWith(verb.toLowerCase())) {
          excludedVerbs.push(verb);
          break;
        }
      }
    }
    
    const availableVerbs = excludedVerbs.length > 0 
      ? levelVerbs.filter(v => !excludedVerbs.includes(v))
      : levelVerbs;

    const systemPrompt = `You are an expert educational content enhancer specializing in making academic questions clearer and more student-friendly.

Your task is to enhance a question to make it:
1. More professionally worded
2. Clearer and easier for students to understand
3. Grammatically perfect
4. Appropriately challenging for the given RBT level (NOT more difficult)

RBT Levels and their Action Verbs:
- L1 (Recall): ${RBT_VERBS.L1.join(', ')}
- L2 (Explain): ${RBT_VERBS.L2.join(', ')}
- L3 (Solve): ${RBT_VERBS.L3.join(', ')}
- L4 (Inspect): ${RBT_VERBS.L4.join(', ')}
- L5 (Judge): ${RBT_VERBS.L5.join(', ')}
- L6 (Build): ${RBT_VERBS.L6.join(', ')}

IMPORTANT RULES:
1. Keep the same RBT level - do NOT increase difficulty
2. Keep the same core topic and subject matter
3. Make the question clearer, not harder
4. Use proper academic language
5. Ensure the question is unambiguous
6. The question MUST start with an action verb from level ${level}: ${availableVerbs.join(', ')}
7. ${useAlternativeVerb ? `IMPORTANT: Use a DIFFERENT action verb than before. Avoid: ${excludedVerbs.join(', ')}` : 'Use DIFFERENT action verbs - vary your choice'}
8. Return ONLY the enhanced question text, nothing else`;

    const userPrompt = `Original question (${level}): "${questionText}"
${subjectContext ? `Subject context: ${subjectContext}` : ''}
${useAlternativeVerb ? `\nThis is a RE-ENHANCEMENT request. Please use a DIFFERENT action verb than the current one.` : ''}

Enhance this question to be clearer, more professional, and student-friendly while keeping it at ${level} level.
Use one of these action verbs: ${availableVerbs.join(', ')}
Do NOT make it more difficult - just improve the wording and clarity.

Return ONLY the enhanced question text.`;

    const result = await callAIWithRetry(LOVABLE_API_KEY, systemPrompt, userPrompt);

    if (!result.success) {
      const status = result.error?.includes("Rate limit") ? 429 : 
                     result.error?.includes("Payment") ? 402 : 500;
      return new Response(
        JSON.stringify({ error: result.error, retryAfter: result.retryAfter }),
        { status, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const data = result.data as { choices?: Array<{ message?: { content?: string } }> };
    const enhancedQuestion = data.choices?.[0]?.message?.content?.trim() || questionText;

    return new Response(
      JSON.stringify({ 
        success: true,
        originalQuestion: questionText,
        enhancedQuestion,
        level
      }),
      { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  } catch (error) {
    console.error("enhance-questions error:", error);
    const errorMessage = error instanceof Error ? error.message : "Unknown error";
    return new Response(
      JSON.stringify({ error: errorMessage }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
