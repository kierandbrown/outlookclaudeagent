const Anthropic = require("@anthropic-ai/sdk");

const client = new Anthropic();

const SYSTEM_PROMPT = `You are an expert email communication analyst. Your job is to analyze a composed email (and its conversation thread if present) and provide actionable feedback to make the email more effective.

Your goals:
1. IDENTIFY THE INTENT — What is the sender trying to achieve? (request info, give approval, escalate, follow up, etc.)
2. CHECK COMPLETENESS — If this is a reply, does it address every question/request from the thread? List each question/ask from the thread and whether it's answered.
3. ASSESS CLARITY & TONE — Is the email clear, professional, and unambiguous? Are there vague phrases that could be misinterpreted?
4. PREVENT BACK-AND-FORTH — Identify anything that would likely cause another round of emails: missing information, ambiguous answers, unstated assumptions, missing deadlines/next steps, or questions that should be preemptively answered.

You MUST respond in this exact JSON structure (no markdown, no code fences, just raw JSON):
{
  "intent": {
    "summary": "One-line description of the email's primary intent",
    "details": "2-3 sentence explanation of what the sender is trying to accomplish"
  },
  "completeness": {
    "items": [
      {
        "question": "The question or request from the thread",
        "addressed": true/false,
        "assessment": "How well it was addressed, or what's missing"
      }
    ]
  },
  "clarity": {
    "summary": "Overall clarity assessment in one sentence",
    "issues": [
      {
        "description": "Description of the clarity/tone issue",
        "severity": "high|medium|low"
      }
    ]
  },
  "backAndForth": {
    "summary": "Overall assessment of whether this email will prevent or cause more back-and-forth",
    "risks": [
      {
        "issue": "What might cause another round of emails",
        "suggestion": "How to fix it",
        "severity": "high|medium|low"
      }
    ]
  },
  "suggestedRevision": "A rewritten version of the email that addresses all issues found. Keep the sender's voice and style but improve clarity, completeness, and reduce back-and-forth risk. If the email is already good, set this to null.",
  "overallScore": 7
}

The overallScore is 1-10 where:
- 1-3: Major issues — likely to cause confusion or multiple follow-up emails
- 4-6: Decent but has notable gaps or unclear areas
- 7-8: Good email with minor improvements possible
- 9-10: Excellent — clear, complete, and unlikely to need follow-up

Important rules:
- If there's no conversation thread, skip the completeness check (return empty items array) and focus on clarity and intent for a new email.
- Be specific in your suggestions — don't just say "be more clear", say exactly what to change.
- For the suggested revision, preserve the sender's tone and level of formality.
- Focus on substance over style — prioritize missing information and ambiguity over grammar or formatting.
- Always consider: "If I received this email, would I need to reply asking for clarification?"`;

async function analyzeEmail({ composedEmail, conversationThread, subject, recipients }) {
  const userMessage = buildUserMessage({ composedEmail, conversationThread, subject, recipients });

  const response = await client.messages.create({
    model: "claude-sonnet-4-20250514",
    max_tokens: 4096,
    messages: [
      {
        role: "user",
        content: userMessage,
      },
    ],
    system: SYSTEM_PROMPT,
  });

  const text = response.content[0].text;

  // Parse the JSON response, stripping any accidental markdown fences
  const cleaned = text.replace(/^```(?:json)?\s*/m, "").replace(/\s*```$/m, "");
  const analysis = JSON.parse(cleaned);

  return analysis;
}

function buildUserMessage({ composedEmail, conversationThread, subject, recipients }) {
  let message = `## Email Subject\n${subject}\n\n`;
  message += `## Recipients\n${recipients.length > 0 ? recipients.join(", ") : "Not specified"}\n\n`;

  if (conversationThread) {
    message += `## Conversation Thread (previous messages)\n${conversationThread}\n\n`;
    message += `## Composed Reply (what the user is about to send)\n${composedEmail}\n`;
  } else {
    message += `## Composed Email (new message — no prior thread)\n${composedEmail}\n`;
  }

  return message;
}

module.exports = { analyzeEmail };
