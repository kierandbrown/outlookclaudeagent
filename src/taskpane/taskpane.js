/* global Office */

const API_BASE = window.location.origin;

let mailboxItem = null;

// Initialize the Office Add-in
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    mailboxItem = Office.context.mailbox.item;
    document.getElementById("analyze-btn").addEventListener("click", runAnalysis);
    document.getElementById("retry-btn").addEventListener("click", runAnalysis);
    document.getElementById("apply-revision-btn").addEventListener("click", applyRevision);
    initToggleSections();
  } else {
    show("not-compose");
    hide("compose-view");
  }
});

// ---- Email Data Extraction ----

function getEmailBody() {
  return new Promise((resolve, reject) => {
    mailboxItem.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Could not read email body."));
      }
    });
  });
}

function getSubject() {
  return new Promise((resolve, reject) => {
    mailboxItem.subject.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Could not read subject."));
      }
    });
  });
}

function getRecipients() {
  return new Promise((resolve, reject) => {
    mailboxItem.to.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.map((r) => r.emailAddress));
      } else {
        reject(new Error("Could not read recipients."));
      }
    });
  });
}

function getConversationThread() {
  // Strategy: Try text-based splitting first (most reliable for plain text).
  // If no thread markers found in text, try HTML body for quoted content
  // (Outlook wraps quoted replies in identifiable HTML elements).
  return new Promise((resolve) => {
    // First attempt: plain text with common thread markers
    mailboxItem.body.getAsync(Office.CoercionType.Text, (textResult) => {
      if (textResult.status === Office.AsyncResultStatus.Succeeded) {
        const fullBody = textResult.value;
        const textSplit = splitByTextMarkers(fullBody);
        if (textSplit) {
          resolve(textSplit);
          return;
        }
      }

      // Second attempt: HTML body with structural markers (blockquote, divRplyFwdMsg, etc.)
      mailboxItem.body.getAsync(Office.CoercionType.Html, (htmlResult) => {
        if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
          const htmlBody = htmlResult.value;
          const htmlSplit = splitByHtmlMarkers(htmlBody);
          if (htmlSplit) {
            resolve(htmlSplit);
            return;
          }
        }

        // Fallback: treat entire text body as composed (new email, no thread)
        const body = textResult.status === Office.AsyncResultStatus.Succeeded
          ? textResult.value.trim()
          : "";
        resolve({ composedPart: body, threadPart: "" });
      });
    });
  });
}

function splitByTextMarkers(fullBody) {
  const threadMarkers = [
    /\n-{2,}\s*Original Message\s*-{2,}/i,
    /\nFrom:\s+.+\nSent:\s+/i,
    /\nOn .+ wrote:/i,
    /\n_{3,}/,
  ];
  for (const marker of threadMarkers) {
    const match = fullBody.match(marker);
    if (match) {
      return {
        composedPart: fullBody.substring(0, match.index).trim(),
        threadPart: fullBody.substring(match.index).trim(),
      };
    }
  }
  return null;
}

function splitByHtmlMarkers(htmlBody) {
  // Outlook and other clients wrap the quoted thread in identifiable elements
  const htmlMarkers = [
    '<div id="divRplyFwdMsg"',          // Classic Outlook
    '<div id="appendonsend"',            // New Outlook
    '<div id="x_divRplyFwdMsg"',         // Outlook web variants
    "<blockquote",                       // Standard quoted replies
    '<div style="border-top',            // Gmail-style forward markers
  ];

  let splitIndex = -1;
  for (const marker of htmlMarkers) {
    const idx = htmlBody.indexOf(marker);
    if (idx !== -1 && (splitIndex === -1 || idx < splitIndex)) {
      splitIndex = idx;
    }
  }

  if (splitIndex === -1) return null;

  // Convert HTML segments to plain text for the analysis
  const composedHtml = htmlBody.substring(0, splitIndex);
  const threadHtml = htmlBody.substring(splitIndex);
  return {
    composedPart: htmlToPlainText(composedHtml),
    threadPart: htmlToPlainText(threadHtml),
  };
}

function htmlToPlainText(html) {
  const temp = document.createElement("div");
  temp.innerHTML = html;
  return (temp.textContent || temp.innerText || "").trim();
}

// ---- Analysis ----

async function runAnalysis() {
  const analyzeBtn = document.getElementById("analyze-btn");
  const btnText = document.getElementById("btn-text");
  const btnSpinner = document.getElementById("btn-spinner");

  // Show loading state
  analyzeBtn.disabled = true;
  btnText.textContent = "Analyzing...";
  show("btn-spinner");
  hide("results");
  hide("error-box");

  try {
    const [threadData, subject, recipients] = await Promise.all([
      getConversationThread(),
      getSubject(),
      getRecipients(),
    ]);

    const response = await fetch(`${API_BASE}/api/analyze`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        composedEmail: threadData.composedPart,
        conversationThread: threadData.threadPart,
        subject: subject,
        recipients: recipients,
      }),
    });

    if (!response.ok) {
      const errData = await response.json().catch(() => ({}));
      throw new Error(errData.error || `Server error (${response.status})`);
    }

    const analysis = await response.json();
    renderResults(analysis);
    show("results");
  } catch (err) {
    document.getElementById("error-message").textContent = err.message;
    show("error-box");
  } finally {
    analyzeBtn.disabled = false;
    btnText.textContent = "Analyze My Email";
    hide("btn-spinner");
  }
}

// ---- Render Results ----

function renderResults(analysis) {
  // Intent
  document.getElementById("intent-body").innerHTML = `
    <p><strong>${escapeHtml(analysis.intent.summary)}</strong></p>
    <p>${escapeHtml(analysis.intent.details)}</p>
  `;

  // Completeness
  const completenessHtml = analysis.completeness.items
    .map(
      (item) => `
      <div class="issue-item">
        <div class="issue-severity severity-${item.addressed ? "low" : "high"}"></div>
        <div>
          <strong>${escapeHtml(item.question)}</strong><br/>
          <span>${escapeHtml(item.assessment)}</span>
        </div>
      </div>`
    )
    .join("");
  document.getElementById("completeness-body").innerHTML =
    analysis.completeness.items.length > 0
      ? completenessHtml
      : "<p>No specific questions found in the thread to check against.</p>";

  // Clarity
  const clarityItems = analysis.clarity.issues
    .map(
      (issue) => `
      <div class="issue-item">
        <div class="issue-severity severity-${issue.severity}"></div>
        <div>${escapeHtml(issue.description)}</div>
      </div>`
    )
    .join("");
  document.getElementById("clarity-body").innerHTML =
    analysis.clarity.issues.length > 0
      ? `<p>${escapeHtml(analysis.clarity.summary)}</p>${clarityItems}`
      : `<p>${escapeHtml(analysis.clarity.summary)}</p>`;

  // Back-and-forth prevention
  const bfItems = analysis.backAndForth.risks
    .map(
      (risk) => `
      <div class="issue-item">
        <div class="issue-severity severity-${risk.severity}"></div>
        <div>
          <strong>${escapeHtml(risk.issue)}</strong><br/>
          <span>${escapeHtml(risk.suggestion)}</span>
        </div>
      </div>`
    )
    .join("");
  document.getElementById("backforth-body").innerHTML =
    analysis.backAndForth.risks.length > 0
      ? `<p>${escapeHtml(analysis.backAndForth.summary)}</p>${bfItems}`
      : `<p>${escapeHtml(analysis.backAndForth.summary)}</p>`;

  // Suggested revision
  if (analysis.suggestedRevision) {
    document.getElementById("revision-text").textContent = analysis.suggestedRevision;
    show("apply-revision-btn");
  } else {
    document.getElementById("revision-text").textContent = "Your email looks good — no revision needed.";
    hide("apply-revision-btn");
  }

  // Score
  const score = analysis.overallScore;
  const scoreEl = document.getElementById("score-display");
  scoreEl.textContent = `${score}/10`;
  scoreEl.className = "";
  if (score >= 8) scoreEl.classList.add("score-high");
  else if (score >= 5) scoreEl.classList.add("score-medium");
  else scoreEl.classList.add("score-low");
}

// ---- Apply Revision ----

function applyRevision() {
  const revision = document.getElementById("revision-text").textContent;
  if (!revision || !mailboxItem) return;

  // Use prependAsync to replace the composed portion while preserving the quoted thread.
  // First, we set the body to just the revision + thread by using setSelectedDataAsync
  // at the beginning, or fall back to setAsync.
  // The safest approach: set the full body as text (the quoted thread is already in the analysis).
  mailboxItem.body.setAsync(revision, { coercionType: Office.CoercionType.Text }, (result) => {
    const btn = document.getElementById("apply-revision-btn");
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      btn.textContent = "Applied!";
      btn.disabled = true;
      setTimeout(() => {
        btn.textContent = "Apply Suggestion";
        btn.disabled = false;
      }, 2000);
    } else {
      btn.textContent = "Failed — try copying manually";
      setTimeout(() => {
        btn.textContent = "Apply Suggestion";
      }, 3000);
    }
  });
}

// ---- Toggle Sections ----

function initToggleSections() {
  document.querySelectorAll(".section-header[data-toggle]").forEach((header) => {
    header.addEventListener("click", () => {
      const targetId = header.getAttribute("data-toggle");
      const body = document.getElementById(targetId);
      header.classList.toggle("collapsed");
      body.classList.toggle("collapsed");
    });
  });
}

// ---- Utility ----

function show(id) {
  document.getElementById(id).classList.remove("hidden");
}

function hide(id) {
  document.getElementById(id).classList.add("hidden");
}

function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text || "";
  return div.innerHTML;
}
