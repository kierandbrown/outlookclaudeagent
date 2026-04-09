require("dotenv").config();
const express = require("express");
const cors = require("cors");
const path = require("path");
const { analyzeEmail } = require("./analyze");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: "500kb" }));

// Serve static files - taskpane, commands, assets
app.use(express.static(path.join(__dirname, "..", "src", "taskpane")));
app.use("/assets", express.static(path.join(__dirname, "..", "src", "assets")));
app.use(
  "/commands.html",
  express.static(path.join(__dirname, "..", "src", "commands", "commands.html"))
);

// Simple in-memory rate limiter (per IP, 10 requests per minute)
const rateLimitMap = new Map();
function rateLimit(req, res, next) {
  const ip = req.ip;
  const now = Date.now();
  const windowMs = 60000;
  const maxRequests = 10;

  if (!rateLimitMap.has(ip)) {
    rateLimitMap.set(ip, []);
  }
  const timestamps = rateLimitMap.get(ip).filter((t) => now - t < windowMs);
  if (timestamps.length >= maxRequests) {
    return res.status(429).json({ error: "Too many requests. Please wait a moment." });
  }
  timestamps.push(now);
  rateLimitMap.set(ip, timestamps);
  next();
}

// Strip script tags and event handlers from email content before analysis
function sanitizeInput(text) {
  if (!text) return text;
  return text
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/\bon\w+\s*=\s*["'][^"']*["']/gi, "");
}

// Health check
app.get("/api/health", (_req, res) => {
  res.json({ status: "ok" });
});

// Main analysis endpoint
app.post("/api/analyze", rateLimit, async (req, res) => {
  const { composedEmail, conversationThread, subject, recipients } = req.body;

  if (!composedEmail || composedEmail.trim().length === 0) {
    return res.status(400).json({ error: "No email content provided." });
  }

  try {
    const analysis = await analyzeEmail({
      composedEmail: sanitizeInput(composedEmail),
      conversationThread: sanitizeInput(conversationThread || ""),
      subject: sanitizeInput(subject || "(No subject)"),
      recipients: recipients || [],
    });
    res.json(analysis);
  } catch (err) {
    console.error("Analysis error:", err.message);
    res.status(500).json({
      error: "Failed to analyze email. Please check your API key and try again.",
    });
  }
});

app.listen(PORT, () => {
  console.log(`Email Analyzer server running on https://localhost:${PORT}`);
  console.log("Ensure you have ANTHROPIC_API_KEY set in your .env file.");
});
