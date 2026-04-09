require("dotenv").config();
const express = require("express");
const https = require("https");
const fs = require("fs");
const os = require("os");
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

// --- HTTPS Setup ---
// Look for certificates in multiple locations:
// 1. Local certs/ directory (manual placement)
// 2. office-addin-dev-certs default location (~/.office-addin-dev-certs/)
function findCerts() {
  const locations = [
    // Local certs/ folder (mkcert or manually placed)
    {
      cert: path.join(__dirname, "..", "certs", "localhost.crt"),
      key: path.join(__dirname, "..", "certs", "localhost.key"),
    },
    {
      cert: path.join(__dirname, "..", "certs", "localhost.pem"),
      key: path.join(__dirname, "..", "certs", "localhost-key.pem"),
    },
    // office-addin-dev-certs default location
    {
      cert: path.join(os.homedir(), ".office-addin-dev-certs", "localhost.crt"),
      key: path.join(os.homedir(), ".office-addin-dev-certs", "localhost.key"),
    },
  ];

  for (const loc of locations) {
    if (fs.existsSync(loc.cert) && fs.existsSync(loc.key)) {
      return {
        cert: fs.readFileSync(loc.cert),
        key: fs.readFileSync(loc.key),
        source: path.dirname(loc.cert),
      };
    }
  }
  return null;
}

const certs = findCerts();

if (certs) {
  console.log(`Using certificates from: ${certs.source}`);
  https.createServer({ cert: certs.cert, key: certs.key }, app).listen(PORT, () => {
    console.log(`Email Analyzer server running on https://localhost:${PORT}`);
    console.log("Ensure you have ANTHROPIC_API_KEY set in your .env file.");
  });
} else {
  console.error("==========================================================");
  console.error("  NO HTTPS CERTIFICATES FOUND");
  console.error("==========================================================");
  console.error("");
  console.error("Outlook Add-ins require HTTPS. Run this to generate certs:");
  console.error("");
  console.error("  npm run setup-certs");
  console.error("");
  console.error("This uses Microsoft's office-addin-dev-certs tool to create");
  console.error("a trusted certificate for localhost on your machine.");
  console.error("You may see a Windows security prompt — click Yes to trust it.");
  console.error("");
  console.error("Then run 'npm start' again.");
  console.error("==========================================================");
  process.exit(1);
}
