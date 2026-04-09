require("dotenv").config();
const express = require("express");
const http = require("http");
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

// Serve the standalone web app at the root
app.use(express.static(path.join(__dirname, "..", "src", "webapp")));

// Serve Outlook add-in taskpane at /taskpane.html
app.use("/taskpane.html", express.static(path.join(__dirname, "..", "src", "taskpane", "taskpane.html")));
app.use("/taskpane.css", express.static(path.join(__dirname, "..", "src", "taskpane", "taskpane.css")));
app.use("/taskpane.js", express.static(path.join(__dirname, "..", "src", "taskpane", "taskpane.js")));
app.use("/assets", express.static(path.join(__dirname, "..", "src", "assets")));
app.use(
  "/commands.html",
  express.static(path.join(__dirname, "..", "src", "commands", "commands.html"))
);
app.use(
  "/commands.js",
  express.static(path.join(__dirname, "..", "src", "commands", "commands.js"))
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

// --- Server startup ---
// Try HTTPS first (needed for Outlook Add-in), fall back to HTTP (works for web app)
function findCerts() {
  const locations = [
    {
      cert: path.join(__dirname, "..", "certs", "localhost.crt"),
      key: path.join(__dirname, "..", "certs", "localhost.key"),
    },
    {
      cert: path.join(__dirname, "..", "certs", "localhost.pem"),
      key: path.join(__dirname, "..", "certs", "localhost-key.pem"),
    },
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
    console.log(`Server running on https://localhost:${PORT}`);
    console.log("Both the web app and Outlook add-in are available.");
  });
} else {
  // No certs — run HTTP. Web app works fine, Outlook add-in won't.
  http.createServer(app).listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    console.log("");
    console.log("Open http://localhost:${PORT} in your browser to use the web app.");
    console.log("");
    console.log("NOTE: The Outlook Add-in requires HTTPS. To enable it later, run:");
    console.log("  npm run setup-certs");
  });
}
