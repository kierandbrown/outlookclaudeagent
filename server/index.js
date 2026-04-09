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

// Health check
app.get("/api/health", (_req, res) => {
  res.json({ status: "ok" });
});

// Main analysis endpoint
app.post("/api/analyze", async (req, res) => {
  const { composedEmail, conversationThread, subject, recipients } = req.body;

  if (!composedEmail || composedEmail.trim().length === 0) {
    return res.status(400).json({ error: "No email content provided." });
  }

  try {
    const analysis = await analyzeEmail({
      composedEmail,
      conversationThread: conversationThread || "",
      subject: subject || "(No subject)",
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
