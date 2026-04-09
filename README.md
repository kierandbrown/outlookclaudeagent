# Outlook Email Analyzer

An Outlook Add-in that uses Claude AI to analyze your composed emails before sending. It identifies your email's intent, checks that you've addressed everything in the conversation thread, and suggests improvements to prevent unnecessary back-and-forth.

## What It Does

When you compose or reply to an email in Outlook, click **"Analyze Email"** in the ribbon to get:

- **Intent Detection** — Identifies what you're trying to accomplish with the email
- **Completeness Check** — Verifies your reply addresses every question/request from the thread
- **Clarity & Tone Analysis** — Flags ambiguous language, vague phrases, or tone issues
- **Back-and-Forth Prevention** — Catches missing info, unstated assumptions, and unclear next steps that would trigger follow-up emails
- **Suggested Revision** — A rewritten version that addresses all identified issues
- **Overall Score** — 1-10 rating of how effective your email is

## Project Structure

```
├── manifest.xml              # Outlook Add-in manifest
├── server/
│   ├── index.js              # Express server (serves UI + API)
│   └── analyze.js            # Claude API integration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # Add-in taskpane UI
│   │   ├── taskpane.css      # Styles
│   │   └── taskpane.js       # Client-side logic (Office.js)
│   ├── commands/
│   │   ├── commands.html     # Ribbon command support
│   │   └── commands.js       # Command handlers
│   └── assets/               # Icons
├── package.json
├── .env.example              # Environment variables template
└── .gitignore
```

## Prerequisites

- **Node.js** 18 or later
- **Anthropic API key** — Get one at [console.anthropic.com](https://console.anthropic.com)
- **Microsoft 365** account (for Outlook Web or Desktop)

## Setup

1. **Clone and install**
   ```bash
   git clone https://github.com/kierandbrown/outlookclaudeagent.git
   cd outlookclaudeagent
   npm install
   ```

2. **Configure your API key**
   ```bash
   cp .env.example .env
   # Edit .env and add your Anthropic API key
   ```

3. **Start the server**
   ```bash
   npm start
   ```
   The server runs on `https://localhost:3000` by default.

4. **Set up HTTPS for local development**

   Outlook Add-ins require HTTPS. For local development, you can use a tool like [mkcert](https://github.com/FiloSottile/mkcert) to create a trusted local certificate, or use [devtunnel](https://learn.microsoft.com/en-us/azure/developer/dev-tunnels/) to expose your local server.

   To add HTTPS with mkcert:
   ```bash
   mkcert -install
   mkcert localhost
   ```
   Then update `server/index.js` to use the generated cert files with `https.createServer()`.

5. **Sideload the add-in in Outlook**

   **Outlook on the Web:**
   - Go to Outlook Web → Settings (gear icon) → "Manage add-ins"
   - Click "My add-ins" → "Add a custom add-in" → "Add from file"
   - Upload `manifest.xml`

   **Outlook Desktop (Windows):**
   - File → Manage Add-ins (opens browser)
   - Same steps as web above

   **Outlook Desktop (Mac):**
   - Go to the "..." menu → "Get Add-ins"
   - "My Add-ins" → "Custom add-ins" → "Add from file"
   - Upload `manifest.xml`

## Usage

1. Open Outlook and compose a new email or reply to an existing thread
2. Click the **"Analyze Email"** button in the ribbon
3. The taskpane opens — click **"Analyze My Email"**
4. Review the analysis results:
   - Check if your intent comes through clearly
   - See if you've missed addressing any questions from the thread
   - Review clarity and tone suggestions
   - See risks for unnecessary back-and-forth
5. If you like the suggested revision, click **"Apply Suggestion"** to replace your draft

## Configuration

| Environment Variable | Description | Default |
|---|---|---|
| `ANTHROPIC_API_KEY` | Your Anthropic API key | (required) |
| `PORT` | Server port | 3000 |

## How the Analysis Works

The add-in extracts your composed email text and the conversation thread (previous messages in the chain). It sends this to Claude with a specialized prompt that:

1. Parses the thread to identify all questions, requests, and action items
2. Checks your draft against each one to verify completeness
3. Analyzes your language for ambiguity and unclear phrasing
4. Identifies anything that would likely trigger another round of emails
5. Produces a scored assessment with specific, actionable feedback

The analysis runs server-side so your API key stays secure.

## Development

```bash
# Start the dev server
npm run dev

# The server serves both the API and the static add-in files
# Changes to src/ files are served immediately (no build step needed)
```

## License

MIT
