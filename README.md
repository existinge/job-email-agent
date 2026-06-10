# Job Hunt Email Agent

A local Python tool that connects to your Outlook inbox via Microsoft Graph, uses an LLM (via OpenRouter) to classify job application emails, and generates an Excel report with follow-up recommendations.

Built because manually triaging hundreds of application emails — receipts, rejections, interviews, spam — was costing real cognitive overhead. This automates the categorization and surfaces only what actually needs attention.

---

## What It Does

- Connects to Outlook via **Microsoft Graph API** (Azure Entra ID OAuth)
- Scans your inbox for job-related emails over a configurable time window
- Classifies each email into one of:
  - `Rejection`
  - `Application Receipt`
  - `Needs Action` (interview invites, assessments, scheduling requests)
  - `Follow Up` (applications due for a check-in)
  - `Job Board Spam`
- Uses an **LLM agent via OpenRouter** for classification, with rule-based fallback if AI is unavailable
- Outputs a **Excel report** (`output/job_hunt_email_report.xlsx`) with classification, confidence score, and follow-up status
- Optionally **auto-moves emails** into Outlook folders (rejections, spam, follow-up) above a confidence threshold
- Includes a **human-approval follow-up system** — drafts follow-up emails, but only sends them after you manually mark `Approved To Send: YES` in the report

---

## Safety Design

- Never deletes emails automatically
- Follow-ups require explicit human approval before sending
- Skips no-reply and ATS domains when drafting follow-ups
- Tracks processed messages to avoid re-classifying on re-runs
- Minimum confidence threshold (configurable) before any auto-move happens
- Fallback to rule-based classification if the AI call fails

---

## Stack

- Python 3.x
- Microsoft Graph API + Azure Entra ID (MSAL OAuth)
- OpenRouter API (model-agnostic; defaults to `openrouter/auto`)
- `config.yaml` for all classification rules, thresholds, and behavior
- `.env` for secrets (never committed)

---

## Setup

### 1. Clone and install dependencies

```bash
git clone https://github.com/YOUR_USERNAME/job-email-agent.git
cd job-email-agent
pip install -r requirements.txt
```

### 2. Register an Azure App

1. Go to [portal.azure.com](https://portal.azure.com) → Azure Active Directory → App Registrations → New Registration
2. Set redirect URI to `http://localhost` (type: Public client/native)
3. Under API Permissions, add Microsoft Graph delegated permissions: `User.Read`, `Mail.ReadWrite`, `Mail.Send`
4. Copy the **Client ID**

### 3. Configure your `.env`

```bash
cp .env.template .env
```

Fill in:
```
MS_CLIENT_ID=your-azure-app-client-id
OPENROUTER_API_KEY=your-openrouter-api-key
OPENROUTER_MODEL=openrouter/auto
```

### 4. Configure `config.yaml`

Edit `config.yaml` to set your name, signature, scan window, and classification behavior. All thresholds and phrase lists are configurable here — no code changes needed for tuning.

---

## Usage

```bash
# Check which messages have already been processed
py job_email_agent.py processed-status

# Safe test run (14 days, 30 messages)
py job_email_agent.py review --days 14 --max 30

# Full scan
py job_email_agent.py review --source all --days 60 --max 250

# Run without AI (rules only)
py job_email_agent.py review --days 30 --max 100 --no-ai

# Draft follow-ups for approved rows in the report
py job_email_agent.py draft-approved --report output/job_hunt_email_report.xlsx

# Send approved follow-ups (requires --yes-really flag)
py job_email_agent.py send-approved --report output/job_hunt_email_report.xlsx --yes-really
```

---

## Output

The Excel report includes:
- Email subject, sender, date
- Classification category and confidence score
- Follow-up eligibility and due date
- `Approved To Send` column (manually set to `YES` to enable sending)

---

## Configuration

All behavior is controlled via `config.yaml`:

| Section | What It Controls |
|---|---|
| `scan` | Days back, max messages, inbox source |
| `ai` | Model, temperature, confidence threshold, fallback behavior |
| `automation` | Auto-move toggles, minimum confidence to move |
| `follow_up` | Wait days per email type before follow-up is due |
| `classification` | Rejection phrases, action phrases, receipt phrases, spam domains |
| `folders` | Outlook folder names to move emails into |

---

## Notes

- Designed for personal use on Windows (`.bat` scripts included for convenience)
- OAuth token is cached locally after first login — subsequent runs don't require re-auth
- Processed message tracking persists in `processed_messages.json`
- The agent sends email body text to OpenRouter for classification — disable with `send_email_content_to_openrouter: false` if preferred

---

## License

MIT
