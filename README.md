# Voice Nav

Hands-free email triage for your commute. Uses OpenAI's Realtime API for voice interaction and Microsoft Graph API to read, reply, skip, and archive Outlook emails — all by voice.

Open the app on your phone, tap **Start**, and your AI assistant walks through your unread emails one by one. No screen needed.

## How it works

- **Voice in/out**: Your browser connects directly to OpenAI's Realtime API via WebRTC for low-latency speech-to-speech conversation
- **Email operations**: The AI calls tools (get emails, reply, archive, skip) that hit your server's Microsoft Graph API integration
- **Driving-friendly UI**: Dark theme, single big button, minimal visual elements

## Prerequisites

- Node.js 18+
- An [OpenAI API key](https://platform.openai.com/api-keys) with access to the Realtime API
- A [Microsoft Entra ID (Azure AD) app registration](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps) with Mail permissions

## Setup

### 1. Clone and install

```bash
git clone https://github.com/tomblomfield/voice-email.git
cd voice-email
npm install
```

### 2. Set up Microsoft Entra ID (Azure AD) credentials

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps)
2. Click **New registration**, set the name, and under **Supported account types** choose "Accounts in any organizational directory and personal Microsoft accounts"
3. Under **Redirect URI**, select **Web** and add `http://localhost:3000/api/auth/callback`
4. After creation, go to **Certificates & secrets** → **New client secret** and copy the value
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** and add: `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `User.Read`, `offline_access`

### 3. Configure environment variables

```bash
cp .env.example .env
```

Edit `.env` and fill in your keys:

| Variable | Description |
|---|---|
| `OPENAI_API_KEY` | Your OpenAI API key |
| `MS_CLIENT_ID` | Application (client) ID from Azure portal |
| `MS_CLIENT_SECRET` | Client secret value from Azure portal |
| `MS_REDIRECT_URI` | `http://localhost:3000/api/auth/callback` (default) |
| `SESSION_SECRET` | Encryption key for cookies (`openssl rand -base64 32`) |

### 4. Run

```bash
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

## Usage

1. **Connect Outlook** — On first visit, you'll be prompted to sign in with your Microsoft account and grant email permissions.
2. **Tap Start** — The AI greets you and tells you how many unread emails you have.
3. **Listen and respond** — For each email, the AI reads a short summary (sender, subject, key points) and asks what you'd like to do:
   - **Reply** — Tell the AI what to say, it drafts and reads it back for confirmation before sending
   - **Skip** — Marks as read, moves to next
   - **Archive** — Removes from inbox, moves to next
4. **Tap Stop** when you're done.

## Tech stack

- [Next.js](https://nextjs.org/) — App framework
- [OpenAI Realtime API](https://platform.openai.com/docs/guides/realtime) — Speech-to-speech via WebRTC
- [@openai/agents](https://github.com/openai/openai-agents-js) — Agent + tool framework
- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/api/overview) — Email operations via OAuth2

## License

MIT
