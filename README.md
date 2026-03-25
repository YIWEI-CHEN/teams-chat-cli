# teams-chat-cli

A Python CLI tool to send and read Microsoft Teams channel messages via the Microsoft Graph API using app-only (daemon) authentication.

---

## Prerequisites

- Python 3.8+
- An Azure subscription with permission to register applications in Microsoft Entra ID (Azure AD)
- Microsoft Teams with an existing team and channel

---

## Azure AD App Registration

Follow these steps to create the app registration that allows the CLI to authenticate and call the Graph API.

### 1. Register a new application

1. Sign in to the [Azure Portal](https://portal.azure.com).
2. Navigate to **Microsoft Entra ID** → **App registrations** → **New registration**.
3. Fill in:
   - **Name**: `teams-chat-cli` (or any name you prefer)
   - **Supported account types**: *Accounts in this organizational directory only*
   - **Redirect URI**: leave blank (not needed for client credentials flow)
4. Click **Register**.
5. Note the **Application (client) ID** and **Directory (tenant) ID** — you'll need these later.

### 2. Create a client secret

1. In your app registration, go to **Certificates & secrets** → **New client secret**.
2. Add a description (e.g., `cli-secret`) and choose an expiry.
3. Click **Add** and immediately copy the **Value** — it is only shown once.

### 3. Grant API permissions

The CLI uses **application permissions** (no user sign-in required).

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**.
2. Search for and add the following permissions:

| Permission | Purpose |
|---|---|
| `ChannelMessage.Read.All` | Read messages from any channel |
| `ChannelMessage.Send` | Send messages to any channel |

3. Click **Grant admin consent for \<your tenant\>** and confirm.
   - Admin consent is required for application permissions.

> **Note:** `ChannelMessage.Send` as an application permission is in [limited access](https://learn.microsoft.com/en-us/graph/teams-protected-apis). You may need to request access via the Graph API protected APIs form, or use a delegated flow for sending in production.

### 4. Find your Team ID and Channel ID

**Using the Teams desktop app:**

1. Open Teams, right-click the team name → **Get link to team**. The URL contains the `groupId` query parameter — that is your **Team ID**.
2. Right-click the channel name → **Get link to channel**. The URL contains the channel ID after `/channel/`.

**Using Graph Explorer:**

```
GET https://graph.microsoft.com/v1.0/me/joinedTeams
GET https://graph.microsoft.com/v1.0/teams/{team-id}/channels
```

---

## Installation

```bash
git clone https://github.com/yiwei-chen/teams-chat-cli.git
cd teams-chat-cli

python -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate

pip install -r requirements.txt
```

---

## Configuration

Copy the example env file and fill in your values:

```bash
cp .env.example .env
```

Edit `.env`:

```ini
TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CLIENT_SECRET=your-client-secret-value
TEAM_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CHANNEL_ID=19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.tacv2
```

> `.env` is listed in `.gitignore` — never commit it.

---

## Usage

### Read messages

```bash
# Read the last 10 messages (default)
python teams_cli.py read

# Read the last 25 messages
python teams_cli.py read --limit 25

# Output raw JSON from the Graph API
python teams_cli.py read --json
```

Example output:

```
============================================================
  Last 3 message(s) in channel
============================================================

[2024-11-01 09:15 UTC] Alice Smith
  Good morning everyone!

[2024-11-01 09:20 UTC] Bob Jones
  Morning! Ready for the standup?

[2024-11-01 09:21 UTC] Alice Smith
  Yes, starting in 5 minutes.
```

### Send a message

```bash
python teams_cli.py send "Hello from the CLI!"
```

Example output:

```
Message sent successfully (id=1730456123456, time=2024-11-01 09:22 UTC)
```

### Help

```bash
python teams_cli.py --help
python teams_cli.py read --help
python teams_cli.py send --help
```

---

## Required Graph API Permissions Summary

| Permission | Type | Description |
|---|---|---|
| `ChannelMessage.Read.All` | Application | Read all messages in all Teams channels |
| `ChannelMessage.Send` | Application | Send messages to any Teams channel |

---

## How It Works

1. **Authentication**: Uses [MSAL for Python](https://github.com/AzureAD/microsoft-authentication-library-for-python) with the client credentials (app-only) flow to obtain a bearer token from Microsoft Entra ID.
2. **Read**: Calls `GET /teams/{teamId}/channels/{channelId}/messages` and displays the results.
3. **Send**: Calls `POST /teams/{teamId}/channels/{channelId}/messages` with a JSON body.

---

## Troubleshooting

| Error | Likely cause |
|---|---|
| `Missing required environment variables` | `.env` file is missing or incomplete |
| `AADSTS700016` | `CLIENT_ID` is wrong or app doesn't exist in the tenant |
| `AADSTS7000215` | `CLIENT_SECRET` is incorrect or expired |
| `403 Forbidden` | API permissions not granted or admin consent not given |
| `404 Not Found` | `TEAM_ID` or `CHANNEL_ID` is wrong |
