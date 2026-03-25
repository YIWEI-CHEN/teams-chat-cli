# teams-chat-cli

A Python CLI tool to send and read Microsoft Teams channel messages via the Microsoft Graph API using interactive user login (delegated auth).

---

## Prerequisites

- [uv](https://docs.astral.sh/uv/getting-started/installation/) (Python package manager)
- An Azure subscription with permission to register applications in Microsoft Entra ID (Azure AD)
- Microsoft Teams with an existing team and channel

---

## Azure AD App Registration

### 1. Register a new application

1. Sign in to the [Azure Portal](https://portal.azure.com).
2. Navigate to **Microsoft Entra ID** → **App registrations** → **New registration**.
3. Fill in:
   - **Name**: `teams-chat-cli` (or any name you prefer)
   - **Supported account types**: *Accounts in this organizational directory only*
   - **Redirect URI**: select **Public client/native (mobile & desktop)** and enter `http://localhost`
4. Click **Register**.
5. Note the **Application (client) ID** — this is the only credential you need.

> No client secret is required. The CLI opens a browser window for interactive login.

### 2. Enable the public client flow

1. In your app registration go to **Authentication**.
2. Under **Advanced settings**, set **Allow public client flows** to **Yes**.
3. Click **Save**.

### 3. Grant API permissions

1. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**.
2. Search for and add:

| Permission | Purpose |
|---|---|
| `ChannelMessage.Read.All` | Read messages from a channel as the signed-in user |
| `ChannelMessage.Send` | Send messages to a channel as the signed-in user |

3. Click **Grant admin consent for \<your tenant\>** and confirm.

### 4. Find your Team ID and Channel ID

**Using the Teams desktop app:**

1. Right-click the **team name** → **Get link to team**. Copy the `groupId=...` value from the URL → `TEAM_ID`.
2. Right-click the **channel name** → **Get link to channel**. Copy the ID between `/channel/` and the next `/` → `CHANNEL_ID`.

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

uv sync
```

---

## Configuration

```bash
cp .env.example .env
```

Edit `.env`:

```ini
CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

# Optional — omit to allow any work/school account to log in
# TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

TEAM_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
CHANNEL_ID=19:xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx@thread.tacv2
```

> `.env` is listed in `.gitignore` — never commit it.

---

## Usage

### First run — login

On the first run the browser opens automatically for Microsoft login. After you sign in, the token is cached at `~/.teams_cli_cache.json` and reused on subsequent runs (no repeated login needed).

### Read messages

```bash
# Read the last 10 messages (default)
uv run teams read

# Read the last 25 messages
uv run teams read --limit 25

# Output raw JSON from the Graph API
uv run teams read --json
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
uv run teams send "Hello from the CLI!"
```

```
Message sent successfully (id=1730456123456, time=2024-11-01 09:22 UTC)
```

### Log out

Clears the cached token. You will be prompted to log in again on the next run.

```bash
uv run teams logout
```

### Help

```bash
uv run teams --help
uv run teams read --help
uv run teams send --help
```

---

## Required Graph API Permissions Summary

| Permission | Type | Description |
|---|---|---|
| `ChannelMessage.Read.All` | Delegated | Read channel messages as the signed-in user |
| `ChannelMessage.Send` | Delegated | Send channel messages as the signed-in user |

---

## How It Works

1. **Authentication**: Uses [MSAL for Python](https://github.com/AzureAD/microsoft-authentication-library-for-python) `PublicClientApplication` with the interactive browser flow. The token is cached at `~/.teams_cli_cache.json` (mode `600`) and refreshed silently on subsequent runs.
2. **Read**: Calls `GET /teams/{teamId}/channels/{channelId}/messages`.
3. **Send**: Calls `POST /teams/{teamId}/channels/{channelId}/messages`.

---

## Troubleshooting

| Error | Likely cause |
|---|---|
| `Missing required environment variables` | `.env` file is missing or incomplete |
| `AADSTS700016` | `CLIENT_ID` is wrong or app doesn't exist in the tenant |
| `AADSTS50011` | Redirect URI `http://localhost` not added in app registration |
| `AADSTS65001` | Admin consent not granted for the required permissions |
| `403 Forbidden` | Permissions not granted or user lacks access to the team/channel |
| `404 Not Found` | `TEAM_ID` or `CHANNEL_ID` is wrong |
