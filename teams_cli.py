#!/usr/bin/env python3
"""Microsoft Teams CLI - Send and read Teams channel messages via Microsoft Graph API."""

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import msal
import requests
from dotenv import load_dotenv
import os

# Delegated permission scopes
SCOPES = [
    "https://graph.microsoft.com/ChannelMessage.Read.All",
    "https://graph.microsoft.com/ChannelMessage.Send",
]

# Token cache stored in the user's home directory
TOKEN_CACHE_FILE = Path.home() / ".teams_cli_cache.json"


def load_config():
    """Load configuration from .env file."""
    load_dotenv()
    required = ["CLIENT_ID", "TEAM_ID", "CHANNEL_ID"]
    config = {}
    missing = []
    for key in required:
        value = os.getenv(key)
        if not value:
            missing.append(key)
        config[key] = value
    if missing:
        print(f"Error: Missing required environment variables: {', '.join(missing)}")
        print("Copy .env.example to .env and fill in your values.")
        sys.exit(1)
    # TENANT_ID is optional — defaults to "organizations" (any work/school account)
    config["TENANT_ID"] = os.getenv("TENANT_ID", "organizations")
    return config


def get_access_token(config):
    """Acquire a token interactively (browser popup), using cache when possible."""
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())

    authority = f"https://login.microsoftonline.com/{config['TENANT_ID']}"
    app = msal.PublicClientApplication(
        client_id=config["CLIENT_ID"],
        authority=authority,
        token_cache=cache,
    )

    # Try to use a cached token silently first
    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # No cached token — open browser for interactive login
    if not result:
        print("Opening browser for login...")
        result = app.acquire_token_interactive(scopes=SCOPES)

    # Persist the updated cache
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())
        TOKEN_CACHE_FILE.chmod(0o600)

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        print(f"Error acquiring token: {error}")
        sys.exit(1)

    return result["access_token"]


def graph_request(method, endpoint, token, **kwargs):
    """Make an authenticated Microsoft Graph API request."""
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    url = f"https://graph.microsoft.com/v1.0{endpoint}"
    response = requests.request(method, url, headers=headers, **kwargs)
    if not response.ok:
        try:
            error_detail = response.json().get("error", {})
            msg = error_detail.get("message", response.text)
            code = error_detail.get("code", response.status_code)
        except Exception:
            msg = response.text
            code = response.status_code
        print(f"Graph API error [{code}]: {msg}")
        sys.exit(1)
    return response


def format_timestamp(iso_str):
    """Format an ISO 8601 timestamp to a human-readable string."""
    if not iso_str:
        return "unknown"
    try:
        dt = datetime.fromisoformat(iso_str.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M UTC")
    except ValueError:
        return iso_str


def cmd_read(args, config, token):
    """Read messages from a Teams channel."""
    endpoint = (
        f"/teams/{config['TEAM_ID']}/channels/{config['CHANNEL_ID']}/messages"
        f"?$top={args.limit}&$orderby=createdDateTime desc"
    )
    response = graph_request("GET", endpoint, token)
    messages = response.json().get("value", [])

    if not messages:
        print("No messages found.")
        return

    if args.json:
        print(json.dumps(messages, indent=2))
        return

    print(f"\n{'='*60}")
    print(f"  Last {len(messages)} message(s) in channel")
    print(f"{'='*60}\n")

    for msg in reversed(messages):
        sender = msg.get("from", {}) or {}
        sender_name = (
            (sender.get("user") or {}).get("displayName")
            or (sender.get("application") or {}).get("displayName")
            or "Unknown"
        )
        timestamp = format_timestamp(msg.get("createdDateTime"))
        body = re.sub(r"<[^>]+>", "", (msg.get("body") or {}).get("content", "")).strip()

        print(f"[{timestamp}] {sender_name}")
        print(f"  {body}")
        print()


def cmd_send(args, config, token):
    """Send a message to a Teams channel."""
    endpoint = f"/teams/{config['TEAM_ID']}/channels/{config['CHANNEL_ID']}/messages"
    payload = {"body": {"contentType": "text", "content": args.message}}
    response = graph_request("POST", endpoint, token, json=payload)
    msg = response.json()
    print(f"Message sent successfully (id={msg.get('id', 'unknown')}, time={format_timestamp(msg.get('createdDateTime'))})")


def cmd_logout(args, config, token):
    """Remove the cached token, forcing a fresh login next time."""
    if TOKEN_CACHE_FILE.exists():
        TOKEN_CACHE_FILE.unlink()
        print("Logged out. You will be prompted to log in again on the next run.")
    else:
        print("No cached session found.")


def main():
    parser = argparse.ArgumentParser(
        description="Microsoft Teams CLI — send and read channel messages via Graph API",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Read the last 10 messages:
    uv run teams read

  Read the last 25 messages in JSON format:
    uv run teams read --limit 25 --json

  Send a message:
    uv run teams send "Hello from the CLI!"

  Log out (clear cached token):
    uv run teams logout
        """,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    read_parser = subparsers.add_parser("read", help="Read messages from the channel")
    read_parser.add_argument(
        "--limit", type=int, default=10, metavar="N",
        help="Number of messages to retrieve (default: 10)",
    )
    read_parser.add_argument(
        "--json", action="store_true",
        help="Output raw JSON from the Graph API",
    )

    send_parser = subparsers.add_parser("send", help="Send a message to the channel")
    send_parser.add_argument("message", help="Text of the message to send")

    subparsers.add_parser("logout", help="Clear the cached login token")

    args = parser.parse_args()
    config = load_config()

    if args.command == "logout":
        cmd_logout(args, config, None)
        return

    token = get_access_token(config)

    if args.command == "read":
        cmd_read(args, config, token)
    elif args.command == "send":
        cmd_send(args, config, token)


if __name__ == "__main__":
    main()
