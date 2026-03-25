#!/usr/bin/env python3
"""Microsoft Teams CLI - Send and read Teams channel messages via Microsoft Graph API."""

import argparse
import json
import sys
from datetime import datetime

import msal
import requests
from dotenv import load_dotenv
import os


def load_config():
    """Load configuration from .env file."""
    load_dotenv()
    required = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "TEAM_ID", "CHANNEL_ID"]
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
    return config


def get_access_token(config):
    """Acquire an access token using the client credentials flow."""
    authority = f"https://login.microsoftonline.com/{config['TENANT_ID']}"
    app = msal.ConfidentialClientApplication(
        client_id=config["CLIENT_ID"],
        client_credential=config["CLIENT_SECRET"],
        authority=authority,
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
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
        sender = (
            msg.get("from", {}) or {}
        )
        sender_name = (
            (sender.get("user") or {}).get("displayName")
            or (sender.get("application") or {}).get("displayName")
            or "Unknown"
        )
        timestamp = format_timestamp(msg.get("createdDateTime"))
        body = (msg.get("body") or {}).get("content", "")

        # Strip simple HTML tags for plain-text display
        import re
        body = re.sub(r"<[^>]+>", "", body).strip()

        print(f"[{timestamp}] {sender_name}")
        print(f"  {body}")
        print()


def cmd_send(args, config, token):
    """Send a message to a Teams channel."""
    endpoint = f"/teams/{config['TEAM_ID']}/channels/{config['CHANNEL_ID']}/messages"
    payload = {
        "body": {
            "contentType": "text",
            "content": args.message,
        }
    }
    response = graph_request("POST", endpoint, token, json=payload)
    msg = response.json()
    msg_id = msg.get("id", "unknown")
    timestamp = format_timestamp(msg.get("createdDateTime"))
    print(f"Message sent successfully (id={msg_id}, time={timestamp})")


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
        """,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # read subcommand
    read_parser = subparsers.add_parser("read", help="Read messages from the channel")
    read_parser.add_argument(
        "--limit",
        type=int,
        default=10,
        metavar="N",
        help="Number of messages to retrieve (default: 10)",
    )
    read_parser.add_argument(
        "--json",
        action="store_true",
        help="Output raw JSON from the Graph API",
    )

    # send subcommand
    send_parser = subparsers.add_parser("send", help="Send a message to the channel")
    send_parser.add_argument("message", help="Text of the message to send")

    args = parser.parse_args()

    config = load_config()
    token = get_access_token(config)

    if args.command == "read":
        cmd_read(args, config, token)
    elif args.command == "send":
        cmd_send(args, config, token)


if __name__ == "__main__":
    main()
