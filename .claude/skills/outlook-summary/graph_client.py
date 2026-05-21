#!/usr/bin/env python3
"""Microsoft Graph helper for the outlook-summary skill (cloud/phone runs).

Credentials come from environment variables (set as secrets in the Claude Code
web environment), NOT from a local file:

  MS_GRAPH_REFRESH_TOKEN  (required) - OAuth refresh token
  MS_GRAPH_CLIENT_ID      (optional) - defaults to the known app registration
  MS_GRAPH_TENANT_ID      (optional) - defaults to the known tenant

Commands:
  fetch                       -> prints JSON {me, inbox[], sent[]} to stdout
  send --subject "..."        -> reads the email body from stdin, sends to self

The refresh response returns a NEW refresh token. Microsoft eventually
invalidates the old one, so the new value is printed to stderr as
NEW_REFRESH_TOKEN=... — update the environment secret with it when prompted.
"""
import json
import os
import sys
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

DEFAULT_CLIENT_ID = "020d9b9d-7438-4783-94c0-8c8430cb1539"
DEFAULT_TENANT_ID = "497f7e96-5fa2-459b-a17a-f33cc41b2131"
GRAPH = "https://graph.microsoft.com/v1.0"


def _post_form(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode())


def _graph(method, path, token, body=None):
    url = path if path.startswith("http") else GRAPH + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", "Bearer " + token)
    if data is not None:
        req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, timeout=60) as r:
        raw = r.read().decode()
        return json.loads(raw) if raw else {}


def get_access_token():
    refresh = os.environ.get("MS_GRAPH_REFRESH_TOKEN")
    if not refresh:
        sys.exit("ERROR: MS_GRAPH_REFRESH_TOKEN is not set. Add it as a "
                 "secret in the Claude Code web environment settings.")
    client_id = os.environ.get("MS_GRAPH_CLIENT_ID", DEFAULT_CLIENT_ID)
    tenant = os.environ.get("MS_GRAPH_TENANT_ID", DEFAULT_TENANT_ID)
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    # No explicit scope: reuse the scopes originally consented (must include
    # Mail.Read and, for the send step, Mail.Send).
    tok = _post_form(url, {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh,
    })
    if "refresh_token" in tok and tok["refresh_token"] != refresh:
        print("NEW_REFRESH_TOKEN=" + tok["refresh_token"], file=sys.stderr)
    return tok["access_token"]


def _collect(token, path):
    items, url = [], path
    while url:
        page = _graph("GET", url, token)
        items.extend(page.get("value", []))
        url = page.get("@odata.nextLink")
    return items


def cmd_fetch(token):
    since = (datetime.now(timezone.utc) - timedelta(days=7)).strftime(
        "%Y-%m-%dT%H:%M:%SZ")
    me = _graph("GET", "/me?$select=mail,userPrincipalName,displayName", token)
    inbox = _collect(token,
        "/me/messages?$filter=receivedDateTime ge " + since +
        "&$top=100&$select=subject,from,receivedDateTime,bodyPreview"
        "&$orderby=receivedDateTime desc")
    sent = _collect(token,
        "/me/mailFolders/sentitems/messages?$filter=sentDateTime ge " + since +
        "&$top=50&$select=subject,toRecipients,sentDateTime,bodyPreview"
        "&$orderby=sentDateTime desc")
    json.dump({"me": me, "inbox": inbox, "sent": sent}, sys.stdout)


def cmd_send(token, subject):
    content = sys.stdin.read()
    me = _graph("GET", "/me?$select=mail,userPrincipalName", token)
    address = me.get("mail") or me.get("userPrincipalName")
    _graph("POST", "/me/sendMail", token, {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": content},
            "toRecipients": [{"emailAddress": {"address": address}}],
        },
        "saveToSentItems": True,
    })
    print("Sent to " + address, file=sys.stderr)


def cmd_scopes(token):
    import base64
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    claims = json.loads(base64.urlsafe_b64decode(payload))
    scopes = claims.get("scp", "").split()
    print("Granted scopes: " + (" ".join(scopes) or "(none)"))
    for need in ("Mail.Read", "Mail.Send"):
        ok = need in scopes or "Mail.ReadWrite" in scopes
        print(("  OK   " if ok else "  MISSING ") + need)


def main():
    args = sys.argv[1:]
    if not args:
        sys.exit("usage: graph_client.py {fetch|send --subject ...|scopes}")
    token = get_access_token()
    cmd = args[0]
    if cmd == "fetch":
        cmd_fetch(token)
    elif cmd == "scopes":
        cmd_scopes(token)
    elif cmd == "send":
        subject = "Weekly Email Summary"
        if "--subject" in args:
            subject = args[args.index("--subject") + 1]
        cmd_send(token, subject)
    else:
        sys.exit("unknown command: " + cmd)


if __name__ == "__main__":
    main()
