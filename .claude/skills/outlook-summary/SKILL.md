---
name: outlook-summary
description: Read Josh's Outlook emails from the last 7 days via the Microsoft Graph API and produce a weekly summary of active situations and action items, grouped by portfolio company then by category. Emails the summary to Josh and shows it in chat. Runs in the cloud so it works from the phone with the laptop off. Use when asked for an email summary, weekly inbox digest, or "what happened in my email this week".
---

# Weekly Outlook Email Summary (cloud / phone)

Read Josh's Outlook emails from the last 7 days and produce a summary of all
active situations and action items. This version runs entirely inside the
Claude Code cloud container, so it works from the phone with no laptop on.

## One-time setup (done by Josh, not Claude)

In the Claude Code **web environment settings**, add these as secrets/env vars:

- `MS_GRAPH_REFRESH_TOKEN` — the OAuth refresh token (required)
- `MS_GRAPH_CLIENT_ID` — optional; defaults to the known app registration
- `MS_GRAPH_TENANT_ID` — optional; defaults to the known tenant

The app registration must have consented **Mail.Read** (to read) and
**Mail.Send** (to email the summary back). The environment's network policy
must allow outbound HTTPS to `login.microsoftonline.com` and
`graph.microsoft.com`.

## Steps

1. **Fetch emails.** Run the helper from the skill directory:
   ```
   python3 .claude/skills/outlook-summary/graph_client.py fetch > /tmp/emails.json
   ```
   It refreshes the access token from `MS_GRAPH_REFRESH_TOKEN`, then pulls the
   last 7 days of inbox (top 100) and sent items (top 50). The JSON contains
   `me`, `inbox[]`, and `sent[]` with subject, from/recipients, dates, and
   `bodyPreview`.

2. **Watch for a rotated token.** If the helper prints `NEW_REFRESH_TOKEN=...`
   on stderr, tell Josh at the end of the run to paste that value into the
   `MS_GRAPH_REFRESH_TOKEN` secret, so the next phone run keeps working.

3. **Read `/tmp/emails.json`** and **filter out** newsletters, marketing,
   spam, and automated notifications (NetSuite, Dakota News, SecondaryLink,
   DFW Business Expo, etc.). `bodyPreview` is usually enough context; only
   request a full body for a thread if genuinely needed.

4. **Summarize**, grouped by portfolio company, then by category:
   - **Primary grouping: Portfolio Company** — assign each email thread to its
     portfolio company where applicable:
     - **Hatch Stamping** (Hatch)
     - **Silo Mills** (SMI, HLI, Joshua Land Farms, Silo)
     - **FDF Energy Services** (FDF)
     - **Keywell Metals** (Keywell)
     - **Cummings Resources** (Cummings)
     - **Hills Holdings** (Hills, Hills HoldCo, HMC, Hills Machinery, RJV)
     - **Groff Tractor** (GTMA, GTE, Groff)
     - **Interstate Waste Services** (Interstate, IWS)
     - **Derby Fabricating** (Derby, DFS)
   - **Cross-portfolio items** — emails that mention multiple portcos but are
     overarching (annual meeting slides, portfolio-wide valuations, portfolio
     performance reports, board decks, fund-level reporting) should NOT be
     split across individual portco sections. Place them in the relevant
     category section below and note which portcos are referenced.
   - **Secondary grouping: Category** — emails not tied to a single portco
     (including cross-portfolio items) grouped by category (Business
     Development, Reporting, Valuations, Sourcing, Fund Operations, Legal,
     HR/Admin, etc.)
   - Within each portco or category section: brief status, what happened this
     week, and action items.
   - Mark action items as completed `[x]` or outstanding `[ ]`.
   - Highlight anything urgent or time-sensitive.
   - Only include portco sections that have activity in the period — skip
     empty ones.

5. **Email the summary to Josh.** Pipe the finished markdown into the helper:
   ```
   python3 .claude/skills/outlook-summary/graph_client.py send \
     --subject "Weekly Email Summary — <date range>" < /tmp/summary.md
   ```
   It sends to Josh's own mailbox and saves to Sent Items.

6. **Also display** the summary directly in the conversation.
