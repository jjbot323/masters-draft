---
name: outlook-summary
description: Read Josh's Outlook emails from the last 7 days via the Microsoft Graph API and produce a weekly summary of active situations and action items, grouped by portfolio company then by category. Use when asked for an email summary, weekly inbox digest, or "what happened in my email this week".
---

# Weekly Outlook Email Summary

Read Josh's Outlook emails from the last 7 days using the Microsoft Graph API and produce a summary of all active situations and action items.

## Instructions

1. **Get access token** from `C:\Users\JoshJohn\.claude\ms_graph_token.json`. If the token is expired, refresh it using the `refresh_token` with:
   - Token endpoint: `https://login.microsoftonline.com/497f7e96-5fa2-459b-a17a-f33cc41b2131/oauth2/v2.0/token`
   - Client ID: `020d9b9d-7438-4783-94c0-8c8430cb1539`
   - Grant_type: `refresh_token`
   - Save the new token back to the same file.

2. **Fetch inbox emails** from the last 7 days (top 100, ordered by date desc):
   ```
   GET https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge {7_DAYS_AGO}&$top=100&$select=subject,from,receivedDateTime,bodyPreview&$orderby=receivedDateTime desc
   ```

3. **Fetch sent items** from the last 7 days (top 50):
   ```
   GET https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$filter=sentDateTime ge {7_DAYS_AGO}&$top=50&$select=subject,toRecipients,sentDateTime,bodyPreview&$orderby=sentDateTime desc
   ```

4. **For key work threads**, search and read the full email body to get complete context:
   ```
   GET https://graph.microsoft.com/v1.0/me/messages?$search="keyword"&$top=3&$select=subject,from,receivedDateTime,body
   ```

5. **Filter out** newsletters, marketing emails, spam, automated notifications (NetSuite, Dakota News, SecondaryLink, DFW Business Expo, etc.)

6. **Summarize** grouped by portfolio company, then by category:
   - **Primary grouping: Portfolio Company** — assign each email thread to its portfolio company where applicable:
     - **Hatch Stamping** (Hatch)
     - **Silo Mills** (SMI, HLI, Joshua Land Farms, Silo)
     - **FDF Energy Services** (FDF)
     - **Keywell Metals** (Keywell)
     - **Cummings Resources** (Cummings)
     - **Hills Holdings** (Hills, Hills HoldCo, HMC, Hills Machinery, RJV)
     - **Groff Tractor** (GTMA, GTE, Groff)
     - **Interstate Waste Services** (Interstate, IWS)
     - **Derby Fabricating** (Derby, DFS)
   - **Cross-portfolio items** — emails that mention multiple portcos but are overarching in nature (e.g., annual meeting slides, portfolio-wide valuations, portfolio performance reports, board decks, fund-level reporting) should NOT be split across individual portco sections. Instead, place them in the relevant category section below and note which portcos are referenced.
   - **Secondary grouping: Category** — emails not tied to a single portco (including cross-portfolio items) should be grouped by category (e.g., Business Development, Reporting, Valuations, Sourcing, Fund Operations, Legal, HR/Admin, etc.)
   - Within each portco or category section: brief status, what happened this week, and action items
   - Mark action items as completed `[x]` or outstanding `[ ]`
   - Highlight anything urgent or time-sensitive
   - Only include portco sections that have activity in the period — skip empty ones

7. **Write the summary** to `C:\Users\JoshJohn\OneDrive\Desktop\Weekly_Email_Summary.md`

8. **Also display** the summary directly in the conversation.

Use Python at `/c/Users/JoshJohn/AppData/Local/Programs/Python/Python312/python.exe` for JSON parsing. Set `PYTHONIOENCODING=utf-8` to avoid encoding errors.
