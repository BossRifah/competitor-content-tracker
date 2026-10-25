# Competitor content tracker

A reusable Google Apps Script automation for monitoring competitor blogs and resource pages, summarizing blog content with Groq, and creating Google Sheets reports with optional Slack notifications.

No company-specific competitor list, production API keys, or webhook credentials are included. Configure it for your organization in your own Google Sheet.

## Setup

1. Create a Google Sheet. Add a tab named **Competitors** with these columns in order: **Competitor**, **Blog URL**, **Resource URL**, **RSS URL**. Add a header row; data starts on row 2. Each row needs a competitor name and at least a blog or resource URL. RSS is optional. See `competitors.example.csv` for a placeholder example; replace example URLs before scanning.
2. In the Sheet, open **Extensions → Apps Script** and paste `Code.gs` into the editor. This script must be bound to the spreadsheet.
3. In Apps Script Project Settings, enable the manifest editor and replace `appsscript.json` with the included manifest. Set the project time zone for your team; the example uses UTC.
4. In **Project Settings → Script Properties**, configure the credentials you want to use:

   | Property | Placeholder | Purpose |
   | --- | --- | --- |
   | `GROQ_API_KEY` | `YOUR_GROQ_API_KEY` | Primary Groq key for AI summaries |
   | `GROQ_API_KEY_SECONDARY` | `YOUR_OPTIONAL_SECONDARY_GROQ_API_KEY` | Optional fallback Groq key |
   | `SLACK_WEBHOOK_URL` | `YOUR_SLACK_WEBHOOK_URL` | Optional Slack incoming webhook |

   Set real values only in Script Properties. The source contains harmless placeholders. Without Groq keys, summaries fall back to text excerpts; without a webhook, Slack sending is skipped. Anyone who can edit the Apps Script project can access its properties.
5. Save and reload the spreadsheet. Use **Blog tracker → Run monthly scan now**, then authorize the script. A configured run sends content excerpts to Groq and a report with the spreadsheet link to your configured Slack destination.
6. To automate, select **Blog tracker → Set up monthly trigger**. It runs on the first day of each month during the 9 AM hour in the project time zone. Use **Remove triggers** to disable it.

## Reports and customization

The script creates or replaces the **Monthly report** tab and appends results to **Archive**. Keep unrelated data out of the report tab because each scan clears its contents and formatting. Repeated runs can append duplicate archive entries.

Change the tab names, extraction limits, summary batch size, and retry settings in `CONFIG`. Competitors are read from the spreadsheet rather than hardcoded. The AI prompt is a general content-analysis prompt with no company or industry assumptions. The Groq model is set in `callGroqWithKey`; use a model available to your Groq account.

The scanner tries RSS, JSON-LD, then HTML extraction. Resource titles and descriptions are classified as guides, templates, reports, webinars, or lead magnets. AI summaries apply to blog posts; resource summaries use available feed descriptions or page metadata.

## Date handling and limitations

- On days 1–3, the target is the previous month. On later days, a manual scan targets the current month.
- Up to five undated items per competitor and source are included as possible recent content. They are marked **Unknown**; publication in the target month is not confirmed.
- HTML parsing is heuristic. Some sites, JavaScript-rendered pages, and protected resources will need site-specific extraction work. The script does not execute page JavaScript.
- Large competitor lists, page fetches, and retries can exceed Apps Script execution quotas. Start with a small list.
- Groq requests use up to 1,500 extracted characters per post. Summaries can be incomplete or inaccurate; review before use.
- Slack delivery does not validate HTTP failure responses; a completion message is not proof of delivery. Check the destination.
- Logs may contain competitor URLs, titles, and diagnostic details. Review logs before sharing.

## Credential handling

Never paste real keys into repository files or commit Script Properties exports. Rotate credentials that have been shared or exposed. Public API endpoint URLs and the Atom XML namespace in the source are service addresses, not credentials.

## Verification

JavaScript syntax and local, mocked behavior checks are run before publication. Live Google Sheets, Groq, Slack, and scheduled-trigger behavior require your own account setup and have not been verified here.
