# Microsoft Outlook MCP Server

A Model Context Protocol (MCP) server that enables AI assistants to interact with Microsoft Outlook email and calendar through the Microsoft Graph API.

[![Download Latest Release](https://img.shields.io/github/v/release/XenoXilus/outlook-mcp?label=Download&color=blue)](https://github.com/XenoXilus/outlook-mcp/releases/latest)
[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

## Features

- **Email Operations**: Read, search, send, reply to emails and download attachments
- **SharePoint Integration**: Access SharePoint files via sharing links or direct file IDs. Download files shared to you via emails. 
- **Calendar Management**: View and manage calendar events and appointments
- **Office Document Processing**: Parse PDF, Word, PowerPoint, and Excel files with extracted text content
- **Receipt & Invoice Collection**: Discover emailed receipts by sender/subject/date, save them as verified and consistently named PDFs (attachment, billing link, or rendered fallback), and stage a review draft — built for scheduled, non-interactive runs
- **Large File Support**: Automatic handling of files that exceed MCP response size limits

## Quick Start

**Choose your installation method:**

| Method | Best For |
|--------|----------|
| [DXT Extension](#installing-as-dxt-extension) | Claude Desktop users |
| [CLI Configuration](#using-with-cli-tools) | Claude Code, mcp CLI, other MCP clients |

> **Prerequisites**: Before installing, you'll need to [set up an Azure application](#azure-setup-guide) to get your Client ID and Tenant ID.

---

## Installation

### Installing as DXT Extension

For Claude Desktop users, DXT extensions provide the simplest installation experience.

**Option 1: Download Pre-built Extension**
1. Download `outlook-mcp.dxt` from the [Releases page](https://github.com/XenoXilus/outlook-mcp/releases)
2. In Claude Desktop, go to **Settings** → **Extensions**
3. Click **Install from file** and select the `.dxt` file
4. Enter your Azure Client ID, Tenant ID, and optional download directory when prompted

**Option 2: Build from Source**
1. Clone and install dependencies:
   ```bash
   git clone https://github.com/XenoXilus/outlook-mcp.git
   cd outlook-mcp
   npm install
   ```
2. Install the DXT CLI: `npm install -g @anthropic-ai/dxt`
3. Pack the extension:
   ```bash
   dxt pack . outlook-mcp.dxt
   ```
4. Install the generated `.dxt` file in Claude Desktop as above

---

### Using with CLI Tools

For CLI-based MCP clients (Claude Code, mcp CLI, etc.), configure the server directly.

**1. Clone and Install:**
```bash
git clone https://github.com/XenoXilus/outlook-mcp.git
cd outlook-mcp
npm install
```

**2. Configure your MCP client:**

Add the following to your MCP servers configuration (location varies by client):

```json
{
  "outlook-mcp": {
    "command": "node",
    "args": ["/absolute/path/to/outlook-mcp/server/index.js"],
    "env": {
      "AZURE_CLIENT_ID": "your-azure-client-id",
      "AZURE_TENANT_ID": "your-azure-tenant-id",
      "MCP_OUTLOOK_WORK_DIR": "/optional/download/directory"
    }
  }
}
```

**Common config file locations:**
- **Claude Code**: `~/.claude.json` or project-level `.mcp.json`
- **mcp CLI**: `~/.config/mcp/servers.json`

**3. Alternative: Use environment variables**

Instead of specifying `env` in the config, you can export the variables in your shell:

```bash
export AZURE_CLIENT_ID="your-azure-client-id"
export AZURE_TENANT_ID="your-azure-tenant-id"
export MCP_OUTLOOK_WORK_DIR="/optional/download/directory"
```

---

## Azure Setup Guide

To use this MCP server, you need to register an application in Microsoft Azure.

### For Business/Work Accounts (Recommended)

1. Go to the [Azure Portal](https://portal.azure.com/) and search for "App registrations".
2. Click **New registration**.
   - Name: `Outlook MCP` (or similar)
   - Supported account types: **Accounts in this organizational directory only** (Single tenant)
   - Redirect URI: Select **Web** and enter `http://localhost/callback`
3. Click **Register**.
4. Go to **Authentication** in the sidebar.
   - Under "Advanced settings", set **Allow public client flows** to **Yes**.
   - Click **Save**.
5. On the Overview page, copy:
   - **Application (client) ID** → This is your `AZURE_CLIENT_ID`
   - **Directory (tenant) ID** → This is your `AZURE_TENANT_ID`
6. Go to **API permissions** in the sidebar.
   - Click **Add a permission** -> **Microsoft Graph** -> **Delegated permissions**.
   - Add these permissions:
     - `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`
     - `Calendars.Read`, `Calendars.ReadWrite`
     - `Contacts.Read`, `Contacts.ReadWrite`
     - `Tasks.Read`, `Tasks.ReadWrite`
     - `User.Read`, `MailboxSettings.Read`
     - `Files.Read.All`, `Files.ReadWrite.All`
     - `Sites.Read.All`, `Sites.ReadWrite.All`
     - `offline_access`
   - Click **Add permissions**.
   - (Optional) If you are an admin, click **Grant admin consent** to suppress consent prompts for users.

**Note:** No client secret is required (PKCE auth flow).

### For Personal Accounts (outlook.com, hotmail.com)

Personal Microsoft accounts can also register apps in Azure:

1. Sign in to the [Azure Portal](https://portal.azure.com/) with your personal Microsoft account (outlook.com, hotmail.com, etc.).
2. If prompted to create a directory, follow the steps to create a free Azure directory.
3. Follow the same steps as above for Business accounts.
4. When configuring, use **Accounts in any organizational directory and personal Microsoft accounts** for supported account types.

---

## Configuration Reference

### Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AZURE_CLIENT_ID` | Yes | Your Azure AD application client ID |
| `AZURE_TENANT_ID` | Yes | Your Azure AD directory (tenant) ID |
| `MCP_OUTLOOK_WORK_DIR` | No | Directory for saving large files (defaults to system temp) |
| `MCP_OUTLOOK_ALLOWED_WRITE_DIRS` | No | Comma-separated extra directories save tools may write into when given an explicit `destDir`. Permission only — defaults are unchanged |
| `MCP_OUTLOOK_SHARED_MAILBOX` | No | Default mailbox for the per-call 'mailbox' argument; empty = own mailbox. |
| `MCP_OUTLOOK_KNOWN_MAILBOXES` | No | Comma-separated shared mailboxes offered by `outlook_list_shared_mailboxes` discovery |

The desktop extension (DXT) exposes only the mail settings above. The receipt/invoice behaviour below is configured **by the calling process** (e.g. a scheduled routine's MCP server config) via environment variables — it is intentionally not part of the extension settings UI:

| Variable | Description |
|----------|-------------|
| `MCP_OUTLOOK_RECEIPTS_DIR` | Directory where receipt/invoice PDFs are saved (falls back to work dir). Also passable per-call as `destDir`. |
| `RECEIPT_RULES_PATH` | Optional JSON file of site-specific vendor rules (see below). Unset = generic heuristics only. If set but missing/invalid, receipt tools fail fast rather than silently degrade. |
| `BILLING_DOMAIN_ALLOWLIST` | Comma-separated hosts `outlook_fetch_billing_pdf` may contact (default: `pay.stripe.com,invoice.stripe.com,files.stripe.com,m.stripe.network`) |
| `RECEIPT_FILENAME_TEMPLATE` | Receipt naming pattern (default: `{vendor} {DDMmmYY} Invoice.pdf`). Also passable per-vendor as `filenameTemplate`. |
| `MCP_OUTLOOK_AUTH_MODE` | `interactive` (default) or `headless` — headless never opens a browser and fails fast if silent refresh is impossible |
| `MCP_OUTLOOK_REFRESH_TOKEN_PATH` | Directory of the encrypted token store for headless runs (defaults to the built-in store) |
| `MCP_OUTLOOK_CHROME_PATH` | Chrome/Chromium binary for `outlook_render_email_pdf` (auto-detected when unset) |

#### Vendor rules (`RECEIPT_RULES_PATH`)

`outlook_extract_receipt` / `outlook_collect_receipts` work with zero configuration for
receipts issued via payment processors that put the vendor in the subject line
("Your receipt from Acme #1234"). For senders that need explicit mapping, or to
capture product labels, point `RECEIPT_RULES_PATH` at a JSON file
(see [`receipt-rules.example.json`](receipt-rules.example.json)):

- `vendorSenders`: `[{ "pattern": "<case-insensitive regex on the from address>", "vendor": "<name>" }]` — checked before the subject heuristic.
- `productLabels`: `["<case-insensitive regex>", ...]` — first match becomes `productLabel`; without a match it falls back to the line-item descriptions (see [Extracted fields](#extracted-fields)).

### Receipt & Invoice Tools

Five tools support autonomous expense-receipt collection (e.g. a scheduled monthly run):

| Tool | Purpose |
|------|---------|
| `outlook_save_attachment` | Save an attachment's **original bytes** to a chosen path/filename. Auto-selects the `Invoice-*.pdf` when a Stripe receipt attaches both Invoice and Receipt PDFs (`prefer: invoice\|receipt\|first`). Validates `%PDF` magic bytes; returns path + SHA-256 + size. |
| `outlook_fetch_billing_pdf` | Fetch the per-receipt invoice/receipt PDF download link in an email body (fallback when a forward loses its attachment). Only real PDF links are used — logo, image and other asset links are never fetched. HTTPS-only, allowlist-only (redirects included), content-type + magic-byte validated, 25 MB / 30 s bounded. |
| `outlook_extract_receipt` | Compact structured summary of a receipt email (vendor, gross amounts, line items, product label, receipt/invoice/order numbers, reply/forward info, PDF links, attachment ids) — never the 60 KB+ HTML body. See [Extracted fields](#extracted-fields). |
| `outlook_render_email_pdf` | Render the sanitised email HTML to PDF via headless Chrome — audit-trail fallback for receipts with no attachment and no link (e.g. app-store order receipts). Under `onExisting: skip`, re-rendering an unchanged email is skipped rather than saved as `(2)`, `(3)`. |
| `outlook_collect_receipts` | One call per period: discovers each vendor's receipts by sender/subject/product/date across the whole mailbox (Sent Items and Drafts are ignored), collapses forwarded/replied copies of one receipt, saves one PDF per receipt (attachment → link → rendered fallback), and returns a manifest with a per-entry `status` plus `missing[]`, `failed[]`, `summary` and `notes`. Idempotent re-runs via `onExisting: skip\|overwrite\|version`. See [Collect results](#collect-results). |

`outlook_create_draft` additionally accepts `attachmentPaths` (absolute local file paths, ≤ 3 MB each) and returns the draft's `webLink` — it stages the email for review and **never sends**.

#### Collect results

`outlook_collect_receipts` returns `{ period, manifest, missing, failed, unmatched, summary, notes }`. Every manifest entry has a `status`:

| Status | Meaning |
|--------|---------|
| `saved` | A new PDF was written. `method` is `attachment`, `link` or `rendered`; `isFormalInvoice` is `false` for a `rendered` copy of the email. |
| `skipped` | The target file already holds the same document (`onExisting: skip`): identical bytes, or a PDF differing only in its `CreationDate`/`ModDate` stamps (the entry then reports that file's own `sha256`); nothing was written. A file holding a *different* document is never skipped in its place: the receipt is saved as `name (2).pdf` (or the numbered version that already holds it) with `existingDiffers: true`. The same rule applies to `outlook_save_attachment`, `outlook_download_attachment`, `outlook_fetch_billing_pdf` and `outlook_render_email_pdf`. |
| `not_found` | Nothing matched the rule in the period. Only these vendors appear in `missing[]`. |
| `fetch_failed` | A receipt was found but could not be searched for, read or saved. Every such entry is also listed in `failed[]`. |
| `ambiguous` | Several rules claim the message, or its product cannot be determined; `candidates` names the rules and `reason` says why. |
| `duplicate` | A forwarded/replied copy of a receipt already covered in the run; `duplicate_of` names it. |

- `failed[]` has one object per `fetch_failed` entry: `attemptedRoute` (`search`, `message`, `attachment`, `link` or `render`), `attemptedRoutes` (the attachment content routes tried, e.g. `["json", "value"]`), `messageId`, `attachmentId`/`attachmentName` (attachment route) and the verbatim `graphError`. "Nothing to collect" and "found it but could not save it" are therefore never confused.
- When a formal PDF (attachment or link) cannot be saved, the entry stays `fetch_failed` and gets a `renderedFallback` — a rendering of the email saved under its own `(rendered)` name — and the entry itself carries `method: rendered` and `isFormalInvoice: false`, so it is never mistaken for the invoice (`savedPath` is only inside `renderedFallback`).
- `unmatched[]` lists every message that matched a rule's sender/subject but was rejected by the product matchers of every such rule: `{ messageId, date, subject, sender, productLabel, rejectedBy }`. It does not block the run.
- `summary` counts entries per status plus `unmatched`; `notes` carries run-level remarks (e.g. a Sent Items/Drafts folder that could not be resolved, or messages excluded by product matchers).
- **Duplicates**: forwarded and replied copies of one receipt are grouped when they share a receipt, invoice or order number — but never when the group would hold two different numbers of one kind or two different known totals (two charges on one order stay two receipts, even with equal totals when their receipt numbers differ; a receipt number shared by different totals adds a `numberConflict` note to both). Identical PDF bytes (`sha256`) catch the rest after saving, under the same rule (otherwise both are kept and `sameFileAs` names the other). The copy carrying the PDF (preferring an `Invoice-*.pdf`) is kept, and a copy with no total never replaces one carrying the PDF and the total; the others become `duplicate` with `duplicate_of` (e.g. `receiptNumber:…` or `sha256:…`). Messages in Sent Items and Drafts are ignored.
- **File names** come from the template (`{vendor} {DDMmmYY} Invoice.pdf`). When several receipts in one run would get the same name (one vendor, one day), each gets its own receipt/invoice/order number appended (`Acme 10Jun26 Invoice 1001-2002.pdf`), so re-runs map every receipt to the same file; an ordinal `(2)` is used only for receipts with no number, assigned oldest first. A receipt saved under the plain name by an earlier run, before a same-day sibling arrived, keeps that file (reported `skipped`) instead of being saved again under its new name.
- Entries carry `sender`, a `forwarded` flag and, for forwards with a quoted header, `originalSender`.
- **Amounts** are structured and gross: `amounts` is `{ subtotal, tax, credits, total, amountPaid, currency, source, totalLabel }`, and `amount`/`taxAmount` are the total charged and the tax sum. For attachment and link saves the figures are re-read from the saved PDF when its total is labelled (`totalLabel` is `amount paid`, `total` or `amount due`; `amountSource: pdf`). An unlabelled (guessed) PDF figure never replaces the email's total — it is used only when the email has none — and sets `pdfTotalGuessed: true`. A PDF total of zero (e.g. only `Balance due $0.00`) never replaces a positive email total. `amountMismatch` (`{ body, pdf }`) appears when the two disagree.
- A billing link is used only when it is unique to one receipt in the run; a link shared by several receipts is skipped (`linkSkipped`).

#### Product matchers

A vendor rule passed to `outlook_collect_receipts` may also carry `productContains` and/or `productExcludes` (a string or an array of strings). Terms are case-insensitive substrings matched against the receipt's product label and its line-item descriptions, so one sender can be split into several labelled streams:

```json
"vendors": [
  { "vendor": "Acme Hosting", "from": "billing@acme.example", "productContains": ["hosting", "domain"] },
  { "vendor": "Acme Support", "from": "billing@acme.example", "productContains": "support plan" },
  { "vendor": "Acme Other",   "from": "billing@acme.example" }
]
```

- A receipt whose product matches a `productContains` rule belongs to that rule; it beats a rule without product matchers (here, `Acme Other` only receives what the other two reject). `productExcludes` can replace the catch-all: `{ "vendor": "Acme Other", "from": "billing@acme.example", "productExcludes": ["hosting", "domain", "support plan"] }`.
- When the email body names no product (no product label or line items) and a rule depends on one, the product is read from the PDF attachment before deciding (`productSource: pdf`).
- If several rules still claim a receipt, or its product cannot be read from the email or the PDF while a rule depends on it, the entry is `ambiguous` with `candidates` (and `pdfProductError` if the PDF could not be read) — it is never guessed or silently accepted. A receipt that every matching rule rejects is listed in `unmatched[]`, never silently dropped.

#### Extracted fields

Besides vendor, `productLabel` and `attachmentIds`, `outlook_extract_receipt` returns:

- `amounts` — `{ subtotal, tax, credits, total, amountPaid, currency, source, totalLabel }` (`totalLabel` says which row gave the total — `amount paid`, `total` or `amount due` — and is `null` when it had to be guessed). `amount` is the gross total charged (not the pre-tax subtotal) and `taxAmount` the tax sum. Credit and discount lines are summed; currency prefixes such as `CA$`, `A$`/`AU$`, `NZ$`, `US$`, `HK$` and ISO codes are recognised.
- `lineItems[]` — `{ description, quantity, amount }`. `productLabel` falls back to their descriptions when no configured `productLabels` pattern matches.
- `receiptNumber`, `invoiceNumber` (also read from an `Invoice-<CODE>-<NNNN>.pdf` attachment name) and `orderNumber`.
- Reply/forward handling: `normalisedSubject` (`Re:`/`Fw:` stripped), `replyOrForward` and `quotedFrom` (`{ name, address }` of the innermost quoted `From:` header).
- `billingPdfUrl` and `receiptPdfUrl` — per-receipt PDF download links on allowlisted hosts only; click-tracking wrappers are unwrapped, and image or logo links are never returned.

#### Matching and download notes

- A vendor rule's `from` and `subjectContains` combine as **OR** — either signal
  matches a receipt (subjects are compared with `Re:`/`Fw:` prefixes ignored), and
  each manifest entry reports `matchedBy` (`from`, `from-normalised`, `subject`,
  and/or `product:…` signals).
- Plus-addressed senders (`invoice+statements+acct_...@stripe.com`) match both
  exactly and via a plus-stripped fallback: when an exact `from` finds nothing,
  the collector retries on the base local part and keeps only messages whose
  normalised sender matches, noting `searchNote` in the manifest.
- `periodStart`/`periodEnd` (and `outlook_search_emails`' `startDate`/`endDate`)
  accept bare dates: `2026-08-31` as an end covers the **whole** last day
  instead of stopping at midnight.
- `outlook_download_attachment` accepts `saveToFile`/`destDir`/`fileName`/`onExisting`
  to write raw bytes server-side and return only `{savedPath, size, sha256}` —
  use this for real invoice PDFs instead of inline base64. The saved file is
  byte-exact with no inline size ceiling.
- `outlook_download_attachment` and `outlook_save_attachment` read the attachment
  content from its JSON item and fall back to the raw `/$value` endpoint, so
  attachment ids containing `-`, `_` and `=` work. Only file attachments have
  downloadable bytes: attached emails (item attachments) and reference
  attachments are rejected with a clear message, and a failure on both routes
  reports the routes tried, the ids and the Graph error.

### Email search

A free-text `outlook_search_emails` query (without a folder restriction) uses KQL, and its date range is sent to Microsoft Graph as a real property restriction with ISO UTC dates, e.g. `"receipt AND received:2026-06-01..2026-06-30"` — not matched as a literal phrase, and independent of the host's locale and time zone. Other searches use `$filter`.

- Results are re-checked against the requested date window, sender and subject. Anything the service failed to filter is dropped and counted in `searchSummary.droppedOutsideFilters`.
- A folder restriction is bound by the endpoint: a multi-folder search queries each folder and merges the results, rather than silently searching the whole mailbox.
- KQL values are quoted only when they contain whitespace or KQL syntax; an unparseable date is a validation error.

### Shared mailboxes (v1.3)

Every mail, folder, and attachment tool — and every receipt tool
(`collect_receipts`, `extract_receipt`, `save_attachment`, `render_email_pdf`,
`fetch_billing_pdf`) — accepts an optional `mailbox`
argument (e.g. `careers@yourcompany.com`). Precedence: per-call `mailbox` →
the `MCP_OUTLOOK_SHARED_MAILBOX` setting → your own mailbox. Requirements:

- Exchange **Full Access** delegation to the shared mailbox for the signed-in
  user (Send As or Send on Behalf additionally governs how sends appear).
  One styling note: signatures and font styling are always the signed-in
  user's own — a send from a shared mailbox carries your personal signature,
  so consider `preserveUserStyling: false` for shared sends.
- The `Mail.*.Shared` delegated scopes — added in v1.3, so **each user must
  re-consent once**: the interactive flow prompts automatically on next
  sign-in; headless setups re-run `npm run auth:bootstrap` once.

Graph offers no delegated API that lists the mailboxes you hold Full Access to,
so discovery is candidate-based: `outlook_list_shared_mailboxes` probes the
addresses in the **Known Shared Mailboxes** setting
(`MCP_OUTLOOK_KNOWN_MAILBOXES`), the Shared Mailbox setting, and any
`candidates` you pass, then reports which ones actually open (with inbox
counts) and which are denied, missing, or malformed. Feed an accessible
address straight back as the `mailbox` argument.

Calendar tools stay on your own calendar in this release.

### Headless (Scheduled) Runs

1. Seed tokens once, interactively: `npm run auth:bootstrap` (opens the browser PKCE flow and stores an encrypted refresh token).
2. Set `MCP_OUTLOOK_AUTH_MODE=headless` for the scheduled run. The server refreshes silently and **never launches a browser**; if re-consent is genuinely required it fails fast with an actionable error telling you to re-run the bootstrap.
3. Optionally set `MCP_OUTLOOK_REFRESH_TOKEN_PATH` to point the run at a specific token store directory.

### Large File Handling

When downloading large attachments or SharePoint files, the server automatically detects when the response would exceed the MCP response cap (~30k characters by default — the practical tool-output token budget; override with `MCP_OUTLOOK_MAX_RESPONSE_CHARS`) and saves the content to local files instead.

- If `MCP_OUTLOOK_WORK_DIR` is set, large files are saved to this directory
- If not set, files are saved to the system temp directory
- Files are automatically named with timestamps to avoid conflicts
- Old files are periodically cleaned up to manage disk space

---

## Example Prompts

Once installed, you can ask the AI assistant things like:

**Email Management**
- "Show me my unread emails from this week"
- "Find all emails from John about the project proposal"
- "Send a reply to the last email from Sarah thanking her for the update"
- "Draft an email to the team summarizing today's meeting"

**Calendar**
- "What meetings do I have tomorrow?"
- "Schedule a 30-minute call with Alex next Tuesday afternoon"
- "Show me my availability for the rest of the week"

**Attachments & SharePoint**
- "Download and summarize the PDF attachment from the latest email from Finance"
- "Get the contents of this SharePoint link: [paste link]"
- "What files were attached to emails from Legal this month?"

**Receipts & Invoices**
- "Collect all my June receipts into my receipts folder and give me the manifest"
- "Save the invoice PDF from the latest Acme receipt email as 'Acme 29Jun26 Invoice.pdf'"
- "Draft an email to my accountant with last month's receipt PDFs attached — don't send it"

**Office Document Processing**

The server automatically parses:
- **PDF files**: Extracts text content
- **Word documents** (.docx): Extracts text content
- **PowerPoint** (.pptx): Extracts slide text
- **Excel** (.xlsx): Parses data into structured format

---

## Authentication

The server uses OAuth 2.0 with PKCE for secure authentication:

1. First run will open a browser for Microsoft authentication
2. Tokens are encrypted and stored locally (uses OS keychain if available, otherwise encrypted file storage)
3. Automatic token refresh for long-term usage
4. No sensitive data stored in plain text

### Required Permissions

The app requests these Microsoft Graph permissions:

- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send` - Email access
- `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared` - Shared/delegated mailbox email access (v1.3)
- `Calendars.Read`, `Calendars.ReadWrite` - Calendar access  
- `Contacts.Read`, `Contacts.ReadWrite` - Contact access
- `Tasks.Read`, `Tasks.ReadWrite` - Task access
- `User.Read`, `MailboxSettings.Read` - User profile
- `Files.Read.All`, `Files.ReadWrite.All` - OneDrive/SharePoint files
- `Sites.Read.All`, `Sites.ReadWrite.All` - SharePoint sites
- `offline_access` - Refresh tokens

---

## Troubleshooting

### Large File Issues
- **Problem**: "Result exceeds maximum length" error
- **Solution**: Ensure `MCP_OUTLOOK_WORK_DIR` is set and writable
- **Alternative**: Files automatically save to system temp if work dir not configured

### Authentication Issues
- **Problem**: Authentication failures
- **Solution**: Verify Azure AD app permissions and client ID
- **Reset**: Clear stored tokens and re-authenticate

### Graph API Errors
- **Problem**: A tool fails with a Microsoft Graph error
- **Details**: The error text includes the HTTP status, Graph error code and message, the `request-id` and `client-request-id`, and the request method and URL as sent (never the auth header), e.g. `HTTP 400 BadRequest: <message> [request-id: …; client-request-id: …] GET https://graph.microsoft.com/…`. The same data is returned structured as `_errorDetails.graphError` (plus `retryAfter` when Graph throttles), and per entry as `graphError` in `outlook_collect_receipts`' `failed[]`.
- **Solution**: Quote the request ids when raising the failure with Microsoft support

### SharePoint Access Issues
- **Problem**: Cannot access SharePoint files
- **Solution**: Ensure sharing links are valid and user has access permissions
- **Alternative**: Use direct file ID access if available

---

## Development

### Project Structure
```
outlook-mcp/
├── server/
│   ├── index.js              # Main MCP server
│   ├── auth/                 # Authentication management
│   ├── graph/                # Microsoft Graph API client
│   ├── schemas/              # MCP tool schemas
│   ├── tools/                # MCP tool implementations
│   │   ├── attachments/      # Attachment tools
│   │   ├── calendar/         # Calendar tools
│   │   ├── email/            # Email tools
│   │   ├── folders/          # Folder management
│   │   ├── receipts/         # Receipt/invoice tools
│   │   └── sharepoint/       # SharePoint tools
│   └── utils/                # Utility modules
└── package.json
```

### Running Tests
```bash
npm test                    # Run all tests
npm run test:watch          # Watch mode
npm run test:benchmark      # Performance benchmarks
```

### Debugging
```bash
npm run test:graph          # Test Graph API connection
```

---

## Support

If this tool saved you time, consider supporting the development!

[![Ko-fi](https://img.shields.io/badge/Ko--fi-Support-ff5f5f?logo=ko-fi)](https://ko-fi.com/xenoxilus)

---

## License

MIT License

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with tests
4. Submit a pull request
