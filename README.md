# Genesys Cloud — Conversation Suggestions Exporter (CLM-safe PowerShell 5.1)

`Get-GcConversationSuggestions.ps1` pulls **Agent Copilot suggestions** for one or
more conversations from:

```
GET /api/v2/conversations/{conversationId}/suggestions
```

and exports them to CSV (plus optional raw JSON), on endpoints locked down with
AppLocker / WDAC where Windows PowerShell 5.1 runs in **Constrained Language
Mode (CLM)**.

## Why this script is different

Standard Genesys samples break in CLM. This script avoids everything CLM blocks:

| CLM restriction | What this script does instead |
|---|---|
| `[Convert]::ToBase64String()`, `[Text.Encoding]::` | Pure-PowerShell Base64 + manual UTF-8 encoding using `-shl`/`-shr`/`-band`/`-bor` over `[int][char]` values |
| `::new()` constructors | Plain arrays with `+=` |
| `[pscustomobject]@{...}` | `New-Object PSObject -Property @{...}` + explicit `Select-Object` column ordering before `Export-Csv` |
| `[uri]::EscapeDataString()` | Minimal `.Replace()` encoding of `% + / = &` (only applied to pagination cursor tokens) |
| Deep .NET exception chains | `[int]$_.Exception.Response.StatusCode` inside try/catch, with fallback `-like` matching on `$_.Exception.Message` |

Every response is validated (e.g. `access_token` must exist before the script
continues), and HTTP 429 rate limits are retried with `Retry-After` /
exponential backoff.

## Prerequisites

1. **OAuth client** (Admin → Integrations → OAuth) using the **Client
   Credentials** grant.
2. The role assigned to that client must:
   - include the Agent Copilot **suggestions view** permission
     (e.g. *Assistants → Suggestion → View*; naming can vary by org),
   - include **Analytics → Conversation Detail → View** (needed by discovery
     mode's conversation-details query), and
   - be **assigned to the division** you query — a `403` almost always means
     the client's role is not in that division.
3. TLS 1.2 must be the OS default (Windows 10/11 / Server 2019+ already is).
   CLM blocks `[Net.ServicePointManager]::SecurityProtocol`, so if you see
   *"Could not create SSL/TLS secure channel"* fix it machine-wide via the
   `SystemDefaultTlsVersions` / `SchUseStrongCrypto` registry values (GPO/admin).

## One-time setup: embed your credentials

The script ships pre-configured for the **Australia (Sydney) region**
(`mypurecloud.com.au`) with the OAuth credentials embedded in a config block
near the top of the file. Edit these two lines before first use:

```powershell
$EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
```

The script refuses to run while the placeholders are still in place.
Command-line `-ClientId` / `-ClientSecret` / `-Region` still override the
embedded values when supplied.

> **Security:** embedded credentials are readable by anyone who can read the
> file. Restrict NTFS permissions on the script, and scope the OAuth client's
> role to just the suggestions-view permission in the one division you query.
> Never commit the file with real credentials to source control.

## Usage

### Discovery mode (recommended) — just give it a division ID

The script finds the conversation IDs itself via
`POST /api/v2/analytics/conversations/details/query` (filtered on the
`divisionId` dimension, newest first), then pulls suggestions for each.
Default window: the last 7 days.

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -DivisionId '11111111-2222-3333-4444-555555555555'
```

Custom date range and cap (ranges over 7 days are split into 7-day analytics
windows automatically):

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -DivisionId '11111111-2222-3333-4444-555555555555' `
    -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
    -MaxConversations 2000 `
    -OutputCsv C:\Reports\suggestions.csv `
    -OutputJson C:\Reports\suggestions.json
```

### Explicit mode — you already have the conversation IDs

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -ConversationId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
```

When both `-ConversationId` and `-DivisionId` are supplied, the division ID
acts as a guard: each conversation's division is verified via
`GET /api/v2/conversations/{id}` and mismatches are skipped with a warning.

Re-use an existing bearer token (skips the token request entirely):

```powershell
.\Get-GcConversationSuggestions.ps1 -Region mypurecloud.com `
    -AccessToken $token -ConversationId $convId
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Region` | no | Region domain only; defaults to embedded `mypurecloud.com.au` (Australia) |
| `-ClientId` / `-ClientSecret` | no | Override the embedded Client Credentials pair |
| `-AccessToken` | no | Existing bearer token; bypasses the token request |
| `-DivisionId` | yes* | Division to pull from. Alone → discovery mode; combined with `-ConversationId` → division guard |
| `-ConversationId` | yes* | Explicit conversation IDs (*provide this or `-DivisionId`) |
| `-StartDate` / `-EndDate` | no | Discovery window (default: last 7 days → now) |
| `-MaxConversations` | no | Discovery cap, newest first (default 500, max 10000) |
| `-PageSize` | no | Suggestions page size, 1–500 (default 100) |
| `-OutputCsv` | no | CSV path (default `.\GcSuggestions_<timestamp>.csv`) |
| `-OutputJson` | no | Also dump the raw suggestion objects as pretty JSON |

## Output

CSV columns (explicitly ordered): `ConversationId, SuggestionId, SuggestionType,
State, DateIssued, Confidence, ResourceId, ResourceTitle, KnowledgeBaseId,
RetrievedAtUtc, RawJson`.

`RawJson` holds the complete, unflattened suggestion entity (depth 15,
compressed), so nothing the API returned is ever lost — new/unknown suggestion
types still land there even if the flattened columns stay empty.

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `401` on token request | Wrong client ID/secret, or wrong region's login host |
| `no access_token was returned` | OAuth client isn't a Client Credentials grant |
| `403` on suggestions call | Role missing suggestions permission, or role not assigned to the conversation's division |
| `404` on a suggestions call | Counted as "no suggestions" and skipped — normal for conversations where Agent Copilot was never active |
| Discovery finds 0 conversations | Wrong division ID, empty date range, or the analytics permission is missing |
| Repeated `429` warnings | Normal — the script honors `Retry-After` and backs off automatically |
