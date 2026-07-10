# Genesys Cloud — Agent Copilot Report Exporter (CLM-safe PowerShell 5.1)

`Get-GcConversationSuggestions.ps1` builds a **Power BI-ready report** of what
Genesys **Agent Copilot** did on your conversations. It is inspired by the
[copilot-conversation-inspector](https://github.com/GenesysCloudBlueprints/copilot-conversation-inspector)
blueprint, but PowerShell-only — no Vue/UI, just clean relational CSVs you
point Power BI at.

APIs used:

```
GET  /api/v2/conversations/{conversationId}/suggestions   Agent Copilot suggestions
GET  /api/v2/conversations/{conversationId}/summaries     Copilot session summaries
POST /api/v2/analytics/conversations/details/query        conversation discovery by division
```

Runs on endpoints locked down with AppLocker / WDAC where Windows
PowerShell 5.1 is in **Constrained Language Mode (CLM)** — see
[CLM safety](#why-this-script-is-clm-safe) below.

## Report output

Each run writes a report **folder** (default `.\GcCopilotReport_<timestamp>`)
containing flat CSVs — no embedded JSON blobs, one fact per column:

| File | Grain | Columns |
|---|---|---|
| `Conversations.csv` | 1 row per conversation | ConversationId, ConversationStart, ConversationEnd, MediaTypes, QueueName, CustomerName, SuggestionCount, SummaryCount |
| `Suggestions.csv` | 1 row per Copilot suggestion | ConversationId, SuggestionId, SuggestionType, State, TriggerType, DateCreated, Confidence, Title, AnswerText, DocumentId, KnowledgeBaseId, ArticleUrl, SearchId, MediaType, QueueId, AgentUserId, ExternalContactId |
| `SuggestionSnippets.csv` | 1 row per knowledge snippet | ConversationId, SuggestionId, SnippetIndex, SnippetText |
| `Summaries.csv` | 1 row per Copilot session summary | ConversationId, SummaryId, MediaType, Language, Status, SummaryText, Confidence, ReasonText, ReasonDescription, ResolutionText, ResolutionDescription, ResolutionOutcome, FollowupText, FollowupDescription, PredictedWrapupCodes |

### Power BI modelling

Load the folder, then relate:

- `Suggestions[ConversationId]` → `Conversations[ConversationId]` (many-to-one)
- `SuggestionSnippets[SuggestionId]` → `Suggestions[SuggestionId]` (many-to-one)
- `Summaries[ConversationId]` → `Conversations[ConversationId]` (many-to-one)

Useful columns for visuals:

- **State** (`Suggested` / `Accepted` / `Dismissed` / `Failed` / `Rated`) — Copilot adoption funnel
- **SuggestionType** (`KnowledgeSearch` / `CannedResponse` / `Script`) — mix of what Copilot surfaces
- **TriggerType** (`Fallback` / `ExplicitQuery`) — how suggestions were raised
- **Confidence** — numeric 0–1, format as percentage in Power BI
- **ArticleUrl** — set the data category to *Web URL* for clickable Knowledge Workbench deep-links
- **ResolutionOutcome** / **PredictedWrapupCodes** — summary quality views

Add `-RawJson` to also drop `RawSuggestions.json` / `RawSummaries.json` in the
folder for full-fidelity debugging (never mixed into the CSVs).

## One-time setup: embed your credentials

The script ships pre-configured for the **Australia (Sydney) region**
(`mypurecloud.com.au`) with the OAuth credentials embedded in a config block
near the top of the file. Edit these two lines before first use:

```powershell
$EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
```

The script refuses to run while the placeholders are still in place.
Command-line `-ClientId` / `-ClientSecret` / `-Region` override the embedded
values when supplied.

> **Security:** embedded credentials are readable by anyone who can read the
> file. Restrict NTFS permissions on the script, scope the OAuth client's role
> to the minimum permissions in the one division you query, and never commit
> the file with real credentials.

### Required permissions (OAuth client's role)

- Agent Copilot **suggestions view** (e.g. *Assistants → Suggestion → View*)
- **Speech and Text Analytics / summaries view** for `Summaries.csv`
  (skip with `-SkipSummaries` if not licensed)
- **Analytics → Conversation Detail → View** (discovery mode)
- Role **assigned to the division** you query — a `403` almost always means it isn't

## Usage

### Discovery mode (recommended) — just give it a division ID

Finds conversations via the analytics details query (filtered on the
`divisionId` conversation dimension, newest first; default window: last
7 days), then pulls suggestions + summaries for each:

```powershell
.\Get-GcConversationSuggestions.ps1 -DivisionId '11111111-2222-3333-4444-555555555555'
```

Custom window, bigger cap, fixed folder for a Power BI refresh:

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -DivisionId '11111111-2222-3333-4444-555555555555' `
    -StartDate (Get-Date).AddDays(-30) `
    -MaxConversations 2000 `
    -OutputFolder 'C:\Reports\CopilotWeekly'
```

Ranges longer than the analytics API's 7-day interval limit are chunked into
7-day windows automatically, walked newest-first.

### Explicit mode — you already have the conversation IDs

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -ConversationId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' -SkipSummaries
```

With both `-ConversationId` and `-DivisionId`, the division acts as a guard:
each conversation's division is verified via `GET /api/v2/conversations/{id}`
and mismatches are skipped with a warning.

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-DivisionId` | yes* | Division to pull from. Alone → discovery mode; with `-ConversationId` → division guard |
| `-ConversationId` | yes* | Explicit conversation IDs (*provide this or `-DivisionId`) |
| `-StartDate` / `-EndDate` | no | Discovery window (default: last 7 days → now) |
| `-MaxConversations` | no | Discovery cap, newest first (default 500, max 10000) |
| `-PageSize` | no | Suggestions page size (default 200, like the blueprint app) |
| `-OutputFolder` | no | Report folder (default `.\GcCopilotReport_<timestamp>`) |
| `-SkipSummaries` | no | Skip the summaries call; `Summaries.csv` omitted |
| `-RawJson` | no | Also write raw API payloads as JSON into the folder |
| `-Region` | no | Region domain; defaults to embedded `mypurecloud.com.au` |
| `-ClientId` / `-ClientSecret` | no | Override the embedded Client Credentials pair |
| `-AccessToken` | no | Existing bearer token; bypasses the token request |

## Why this script is CLM-safe

Standard Genesys samples break in Constrained Language Mode. This script
avoids everything CLM blocks:

| CLM restriction | What this script does instead |
|---|---|
| `[Convert]::ToBase64String()`, `[Text.Encoding]::` | Pure-PowerShell Base64 + manual UTF-8 encoding using `-shl`/`-shr`/`-band`/`-bor` over `[int][char]` values |
| `::new()` constructors | Plain arrays with `+=` |
| `[pscustomobject]@{...}` | `New-Object PSObject -Property @{...}` + explicit `Select-Object` column ordering before `Export-Csv` |
| `[uri]::EscapeDataString()` | Minimal `.Replace()` encoding of `% + / = &` (pagination cursors only) |
| Deep .NET exception chains | `[int]$_.Exception.Response.StatusCode` in try/catch, with fallback `-like` matching on `$_.Exception.Message` |

Every response is validated (e.g. `access_token` must exist before the script
continues), HTTP 429 rate limits retry with `Retry-After`/exponential backoff,
and error messages include the Genesys API's own error body.

TLS note: CLM blocks `[Net.ServicePointManager]::SecurityProtocol`; on
Windows 10/11 / Server 2019+ TLS 1.2 is already the OS default. If you see
*"Could not create SSL/TLS secure channel"*, fix it machine-wide via the
`SystemDefaultTlsVersions` / `SchUseStrongCrypto` registry values (GPO/admin).

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| `401` on token request | Wrong client ID/secret, or wrong region's login host |
| `no access_token was returned` | OAuth client isn't a Client Credentials grant |
| `403` on suggestions/summaries | Role missing the permission, or role not in the division (message includes the API's own explanation) |
| `404` on a suggestions call | Counted as "no suggestions" — normal for conversations where Agent Copilot was never active |
| Discovery finds 0 conversations | Wrong division ID, empty date range, or missing analytics permission |
| Repeated `429` warnings | Normal — the script honors `Retry-After` and backs off automatically |
