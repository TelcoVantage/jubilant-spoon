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
     (e.g. *Assistants → Suggestion → View*; naming can vary by org), and
   - be **assigned to the division** that owns the conversations. Conversation
     IDs you pass are expected to come from a single division — a `403` on the
     suggestions call almost always means the client's role is not in that
     division.
3. TLS 1.2 must be the OS default (Windows 10/11 / Server 2019+ already is).
   CLM blocks `[Net.ServicePointManager]::SecurityProtocol`, so if you see
   *"Could not create SSL/TLS secure channel"* fix it machine-wide via the
   `SystemDefaultTlsVersions` / `SchUseStrongCrypto` registry values (GPO/admin).

## Usage

Single conversation:

```powershell
.\Get-GcConversationSuggestions.ps1 `
    -Region usw2.pure.cloud `
    -ClientId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx' `
    -ClientSecret 'your-secret' `
    -ConversationId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
```

Multiple conversations, with a division guard and both exports:

```powershell
$convIds = @(
    'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
    'ffffffff-1111-2222-3333-444444444444'
)

.\Get-GcConversationSuggestions.ps1 `
    -Region mypurecloud.ie `
    -ClientId $id -ClientSecret $secret `
    -ConversationId $convIds `
    -DivisionId '11111111-2222-3333-4444-555555555555' `
    -OutputCsv C:\Reports\suggestions.csv `
    -OutputJson C:\Reports\suggestions.json
```

Re-use an existing bearer token (skips the token request entirely):

```powershell
.\Get-GcConversationSuggestions.ps1 -Region mypurecloud.com `
    -AccessToken $token -ConversationId $convId
```

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-Region` | no | Region domain only, e.g. `mypurecloud.com`, `usw2.pure.cloud`, `mypurecloud.ie` (default `mypurecloud.com`) |
| `-ClientId` / `-ClientSecret` | yes* | Client Credentials OAuth pair (*unless `-AccessToken` is given) |
| `-AccessToken` | no | Existing bearer token; bypasses the token request |
| `-ConversationId` | yes | One or more conversation IDs (same division) |
| `-DivisionId` | no | Verifies each conversation's division via `GET /api/v2/conversations/{id}` first; mismatches are skipped with a warning |
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
| `404` | Bad conversation ID, or the conversation lives in a different region/org |
| Repeated `429` warnings | Normal — the script honors `Retry-After` and backs off automatically |
