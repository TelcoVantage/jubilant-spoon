<#
.SYNOPSIS
    Genesys Cloud Agent Copilot report exporter (Power BI friendly).

    Inspired by the GenesysCloudBlueprints/copilot-conversation-inspector
    blueprint, but PowerShell-only: no UI, just clean relational CSVs.

.DESCRIPTION
    For every conversation (explicit IDs, or discovered dynamically from a
    division ID) the script calls:

        GET  /api/v2/conversations/{conversationId}/suggestions   (Agent Copilot suggestions)
        GET  /api/v2/conversations/{conversationId}/summaries     (Copilot session summaries)
        POST /api/v2/analytics/conversations/details/query        (discovery mode only)

    and writes a report FOLDER containing flat, star-schema style CSVs that
    load straight into Power BI (relate the child tables on ConversationId /
    SuggestionId):

        Conversations.csv       1 row per conversation  (dates, media, queue,
                                customer, suggestion/summary counts)
        Suggestions.csv         1 row per Copilot suggestion, fully flattened
                                (type, state, trigger, confidence, article
                                link, extracted answer, context)
        SuggestionSnippets.csv  1 row per knowledge snippet (a suggestion can
                                carry several) - child of Suggestions.csv
        Summaries.csv           1 row per Copilot session summary (summary
                                text, reason, resolution, follow-up, wrap-ups)

    No embedded JSON blobs in any CSV column. Add -RawJson if you also want
    the untouched API payloads saved alongside for debugging.

    Designed for Windows PowerShell 5.1 in CONSTRAINED LANGUAGE MODE
    (AppLocker / WDAC locked-down endpoints). The entire script is CLM-safe:

      * No .NET static method calls  (no [Convert]::, [Text.Encoding]::, [uri]::, [math]::)
      * No ::new() constructors      (plain arrays with += instead)
      * No [pscustomobject]@{} casts (New-Object PSObject -Property @{} instead,
                                      with explicit Select-Object column ordering)
      * Base64 for the OAuth Basic header is implemented in pure PowerShell
        using bit operators (-shl / -shr / -band / -bor) over [int][char] values.
      * HTTP status codes are read via [int]$_.Exception.Response.StatusCode inside
        try/catch, with fallback string matching on $_.Exception.Message.
      * Every API response is validated explicitly before being used.

    TLS NOTE: CLM blocks setting [Net.ServicePointManager]::SecurityProtocol.
    On a current Windows 10/11 or Server 2019+ endpoint TLS 1.2 is negotiated
    by default. If you hit "Could not create SSL/TLS secure channel", enable
    strong crypto machine-wide via registry (SystemDefaultTlsVersions /
    SchUseStrongCrypto) - an admin/GPO change, not a script change.

.PARAMETER Region
    Genesys Cloud region domain (NOT the full URL). Defaults to the embedded
    value 'mypurecloud.com.au' (Australia / Sydney).

.PARAMETER ClientId
    OAuth client ID (Client Credentials grant). Defaults to the embedded
    $EmbeddedClientId value in the configuration block below.

.PARAMETER ClientSecret
    OAuth client secret. Defaults to the embedded $EmbeddedClientSecret value.

.PARAMETER AccessToken
    Optional. Supply an existing bearer token to skip the token request.

.PARAMETER ConversationId
    Optional. Explicit conversation IDs. When omitted, supply -DivisionId and
    the script discovers conversations dynamically.

.PARAMETER DivisionId
    The division to work with. Alone -> discovery mode (analytics details
    query filtered on the divisionId conversation dimension, newest first).
    Combined with -ConversationId -> division guard on each conversation.

.PARAMETER StartDate
    Discovery window start (default: 7 days ago). Windows longer than the
    analytics API's 7-day interval limit are chunked automatically.

.PARAMETER EndDate
    Discovery window end (default: now).

.PARAMETER MaxConversations
    Discovery cap, newest first (default 500).

.PARAMETER PageSize
    Suggestions page size (default 200, same as the blueprint app).

.PARAMETER OutputFolder
    Report folder to create/write. Default: .\GcCopilotReport_<timestamp>

.PARAMETER SkipSummaries
    Skip the per-conversation summaries call (faster; Summaries.csv omitted).

.PARAMETER RawJson
    Also write RawSuggestions.json / RawSummaries.json into the report folder.

.EXAMPLE
    # Discovery mode: last 7 days of one division, full report folder
    .\Get-GcConversationSuggestions.ps1 -DivisionId '11111111-2222-3333-4444-555555555555'

.EXAMPLE
    # 30-day window, bigger cap, custom folder for the Power BI refresh
    .\Get-GcConversationSuggestions.ps1 `
        -DivisionId '11111111-2222-3333-4444-555555555555' `
        -StartDate (Get-Date).AddDays(-30) `
        -MaxConversations 2000 `
        -OutputFolder 'C:\Reports\CopilotWeekly'

.EXAMPLE
    # Explicit conversation IDs, suggestions only (no summaries)
    .\Get-GcConversationSuggestions.ps1 `
        -ConversationId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' -SkipSummaries
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Region = '',

    [Parameter(Mandatory = $false)]
    [string]$ClientId = '',

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = '',

    [Parameter(Mandatory = $false)]
    [string]$AccessToken = '',

    [Parameter(Mandatory = $false)]
    [string[]]$ConversationId = @(),

    [Parameter(Mandatory = $false)]
    [string]$DivisionId = '',

    [Parameter(Mandatory = $false)]
    [datetime]$StartDate = (Get-Date).AddDays(-7),

    [Parameter(Mandatory = $false)]
    [datetime]$EndDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10000)]
    [int]$MaxConversations = 500,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 500)]
    [int]$PageSize = 200,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = '',

    [Parameter(Mandatory = $false)]
    [switch]$SkipSummaries,

    [Parameter(Mandatory = $false)]
    [switch]$RawJson
)

$ErrorActionPreference = 'Stop'

# ===========================================================================
# EMBEDDED CONFIGURATION - Australia (Sydney) region
# Paste your OAuth Client Credentials pair below. Command-line -ClientId /
# -ClientSecret / -Region parameters still work and override these values.
#
# SECURITY: anyone who can read this file can read these credentials. Restrict
# NTFS permissions on the script and scope the OAuth client's role to the
# minimum permissions in the one division you query.
# ===========================================================================
$EmbeddedRegion       = 'mypurecloud.com.au'
$EmbeddedClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$EmbeddedClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'

if ($Region -eq '')       { $Region       = $EmbeddedRegion }
if ($ClientId -eq '')     { $ClientId     = $EmbeddedClientId }
if ($ClientSecret -eq '') { $ClientSecret = $EmbeddedClientSecret }

if (($AccessToken -eq '') -and (($ClientId -like 'PASTE-YOUR-*') -or ($ClientSecret -like 'PASTE-YOUR-*'))) {
    throw 'Embedded credentials are still placeholders. Edit $EmbeddedClientId / $EmbeddedClientSecret at the top of the script (or pass -ClientId / -ClientSecret).'
}

$loginBase = 'https://login.' + $Region
$apiBase   = 'https://api.' + $Region

# ---------------------------------------------------------------------------
# Pure-PowerShell Base64 (CLM-safe: no [Convert]:: / [Text.Encoding]::)
# UTF-8 encodes the string manually, then packs 3 bytes -> 4 base64 chars
# using only -shl / -shr / -band / -bor over [int] values.
# ---------------------------------------------------------------------------
function ConvertTo-GcBase64 {
    param([Parameter(Mandatory = $true)][string]$Text)

    $b64Alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

    # --- Manual UTF-8 encoding ---
    $bytes = @()
    foreach ($ch in $Text.ToCharArray()) {
        $cp = [int]$ch
        if ($cp -le 0x7F) {
            $bytes += $cp
        }
        elseif ($cp -le 0x7FF) {
            $bytes += (0xC0 -bor (($cp -shr 6) -band 0x1F))
            $bytes += (0x80 -bor ($cp -band 0x3F))
        }
        elseif (($cp -ge 0xD800) -and ($cp -le 0xDFFF)) {
            # Surrogate pair (emoji etc.) - not expected in OAuth credentials.
            throw 'ConvertTo-GcBase64: characters outside the Basic Multilingual Plane are not supported.'
        }
        else {
            $bytes += (0xE0 -bor (($cp -shr 12) -band 0x0F))
            $bytes += (0x80 -bor (($cp -shr 6) -band 0x3F))
            $bytes += (0x80 -bor ($cp -band 0x3F))
        }
    }

    # --- 3-byte -> 4-char packing ---
    $out   = ''
    $i     = 0
    $total = $bytes.Count
    while ($i -lt $total) {
        $b0 = [int]$bytes[$i]
        $hasB1 = (($i + 1) -lt $total)
        $hasB2 = (($i + 2) -lt $total)
        $b1 = 0
        $b2 = 0
        if ($hasB1) { $b1 = [int]$bytes[$i + 1] }
        if ($hasB2) { $b2 = [int]$bytes[$i + 2] }

        $out += [string]$b64Alphabet[(($b0 -shr 2) -band 0x3F)]
        $out += [string]$b64Alphabet[((($b0 -band 0x03) -shl 4) -bor (($b1 -shr 4) -band 0x0F))]
        if ($hasB1) {
            $out += [string]$b64Alphabet[((($b1 -band 0x0F) -shl 2) -bor (($b2 -shr 6) -band 0x03))]
        } else {
            $out += '='
        }
        if ($hasB2) {
            $out += [string]$b64Alphabet[($b2 -band 0x3F)]
        } else {
            $out += '='
        }
        $i += 3
    }
    return $out
}

# ---------------------------------------------------------------------------
# Safe property reader for API response objects (PSObject access is CLM-safe).
# Returns $null instead of throwing when a property is absent.
# ---------------------------------------------------------------------------
function Get-Prop {
    param($InputObject, [Parameter(Mandatory = $true)][string]$Name)

    if ($null -eq $InputObject) { return $null }
    $p = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

# Reads the first non-empty property among $Names - handy for the legacy /
# alternate payload shapes the blueprint also tolerates.
function Get-FirstProp {
    param($InputObject, [string[]]$Names)

    foreach ($n in $Names) {
        $v = Get-Prop -InputObject $InputObject -Name $n
        if (($null -ne $v) -and ([string]$v -ne '')) { return $v }
    }
    return $null
}

# Summaries use "confidence fields": either a plain string, or an object with
# text/value/content + description + confidence + outcome. Flatten safely.
function Get-CFText {
    param($Field)

    if ($null -eq $Field) { return '' }
    if ($Field -is [string]) { return $Field }
    $v = Get-FirstProp -InputObject $Field -Names @('text', 'value', 'content')
    if ($null -ne $v) { return [string]$v }
    return ''
}

function Get-CFDetail {
    param($Field, [string]$Name)

    if ($null -eq $Field) { return '' }
    if ($Field -is [string]) { return '' }
    $v = Get-Prop -InputObject $Field -Name $Name
    if ($null -ne $v) { return [string]$v }
    return ''
}

# ---------------------------------------------------------------------------
# HTTP status extraction per the CLM error-handling pattern:
# [int]$_.Exception.Response.StatusCode in try/catch, then message fallback.
# ---------------------------------------------------------------------------
function Get-HttpStatusFromError {
    param($ErrorRecord)

    $status = 0
    try { $status = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $status = 0 }

    if ($status -eq 0) {
        $msg = ''
        try { $msg = [string]$ErrorRecord.Exception.Message } catch { $msg = '' }
        if     ($msg -like '*429*')                                { $status = 429 }
        elseif ($msg -like '*401*' -or $msg -like '*Unauthorized*'){ $status = 401 }
        elseif ($msg -like '*403*' -or $msg -like '*Forbidden*')   { $status = 403 }
        elseif ($msg -like '*404*' -or $msg -like '*Not Found*')   { $status = 404 }
    }
    return $status
}

# ---------------------------------------------------------------------------
# Generic API caller with 429 (rate limit) retry + exponential backoff.
# ---------------------------------------------------------------------------
function Invoke-GcApi {
    param(
        [string]$Method = 'Get',
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        $Body = $null,
        [string]$ContentType = 'application/json',
        [int]$MaxAttempts = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            if ($null -ne $Body) {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -Body $Body -ContentType $ContentType -ErrorAction Stop
            }
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ErrorAction Stop
        }
        catch {
            $status = Get-HttpStatusFromError -ErrorRecord $_

            if (($status -eq 429) -and ($attempt -lt $MaxAttempts)) {
                $retryAfter = 0
                try { $retryAfter = [int]$_.Exception.Response.Headers['Retry-After'] } catch { $retryAfter = 0 }
                if ($retryAfter -le 0) { $retryAfter = 2 * $attempt * $attempt }   # 2, 8, 18, 32 s
                Write-Warning ('Rate limited (429). Waiting ' + $retryAfter + 's before retry ' + ($attempt + 1) + '/' + $MaxAttempts + ' ...')
                Start-Sleep -Seconds $retryAfter
                continue
            }

            $detail = ''
            try { $detail = [string]$_.Exception.Message } catch { $detail = '' }

            # Genesys returns a JSON body explaining WHY a request failed
            # (e.g. which query field a 400 rejected). PowerShell exposes it
            # on ErrorDetails - CLM-safe, and far more useful than the
            # generic 'Bad Request' text.
            $apiBody = ''
            try { $apiBody = [string]$_.ErrorDetails.Message } catch { $apiBody = '' }
            if ($apiBody -ne '') { $detail = $detail + ' API says: ' + $apiBody }

            if     ($status -eq 400) { throw ('Bad request (400). The API rejected the request body/parameters. ' + $detail + ' URI: ' + $Uri) }
            elseif ($status -eq 401) { throw ('API call unauthorized (401). Token expired or invalid. URI: ' + $Uri) }
            elseif ($status -eq 403) { throw ('API call forbidden (403). Check the OAuth client role has the required permission AND is assigned to the division. ' + $apiBody + ' URI: ' + $Uri) }
            elseif ($status -eq 404) { throw ('Resource not found (404). URI: ' + $Uri) }
            elseif ($status -gt 0)   { throw ('API call failed (HTTP ' + $status + '): ' + $detail + ' URI: ' + $Uri) }
            else                     { throw ('API call failed (network/unknown): ' + $detail + ' URI: ' + $Uri) }
        }
    }
}

# ---------------------------------------------------------------------------
# OAuth Client Credentials token (validated before use).
# ---------------------------------------------------------------------------
function Get-GcAccessToken {
    param(
        [Parameter(Mandatory = $true)][string]$LoginBase,
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Secret
    )

    $basic = ConvertTo-GcBase64 -Text ($Id + ':' + $Secret)
    $tokenHeaders = @{
        'Authorization' = 'Basic ' + $basic
    }

    Write-Host ('Requesting OAuth token from ' + $LoginBase + ' ...')
    $resp = Invoke-GcApi -Method 'Post' `
                         -Uri ($LoginBase + '/oauth/token') `
                         -Headers $tokenHeaders `
                         -Body 'grant_type=client_credentials' `
                         -ContentType 'application/x-www-form-urlencoded'

    $token = Get-Prop -InputObject $resp -Name 'access_token'
    if (($null -eq $token) -or ([string]$token -eq '')) {
        throw 'Token endpoint responded but no access_token was returned. Verify the OAuth client uses the Client Credentials grant.'
    }

    $ttl = Get-Prop -InputObject $resp -Name 'expires_in'
    if ($null -ne $ttl) {
        Write-Host ('Token acquired (expires in ' + [string]$ttl + 's).')
    } else {
        Write-Host 'Token acquired.'
    }
    return [string]$token
}

# ---------------------------------------------------------------------------
# Optional division guard: confirm the conversation belongs to -DivisionId.
# ---------------------------------------------------------------------------
function Test-GcConversationDivision {
    param(
        [Parameter(Mandatory = $true)][string]$ApiBase,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][string]$ConvId,
        [Parameter(Mandatory = $true)][string]$ExpectedDivisionId
    )

    $conv = Invoke-GcApi -Uri ($ApiBase + '/api/v2/conversations/' + $ConvId) -Headers $Headers

    $divisionIds = @()
    $divs = Get-Prop -InputObject $conv -Name 'divisions'
    if ($null -ne $divs) {
        foreach ($d in @($divs)) {
            $dObj = Get-Prop -InputObject $d -Name 'division'
            $dId  = Get-Prop -InputObject $dObj -Name 'id'
            if (($null -ne $dId) -and ([string]$dId -ne '')) { $divisionIds += [string]$dId }
        }
    }

    if ($divisionIds.Count -eq 0) {
        Write-Warning ('Conversation ' + $ConvId + ': no division information returned; proceeding anyway.')
        return $true
    }

    foreach ($dId in $divisionIds) {
        if ($dId -eq $ExpectedDivisionId) { return $true }
    }

    Write-Warning ('Conversation ' + $ConvId + ' belongs to division(s) [' + ($divisionIds -join ', ') + '], not the expected division ' + $ExpectedDivisionId + '. Skipping.')
    return $false
}

# ---------------------------------------------------------------------------
# Discovery mode: find conversations in a division dynamically via
# POST /api/v2/analytics/conversations/details/query, filtered on the
# divisionId CONVERSATION dimension (conversationFilters), newest first.
# The analytics API caps a single query interval at 7 days and a page at
# 100 rows, so wider ranges are split into 7-day windows and paged.
#
# Returns rich objects (id + start/end + media + queue + customer) so
# Conversations.csv carries useful reporting columns, not just IDs.
# ---------------------------------------------------------------------------
function Get-GcConversationsByDivision {
    param(
        [Parameter(Mandatory = $true)][string]$ApiBase,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][string]$DivId,
        [Parameter(Mandatory = $true)][datetime]$From,
        [Parameter(Mandatory = $true)][datetime]$To,
        [int]$Cap = 500
    )

    if ($From -ge $To) {
        throw ('-StartDate (' + [string]$From + ') must be earlier than -EndDate (' + [string]$To + ').')
    }

    $found = @()
    $seen  = @{}   # de-dupe: a conversation can span analytics windows

    # Walk BACKWARD from -EndDate in 7-day windows so "newest first" holds
    # across windows, not just inside one - the cap then keeps the most
    # recent conversations.
    $windowEnd = $To
    while (($windowEnd -gt $From) -and ($found.Count -lt $Cap)) {
        $windowStart = $windowEnd.AddDays(-7)
        if ($windowStart -lt $From) { $windowStart = $From }

        $interval = $windowStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z/' + `
                    $windowEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z'
        Write-Host ('  Analytics window: ' + $interval)

        $pageNumber = 0
        $maxPagesPerWindow = 100
        while (($pageNumber -lt $maxPagesPerWindow) -and ($found.Count -lt $Cap)) {
            $pageNumber++

            # divisionId is a CONVERSATION-level dimension, so the predicate
            # lives in conversationFilters (segmentFilters -> HTTP 400).
            $queryBody = @{
                interval            = $interval
                order               = 'desc'
                orderBy             = 'conversationStart'
                paging              = @{
                    pageSize   = 100
                    pageNumber = $pageNumber
                }
                conversationFilters = @(
                    @{
                        type       = 'or'
                        predicates = @(
                            @{
                                type      = 'dimension'
                                dimension = 'divisionId'
                                operator  = 'matches'
                                value     = $DivId
                            }
                        )
                    }
                )
            }
            $queryJson = ConvertTo-Json -InputObject $queryBody -Depth 10

            $resp = Invoke-GcApi -Method 'Post' `
                                 -Uri ($ApiBase + '/api/v2/analytics/conversations/details/query') `
                                 -Headers $Headers `
                                 -Body $queryJson

            $convs = $null
            if ($null -ne $resp) { $convs = Get-Prop -InputObject $resp -Name 'conversations' }
            $convArr = @()
            if ($null -ne $convs) { $convArr = @($convs) }
            if ($convArr.Count -eq 0) { break }   # window exhausted

            foreach ($c in $convArr) {
                $cId = Get-Prop -InputObject $c -Name 'conversationId'
                if (($null -eq $cId) -or ([string]$cId -eq '')) { continue }
                $key = [string]$cId
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true

                # --- Flatten reporting metadata from the analytics record ---
                $mediaTypes   = @()
                $customerName = ''
                $queueName    = ''
                $participants = Get-Prop -InputObject $c -Name 'participants'
                if ($null -ne $participants) {
                    foreach ($p in @($participants)) {
                        $purpose = [string](Get-Prop -InputObject $p -Name 'purpose')
                        $pName   = [string](Get-Prop -InputObject $p -Name 'participantName')

                        if (($customerName -eq '') -and (($purpose -eq 'customer') -or ($purpose -eq 'external'))) {
                            $customerName = $pName
                        }
                        if (($queueName -eq '') -and ($purpose -eq 'acd')) {
                            $queueName = $pName
                        }

                        $sessions = Get-Prop -InputObject $p -Name 'sessions'
                        if ($null -ne $sessions) {
                            foreach ($s in @($sessions)) {
                                $mt = [string](Get-Prop -InputObject $s -Name 'mediaType')
                                if (($mt -ne '') -and (-not ($mediaTypes -contains $mt))) { $mediaTypes += $mt }
                            }
                        }
                    }
                }

                $found += New-Object PSObject -Property @{
                    ConversationId    = $key
                    ConversationStart = [string](Get-Prop -InputObject $c -Name 'conversationStart')
                    ConversationEnd   = [string](Get-Prop -InputObject $c -Name 'conversationEnd')
                    MediaTypes        = ($mediaTypes -join ';')
                    QueueName         = $queueName
                    CustomerName      = $customerName
                }
                if ($found.Count -ge $Cap) { break }
            }

            Write-Host ('    Page ' + $pageNumber + ': +' + $convArr.Count + ' rows (unique so far: ' + $found.Count + ')')
            if ($convArr.Count -lt 100) { break }   # short page = last page
        }

        $windowEnd = $windowStart
    }

    if ($found.Count -ge $Cap) {
        Write-Warning ('Hit -MaxConversations cap (' + $Cap + '); older conversations in the range were not fetched. Raise -MaxConversations or narrow the date range.')
    }
    return ,$found
}

# ---------------------------------------------------------------------------
# Fetch every suggestion page for one conversation.
# Handles both nextUri-style and cursor ("after") style pagination.
# ---------------------------------------------------------------------------
function Get-GcConversationSuggestions {
    param(
        [Parameter(Mandatory = $true)][string]$ApiBase,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][string]$ConvId,
        [int]$Size = 200
    )

    $all = @()
    $baseUri = $ApiBase + '/api/v2/conversations/' + $ConvId + '/suggestions'
    $uri = $baseUri + '?pageSize=' + [string]$Size
    $page = 0
    $maxPages = 100   # hard safety cap

    while (($uri -ne '') -and ($page -lt $maxPages)) {
        $page++
        $resp = Invoke-GcApi -Uri $uri -Headers $Headers

        if ($null -eq $resp) { break }

        # Current payloads use 'entities'; some older shapes used 'suggestions'.
        $entities = Get-Prop -InputObject $resp -Name 'entities'
        if ($null -eq $entities) { $entities = Get-Prop -InputObject $resp -Name 'suggestions' }
        if ($null -ne $entities) {
            foreach ($e in @($entities)) {
                if ($null -ne $e) { $all += $e }
            }
        }

        # --- Work out the next page, if any ---
        $uri = ''

        $nextUri = Get-Prop -InputObject $resp -Name 'nextUri'
        if (($null -ne $nextUri) -and ([string]$nextUri -ne '')) {
            $n = [string]$nextUri
            if ($n -like 'http*') { $uri = $n } else { $uri = $ApiBase + $n }
            continue
        }

        $after = $null
        $cursors = Get-Prop -InputObject $resp -Name 'cursors'
        if ($null -ne $cursors) { $after = Get-Prop -InputObject $cursors -Name 'after' }
        if ($null -eq $after)   { $after = Get-Prop -InputObject $resp -Name 'after' }

        if (($null -ne $after) -and ([string]$after -ne '')) {
            # Minimal CLM-safe encoding of cursor token (no [uri]:: methods).
            $safeAfter = [string]$after
            $safeAfter = $safeAfter.Replace('%', '%25').Replace('+', '%2B').Replace('/', '%2F').Replace('=', '%3D').Replace('&', '%26')
            $uri = $baseUri + '?pageSize=' + [string]$Size + '&after=' + $safeAfter
        }
    }

    if ($page -ge $maxPages) {
        Write-Warning ('Conversation ' + $ConvId + ': stopped at safety cap of ' + $maxPages + ' pages.')
    }
    return ,$all
}

# ---------------------------------------------------------------------------
# Fetch Copilot session summaries for one conversation.
# Current shape: { sessionSummaries: [...] }; older shapes used entities /
# summaries arrays - all tolerated, same as the blueprint app.
# ---------------------------------------------------------------------------
function Get-GcConversationSummaries {
    param(
        [Parameter(Mandatory = $true)][string]$ApiBase,
        [Parameter(Mandatory = $true)][hashtable]$Headers,
        [Parameter(Mandatory = $true)][string]$ConvId
    )

    $resp = Invoke-GcApi -Uri ($ApiBase + '/api/v2/conversations/' + $ConvId + '/summaries') -Headers $Headers
    if ($null -eq $resp) { return ,@() }

    $entries = Get-Prop -InputObject $resp -Name 'sessionSummaries'
    if ($null -eq $entries) { $entries = Get-Prop -InputObject $resp -Name 'entities' }
    if ($null -eq $entries) { $entries = Get-Prop -InputObject $resp -Name 'summaries' }

    $all = @()
    if ($null -ne $entries) {
        foreach ($e in @($entries)) {
            if ($null -ne $e) { $all += $e }
        }
    }
    return ,$all
}

# ---------------------------------------------------------------------------
# Flatten one suggestion into a clean row + snippet child rows.
# Field layout follows the blueprint's SuggestionEntry model: current shape
# (knowledgeSearch payload + context) with the legacy fallbacks it also keeps.
# Returns a hashtable: @{ Row = <PSObject>; Snippets = <PSObject[]> }
# ---------------------------------------------------------------------------
function ConvertTo-SuggestionRow {
    param(
        [Parameter(Mandatory = $true)][string]$ConvId,
        [Parameter(Mandatory = $true)]$Suggestion,
        [Parameter(Mandatory = $true)][string]$RegionDomain
    )

    $suggId = [string](Get-Prop -InputObject $Suggestion -Name 'id')

    $ks      = Get-Prop -InputObject $Suggestion -Name 'knowledgeSearch'
    $context = Get-Prop -InputObject $Suggestion -Name 'context'

    # --- Title: current shape first, then every legacy container ---
    $title = [string](Get-FirstProp -InputObject $ks -Names @('title'))
    if ($title -eq '') {
        $ka = Get-Prop -InputObject $Suggestion -Name 'knowledgeArticle'
        $title = [string](Get-FirstProp -InputObject $ka -Names @('title'))
    }
    if ($title -eq '') {
        $sg = Get-Prop -InputObject $Suggestion -Name 'suggestion'
        $title = [string](Get-FirstProp -InputObject $sg -Names @('title'))
    }
    if ($title -eq '') {
        $cr = Get-Prop -InputObject $Suggestion -Name 'cannedResponse'
        $title = [string](Get-FirstProp -InputObject $cr -Names @('name'))
    }
    if ($title -eq '') {
        $sc = Get-Prop -InputObject $Suggestion -Name 'script'
        $title = [string](Get-FirstProp -InputObject $sc -Names @('name'))
    }
    if ($title -eq '') {
        $title = [string](Get-FirstProp -InputObject $Suggestion -Names @('title', 'name'))
    }

    # --- Extracted answer (the highlighted passage Copilot surfaced) ---
    $answerText = ''
    $kAnswer = Get-Prop -InputObject $ks -Name 'knowledgeAnswer'
    if ($null -ne $kAnswer) { $answerText = [string](Get-FirstProp -InputObject $kAnswer -Names @('answer')) }
    if ($answerText -eq '') {
        $ans = Get-Prop -InputObject $Suggestion -Name 'answer'
        $answerText = [string](Get-FirstProp -InputObject $ans -Names @('text'))
    }
    if ($answerText -eq '') {
        $answerText = [string](Get-FirstProp -InputObject $Suggestion -Names @('snippet', 'body'))
    }

    # --- Confidence: numeric 0..1 so Power BI can aggregate/format it ---
    $confidence = Get-Prop -InputObject $ks -Name 'confidence'
    if ($null -eq $confidence) { $confidence = Get-Prop -InputObject $Suggestion -Name 'confidence' }
    $confidenceStr = ''
    if ($null -ne $confidence) { $confidenceStr = [string]$confidence }

    # --- Knowledge document + Workbench deep-link (same URL scheme as the
    #     blueprint's getKnowledgeArticleUrl helper) ---
    $documentId      = ''
    $knowledgeBaseId = ''
    $articleUrl      = ''
    $doc = Get-Prop -InputObject $ks -Name 'document'
    if ($null -ne $doc) {
        $documentId = [string](Get-Prop -InputObject $doc -Name 'id')
        $selfUri    = [string](Get-Prop -InputObject $doc -Name 'selfUri')
        if ($selfUri -match '/knowledgebases/([^/]+)/documents/([^/?#]+)') {
            $knowledgeBaseId = [string]$Matches[1]
            $articleUrl = 'https://apps.' + $RegionDomain + '/directory/#/admin/knowledge/v2/knowledge-bases/' + $Matches[1] + '/articles/' + $Matches[2]
        }
    }

    # --- Context: where/for whom Copilot raised the suggestion ---
    $queueId = ''; $agentUserId = ''; $mediaType = ''; $externalContactId = ''
    if ($null -ne $context) {
        $q = Get-Prop -InputObject $context -Name 'queue'
        $queueId = [string](Get-FirstProp -InputObject $q -Names @('id'))
        $u = Get-Prop -InputObject $context -Name 'user'
        $agentUserId = [string](Get-FirstProp -InputObject $u -Names @('id'))
        $mediaType = [string](Get-FirstProp -InputObject $context -Names @('mediaType'))
        $ec = Get-Prop -InputObject $context -Name 'externalContact'
        $externalContactId = [string](Get-FirstProp -InputObject $ec -Names @('id'))
    }

    $row = New-Object PSObject -Property @{
        ConversationId    = $ConvId
        SuggestionId      = $suggId
        SuggestionType    = [string](Get-Prop -InputObject $Suggestion -Name 'type')
        State             = [string](Get-Prop -InputObject $Suggestion -Name 'state')
        TriggerType       = [string](Get-Prop -InputObject $Suggestion -Name 'triggerType')
        DateCreated       = [string](Get-FirstProp -InputObject $Suggestion -Names @('dateCreated', 'dateIssued'))
        Confidence        = $confidenceStr
        Title             = $title
        AnswerText        = $answerText
        DocumentId        = $documentId
        KnowledgeBaseId   = $knowledgeBaseId
        ArticleUrl        = $articleUrl
        SearchId          = [string](Get-FirstProp -InputObject $ks -Names @('searchId'))
        MediaType         = $mediaType
        QueueId           = $queueId
        AgentUserId       = $agentUserId
        ExternalContactId = $externalContactId
    }

    # --- Snippets: one child row each (a suggestion can carry several) ---
    $snippetRows = @()
    $snips = Get-Prop -InputObject $ks -Name 'snippets'
    if ($null -ne $snips) {
        $ix = 0
        foreach ($sn in @($snips)) {
            if ($null -eq $sn) { continue }
            $ix++
            $snippetRows += New-Object PSObject -Property @{
                ConversationId = $ConvId
                SuggestionId   = $suggId
                SnippetIndex   = $ix
                SnippetText    = [string]$sn
            }
        }
    }

    return @{ Row = $row; Snippets = $snippetRows }
}

# ---------------------------------------------------------------------------
# Flatten one Copilot session summary into a clean row.
# ---------------------------------------------------------------------------
function ConvertTo-SummaryRow {
    param(
        [Parameter(Mandatory = $true)][string]$ConvId,
        [Parameter(Mandatory = $true)]$Summary
    )

    $reason     = Get-Prop -InputObject $Summary -Name 'reason'
    $resolution = Get-Prop -InputObject $Summary -Name 'resolution'
    $followup   = Get-Prop -InputObject $Summary -Name 'followup'

    # Summary text: current shape has it at top level; legacy under .summary
    $text = [string](Get-FirstProp -InputObject $Summary -Names @('text'))
    if ($text -eq '') {
        $legacy = Get-Prop -InputObject $Summary -Name 'summary'
        $text = Get-CFText -Field $legacy
    }

    # Predicted wrap-up codes: join names for a single tidy column
    $wrapups = @()
    $pwc = Get-Prop -InputObject $Summary -Name 'predictedWrapupCodes'
    if ($null -ne $pwc) {
        foreach ($w in @($pwc)) {
            $wName = [string](Get-FirstProp -InputObject $w -Names @('name', 'id'))
            if ($wName -ne '') { $wrapups += $wName }
        }
    }
    $suggested = Get-Prop -InputObject $Summary -Name 'suggestedWrapUpCode'
    if ($null -ne $suggested) {
        $wName = [string](Get-FirstProp -InputObject $suggested -Names @('name'))
        if (($wName -ne '') -and (-not ($wrapups -contains $wName))) { $wrapups += $wName }
    }

    $confidence = Get-Prop -InputObject $Summary -Name 'confidence'
    $confidenceStr = ''
    if ($null -ne $confidence) { $confidenceStr = [string]$confidence }

    return New-Object PSObject -Property @{
        ConversationId        = $ConvId
        SummaryId             = [string](Get-Prop -InputObject $Summary -Name 'id')
        MediaType             = [string](Get-Prop -InputObject $Summary -Name 'mediaType')
        Language              = [string](Get-Prop -InputObject $Summary -Name 'language')
        Status                = [string](Get-Prop -InputObject $Summary -Name 'status')
        SummaryText           = $text
        Confidence            = $confidenceStr
        ReasonText            = (Get-CFText -Field $reason)
        ReasonDescription     = (Get-CFDetail -Field $reason -Name 'description')
        ResolutionText        = (Get-CFText -Field $resolution)
        ResolutionDescription = (Get-CFDetail -Field $resolution -Name 'description')
        ResolutionOutcome     = (Get-CFDetail -Field $resolution -Name 'outcome')
        FollowupText          = (Get-CFText -Field $followup)
        FollowupDescription   = (Get-CFDetail -Field $followup -Name 'description')
        PredictedWrapupCodes  = ($wrapups -join ';')
    }
}

# ===========================================================================
# MAIN
# ===========================================================================

if (($AccessToken -eq '') -and (($ClientId -eq '') -or ($ClientSecret -eq ''))) {
    throw 'Provide either -AccessToken, or both -ClientId and -ClientSecret.'
}

$explicitIds = (@($ConversationId).Count -gt 0)
if ((-not $explicitIds) -and ($DivisionId -eq '')) {
    throw 'Provide -DivisionId to discover conversations dynamically, or -ConversationId for explicit conversations.'
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if ($OutputFolder -eq '') { $OutputFolder = '.\GcCopilotReport_' + $timestamp }
$null = New-Item -ItemType Directory -Path $OutputFolder -Force

$token = $AccessToken
if ($token -eq '') {
    $token = Get-GcAccessToken -LoginBase $loginBase -Id $ClientId -Secret $ClientSecret
}

$apiHeaders = @{
    'Authorization' = 'Bearer ' + $token
    'Accept'        = 'application/json'
}

# --- Build the target conversation list (rich objects) ---------------------
$targets = @()
if ($explicitIds) {
    foreach ($convId in $ConversationId) {
        $cid = ([string]$convId).Trim()
        if ($cid -eq '') { continue }
        $targets += New-Object PSObject -Property @{
            ConversationId    = $cid
            ConversationStart = ''
            ConversationEnd   = ''
            MediaTypes        = ''
            QueueName         = ''
            CustomerName      = ''
        }
    }
}
else {
    Write-Host ''
    Write-Host ('Discovering conversations in division ' + $DivisionId + ' from ' + $StartDate.ToString('yyyy-MM-dd HH:mm') + ' to ' + $EndDate.ToString('yyyy-MM-dd HH:mm') + ' ...')
    $targets = Get-GcConversationsByDivision -ApiBase $apiBase -Headers $apiHeaders -DivId $DivisionId -From $StartDate -To $EndDate -Cap $MaxConversations
    Write-Host ('Conversations discovered: ' + $targets.Count)
    if ($targets.Count -eq 0) {
        Write-Host 'No conversations found in that division/date range - nothing to do.'
        return
    }
}

# --- Pull suggestions (+ summaries) per conversation ------------------------
$suggestionRows = @()
$snippetRows    = @()
$summaryRows    = @()
$rawSuggestions = @()
$rawSummaries   = @()
$suggCountByConv = @{}
$summCountByConv = @{}

$okCount     = 0
$skipCount   = 0
$noSuggCount = 0
$failCount   = 0
$idx         = 0

foreach ($t in $targets) {
    $cid = [string](Get-Prop -InputObject $t -Name 'ConversationId')
    $idx++
    Write-Host ('[' + $idx + '/' + $targets.Count + '] Conversation ' + $cid)

    try {
        # Division guard only applies to explicit IDs - discovered ones were
        # already filtered on divisionId by the analytics query itself.
        if ($explicitIds -and ($DivisionId -ne '')) {
            $inDivision = Test-GcConversationDivision -ApiBase $apiBase -Headers $apiHeaders -ConvId $cid -ExpectedDivisionId $DivisionId
            if (-not $inDivision) {
                $skipCount++
                continue
            }
        }

        # ---- Suggestions ----
        $suggestions = Get-GcConversationSuggestions -ApiBase $apiBase -Headers $apiHeaders -ConvId $cid -Size $PageSize
        $suggCountByConv[$cid] = $suggestions.Count
        if ($suggestions.Count -gt 0) {
            Write-Host ('    Suggestions: ' + $suggestions.Count)
        }

        foreach ($s in $suggestions) {
            $rawSuggestions += $s
            $flat = ConvertTo-SuggestionRow -ConvId $cid -Suggestion $s -RegionDomain $Region
            $suggestionRows += $flat.Row
            foreach ($snRow in $flat.Snippets) { $snippetRows += $snRow }
        }

        # ---- Summaries ----
        if (-not $SkipSummaries) {
            try {
                $summaries = Get-GcConversationSummaries -ApiBase $apiBase -Headers $apiHeaders -ConvId $cid
                $summCountByConv[$cid] = $summaries.Count
                if ($summaries.Count -gt 0) {
                    Write-Host ('    Summaries:   ' + $summaries.Count)
                }
                foreach ($sm in $summaries) {
                    $rawSummaries += $sm
                    $summaryRows += ConvertTo-SummaryRow -ConvId $cid -Summary $sm
                }
            }
            catch {
                $smsg = ''
                try { $smsg = [string]$_.Exception.Message } catch { $smsg = '' }
                # 404 = no Copilot summary for this conversation; anything else is worth seeing
                if (-not ($smsg -like '*404*')) {
                    Write-Warning ('    Summaries failed for ' + $cid + ': ' + $smsg)
                }
            }
        }

        $okCount++
    }
    catch {
        $emsg = ''
        try { $emsg = [string]$_.Exception.Message } catch { $emsg = 'unknown error' }

        # Many conversations never had Agent Copilot active; the suggestions
        # endpoint answers 404 for those. Treat as "none", not a failure.
        if ($emsg -like '*404*') {
            $noSuggCount++
            $suggCountByConv[$cid] = 0
        }
        else {
            $failCount++
            Write-Warning ('    Conversation ' + $cid + ' failed: ' + $emsg)
        }
    }
}

# --- Stamp per-conversation counts, then export the report -----------------
$conversationRows = @()
foreach ($t in $targets) {
    $cid = [string](Get-Prop -InputObject $t -Name 'ConversationId')
    $sCount = 0
    if ($suggCountByConv.ContainsKey($cid)) { $sCount = [int]$suggCountByConv[$cid] }
    $smCount = 0
    if ($summCountByConv.ContainsKey($cid)) { $smCount = [int]$summCountByConv[$cid] }

    $conversationRows += New-Object PSObject -Property @{
        ConversationId    = $cid
        ConversationStart = [string](Get-Prop -InputObject $t -Name 'ConversationStart')
        ConversationEnd   = [string](Get-Prop -InputObject $t -Name 'ConversationEnd')
        MediaTypes        = [string](Get-Prop -InputObject $t -Name 'MediaTypes')
        QueueName         = [string](Get-Prop -InputObject $t -Name 'QueueName')
        CustomerName      = [string](Get-Prop -InputObject $t -Name 'CustomerName')
        SuggestionCount   = $sCount
        SummaryCount      = $smCount
    }
}

Write-Host ''
Write-Host ('Done. Conversations OK: ' + $okCount + '  without suggestions (404): ' + $noSuggCount + '  skipped (division): ' + $skipCount + '  failed: ' + $failCount)
Write-Host ('Totals - suggestions: ' + $suggestionRows.Count + '  snippets: ' + $snippetRows.Count + '  summaries: ' + $summaryRows.Count)
Write-Host ''

# Explicit column order on every export (New-Object -Property hashtables do
# not preserve key order).
$convCsv = Join-Path -Path $OutputFolder -ChildPath 'Conversations.csv'
$conversationRows |
    Select-Object ConversationId, ConversationStart, ConversationEnd, MediaTypes,
                  QueueName, CustomerName, SuggestionCount, SummaryCount |
    Export-Csv -Path $convCsv -NoTypeInformation -Encoding UTF8
Write-Host ('Written: ' + $convCsv + '  (' + $conversationRows.Count + ' rows)')

$suggCsv = Join-Path -Path $OutputFolder -ChildPath 'Suggestions.csv'
if ($suggestionRows.Count -gt 0) {
    $suggestionRows |
        Select-Object ConversationId, SuggestionId, SuggestionType, State, TriggerType,
                      DateCreated, Confidence, Title, AnswerText, DocumentId,
                      KnowledgeBaseId, ArticleUrl, SearchId, MediaType, QueueId,
                      AgentUserId, ExternalContactId |
        Export-Csv -Path $suggCsv -NoTypeInformation -Encoding UTF8
    Write-Host ('Written: ' + $suggCsv + '  (' + $suggestionRows.Count + ' rows)')
} else {
    Write-Host 'No suggestions found - Suggestions.csv not written.'
}

if ($snippetRows.Count -gt 0) {
    $snipCsv = Join-Path -Path $OutputFolder -ChildPath 'SuggestionSnippets.csv'
    $snippetRows |
        Select-Object ConversationId, SuggestionId, SnippetIndex, SnippetText |
        Export-Csv -Path $snipCsv -NoTypeInformation -Encoding UTF8
    Write-Host ('Written: ' + $snipCsv + '  (' + $snippetRows.Count + ' rows)')
}

if ($summaryRows.Count -gt 0) {
    $summCsv = Join-Path -Path $OutputFolder -ChildPath 'Summaries.csv'
    $summaryRows |
        Select-Object ConversationId, SummaryId, MediaType, Language, Status,
                      SummaryText, Confidence, ReasonText, ReasonDescription,
                      ResolutionText, ResolutionDescription, ResolutionOutcome,
                      FollowupText, FollowupDescription, PredictedWrapupCodes |
        Export-Csv -Path $summCsv -NoTypeInformation -Encoding UTF8
    Write-Host ('Written: ' + $summCsv + '  (' + $summaryRows.Count + ' rows)')
}

if ($RawJson) {
    if ($rawSuggestions.Count -gt 0) {
        $rawSuggPath = Join-Path -Path $OutputFolder -ChildPath 'RawSuggestions.json'
        ConvertTo-Json -InputObject $rawSuggestions -Depth 15 | Out-File -FilePath $rawSuggPath -Encoding UTF8
        Write-Host ('Written: ' + $rawSuggPath)
    }
    if ($rawSummaries.Count -gt 0) {
        $rawSummPath = Join-Path -Path $OutputFolder -ChildPath 'RawSummaries.json'
        ConvertTo-Json -InputObject $rawSummaries -Depth 15 | Out-File -FilePath $rawSummPath -Encoding UTF8
        Write-Host ('Written: ' + $rawSummPath)
    }
}

Write-Host ''
Write-Host ('Report folder ready for Power BI: ' + $OutputFolder)
