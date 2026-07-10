<#
.SYNOPSIS
    Retrieves Agent Copilot suggestions for a Genesys Cloud conversation:
        GET /api/v2/conversations/{conversationId}/suggestions

.DESCRIPTION
    Designed for Windows PowerShell 5.1 running in CONSTRAINED LANGUAGE MODE
    (AppLocker / WDAC locked-down endpoints). The entire script is CLM-safe:

      * No .NET static method calls  (no [Convert]::, [Text.Encoding]::, [uri]::, [Guid]::)
      * No ::new() constructors      (plain arrays with += instead)
      * No [pscustomobject]@{} casts (New-Object PSObject -Property @{} instead,
                                      with explicit Select-Object column ordering)
      * Base64 for the OAuth Basic header is implemented in pure PowerShell
        using bit operators (-shl / -shr / -band / -bor) over [int][char] values.
      * HTTP status codes are read via [int]$_.Exception.Response.StatusCode inside
        try/catch, with fallback string matching on $_.Exception.Message.
      * Every API response is validated explicitly before being used.

    Authentication uses the OAuth Client Credentials grant. The OAuth client's
    role must be assigned to the division that owns the conversation (the
    conversation IDs you pass come from a single division), and the role needs
    the Agent Copilot / suggestions view permission (e.g. assistants > suggestion
    > view, or conversation view depending on your org's permission model).

    TLS NOTE: CLM blocks setting [Net.ServicePointManager]::SecurityProtocol.
    On a current Windows 10/11 or Server 2019+ endpoint TLS 1.2 is negotiated by
    default. If you hit "Could not create SSL/TLS secure channel", enable strong
    crypto machine-wide via registry (SystemDefaultTlsVersions /
    SchUseStrongCrypto) - that is an admin/GPO change, not a script change.

.PARAMETER Region
    Genesys Cloud region domain (NOT the full URL). Defaults to the embedded
    value 'mypurecloud.com.au' (Australia / Sydney). Other examples:
      mypurecloud.com  mypurecloud.ie  mypurecloud.de  mypurecloud.jp
      usw2.pure.cloud  cac1.pure.cloud  euw2.pure.cloud  euc2.pure.cloud
      aps1.pure.cloud  apne2.pure.cloud  sae1.pure.cloud  mec1.pure.cloud

.PARAMETER ClientId
    OAuth client ID (Client Credentials grant). Defaults to the embedded
    $EmbeddedClientId value in the configuration block below.

.PARAMETER ClientSecret
    OAuth client secret. Defaults to the embedded $EmbeddedClientSecret value
    in the configuration block below.

.PARAMETER AccessToken
    Optional. Supply an existing bearer token to skip the token request
    (ClientId/ClientSecret are then ignored).

.PARAMETER ConversationId
    Optional. One or more explicit conversation IDs (all from the same
    division). When omitted, supply -DivisionId instead and the script
    discovers the conversation IDs dynamically.

.PARAMETER DivisionId
    The division ID to work with. Two behaviors:
      * -ConversationId OMITTED: conversation IDs are discovered dynamically
        for this division via POST /api/v2/analytics/conversations/details/query
        (filtered on the divisionId dimension, newest first, within
        -StartDate/-EndDate). Requires the analytics conversationDetail view
        permission on the OAuth client's role.
      * -ConversationId SUPPLIED: acts as a guard - each conversation is
        fetched from GET /api/v2/conversations/{id} and its division verified;
        mismatches are skipped with a warning.

.PARAMETER StartDate
    Discovery-mode only: start of the conversation search window
    (default: 7 days ago). Ranges longer than 7 days are automatically
    split into 7-day analytics queries.

.PARAMETER EndDate
    Discovery-mode only: end of the conversation search window (default: now).

.PARAMETER MaxConversations
    Discovery-mode only: safety cap on how many conversations are pulled from
    analytics before fetching suggestions (default 500, newest first).

.PARAMETER PageSize
    Suggestions page size (default 100).

.PARAMETER OutputCsv
    Path of the CSV export. Default: .\GcSuggestions_<timestamp>.csv

.PARAMETER OutputJson
    Optional path; when supplied the raw suggestion objects are also written
    as pretty-printed JSON for full-fidelity inspection.

.EXAMPLE
    # DISCOVERY MODE: give it a division ID, it finds the conversations itself
    # (last 7 days by default) and pulls suggestions for each.
    .\Get-GcConversationSuggestions.ps1 `
        -DivisionId '11111111-2222-3333-4444-555555555555'

.EXAMPLE
    # Discovery over a custom window with a bigger cap and both exports
    .\Get-GcConversationSuggestions.ps1 `
        -DivisionId '11111111-2222-3333-4444-555555555555' `
        -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) `
        -MaxConversations 2000 `
        -OutputCsv C:\Reports\suggestions.csv `
        -OutputJson C:\Reports\suggestions.json

.EXAMPLE
    # Explicit conversation ID (embedded region + credentials)
    .\Get-GcConversationSuggestions.ps1 `
        -ConversationId 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
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
    [int]$PageSize = 100,

    [Parameter(Mandatory = $false)]
    [string]$OutputCsv = '',

    [Parameter(Mandatory = $false)]
    [string]$OutputJson = ''
)

$ErrorActionPreference = 'Stop'

# ===========================================================================
# EMBEDDED CONFIGURATION - Australia (Sydney) region
# Paste your OAuth Client Credentials pair below. Command-line -ClientId /
# -ClientSecret / -Region parameters still work and override these values.
#
# SECURITY: anyone who can read this file can read these credentials. Restrict
# NTFS permissions on the script and scope the OAuth client's role to the
# minimum permissions (suggestions view) in the one division you query.
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
            elseif ($status -eq 404) { throw ('Resource not found (404). Check the conversation ID and region. URI: ' + $Uri) }
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
# Discovery mode: find conversation IDs in a division dynamically via
# POST /api/v2/analytics/conversations/details/query, filtered on the
# divisionId dimension, newest first. The analytics API caps a single query
# interval at 7 days and a page at 100 rows, so wider date ranges are split
# into 7-day windows and each window is paged until exhausted.
# ---------------------------------------------------------------------------
function Get-GcConversationIdsByDivision {
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

    $ids  = @()
    $seen = @{}   # de-dupe: a conversation can span analytics windows

    # Walk BACKWARD from -EndDate in 7-day windows so "newest first" holds
    # across windows, not just inside one - the cap then keeps the most
    # recent conversations.
    $windowEnd = $To
    while (($windowEnd -gt $From) -and ($ids.Count -lt $Cap)) {
        $windowStart = $windowEnd.AddDays(-7)
        if ($windowStart -lt $From) { $windowStart = $From }

        $interval = $windowStart.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z/' + `
                    $windowEnd.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z'
        Write-Host ('  Analytics window: ' + $interval)

        $pageNumber = 0
        $maxPagesPerWindow = 100
        while (($pageNumber -lt $maxPagesPerWindow) -and ($ids.Count -lt $Cap)) {
            $pageNumber++

            # divisionId is a CONVERSATION-level dimension, so the predicate
            # must live in conversationFilters - putting it in segmentFilters
            # makes the API reject the query with HTTP 400.
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
                if (($null -ne $cId) -and ([string]$cId -ne '')) {
                    $key = [string]$cId
                    if (-not $seen.ContainsKey($key)) {
                        $seen[$key] = $true
                        $ids += $key
                        if ($ids.Count -ge $Cap) { break }
                    }
                }
            }

            Write-Host ('    Page ' + $pageNumber + ': +' + $convArr.Count + ' rows (unique so far: ' + $ids.Count + ')')
            if ($convArr.Count -lt 100) { break }   # short page = last page
        }

        $windowEnd = $windowStart
    }

    if ($ids.Count -ge $Cap) {
        Write-Warning ('Hit -MaxConversations cap (' + $Cap + '); older conversations in the range were not fetched. Raise -MaxConversations or narrow the date range.')
    }
    return ,$ids
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
        [int]$Size = 100
    )

    $all = @()
    $baseUri = $ApiBase + '/api/v2/conversations/' + $ConvId + '/suggestions'
    $uri = $baseUri + '?pageSize=' + [string]$Size
    $page = 0
    $maxPages = 100   # hard safety cap

    while (($uri -ne '') -and ($page -lt $maxPages)) {
        $page++
        Write-Host ('  Page ' + $page + ' -> ' + $uri)
        $resp = Invoke-GcApi -Uri $uri -Headers $Headers

        if ($null -eq $resp) {
            Write-Warning ('Conversation ' + $ConvId + ': empty response on page ' + $page + '.')
            break
        }

        $entities = Get-Prop -InputObject $resp -Name 'entities'
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
# Flatten one suggestion entity into a CSV-friendly row.
# The Suggestion model nests its payload under a type-specific container
# (knowledgeArticle / knowledgeAnswer / cannedResponse / script), so each
# container is probed defensively - unknown future types still land in RawJson.
# ---------------------------------------------------------------------------
function ConvertTo-SuggestionRow {
    param(
        [Parameter(Mandatory = $true)][string]$ConvId,
        [Parameter(Mandatory = $true)]$Suggestion,
        [Parameter(Mandatory = $true)][string]$RetrievedAtUtc
    )

    $resourceId      = ''
    $resourceTitle   = ''
    $knowledgeBaseId = ''

    foreach ($containerName in @('knowledgeArticle', 'knowledgeAnswer', 'knowledgeSearch', 'cannedResponse', 'script')) {
        $container = Get-Prop -InputObject $Suggestion -Name $containerName
        if ($null -eq $container) { continue }

        $v = Get-Prop -InputObject $container -Name 'id'
        if (($null -ne $v) -and ($resourceId -eq '')) { $resourceId = [string]$v }

        foreach ($titleProp in @('title', 'name')) {
            $v = Get-Prop -InputObject $container -Name $titleProp
            if (($null -ne $v) -and ($resourceTitle -eq '')) { $resourceTitle = [string]$v }
        }

        $kb = Get-Prop -InputObject $container -Name 'knowledgeBase'
        $v  = Get-Prop -InputObject $kb -Name 'id'
        if (($null -ne $v) -and ($knowledgeBaseId -eq '')) { $knowledgeBaseId = [string]$v }

        # Some payloads nest the document one level deeper (e.g. .article / .document)
        foreach ($innerName in @('article', 'document')) {
            $inner = Get-Prop -InputObject $container -Name $innerName
            if ($null -eq $inner) { continue }
            $v = Get-Prop -InputObject $inner -Name 'id'
            if (($null -ne $v) -and ($resourceId -eq '')) { $resourceId = [string]$v }
            $v = Get-Prop -InputObject $inner -Name 'title'
            if (($null -ne $v) -and ($resourceTitle -eq '')) { $resourceTitle = [string]$v }
        }
    }

    $confidence = Get-Prop -InputObject $Suggestion -Name 'confidence'
    $confidenceStr = ''
    if ($null -ne $confidence) { $confidenceStr = [string]$confidence }

    # CLM-safe object creation (no [pscustomobject] cast). Column order is NOT
    # preserved by -Property hashtables; Select-Object downstream fixes that.
    return New-Object PSObject -Property @{
        ConversationId  = $ConvId
        SuggestionId    = [string](Get-Prop -InputObject $Suggestion -Name 'id')
        SuggestionType  = [string](Get-Prop -InputObject $Suggestion -Name 'type')
        State           = [string](Get-Prop -InputObject $Suggestion -Name 'state')
        DateIssued      = [string](Get-Prop -InputObject $Suggestion -Name 'dateIssued')
        Confidence      = $confidenceStr
        ResourceId      = $resourceId
        ResourceTitle   = $resourceTitle
        KnowledgeBaseId = $knowledgeBaseId
        RetrievedAtUtc  = $RetrievedAtUtc
        RawJson         = (ConvertTo-Json -InputObject $Suggestion -Depth 15 -Compress)
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
if ($OutputCsv -eq '') { $OutputCsv = '.\GcSuggestions_' + $timestamp + '.csv' }
$retrievedAtUtc = [string](Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

$token = $AccessToken
if ($token -eq '') {
    $token = Get-GcAccessToken -LoginBase $loginBase -Id $ClientId -Secret $ClientSecret
}

$apiHeaders = @{
    'Authorization' = 'Bearer ' + $token
    'Accept'        = 'application/json'
}

# --- Build the target conversation list -----------------------------------
$targetIds = @()
if ($explicitIds) {
    foreach ($convId in $ConversationId) {
        $cid = ([string]$convId).Trim()
        if ($cid -ne '') { $targetIds += $cid }
    }
}
else {
    Write-Host ''
    Write-Host ('Discovering conversations in division ' + $DivisionId + ' from ' + $StartDate.ToString('yyyy-MM-dd HH:mm') + ' to ' + $EndDate.ToString('yyyy-MM-dd HH:mm') + ' ...')
    $targetIds = Get-GcConversationIdsByDivision -ApiBase $apiBase -Headers $apiHeaders -DivId $DivisionId -From $StartDate -To $EndDate -Cap $MaxConversations
    Write-Host ('Conversations discovered: ' + $targetIds.Count)
    if ($targetIds.Count -eq 0) {
        Write-Host 'No conversations found in that division/date range - nothing to do.'
        return
    }
}

# --- Pull suggestions per conversation -------------------------------------
$rows           = @()
$rawSuggestions = @()
$okCount        = 0
$skipCount      = 0
$noSuggCount    = 0
$failCount      = 0

foreach ($cid in $targetIds) {
    Write-Host ''
    Write-Host ('Conversation: ' + $cid)

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

        $suggestions = Get-GcConversationSuggestions -ApiBase $apiBase -Headers $apiHeaders -ConvId $cid -Size $PageSize
        Write-Host ('  Suggestions found: ' + $suggestions.Count)

        foreach ($s in $suggestions) {
            $rawSuggestions += $s
            $rows += ConvertTo-SuggestionRow -ConvId $cid -Suggestion $s -RetrievedAtUtc $retrievedAtUtc
        }
        $okCount++
    }
    catch {
        $emsg = ''
        try { $emsg = [string]$_.Exception.Message } catch { $emsg = 'unknown error' }

        # In discovery mode many conversations never had Agent Copilot active;
        # the suggestions endpoint answers 404 for those. Treat as "none".
        if ($emsg -like '*404*') {
            $noSuggCount++
            Write-Host '  No suggestions available for this conversation (404).'
        }
        else {
            $failCount++
            Write-Warning ('Conversation ' + $cid + ' failed: ' + $emsg)
        }
    }
}

Write-Host ''
Write-Host ('Done. Conversations OK: ' + $okCount + '  without suggestions (404): ' + $noSuggCount + '  skipped (division): ' + $skipCount + '  failed: ' + $failCount + '  total suggestions: ' + $rows.Count)

if ($rows.Count -gt 0) {
    # Explicit column order (New-Object -Property hashtables do not preserve it).
    $rows |
        Select-Object ConversationId, SuggestionId, SuggestionType, State, DateIssued,
                      Confidence, ResourceId, ResourceTitle, KnowledgeBaseId,
                      RetrievedAtUtc, RawJson |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host ('CSV written: ' + $OutputCsv)

    if ($OutputJson -ne '') {
        ConvertTo-Json -InputObject $rawSuggestions -Depth 15 | Out-File -FilePath $OutputJson -Encoding UTF8
        Write-Host ('Raw JSON written: ' + $OutputJson)
    }
}
else {
    Write-Host 'No suggestions returned - nothing exported.'
}
