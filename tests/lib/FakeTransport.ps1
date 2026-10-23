# Replaces the network seam and environment functions of powerbi2csv.ps1 with fixture-backed fakes.
# Dot-source AFTER powerbi2csv.ps1 so that these definitions win. Same idea as azure2csv tests\lib\FakeTransport.ps1.

$global:TWRequests = [System.Collections.Generic.List[object]]::new()
$global:TWUnrouted = [System.Collections.Generic.List[string]]::new()
$global:TWSleeps = [System.Collections.Generic.List[double]]::new()
$global:TWImportExcel = $false

function New-TWResponse {
    param($Obj, [int] $Status = 200, $Headers = $null)
    if ($null -eq $Headers) { $Headers = @{} }
    $content = $Obj
    if (-not ($Obj -is [string])) { $content = ConvertTo-Json -InputObject $Obj -Depth 40 -Compress }
    return @{ StatusCode = $Status; Content = $content; Headers = $Headers }
}

function New-TWError {
    param([int] $Status, [string] $Code, [string] $Message)
    return (New-TWResponse -Obj ([ordered]@{ error = [ordered]@{ code = $Code; message = $Message } }) -Status $Status)
}

function Resolve-TWRoute {
    param($Routes, [string] $Method, [string] $Path, [string] $Body, [string] $Kind)
    foreach ($route in $Routes) {
        if ($Method -eq $route.Method -and $Path -match $route.Pattern) {
            $m = $Matches
            $res = & $route.Handler $Path $Body $m
            if ($res -is [hashtable] -and $res.ContainsKey('StatusCode')) { return $res }
            return (New-TWResponse -Obj $res)
        }
    }
    $global:TWUnrouted.Add(('{0} {1} {2}' -f $Kind, $Method, $Path))
    return (New-TWError 404 'NoFixture' ('No {0} fixture for {1} {2}' -f $Kind, $Method, $Path))
}

function Invoke-A2CHttp {
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, [string] $Body, [string] $Kind = 'pbi', [string] $OutFile)
    $global:TWRequests.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body; Kind = $Kind })
    foreach ($t in $global:TW.Throttle) {
        if ($Uri -match $t.Pattern -and $t.Remaining -gt 0) {
            $t.Remaining--
            return @{ StatusCode = 429; Content = '{"error":{"code":"TooManyRequests","message":"throttled"}}'; Headers = @{ 'Retry-After' = [string]$t.RetryAfter } }
        }
    }
    switch ($Kind) {
        'pbi' {
            $path = $Uri -replace '^https://api\.powerbi\.com/v1\.0/myorg', ''
            if ($path -like '/admin/*' -and -not $global:TW.Admin) { return (New-TWError 401 'PowerBINotAuthorizedException' 'Unauthorized') }
            $res = Resolve-TWRoute -Routes $global:TW.PbiRoutes -Method $Method -Path $path -Body $Body -Kind 'pbi'
            # downloads: a route answers with File = text and the file is written like Invoke-WebRequest -OutFile
            if ($OutFile -and $res.StatusCode -eq 200 -and $res.ContainsKey('File')) { [IO.File]::WriteAllText($OutFile, [string]$res.File); return @{ StatusCode = 200; Content = ''; Headers = @{} } }
            return $res
        }
        'fabric' {
            $path = $Uri -replace '^https://api\.fabric\.microsoft\.com/v1', ''
            if ($path -like '/admin/*' -and -not $global:TW.Admin) {
                return (New-TWResponse -Obj ([ordered]@{ errorCode = 'InsufficientPrivileges'; message = 'The caller does not have sufficient permissions' }) -Status 403)
            }
            return (Resolve-TWRoute -Routes $global:TW.FabricRoutes -Method $Method -Path $path -Body $Body -Kind 'fabric')
        }
    }
    if ($Kind -eq 'onelake') { return (Resolve-TWRoute -Routes $global:TW.OneLakeRoutes -Method $Method -Path ($Uri -replace '^https://onelake\.dfs\.fabric\.microsoft\.com', '') -Body $Body -Kind 'onelake') }
    if ($Kind -eq 'flow') { return (Resolve-TWRoute -Routes $global:TW.FlowRoutes -Method $Method -Path ($Uri -replace '^https://api\.flow\.microsoft\.com', '') -Body $Body -Kind 'flow') }
    return @{ StatusCode = 0; Content = 'Unknown API kind: ' + $Kind; Headers = @{} }
}

function Invoke-A2CSqlQuery {
    # TW.Sql: database -> rows of the "objects" query, or a text = the connection error.
    # TW.SqlSets: database -> @{ columns / modules / references = rows }; a text there = that query is refused.
    param([string] $Server, [string] $Database, [System.Collections.IDictionary] $Queries, [int] $TimeoutSec = 300)
    $global:TWRequests.Add([pscustomobject]@{ Method = 'SQL'; Uri = $Server + '/' + $Database; Body = (@($Queries.Keys) -join ','); Kind = 'sql' })
    $script:A2C.Stats['sql'] = [int]$script:A2C.Stats['sql'] + 1
    $x = $global:TW.Sql[$Database]
    if ($x -is [string]) { return [pscustomobject]@{ Ok = $false; Error = $x; Sets = @{}; SetErrors = @{} } }
    if ($null -eq $x) { return [pscustomobject]@{ Ok = $false; Error = 'Cannot open database "' + $Database + '"'; Sets = @{}; SetErrors = @{} } }
    $sets = @{}; $errs = @{}
    foreach ($q in $Queries.Keys) {
        $rows = @()
        if ($q -eq 'objects') { $rows = @($x) }
        elseif ($null -ne $global:TW.SqlSets[$Database] -and $global:TW.SqlSets[$Database].ContainsKey($q)) {
            $v = $global:TW.SqlSets[$Database][$q]
            if ($v -is [string]) { $errs[$q] = $v } else { $rows = @($v) }
        }
        $sets[$q] = $rows
    }
    return [pscustomobject]@{ Ok = $true; Error = ''; Sets = $sets; SetErrors = $errs }
}

function Start-A2CSleep {
    param([double] $Seconds)
    $global:TWSleeps.Add($Seconds)
    if ($null -ne $script:A2C) { $script:A2C.Stats.SleepSec += $Seconds }
}

function Get-A2CAzContext { return $global:TW.Context }
function Connect-A2CAzure {
    param($Options)
    return $global:TW.Context
}
function Assert-A2CPrerequisites { return '9.9.9-test' }
function Test-A2CCanPrompt { return $false }
function Invoke-A2CConnectAzAccount { param([hashtable] $Params) throw 'No sign-in in tests.' }
function Test-A2CModuleAvailable { param([string] $Name) if ($Name -eq 'ImportExcel') { return [bool]$global:TWImportExcel }; return $false }
function Get-A2CBearerToken { param([string] $Resource) return 'eyJ0eXAiOiJKV1QifQ.fake-entra-token.' + $Resource }

function Reset-TWState {
    $global:TW = New-TWWorld
    $global:TWRequests.Clear()
    $global:TWUnrouted.Clear()
    $global:TWSleeps.Clear()
}

function Invoke-TWRun {
    # Runs the full pipeline against the fake tenant; returns the run context.
    param([hashtable] $Params, [switch] $KeepWorld)
    if (-not $KeepWorld) { Reset-TWState }
    $p = @{ OutputPath = $global:TWOut; NoExcel = $true }
    foreach ($k in $Params.Keys) { $p[$k] = $Params[$k] }
    $opt = New-A2COptions -Params $p
    $opt.Now = [DateTime]::SpecifyKind([DateTime]'2026-09-29T00:00:00', [DateTimeKind]::Utc)
    Invoke-A2CRun -Options $opt 6>$null | Out-Null
    return $script:A2C
}

function Import-TWCsv {
    param($Run, [string] $File)
    $path = [IO.Path]::Combine($Run.OutDir, $File)
    if (-not (Test-Path -LiteralPath $path)) { throw ('File not written: ' + $File) }
    return @(Import-Csv -LiteralPath $path -Encoding UTF8)
}
