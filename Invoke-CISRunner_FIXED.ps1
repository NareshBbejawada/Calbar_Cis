[CmdletBinding()]
param(
    [string]$RepositoryServer = 'LASQLCL3\ODYDEV',
    [string]$RepositoryDatabase = 'CALBAR_SQL_CIS',
    [string]$LogFolder = 'E:\CALBarCISGovernance\Runner\Logs'
)

$ErrorActionPreference = 'Stop'
Import-Module SqlServer -ErrorAction Stop
New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
$LogFile = Join-Path $LogFolder ("CISRunner_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $LogFile

function ConvertFrom-DbValue {
    param([AllowNull()]$Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $null }
    return $Value
}

function ConvertTo-SqlUnicodeLiteral {
    param([AllowNull()]$Value, [int]$MaximumLength = 4000)
    $Value = ConvertFrom-DbValue $Value
    if ($null -eq $Value) { return 'NULL' }
    $Text = [string]$Value
    if ($Text.Length -gt $MaximumLength) { $Text = $Text.Substring(0, $MaximumLength) }
    return "N'" + $Text.Replace("'", "''") + "'"
}

function Get-FirstRow {
    param([AllowNull()]$Result)
    if ($null -eq $Result) { return $null }
    if ($Result -is [System.Array]) { return $Result | Select-Object -First 1 }
    return $Result
}

$RunId = $null
try {
    $RunRow = Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
        -TrustServerCertificate -Query @"
INSERT audit.CIS_Run(StartedBy)
OUTPUT inserted.RunId
VALUES(SUSER_SNAME());
"@ -ErrorAction Stop

    $RunId = [long](Get-FirstRow $RunRow).RunId

    $Servers = @(Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
        -TrustServerCertificate -Query @"
SELECT ServerId, ServerInstance
FROM inv.SQL_Server_Inventory
WHERE IsEnabled = 1
ORDER BY ServerInstance;
"@ -ErrorAction Stop)

    $Controls = @(Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
        -TrustServerCertificate -Query @"
SELECT ControlId, ControlCode, ExpectedValue, CheckType, CheckQuery, MinimumMajorVersion
FROM cfg.CIS_Control_Definition
WHERE IsEnabled = 1
ORDER BY SortOrder;
"@ -ErrorAction Stop)

    $ResultCount = 0

    foreach ($Server in $Servers) {
        $ServerInstance = [string]$Server.ServerInstance
        Write-Host "Scanning $ServerInstance"

        $MajorVersion = $null
        $VersionError = $null
        try {
            $VersionRow = Get-FirstRow (Invoke-Sqlcmd -ServerInstance $ServerInstance -Database master `
                -TrustServerCertificate -Query @"
SELECT TRY_CONVERT(int, SERVERPROPERTY('ProductMajorVersion')) AS Major;
"@ -QueryTimeout 20 -ErrorAction Stop)

            $RawMajor = ConvertFrom-DbValue $VersionRow.Major
            if ($null -ne $RawMajor -and -not [string]::IsNullOrWhiteSpace([string]$RawMajor)) {
                $ParsedMajor = 0
                if ([int]::TryParse([string]$RawMajor, [ref]$ParsedMajor)) {
                    $MajorVersion = $ParsedMajor
                }
            }
        }
        catch {
            $VersionError = $_.Exception.Message
        }

        foreach ($Control in $Controls) {
            $Status = 'ERROR'
            $ActualValue = $null
            $ErrorMessage = $null

            try {
                if ($VersionError) {
                    throw "Unable to determine SQL Server version: $VersionError"
                }

                $MinimumMajorVersion = ConvertFrom-DbValue $Control.MinimumMajorVersion
                if ($null -ne $MinimumMajorVersion -and -not [string]::IsNullOrWhiteSpace([string]$MinimumMajorVersion)) {
                    $MinimumVersionNumber = 0
                    if (-not [int]::TryParse([string]$MinimumMajorVersion, [ref]$MinimumVersionNumber)) {
                        throw "Invalid MinimumMajorVersion '$MinimumMajorVersion' for control $($Control.ControlCode)."
                    }

                    if ($null -eq $MajorVersion) {
                        throw "SQL Server major version was NULL for $ServerInstance."
                    }

                    if ($MajorVersion -lt $MinimumVersionNumber) {
                        $Status = 'PASS'
                        $ActualValue = "Not applicable (SQL major version $MajorVersion)"
                    }
                }

                if ($Status -eq 'ERROR') {
                    $CheckQuery = ConvertFrom-DbValue $Control.CheckQuery
                    if ([string]::IsNullOrWhiteSpace([string]$CheckQuery)) {
                        throw "CheckQuery is blank for control $($Control.ControlCode)."
                    }

                    $CheckRow = Get-FirstRow (Invoke-Sqlcmd -ServerInstance $ServerInstance -Database master `
                        -TrustServerCertificate -Query ([string]$CheckQuery) -QueryTimeout 30 -ErrorAction Stop)

                    if ($null -eq $CheckRow) {
                        throw "The check returned no row."
                    }

                    $RawActualValue = ConvertFrom-DbValue $CheckRow.ActualValue
                    if ($null -eq $RawActualValue -or [string]::IsNullOrWhiteSpace([string]$RawActualValue)) {
                        throw "The check returned NULL or blank ActualValue."
                    }

                    $ActualValue = [string]$RawActualValue
                    $ExpectedValue = [string](ConvertFrom-DbValue $Control.ExpectedValue)
                    $CheckType = ([string](ConvertFrom-DbValue $Control.CheckType)).Trim().ToUpperInvariant()

                    switch ($CheckType) {
                        'MINIMUM' {
                            $ActualNumber = [decimal]0
                            $ExpectedNumber = [decimal]0
                            if (-not [decimal]::TryParse($ActualValue, [ref]$ActualNumber)) {
                                throw "ActualValue '$ActualValue' is not numeric."
                            }
                            if (-not [decimal]::TryParse($ExpectedValue, [ref]$ExpectedNumber)) {
                                throw "ExpectedValue '$ExpectedValue' is not numeric."
                            }
                            $Status = if ($ActualNumber -ge $ExpectedNumber) { 'PASS' } else { 'FAIL' }
                        }
                        'SET' {
                            $AllowedValues = @($ExpectedValue -split ',' | ForEach-Object { $_.Trim() })
                            $Status = if ($AllowedValues -contains $ActualValue.Trim()) { 'PASS' } else { 'FAIL' }
                        }
                        default {
                            $Status = if ($ActualValue.Trim() -eq $ExpectedValue.Trim()) { 'PASS' } else { 'FAIL' }
                        }
                    }
                }
            }
            catch {
                $Status = 'ERROR'
                $ErrorMessage = $_.Exception.Message
            }

            $ActualSql = ConvertTo-SqlUnicodeLiteral $ActualValue
            $ExpectedSql = ConvertTo-SqlUnicodeLiteral (ConvertFrom-DbValue $Control.ExpectedValue) 100
            $ErrorSql = ConvertTo-SqlUnicodeLiteral $ErrorMessage

            $InsertQuery = @"
INSERT audit.CIS_Compliance_Result
(
    RunId, ServerId, ControlId, Status,
    ActualValue, ExpectedValue, ErrorMessage
)
VALUES
(
    $RunId,
    $($Server.ServerId),
    $($Control.ControlId),
    '$Status',
    $ActualSql,
    $ExpectedSql,
    $ErrorSql
);
"@

            Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
                -TrustServerCertificate -Query $InsertQuery -ErrorAction Stop
            $ResultCount++
        }
    }

    $ServerCount = @($Servers).Count
    Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
        -TrustServerCertificate -Query @"
UPDATE audit.CIS_Run
SET CompletedAtUtc = SYSUTCDATETIME(),
    ServerCount = $ServerCount,
    ResultCount = $ResultCount,
    RunStatus = 'COMPLETED'
WHERE RunId = $RunId;
"@ -ErrorAction Stop

    Write-Host "Completed run $RunId with $ResultCount results across $ServerCount servers."
}
catch {
    if ($RunId) {
        $Failure = ConvertTo-SqlUnicodeLiteral $_.Exception.Message
        Invoke-Sqlcmd -ServerInstance $RepositoryServer -Database $RepositoryDatabase `
            -TrustServerCertificate -Query @"
UPDATE audit.CIS_Run
SET CompletedAtUtc = SYSUTCDATETIME(),
    RunStatus = 'FAILED'
WHERE RunId = $RunId;
"@ -ErrorAction SilentlyContinue
    }
    throw
}
finally {
    Stop-Transcript
}
