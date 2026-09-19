#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ServerListPath = (Join-Path $PSScriptRoot 'servers.txt'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot 'Reports'),
    [ValidateRange(1,99)][decimal]$WarningThreshold = 20,
    [ValidateRange(0,98)][decimal]$CriticalThreshold = 10,
    [ValidateSet('Auto','WSMan','DCOM')][string]$ConnectionProtocol = 'Auto',
    [ValidateRange(5,600)][int]$OperationTimeoutSec = 30,
    [System.Management.Automation.PSCredential]$CimCredential,
    [switch]$TestEmailOnly,
    [switch]$NoEmail,
    [switch]$OnlySendOnAlert,
    [switch]$OpenReport
)

# Database Mail report. Run with Windows PowerShell 5.1.
# Keep the script parameter block above all executable statements.
# Configure the existing Database Mail instance, profile, and recipients below.
# No SMTP configuration or credentials are required in this script.

# ========================= EDIT THESE SETTINGS =========================
$Config = @{
    ReportTitle               = 'SQL Server Drive Capacity Report'
    EnvironmentName           = 'SQL Server Estate'
    TeamName                  = 'Database Administration'

    # The ONE SQL instance with an existing, working Database Mail profile.
    # Default instance: SQLMAIL01.yourcompany.com
    # Named instance:   SQLMAIL01.yourcompany.com\DBAPPS1
    # Explicit TCP:     tcp:SQLMAIL01.yourcompany.com,49285
    # Use the host name covered by that SQL instance's trusted certificate.
    MailSqlInstance           = 'SQLMAIL01.yourcompany.com\DBAPPS1'
    MailProfileName           = 'DBA_Mail_Profile'
    MailTo                    = @('DBA-Team@yourcompany.com')
    MailCc                    = @()

    # Start with HTML only. All four CSV files are still saved locally.
    # Enable only after confirming SQL-host locality, file access, and limits.
    AttachCsvFiles            = $false

    SqlConnectTimeoutSeconds  = 15
    SqlCommandTimeoutSeconds  = 60

    # 0 = all report rows in email. Set a positive cap to reduce email size;
    # capped sections are labeled. Local HTML and CSV always include all rows.
    MaxInlineRowsPerSection   = 0
}
# No SMTP server, SMTP password, or MailFrom setting is needed here.
# The EXISTING Database Mail profile controls sender / SMTP / TLS settings.
# This script does not change SQL Server, Database Mail, or certificate trust.
# ======================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:LogPath = $null

function Write-RunLog {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f [DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz'), $Level, $Message
    Write-Host $line
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function ConvertTo-HtmlText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Format-ReportNumber {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or $Value -is [DBNull]) { return 'N/A' }
    return ([decimal]$Value).ToString('N2', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Get-PropertyValue {
    param([AllowNull()][object]$InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Get-ReferenceKey {
    param([AllowNull()][object]$Reference, [string]$Key)
    $value = Get-PropertyValue -InputObject $Reference -Name $Key
    if ($null -ne $value) { return [string]$value }
    # Some providers serialize WMI reference properties as object-path strings.
    $pattern = '(?i)\b' + [regex]::Escape($Key) + '\s*=\s*"((?:\\.|[^"])*)"'
    $match = [regex]::Match([string]$Reference, $pattern)
    if ($match.Success) {
        return $match.Groups[1].Value.Replace('\"','"').Replace('\\','\')
    }
    return ''
}

function Read-ServerNames {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Server list not found: $Path"
    }
    $names = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    $lineNumber = 0
    foreach ($line in Get-Content -LiteralPath $Path) {
        $lineNumber++
        $text = ([string]$line -replace '#.*$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { continue }
        $text = $text -replace '^(?i)tcp:', ''
        # Host\Instance and Host,Port are normalized to the Windows host.
        $computer = (($text -split '\\', 2)[0] -split ',', 2)[0].Trim().TrimEnd('.')
        if ($computer -notmatch '^[A-Za-z0-9_][A-Za-z0-9_.-]*$') {
            throw "Invalid host at line ${lineNumber}: '$text'. Use a Windows computer name or FQDN."
        }
        if (-not $seen.ContainsKey($computer)) {
            $seen[$computer] = $true
            $names.Add($computer)
        }
    }
    if ($names.Count -eq 0) { throw 'The server list contains no valid host names.' }
    return $names.ToArray()
}

function New-CollectionIssue {
    param([string]$Server, [string]$Stage, [string]$Volume, [string]$Details)
    [pscustomobject][ordered]@{
        Server      = $Server
        Stage       = $Stage
        Volume      = $Volume
        Details     = $Details
        ObservedAt  = [DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz')
    }
}

function Get-ServerVolumes {
    param([string]$Computer)
    $rows = [System.Collections.Generic.List[object]]::new()
    $issues = [System.Collections.Generic.List[object]]::new()
    $attempts = [System.Collections.Generic.List[string]]::new()
    $session = $null
    $transport = ''
    $volumes = @()
    $protocols = @($ConnectionProtocol)
    if ($ConnectionProtocol -eq 'Auto') { $protocols = @('WSMan','DCOM') }

    foreach ($protocol in $protocols) {
        try {
            $sessionParameters = @{
                ComputerName        = $Computer
                SessionOption       = (New-CimSessionOption -Protocol $protocol)
                OperationTimeoutSec = $OperationTimeoutSec
                ErrorAction         = 'Stop'
            }
            if ($null -ne $CimCredential) { $sessionParameters.Credential = $CimCredential }
            $session = New-CimSession @sessionParameters
            $volumes = @(Get-CimInstance -CimSession $session -Namespace 'root/cimv2' `
                -ClassName Win32_Volume -Filter 'DriveType = 3' `
                -Property DeviceID,DriveLetter,Name,Label,FileSystem,Capacity,FreeSpace,Availability `
                -OperationTimeoutSec $OperationTimeoutSec -ErrorAction Stop)
            if ($volumes.Count -eq 0) { throw 'The provider returned no local fixed volumes; coverage cannot be confirmed.' }
            $transport = $protocol
            break
        }
        catch {
            $attempts.Add(('{0}: {1}' -f $protocol, $_.Exception.Message))
            if ($null -ne $session) {
                Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
                $session = $null
            }
        }
    }

    if ($null -eq $session) {
        $issues.Add((New-CollectionIssue -Server $Computer -Stage 'Connection / volume query' `
            -Volume '' -Details ($attempts.ToArray() -join ' | ')))
        return [pscustomobject]@{
            Rows = $rows.ToArray(); Issues = $issues.ToArray()
            Coverage = [pscustomobject][ordered]@{
                Server = $Computer; CollectionStatus = 'FAILED'; Transport = ''
                VolumeCount = 0; Notes = ($attempts.ToArray() -join ' | ')
            }
        }
    }

    try {
        $mountMap = @{}
        try {
            $mounts = @(Get-CimInstance -CimSession $session -Namespace 'root/cimv2' `
                -ClassName Win32_MountPoint -OperationTimeoutSec $OperationTimeoutSec -ErrorAction Stop)
            foreach ($mount in $mounts) {
                $volumeId = Get-ReferenceKey -Reference $mount.Volume -Key 'DeviceID'
                $directory = Get-ReferenceKey -Reference $mount.Directory -Key 'Name'
                if ([string]::IsNullOrWhiteSpace($volumeId) -or [string]::IsNullOrWhiteSpace($directory)) {
                    throw 'A mount-point reference could not be decoded. Some paths may be shown as volume GUIDs.'
                }
                if (-not $mountMap.ContainsKey($volumeId)) {
                    $mountMap[$volumeId] = [System.Collections.Generic.List[string]]::new()
                }
                $mountMap[$volumeId].Add($directory.TrimEnd('\') + '\')
            }
        }
        catch {
            $issues.Add((New-CollectionIssue -Server $Computer -Stage 'Mount-point lookup' -Volume '' `
                -Details $_.Exception.Message))
        }

        $seenVolumes = @{}
        foreach ($volume in $volumes) {
            $deviceId = [string](Get-PropertyValue $volume 'DeviceID')
            $name = [string](Get-PropertyValue $volume 'Name')
            $letter = [string](Get-PropertyValue $volume 'DriveLetter')
            $paths = [System.Collections.Generic.List[string]]::new()
            if ($letter) { $paths.Add($letter.TrimEnd('\') + '\') }
            if ($deviceId -and $mountMap.ContainsKey($deviceId)) {
                foreach ($directory in $mountMap[$deviceId]) { $paths.Add($directory) }
            }
            if ($name -and $name -match '^[A-Za-z]:\\') { $paths.Add($name.TrimEnd('\') + '\') }
            $pathArray = @($paths.ToArray() | Sort-Object -Unique)
            $displayPath = $pathArray -join '; '
            if (-not $displayPath) { $displayPath = $name }
            if (-not $displayPath) { $displayPath = $deviceId }
            if (-not $displayPath) { $displayPath = '(unidentified volume)' }
            # One row per volume on a host, not per mount path.
            if ($deviceId -and $seenVolumes.ContainsKey($deviceId)) { continue }
            if ($deviceId) { $seenVolumes[$deviceId] = $true }

            $status = 'UNKNOWN'
            $totalGB = $null; $usedGB = $null; $freeGB = $null
            $freePercent = $null; $usedPercent = $null; $reclaimGB = $null
            $capacityBytes = $null; $freeBytes = $null
            $notes = ''
            try {
                $rawCapacity = Get-PropertyValue $volume 'Capacity'
                $rawFree = Get-PropertyValue $volume 'FreeSpace'
                if ($null -eq $rawCapacity -or $rawCapacity -is [DBNull] -or
                    $null -eq $rawFree -or $rawFree -is [DBNull]) {
                    throw 'Capacity or free-space measurement is unavailable (null).'
                }
                $capacityBytes = [decimal]$rawCapacity
                $freeBytes = [decimal]$rawFree
                if ($capacityBytes -le 0 -or $freeBytes -lt 0 -or $freeBytes -gt $capacityBytes) {
                    throw 'The provider returned invalid capacity/free-space measurements.'
                }
                if ((Get-PropertyValue $volume 'Availability') -eq 8) {
                    throw 'The provider reports this volume as offline; free space is not considered verified.'
                }
                # Evaluate thresholds using full precision, BEFORE display rounding.
                $freePercent = ($freeBytes / $capacityBytes) * 100
                $usedPercent = 100 - $freePercent
                $totalGB = $capacityBytes / 1GB
                $freeGB = $freeBytes / 1GB
                $usedGB = ($capacityBytes - $freeBytes) / 1GB
                $shortfall = [Math]::Max([decimal]0, (($WarningThreshold / 100) * $capacityBytes) - $freeBytes)
                $reclaimGB = [Math]::Ceiling(($shortfall / 1GB) * 100) / 100
                if ($freePercent -lt $CriticalThreshold) { $status = 'CRITICAL' }
                elseif ($freePercent -lt $WarningThreshold) { $status = 'WARNING' }
                else { $status = 'HEALTHY' }
            }
            catch {
                $notes = $_.Exception.Message
                $issues.Add((New-CollectionIssue -Server $Computer -Stage 'Volume measurement' `
                    -Volume $displayPath -Details $notes))
            }

            $rows.Add([pscustomobject][ordered]@{
                Server                 = $Computer
                DriveOrMountPoint      = $displayPath
                VolumeLabel            = [string](Get-PropertyValue $volume 'Label')
                FileSystem             = [string](Get-PropertyValue $volume 'FileSystem')
                TotalGB                = $totalGB
                UsedGB                 = $usedGB
                FreeGB                 = $freeGB
                FreePercent            = $freePercent
                UsedPercent            = $usedPercent
                Status                 = $status
                ReclaimToThresholdGB   = $reclaimGB
                ThresholdPercent       = $WarningThreshold
                CapacityBytes          = $capacityBytes
                FreeBytes              = $freeBytes
                VolumeId               = $deviceId
                Transport              = $transport
                ObservedAt             = [DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm:ss zzz')
                Notes                  = $notes
            })
        }
    }
    catch {
        $issues.Add((New-CollectionIssue -Server $Computer -Stage 'Unexpected collection failure' `
            -Volume '' -Details $_.Exception.Message))
    }
    finally {
        Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
    }
    $collectionStatus = 'COMPLETE'
    if ($issues.Count -gt 0) { $collectionStatus = 'PARTIAL' }
    if ($rows.Count -eq 0) { $collectionStatus = 'FAILED' }
    [pscustomobject]@{
        Rows = $rows.ToArray(); Issues = $issues.ToArray()
        Coverage = [pscustomobject][ordered]@{
            Server = $Computer; CollectionStatus = $collectionStatus; Transport = $transport
            VolumeCount = $rows.Count; Notes = ($attempts.ToArray() -join ' | ')
        }
    }
}

function Export-ReportCsv {
    param([AllowEmptyCollection()][object[]]$Rows, [string[]]$Columns, [string]$Path)
    $safeRows = @(
        foreach ($row in $Rows) {
            $copy = [ordered]@{}
            foreach ($column in $Columns) {
                $value = Get-PropertyValue -InputObject $row -Name $column
                # Prevent provider-supplied text from becoming an Excel formula.
                if ($value -is [string] -and $value -match '^[\s]*[=+@-]') { $value = "'" + $value }
                $copy[$column] = $value
            }
            [pscustomobject]$copy
        }
    )
    if ($safeRows.Count -gt 0) {
        $safeRows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }
    else {
        $empty = [ordered]@{}
        foreach ($column in $Columns) { $empty[$column] = $null }
        $header = @([pscustomobject]$empty | ConvertTo-Csv -NoTypeInformation)[0]
        Set-Content -LiteralPath $Path -Value $header -Encoding UTF8
    }
}

function New-VolumeTableHtml {
    param([AllowEmptyCollection()][object[]]$Rows, [switch]$ActionTable, [int]$Limit = 0)
    if ($Rows.Count -eq 0) {
        return '<p style="padding:14px;background:#f1f5f9;color:#475569;">No rows in this section. Check collection coverage before concluding that the estate is healthy.</p>'
    }
    $displayRows = @($Rows)
    if ($Limit -gt 0) { $displayRows = @($Rows | Select-Object -First $Limit) }
    $builder = [Text.StringBuilder]::new()
    $th = 'padding:11px 9px;text-align:left;font-size:11px;line-height:16px;color:#fff;background:#20384f;border:1px solid #20384f;'
    $td = 'padding:11px 9px;border-bottom:1px solid #dce4eb;vertical-align:top;font-size:12px;line-height:18px;'
    [void]$builder.Append('<table width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><thead><tr>')
    foreach ($heading in @('Server','Drive / mount point','Volume label','Total GB','Used GB','Free GB','Free %','Status')) {
        [void]$builder.Append('<th style="' + $th + '">' + $heading + '</th>')
    }
    if ($ActionTable) {
        [void]$builder.Append('<th style="' + $th + '">Reclaim to ' + (ConvertTo-HtmlText $WarningThreshold) + '% (GB)</th>')
    }
    [void]$builder.Append('</tr></thead><tbody>')
    foreach ($row in $displayRows) {
        $background = '#ffffff'; $ink = '#17603a'; $badge = '#e7f4ec'
        switch ($row.Status) {
            'CRITICAL' { $background = '#fff2f2'; $ink = '#a61b29'; $badge = '#fce0e3' }
            'WARNING'  { $background = '#fffaf0'; $ink = '#885400'; $badge = '#ffedc2' }
            'UNKNOWN'  { $background = '#f4f6f8'; $ink = '#475569'; $badge = '#e2e8f0' }
        }
        [void]$builder.Append('<tr bgcolor="' + $background + '" style="background:' + $background + ';">')
        [void]$builder.Append('<td style="' + $td + 'font-weight:600;">' + (ConvertTo-HtmlText $row.Server) + '</td>')
        [void]$builder.Append('<td style="' + $td + 'word-break:break-all;">' + (ConvertTo-HtmlText $row.DriveOrMountPoint) + '</td>')
        [void]$builder.Append('<td style="' + $td + '">' + (ConvertTo-HtmlText $row.VolumeLabel) + '</td>')
        foreach ($value in @($row.TotalGB, $row.UsedGB, $row.FreeGB)) {
            [void]$builder.Append('<td style="' + $td + 'text-align:right;white-space:nowrap;">' + (Format-ReportNumber $value) + '</td>')
        }
        $percentText = Format-ReportNumber $row.FreePercent
        if ($null -ne $row.FreePercent) { $percentText += '%' }
        [void]$builder.Append('<td style="' + $td + 'color:' + $ink + ';font-weight:700;white-space:nowrap;">' + $percentText + '</td>')
        [void]$builder.Append('<td style="' + $td + '"><span style="display:inline-block;padding:3px 7px;font-size:10px;font-weight:700;color:' + $ink + ';background:' + $badge + ';">' + (ConvertTo-HtmlText $row.Status) + '</span></td>')
        if ($ActionTable) {
            [void]$builder.Append('<td style="' + $td + 'text-align:right;">' + (Format-ReportNumber $row.ReclaimToThresholdGB) + '</td>')
        }
        [void]$builder.Append('</tr>')
    }
    [void]$builder.Append('</tbody></table>')
    if ($displayRows.Count -lt $Rows.Count) {
        [void]$builder.Append('<p style="font-size:12px;color:#885400;">Showing ' + $displayRows.Count + ' of ' + $Rows.Count + ' rows. The complete data is in the CSV files and the saved HTML report.</p>')
    }
    return $builder.ToString()
}

function New-IssueTableHtml {
    param([AllowEmptyCollection()][object[]]$Rows, [int]$Limit = 0)
    if ($Rows.Count -eq 0) { return '<p style="color:#17603a;font-size:13px;">No collection errors were reported.</p>' }
    $shown = @($Rows)
    if ($Limit -gt 0) { $shown = @($Rows | Select-Object -First $Limit) }
    $builder = [Text.StringBuilder]::new()
    [void]$builder.Append('<table width="100%" cellpadding="10" cellspacing="0" style="border-collapse:collapse;font-size:12px;line-height:18px;"><tr bgcolor="#475569" style="color:#ffffff;text-align:left;"><th>Server</th><th>Stage / volume</th><th>Issue requiring review</th></tr>')
    foreach ($row in $shown) {
        [void]$builder.Append('<tr bgcolor="#f4f6f8"><td style="border-bottom:1px solid #dce4eb;vertical-align:top;">' + (ConvertTo-HtmlText $row.Server) + '</td>')
        [void]$builder.Append('<td style="border-bottom:1px solid #dce4eb;vertical-align:top;word-break:break-all;">' + (ConvertTo-HtmlText $row.Stage) + '<br>' + (ConvertTo-HtmlText $row.Volume) + '</td>')
        [void]$builder.Append('<td style="border-bottom:1px solid #dce4eb;word-break:break-word;">' + (ConvertTo-HtmlText $row.Details) + '</td></tr>')
    }
    [void]$builder.Append('</table>')
    if ($shown.Count -lt $Rows.Count) {
        [void]$builder.Append('<p>Showing ' + $shown.Count + ' of ' + $Rows.Count + ' issues. See CollectionIssues.csv for all issues.</p>')
    }
    return $builder.ToString()
}

function New-ReportHtml {
    param(
        [AllowEmptyCollection()][object[]]$Drives,
        [AllowEmptyCollection()][object[]]$Alerts,
        [AllowEmptyCollection()][object[]]$Issues,
        [object]$Summary,
        [int]$Limit = 0
    )
    $bannerColor = '#17603a'; $bannerBackground = '#e7f4ec'
    $bannerTitle = 'NO LOW-SPACE ALERTS DETECTED'
    $bannerText = 'All successfully measured volumes meet the configured free-space threshold.'
    if ($Issues.Count -gt 0) {
        $bannerColor = '#885400'; $bannerBackground = '#fff3d6'
        $bannerTitle = 'COLLECTION INCOMPLETE - REVIEW REQUIRED'
        $bannerText = 'One or more servers or volumes could not be fully assessed. Unknown results are not healthy results.'
    }
    if ($Alerts.Count -gt 0) {
        $bannerColor = '#a61b29'; $bannerBackground = '#fdebed'
        $bannerTitle = 'ACTION REQUIRED - LOW FREE SPACE'
        $bannerText = '{0} volume(s) are below {1}% free space. Prioritize the lowest free-space percentages.' -f $Alerts.Count, $WarningThreshold
        if ($Issues.Count -gt 0) { $bannerText += ' Collection is also incomplete; review the collection issues below.' }
    }
    $actionHtml = New-VolumeTableHtml -Rows $Alerts -ActionTable -Limit $Limit
    $allHtml = New-VolumeTableHtml -Rows $Drives -Limit $Limit
    $issuesHtml = New-IssueTableHtml -Rows $Issues -Limit $Limit
    $tokens = [ordered]@{
        '{{TITLE}}' = (ConvertTo-HtmlText $Config.ReportTitle)
        '{{ENVIRONMENT}}' = (ConvertTo-HtmlText $Config.EnvironmentName)
        '{{TEAM}}' = (ConvertTo-HtmlText $Config.TeamName)
        '{{BANNER_COLOR}}' = $bannerColor
        '{{BANNER_BACKGROUND}}' = $bannerBackground
        '{{BANNER_TITLE}}' = $bannerTitle
        '{{BANNER_TEXT}}' = (ConvertTo-HtmlText $bannerText)
        '{{REQUESTED}}' = $Summary.Requested
        '{{COMPLETE}}' = $Summary.Complete
        '{{PARTIAL}}' = $Summary.Partial
        '{{FAILED}}' = $Summary.Failed
        '{{VOLUMES}}' = $Drives.Count
        '{{ALERTS}}' = $Alerts.Count
        '{{CRITICAL}}' = $Summary.Critical
        '{{ISSUES}}' = $Issues.Count
        '{{STARTED}}' = (ConvertTo-HtmlText $Summary.Started)
        '{{FINISHED}}' = (ConvertTo-HtmlText $Summary.Finished)
        '{{TIMEZONE}}' = (ConvertTo-HtmlText $Summary.TimeZone)
        '{{COLLECTOR}}' = (ConvertTo-HtmlText $Summary.Collector)
        '{{WARNING_THRESHOLD}}' = (ConvertTo-HtmlText $WarningThreshold)
        '{{CRITICAL_THRESHOLD}}' = (ConvertTo-HtmlText $CriticalThreshold)
        '{{ACTION_TABLE}}' = $actionHtml
        '{{ISSUE_TABLE}}' = $issuesHtml
        '{{ALL_TABLE}}' = $allHtml
        '{{ATTACHMENT_NOTE}}' = $(if ($Config.AttachCsvFiles) {
            'CSV email attachments are enabled.'
        } else {
            'CSV email attachments are disabled; this message contains the HTML report.'
        })
    }
    $html = @'
<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{TITLE}}</title></head>
<body style="margin:0;padding:24px 12px;background:#eef2f6;font-family:Segoe UI,Arial,sans-serif;color:#223246;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:1200px;margin:0 auto;background:#ffffff;border:1px solid #dce4eb;">
<tr><td bgcolor="#122b43" style="background:#122b43;padding:28px 30px;color:#ffffff;border-top:5px solid #239a9b;">
<div style="font-size:11px;letter-spacing:2px;color:#b3d2df;font-weight:700;">DATABASE OPERATIONS &nbsp; / &nbsp; CAPACITY MONITORING</div>
<h1 style="font-size:27px;line-height:35px;margin:10px 0 5px;font-weight:600;">{{TITLE}}</h1>
<div style="font-size:13px;line-height:20px;color:#ccdae7;">{{ENVIRONMENT}} &nbsp; | &nbsp; {{TEAM}} &nbsp; | &nbsp; Read-only inventory</div>
</td></tr>
<tr><td style="padding:20px 30px 8px;font-size:12px;line-height:20px;color:#596d81;">
<strong>Scan window:</strong> {{STARTED}} to {{FINISHED}}<br>
<strong>Collector:</strong> {{COLLECTOR}} &nbsp; | &nbsp; <strong>Time zone:</strong> {{TIMEZONE}}
</td></tr>
<tr><td style="padding:14px 30px;">
<table role="presentation" width="100%" cellspacing="0" cellpadding="16" style="background:{{BANNER_BACKGROUND}};border-left:4px solid {{BANNER_COLOR}};">
<tr><td><div style="font-size:13px;font-weight:700;color:{{BANNER_COLOR}};letter-spacing:.4px;">{{BANNER_TITLE}}</div>
<div style="font-size:13px;line-height:21px;margin-top:5px;">{{BANNER_TEXT}}</div></td></tr></table>
</td></tr>
<tr><td style="padding:4px 30px 20px;">
<table role="presentation" width="100%" cellpadding="14" cellspacing="0" style="border-collapse:collapse;background:#f6f8fa;">
<tr>
<td width="20%" style="border:1px solid #e1e7ed;"><div style="font-size:11px;color:#61758a;">SERVERS LISTED</div><div style="font-size:27px;font-weight:700;margin-top:5px;">{{REQUESTED}}</div></td>
<td width="20%" style="border:1px solid #e1e7ed;"><div style="font-size:11px;color:#61758a;">VOLUMES FOUND</div><div style="font-size:27px;font-weight:700;margin-top:5px;">{{VOLUMES}}</div></td>
<td width="20%" style="border:1px solid #e1e7ed;"><div style="font-size:11px;color:#61758a;">BELOW {{WARNING_THRESHOLD}}% FREE</div><div style="font-size:27px;font-weight:700;color:#a61b29;margin-top:5px;">{{ALERTS}}</div></td>
<td width="20%" style="border:1px solid #e1e7ed;"><div style="font-size:11px;color:#61758a;">CRITICAL VOLUMES</div><div style="font-size:27px;font-weight:700;color:#a61b29;margin-top:5px;">{{CRITICAL}}</div></td>
<td width="20%" style="border:1px solid #e1e7ed;"><div style="font-size:11px;color:#61758a;">COLLECTION ISSUES</div><div style="font-size:27px;font-weight:700;color:#475569;margin-top:5px;">{{ISSUES}}</div></td>
</tr></table>
<p style="font-size:12px;line-height:20px;margin:9px 0 0;color:#596d81;"><strong>Server coverage:</strong> {{COMPLETE}} complete &nbsp; / &nbsp; {{PARTIAL}} partial &nbsp; / &nbsp; {{FAILED}} failed. Complete means the volume and mount-point queries succeeded and returned valid measurements; it does not mean free space is healthy.</p>
</td></tr>
<tr><td style="padding:0 30px 22px;">
<h2 style="font-size:18px;margin:0 0 8px;color:#122b43;">01 &nbsp; Drives requiring DBA action</h2>
<p style="font-size:13px;line-height:21px;margin:0 0 14px;color:#596d81;">Only volumes below {{WARNING_THRESHOLD}}% free space appear here, ordered by lowest free percentage. Please review the cause, assign an owner, and record the remediation plan.</p>
{{ACTION_TABLE}}
<p style="font-size:11px;line-height:18px;color:#6b7b8c;margin:9px 0 0;">Reclaim to {{WARNING_THRESHOLD}}% is the minimum space to free at the current volume size, rounded upward to 0.01 GB. It is not the amount of additional disk capacity to provision. Plan additional headroom for growth.</p>
</td></tr>
<tr><td style="padding:0 30px 24px;">
<table role="presentation" width="100%" cellpadding="16" cellspacing="0" style="background:#edf5f8;border-left:4px solid #239a9b;"><tr><td style="font-size:12px;line-height:21px;">
<strong style="color:#122b43;">REQUIRED FOLLOW-UP</strong><br>
<strong>Critical:</strong> prioritize immediate investigation and confirm SQL workload impact.<br>
<strong>Warning:</strong> review growth, retention-compliant cleanup, and capacity requirements.<br>
<strong>Safety:</strong> do not delete active database or log files, restart SQL, or shrink databases as an automatic response. This script performs no remediation.
</td></tr></table>
</td></tr>
<tr><td style="padding:0 30px 24px;">
<h2 style="font-size:18px;margin:0 0 8px;color:#122b43;">02 &nbsp; Collection issues</h2>
<p style="font-size:13px;line-height:21px;margin:0 0 12px;color:#596d81;">Resolve connectivity, permissions, mount-path, or volume-visibility issues. Unmeasured volumes and unreachable hosts are never counted as healthy.</p>
{{ISSUE_TABLE}}
</td></tr>
<tr><td style="padding:0 30px 24px;">
<h2 style="font-size:18px;margin:0 0 8px;color:#122b43;">03 &nbsp; Complete drive inventory</h2>
<p style="font-size:13px;line-height:21px;margin:0 0 12px;color:#596d81;">All returned local fixed volumes, including drive letters and mounted-folder volumes. The complete CSV inventory also includes filesystem, volume ID, transport, raw bytes, and collection time.</p>
{{ALL_TABLE}}
</td></tr>
<tr><td bgcolor="#f4f7fa" style="padding:20px 30px;border-top:1px solid #dce4eb;font-size:11px;line-height:19px;color:#61758a;">
<strong style="color:#a61b29;">CRITICAL:</strong> below {{CRITICAL_THRESHOLD}}% &nbsp; | &nbsp;
<strong style="color:#885400;">WARNING:</strong> {{CRITICAL_THRESHOLD}}% to below {{WARNING_THRESHOLD}}% &nbsp; | &nbsp;
<strong style="color:#17603a;">HEALTHY:</strong> {{WARNING_THRESHOLD}}% or more &nbsp; | &nbsp;
<strong>UNKNOWN:</strong> measurement unavailable<br>
Capacity units labeled GB use 1,073,741,824 bytes (GiB). Threshold decisions use unrounded percentages; displayed percentages use two decimals. This is a point-in-time Windows volume report, not a SQL file-growth forecast or a physical-storage deduplication report. Shared storage may appear on multiple nodes.<br>
CSV files are saved on the collector: AllDrives.csv, ActionRequired.csv, CollectionIssues.csv, ServerCoverage.csv. {{ATTACHMENT_NOTE}} No monitored-server configuration changes, file cleanup, service restarts, or reboots are performed. Email requests are queued in msdb through Database Mail.
</td></tr>
</table></body></html>
'@
    foreach ($key in $tokens.Keys) { $html = $html.Replace($key, [string]$tokens[$key]) }
    return $html
}

function Send-ReportEmail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Html,
        [AllowEmptyCollection()][string[]]$Attachments = @(),
        [bool]$HighPriority = $false
    )

    if ([string]::IsNullOrWhiteSpace([string]$Config.MailSqlInstance) -or
        $Config.MailSqlInstance -like '*yourcompany.com*') {
        throw 'Set MailSqlInstance to the SQL instance with your EXISTING Database Mail profile.'
    }
    if ([string]::IsNullOrWhiteSpace([string]$Config.MailProfileName) -or
        ([string]$Config.MailProfileName).Length -gt 128) {
        throw 'MailProfileName must be the existing profile name (1 to 128 characters).'
    }
    $toList = @($Config.MailTo | ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $ccList = @($Config.MailCc | ForEach-Object { ([string]$_).Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($toList.Count -eq 0) { throw 'Configure at least one recipient in MailTo.' }
    foreach ($address in @($toList + $ccList)) {
        if ($address -like '*yourcompany.com*' -or $address -match '[\r\n;,]') {
            throw 'MailTo/MailCc must contain real, individual email addresses: one address per array element.'
        }
        # Validate format. This does not prove that the mailbox exists.
        [void][System.Net.Mail.MailAddress]::new($address)
    }
    if ($Subject.Length -gt 255) {
        Write-RunLog -Level 'WARN' -Message 'Email subject shortened to the Database Mail limit of 255 characters.'
        $Subject = $Subject.Substring(0, 255)
    }
    if ([int]$Config.SqlConnectTimeoutSeconds -le 0 -or [int]$Config.SqlCommandTimeoutSeconds -le 0) {
        throw 'SQL connection and command timeouts must be greater than zero.'
    }

    $connection = $null
    $command = $null
    $hostCommand = $null
    $submitAttempted = $false
    $mailItemId = $null
    try {
        # Use explicit connection-string keys, not dot-property assignment.
        # PowerShell can route property writes on this IDictionary through
        # its indexer: DataSource is not the supported key 'Data Source'.
        $builder = [System.Data.SqlClient.SqlConnectionStringBuilder]::new()
        $builder['Data Source'] = [string]$Config.MailSqlInstance
        $builder['Initial Catalog'] = 'msdb'
        $builder['Integrated Security'] = $true
        $builder['Encrypt'] = $true
        $builder['TrustServerCertificate'] = $false
        $builder['Connect Timeout'] = [int]$Config.SqlConnectTimeoutSeconds
        $builder['Application Name'] = 'DBA Drive Capacity Report - Database Mail'
        $connection = [System.Data.SqlClient.SqlConnection]::new($builder.get_ConnectionString())
        Write-RunLog -Message ('Connecting to mail SQL instance {0} using the current Windows identity.' -f $Config.MailSqlInstance)
        $connection.Open()

        $attachmentPaths = @()
        if ($Config.AttachCsvFiles -and @($Attachments).Count -gt 0) {
            # Guard against passing the collector's D:\ path to a different host.
            # This implementation supports local files on the SQL host only.
            # Remote sending remains supported when CSV attachments are OFF.
            $hostCommand = $connection.CreateCommand()
            $hostCommand.CommandTimeout = [int]$Config.SqlCommandTimeoutSeconds
            $hostCommand.CommandText = "SELECT CONVERT(nvarchar(128), SERVERPROPERTY('ComputerNamePhysicalNetBIOS'));"
            $sqlPhysicalHost = [string]$hostCommand.ExecuteScalar()
            if ([string]::IsNullOrWhiteSpace($sqlPhysicalHost) -or
                $sqlPhysicalHost -ine [Environment]::MachineName) {
                throw ('CSV attachment mode requires this script to run on the Windows host currently running the mail SQL instance. SQL host: {0}; collector: {1}. Set AttachCsvFiles = $false for remote HTML-only sending.' -f $sqlPhysicalHost, [Environment]::MachineName)
            }
            foreach ($file in $Attachments) {
                if ([string]::IsNullOrWhiteSpace($file) -or $file.Contains(';')) {
                    throw 'An attachment path is empty or contains a semicolon, which Database Mail uses as a separator.'
                }
                $item = Get-Item -LiteralPath $file -ErrorAction Stop
                if ($item.PSIsContainer -or $item.PSProvider.Name -ne 'FileSystem' -or
                    $item.FullName -notmatch '^[A-Za-z]:\\') {
                    throw ('CSV attachment mode supports only local, absolute drive-letter file paths on the mail SQL host: {0}' -f $file)
                }
                $attachmentPaths += $item.FullName
            }
            Write-RunLog -Message ('Submitting {0} CSV attachment(s). SQL Server must have file read access; Database Mail size/extension limits apply.' -f $attachmentPaths.Count)
        }
        elseif (@($Attachments).Count -gt 0) {
            Write-RunLog -Message 'HTML-only email selected. All CSV files remain saved in the report folder.'
        }

        # Parameters preserve quotes and Unicode HTML without SQL concatenation.
        $command = $connection.CreateCommand()
        $command.CommandType = [System.Data.CommandType]::StoredProcedure
        $command.CommandText = 'msdb.dbo.sp_send_dbmail'
        $command.CommandTimeout = [int]$Config.SqlCommandTimeoutSeconds
        [void]$command.Parameters.Add('@profile_name', [System.Data.SqlDbType]::NVarChar, 128)
        $command.Parameters['@profile_name'].Value = [string]$Config.MailProfileName
        [void]$command.Parameters.Add('@recipients', [System.Data.SqlDbType]::VarChar, -1)
        $command.Parameters['@recipients'].Value = $toList -join ';'
        if ($ccList.Count -gt 0) {
            [void]$command.Parameters.Add('@copy_recipients', [System.Data.SqlDbType]::VarChar, -1)
            $command.Parameters['@copy_recipients'].Value = $ccList -join ';'
        }
        [void]$command.Parameters.Add('@subject', [System.Data.SqlDbType]::NVarChar, 255)
        $command.Parameters['@subject'].Value = $Subject
        [void]$command.Parameters.Add('@body', [System.Data.SqlDbType]::NVarChar, -1)
        $command.Parameters['@body'].Value = $Html
        [void]$command.Parameters.Add('@body_format', [System.Data.SqlDbType]::VarChar, 20)
        $command.Parameters['@body_format'].Value = 'HTML'
        [void]$command.Parameters.Add('@importance', [System.Data.SqlDbType]::VarChar, 6)
        $command.Parameters['@importance'].Value = $(if ($HighPriority) { 'High' } else { 'Normal' })
        if ($attachmentPaths.Count -gt 0) {
            [void]$command.Parameters.Add('@file_attachments', [System.Data.SqlDbType]::NVarChar, -1)
            $command.Parameters['@file_attachments'].Value = $attachmentPaths -join ';'
        }
        $idParameter = $command.Parameters.Add('@mailitem_id', [System.Data.SqlDbType]::Int)
        $idParameter.Direction = [System.Data.ParameterDirection]::Output
        $returnParameter = $command.Parameters.Add('@RETURN_VALUE', [System.Data.SqlDbType]::Int)
        $returnParameter.Direction = [System.Data.ParameterDirection]::ReturnValue

        # No automatic retry: on a timeout, SQL may already have queued the mail.
        $submitAttempted = $true
        [void]$command.ExecuteNonQuery()
        if ($null -eq $returnParameter.Value -or $returnParameter.Value -is [DBNull] -or
            [int]$returnParameter.Value -ne 0) {
            throw ('sp_send_dbmail returned a failure or no return code: {0}' -f $returnParameter.Value)
        }
        if ($null -eq $idParameter.Value -or $idParameter.Value -is [DBNull] -or [int]$idParameter.Value -le 0) {
            throw 'sp_send_dbmail did not return a valid mailitem_id. Check Database Mail history before retrying.'
        }
        $mailItemId = [int]$idParameter.Value
    }
    catch {
        try {
            Write-RunLog -Level 'ERROR' -Message ('Database Mail submission error: ' + $_.Exception.ToString())
            if ($submitAttempted) {
                Write-RunLog -Level 'WARN' -Message 'Check msdb Database Mail history before retrying; a failed or timed-out client call does not always prove that no message was queued.'
            }
        } catch { }
        throw
    }
    finally {
        if ($null -ne $hostCommand) { $hostCommand.Dispose() }
        if ($null -ne $command) { $command.Dispose() }
        if ($null -ne $connection) { $connection.Dispose() }
    }
    # Do not interpret successful queue insertion as delivery confirmation.
    Write-RunLog -Message ('Database Mail QUEUED mailitem_id={0}; profile={1}; recipients={2}. Check msdb.dbo.sysmail_allitems for send status.' -f $mailItemId, $Config.MailProfileName, ($toList -join ';'))
    return $mailItemId
}

function Save-MailSubmissionReceipt {
    param([int]$MailItemId, [string]$Folder, [string]$Subject)
    # Failure to save a receipt must not imply that an already queued email failed.
    try {
        [pscustomobject]@{
            MailItemId = $MailItemId
            SqlInstance = [string]$Config.MailSqlInstance
            Profile = [string]$Config.MailProfileName
            Recipients = @($Config.MailTo)
            Subject = $Subject
            SubmittedAt = [DateTimeOffset]::Now.ToString('o')
            SubmissionStatus = 'QUEUED - not a delivery confirmation'
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $Folder 'DatabaseMailSubmission.json') -Encoding UTF8
        $statusSql = @"
-- Connect SSMS to the MailSqlInstance shown in DatabaseMailSubmission.json.
-- Read-only status check for this submission.
USE msdb;
DECLARE @MailItemId int = $MailItemId;
SELECT mailitem_id, subject, recipients, sent_status, send_request_date, sent_date
FROM dbo.sysmail_allitems WHERE mailitem_id = @MailItemId;
SELECT log_date, event_type, description
FROM dbo.sysmail_event_log WHERE mailitem_id = @MailItemId ORDER BY log_date DESC;
"@
        Set-Content -LiteralPath (Join-Path $Folder 'CheckMailStatus.sql') -Value $statusSql -Encoding UTF8
    }
    catch {
        Write-Warning ('Mail was queued as mailitem_id={0}, but its local receipt could not be saved: {1}' -f $MailItemId, $_.Exception.Message)
    }
}

# =============================== MAIN ===============================
try {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'Run this script on a Windows management server with the CimCmdlets module.'
    }
    if ($CriticalThreshold -ge $WarningThreshold) {
        throw 'CriticalThreshold must be lower than WarningThreshold.'
    }
    if ([int]$Config.MaxInlineRowsPerSection -lt 0) { throw 'MaxInlineRowsPerSection cannot be negative.' }
    $started = [DateTimeOffset]::Now
    $runId = '{0}_{1}' -f $started.ToString('yyyyMMdd_HHmmss'), $PID
    $runFolder = Join-Path $OutputFolder $runId
    [void](New-Item -ItemType Directory -Path $runFolder -Force)
    $runFolder = (Resolve-Path -LiteralPath $runFolder).ProviderPath
    $script:LogPath = Join-Path $runFolder 'Run.log'
    if ($TestEmailOnly) {
        if ($NoEmail -or $OnlySendOnAlert -or $OpenReport) {
            throw 'Use -TestEmailOnly by itself; do not combine it with -NoEmail, -OnlySendOnAlert, or -OpenReport.'
        }
        Write-RunLog -Message 'TEST EMAIL ONLY: no drive collection and no CSV attachments.'
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        $testHtml = '<html><body style="font-family:Segoe UI,Arial,sans-serif;color:#223246;">' +
            '<div style="background:#122b43;color:white;padding:20px;font-size:22px;">DBA Capacity Report - Database Mail Test</div>' +
            '<div style="padding:20px;border:1px solid #dce4eb;"><p>This is a mail-routing test. No drive inventory was collected.</p>' +
            '<p><strong>Collector:</strong> ' + (ConvertTo-HtmlText ([Environment]::MachineName)) +
            '<br><strong>Windows identity:</strong> ' + (ConvertTo-HtmlText $identity) +
            '<br><strong>Mail SQL instance:</strong> ' + (ConvertTo-HtmlText $Config.MailSqlInstance) +
            '<br><strong>Profile:</strong> ' + (ConvertTo-HtmlText $Config.MailProfileName) +
            '</p><p>If you can read this email, the test reached this mailbox.</p></div></body></html>'
        $testSubject = '[TEST] DBA Drive Capacity - Database Mail - ' + [DateTimeOffset]::Now.ToString('yyyy-MM-dd HH:mm zzz')
        $testMailId = Send-ReportEmail -Subject $testSubject -Html $testHtml -Attachments @() -HighPriority $false
        Save-MailSubmissionReceipt -MailItemId $testMailId -Folder $runFolder -Subject $testSubject
        Write-RunLog -Message ('Test submission complete. CheckMailStatus.sql is saved in: {0}' -f $runFolder)
        exit 0
    }
    $computers = @(Read-ServerNames -Path $ServerListPath)
    Write-RunLog -Message ('Starting read-only collection from {0} host(s). Alert threshold: below {1}%.' -f $computers.Count, $WarningThreshold)

    $allRows = [System.Collections.Generic.List[object]]::new()
    $allIssues = [System.Collections.Generic.List[object]]::new()
    $coverageRows = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($computer in $computers) {
        $index++
        Write-Progress -Activity 'Collecting SQL host drive capacity' -Status ("$index / $($computers.Count): $computer") `
            -PercentComplete (($index / $computers.Count) * 100)
        $result = Get-ServerVolumes -Computer $computer
        foreach ($row in $result.Rows) { $allRows.Add($row) }
        foreach ($issue in $result.Issues) {
            $allIssues.Add($issue)
            Write-RunLog -Level 'WARN' -Message ('{0} | {1} | {2}' -f $issue.Server, $issue.Stage, $issue.Details)
        }
        $coverageRows.Add($result.Coverage)
        Write-RunLog -Message ('{0}: {1}; {2} volume(s); {3}' -f $computer, $result.Coverage.CollectionStatus, $result.Coverage.VolumeCount, $result.Coverage.Transport)
    }
    Write-Progress -Activity 'Collecting SQL host drive capacity' -Completed
    $drives = @($allRows.ToArray() | Sort-Object Server,DriveOrMountPoint)
    $alerts = @($drives | Where-Object { $_.Status -in @('CRITICAL','WARNING') } | Sort-Object FreePercent,Server,DriveOrMountPoint)
    $issues = @($allIssues.ToArray() | Sort-Object Server,Stage,Volume)
    $coverage = @($coverageRows.ToArray() | Sort-Object Server)
    $finished = [DateTimeOffset]::Now
    $summary = [pscustomobject]@{
        Requested = $computers.Count
        Complete = @($coverage | Where-Object { $_.CollectionStatus -eq 'COMPLETE' }).Count
        Partial = @($coverage | Where-Object { $_.CollectionStatus -eq 'PARTIAL' }).Count
        Failed = @($coverage | Where-Object { $_.CollectionStatus -eq 'FAILED' }).Count
        Critical = @($alerts | Where-Object { $_.Status -eq 'CRITICAL' }).Count
        Started = $started.ToString('yyyy-MM-dd HH:mm:ss zzz')
        Finished = $finished.ToString('yyyy-MM-dd HH:mm:ss zzz')
        TimeZone = [TimeZoneInfo]::Local.Id
        Collector = [Environment]::MachineName
    }
    $driveColumns = @('Server','DriveOrMountPoint','VolumeLabel','FileSystem','TotalGB','UsedGB','FreeGB',
        'FreePercent','UsedPercent','Status','ReclaimToThresholdGB','ThresholdPercent','CapacityBytes',
        'FreeBytes','VolumeId','Transport','ObservedAt','Notes')
    $allCsv = Join-Path $runFolder 'AllDrives.csv'
    $alertCsv = Join-Path $runFolder 'ActionRequired.csv'
    $issueCsv = Join-Path $runFolder 'CollectionIssues.csv'
    $coverageCsv = Join-Path $runFolder 'ServerCoverage.csv'
    Export-ReportCsv -Rows $drives -Columns $driveColumns -Path $allCsv
    Export-ReportCsv -Rows $alerts -Columns $driveColumns -Path $alertCsv
    Export-ReportCsv -Rows $issues -Columns @('Server','Stage','Volume','Details','ObservedAt') -Path $issueCsv
    Export-ReportCsv -Rows $coverage -Columns @('Server','CollectionStatus','Transport','VolumeCount','Notes') -Path $coverageCsv
    $htmlPath = Join-Path $runFolder 'DriveCapacityReport.html'
    $fullHtml = New-ReportHtml -Drives $drives -Alerts $alerts -Issues $issues -Summary $summary
    Set-Content -LiteralPath $htmlPath -Value $fullHtml -Encoding UTF8
    Write-RunLog -Message ("Report saved: $htmlPath")

    $prefix = 'HEALTHY'
    if ($issues.Count -gt 0) { $prefix = 'INCOMPLETE' }
    if ($alerts.Count -gt 0) { $prefix = 'ACTION REQUIRED' }
    if ($summary.Critical -gt 0) { $prefix = 'CRITICAL' }
    $subject = '[{0}] SQL Drive Capacity | {1} drive(s) below {2}% free | {3}' -f $prefix, $alerts.Count, $WarningThreshold, $finished.ToString('yyyy-MM-dd HH:mm zzz')
    if ($issues.Count -gt 0) { $subject += ' | Collection incomplete' }
    $shouldSend = -not $NoEmail
    if ($OnlySendOnAlert -and $alerts.Count -eq 0 -and $issues.Count -eq 0) { $shouldSend = $false }
    if ($shouldSend) {
        $emailHtml = New-ReportHtml -Drives $drives -Alerts $alerts -Issues $issues -Summary $summary `
            -Limit ([int]$Config.MaxInlineRowsPerSection)
        $mailItemId = Send-ReportEmail -Subject $subject -Html $emailHtml -Attachments @($allCsv,$alertCsv,$issueCsv,$coverageCsv) `
            -HighPriority ($alerts.Count -gt 0 -or $issues.Count -gt 0)
        Save-MailSubmissionReceipt -MailItemId $mailItemId -Folder $runFolder -Subject $subject
    }
    else { Write-RunLog -Message 'Email not sent (-NoEmail or a clean -OnlySendOnAlert run).' }
    if ($OpenReport) {
        try { Invoke-Item -LiteralPath $htmlPath }
        catch { Write-RunLog -Level 'WARN' -Message ('Could not open the report viewer: ' + $_.Exception.Message) }
    }
    Write-RunLog -Message ('Finished. Volumes: {0}; alerts: {1}; collection issues: {2}.' -f $drives.Count, $alerts.Count, $issues.Count)
    if ($issues.Count -gt 0) { exit 2 }
    exit 0
}
catch {
    $failureMessage = $_.Exception.ToString()
    try { Write-RunLog -Level 'ERROR' -Message $failureMessage } catch { Write-Host $failureMessage }
    Write-Error -Message ('Report run failed: ' + $failureMessage) -ErrorAction Continue
    exit 1
}
