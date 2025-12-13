[CmdletBinding()]
param(
    [string]$RootFolder
)

$ErrorActionPreference       = "Stop"
$Script:ShownBackupInfoShown = $false
$Script:SkipExitPause        = $false
$Script:SkipUpdateCheck      = $false
$Script:VersionStatus        = "Unknown"
$Script:AutoUseLastFolder    = $false
$Script:RootFolderForLogs    = $null
$Script:LogFilePath          = $null
$Script:ScriptPath           = $MyInvocation.MyCommand.Path
$Script:AppDataDir           = Join-Path $env:LOCALAPPDATA "LSR-XML-Helper"
$Script:LocalVersionFile     = Join-Path $Script:AppDataDir "version.txt"
$Script:RemoteVersionUrl     = "https://pastebin.com/raw/56yTg6aw"
$Script:RemoteScriptUrl      = "https://drive.usercontent.google.com/download?id=1uunlxT5bV5sXCDO-OGaIOT4QFuuIF4G2&export=download&confirm=t"


if (-not (Test-Path $Script:AppDataDir)) {
    New-Item -ItemType Directory -Path $Script:AppDataDir -Force | Out-Null
}

function Write-LogError {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ErrorObject
    )

    try {
        $logPath = Get-LogFilePath
        $time    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $lines   = @()

        if ($ErrorObject -isnot [System.Management.Automation.ErrorRecord]) {
            $lines += "[{0}] ERROR: {1}" -f $time, ([string]$ErrorObject)
        }
        else {
            $err = [System.Management.Automation.ErrorRecord]$ErrorObject
            $ex  = $err.Exception
            $inv = $err.InvocationInfo

            $funcName = if ($inv -and $inv.MyCommand -and $inv.MyCommand.Name) {
                $inv.MyCommand.Name
            }
            elseif ($inv -and $inv.InvocationName) {
                $inv.InvocationName
            }
            else {
                '<global>'
            }

            $lines += "[{0}] ERROR: {1}" -f $time, $ex.Message
            $lines += "    Function : {0}" -f $funcName

            if ($inv) {
                if ($inv.ScriptLineNumber) {
                    $lines += "    Line     : {0}" -f $inv.ScriptLineNumber
                }
                if ($inv.OffsetInLine) {
                    $lines += "    Column   : {0}" -f $inv.OffsetInLine
                }
                if ($inv.Line) {
                    $lines += "    Code     : {0}" -f $inv.Line.Trim()
                }
            }

            if ($err.CategoryInfo) {
                $lines += "    Category : {0}" -f ($err.CategoryInfo.ToString())
            }
            if ($err.FullyQualifiedErrorId) {
                $lines += "    FullyQualifiedErrorId : {0}" -f $err.FullyQualifiedErrorId
            }
        }

        $lines += ""

        Add-Content -Path $logPath -Value $lines -Encoding UTF8
    }
    catch {
    }
}

function Get-LogFilePath {
    if ($Script:LogFilePath) {
        $existingDir = Split-Path $Script:LogFilePath -Parent
        if (Test-Path $existingDir) {
            return $Script:LogFilePath
        }
    }

    $helperRoot = $null

    if ($Script:RootFolderForLogs -and (Test-Path $Script:RootFolderForLogs)) {
        $helperRoot = Join-Path $Script:RootFolderForLogs "LSR-XML-Helper"
    }
    elseif ($Script:ScriptPath -and (Test-Path $Script:ScriptPath)) {
        $scriptDir  = Split-Path $Script:ScriptPath -Parent
        $helperRoot = Join-Path $scriptDir "LSR-XML-Helper"
    }
    else {
        $helperRoot = $Script:AppDataDir
    }

    if (-not (Test-Path $helperRoot)) {
        New-Item -ItemType Directory -Path $helperRoot -Force | Out-Null
    }

    $logsDir = Join-Path $helperRoot "Logs"
    if (-not (Test-Path $logsDir)) {
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }

    $Script:LogFilePath = Join-Path $logsDir "LSR-XML-Helper.log"
    return $Script:LogFilePath
}

function Get-LocalVersion {
    if (Test-Path $Script:LocalVersionFile) {
        return (Get-Content -Path $Script:LocalVersionFile -Raw).Trim()
    }
    return "0.0.0"
}

function Set-LocalVersion {
    param([string]$Version)
    $Version = $Version.Trim()
    Set-Content -Path $Script:LocalVersionFile -Value $Version -Encoding ASCII
}

function Get-RemoteVersion {
    try {
        $content = Invoke-WebRequest -Uri $Script:RemoteVersionUrl -UseBasicParsing -ErrorAction Stop
        $ver = $content.Content.Trim()

        if ($ver -notmatch '^\d+(\.\d+){1,2}$') {
            Write-Host "[!] Remote version format looks invalid ($ver)." -ForegroundColor Yellow
            Write-Host "[!] Update check cancelled." -ForegroundColor Yellow
            Start-Sleep -Seconds 3
            return $null
        }

        return $ver
    } catch {
        Write-Host "[i] Could not check online version." -ForegroundColor Yellow
        Write-Host "[i] This may be due to:" -ForegroundColor Yellow
        Write-Host "    • No internet connection" -ForegroundColor Yellow
        Write-Host "    • Antivirus blocking PowerShell" -ForegroundColor Yellow
        Write-Host "    • Firewall blocking downloads" -ForegroundColor Yellow
        Write-Host "    • Network admin restrictions" -ForegroundColor Yellow
        Start-Sleep -Seconds 4
        return $null
    }
}

function Compare-Version {
    param(
        [string]$A,
        [string]$B
    )

    try {
        $vA = [version]$A
        $vB = [version]$B
    } catch {
        return 0
    }

    return $vA.CompareTo($vB)
}

function Perform-Update {
    param([string]$NewVersion)

    Write-Host ""
    Write-Host "Downloading new version $NewVersion..." -ForegroundColor Cyan
    Start-Sleep -Seconds 2

    $tempFile = Join-Path ([System.IO.Path]::GetTempPath()) "LSR-XML-Helper.exe"

    try {
        Invoke-WebRequest -Uri $Script:RemoteScriptUrl -OutFile $tempFile -UseBasicParsing -ErrorAction Stop
    } catch {
        Write-Host "[!] Failed to download new version: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[i] This could be caused by:" -ForegroundColor Yellow
        Write-Host "    • Internet connection issues" -ForegroundColor Yellow
        Write-Host "    • Antivirus blocking downloads" -ForegroundColor Yellow
        Write-Host "    • Firewall restrictions" -ForegroundColor Yellow
        Start-Sleep -Seconds 4
        return
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($tempFile)
    } catch {
        Write-Host "[!] Failed to read downloaded file." -ForegroundColor Yellow
        Write-Host "[i] Your antivirus may have removed or quarantined it." -ForegroundColor Yellow
        Start-Sleep -Seconds 4
        return
    }

    if ($bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) {
        Write-Host "[!] Downloaded file is not a valid EXE." -ForegroundColor Yellow
        Write-Host "[i] The download was blocked, corrupted, or replaced by AV." -ForegroundColor Yellow
        Start-Sleep -Seconds 4
        return
    }

    try {
        Copy-Item -Path $tempFile -Destination $Script:ScriptPath -Force
    } catch {
        Write-Host "[!] Failed to replace the application file." -ForegroundColor Yellow
        Write-Host "[i] Possible causes:" -ForegroundColor Yellow
        Write-Host "    • The app is not in a writable location" -ForegroundColor Yellow
        Write-Host "    • Antivirus prevented file updates" -ForegroundColor Yellow
        Write-Host "    • File is currently locked/open" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Move the tool into Desktop or Documents and try again." -ForegroundColor Cyan
        Start-Sleep -Seconds 4
        return
    }

    Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    Set-LocalVersion -Version $NewVersion

    Write-Host ""
    Write-Host "[+] Successfully updated to version $NewVersion!" -ForegroundColor Green
    Write-Host ""
    Write-Host "[i] If the tool still shows an older version when opened:" -ForegroundColor Yellow
    Write-Host "    • Your antivirus may have quarantined the updated file" -ForegroundColor Yellow
    Write-Host "    • Your firewall may have blocked the download mid-way" -ForegroundColor Yellow
    Write-Host "    • Some security tools silently roll files back" -ForegroundColor Yellow
    Write-Host "    • Add this tool to antivirus 'allowed list / exclusions'" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "[i] Restart now to use the updated version." -ForegroundColor Cyan
    Start-Sleep -Seconds 4

    exit
}

function Check-ForUpdate {
    $local  = Get-LocalVersion
    $remote = Get-RemoteVersion

    $Script:VersionStatus     = "Unknown"
    $Script:LatestRemoteVersion = $null

    if (-not $remote) {
        Write-Host "[i] Could not retrieve version information." -ForegroundColor Yellow
        Write-Host "[i] This may be caused by:" -ForegroundColor Yellow
        Write-Host "    • Internet connection issues" -ForegroundColor Yellow
        Write-Host "    • Firewall blocking downloads" -ForegroundColor Yellow
        Write-Host "    • Antivirus blocking update check" -ForegroundColor Yellow
        Start-Sleep -Seconds 3
        return
    }

    $cmp = Compare-Version -A $local -B $remote
    $Script:LatestRemoteVersion = $remote

    if ($cmp -ge 0) {
        $Script:VersionStatus = "Latest"
    } else {
        $Script:VersionStatus = "Outdated"
    }

    Write-Host ""
    Write-Host ""
    Write-Host "Version info:" -ForegroundColor Cyan

    Write-Host -NoNewline "Current: "
    if ($cmp -lt 0) {
        Write-Host $local -ForegroundColor Yellow
    } else {
        Write-Host $local -ForegroundColor Green
    }

    Write-Host -NoNewline "Latest : "
    Write-Host $remote -ForegroundColor Green
    Write-Host ""

    if ($cmp -ge 0) {
        Write-Host "[+] You are already on the latest version." -ForegroundColor Green
        Start-Sleep -Seconds 2
        return
    }

    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host "     New version available!" -ForegroundColor Green
    Write-Host "===============================================" -ForegroundColor Cyan
    Write-Host ""

    $answer = Read-Host "Update now? (Y/N)"
    if ($answer -match '^[Yy]$') {
        Perform-Update -NewVersion $remote
    }
}

function Clear-Screen {
    try {
        $h = [System.Console]::WindowHeight
        if ($h -gt 0 -and [System.Console]::BufferHeight -ne $h) {
            [System.Console]::BufferHeight = $h
        }
    } catch { }

    try {
        $esc = [char]27
        Write-Host "$esc[3J$esc[2J$esc[H" -NoNewline
    } catch { }

    try { [System.Console]::Clear() } catch { Clear-Host }
}

function Write-Title {
    param([string]$Text)
    Write-Host ""
    Write-Host "================ $Text ================" -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Text)
    Write-Host "[i] $Text"
}

function Write-Good {
    param([string]$Text)
    Write-Host "[+] $Text" -ForegroundColor Green
}

function Write-Bad {
    param([string]$Text)
    Write-Host "[!] $Text" -ForegroundColor Yellow
}

function Open-Folder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    try {
        Start-Process explorer.exe $Path | Out-Null
    } catch {
        Write-Bad "Could not open folder: $Path"
        Start-Sleep -Seconds 2
    }
}

function Get-HelperRootForXmlPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $dir        = Split-Path $XmlPath -Parent
    $helperRoot = Join-Path $dir "LSR-XML-Helper"

    if (-not (Test-Path $helperRoot)) {
        New-Item -ItemType Directory -Path $helperRoot -Force | Out-Null
    }

    return $helperRoot
}

function Get-ChangeLogPathForXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $fileName = Split-Path $XmlPath -Leaf
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

    $helperRoot = Get-HelperRootForXmlPath -XmlPath $XmlPath
    $changesDir = Join-Path $helperRoot "XML-Edits"

    if (-not (Test-Path $changesDir)) {
        New-Item -ItemType Directory -Path $changesDir -Force | Out-Null
    }

    return (Join-Path $changesDir ("{0}.changes.json" -f $baseName))
}

function Load-ChangesForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $changesPath = Get-ChangeLogPathForXml -XmlPath $XmlPath
    if (-not (Test-Path $changesPath)) {
        return @()
    }

    try {
        $json = Get-Content -Path $changesPath -Raw
        if ([string]::IsNullOrWhiteSpace($json)) {
            return @()
        }

        $data = $json | ConvertFrom-Json

        if ($null -eq $data) {
            return @()
        }

        if ($data -is [System.Collections.IEnumerable]) {
            $clean = @()
            foreach ($item in $data) {
                if ($null -ne $item) {
                    $clean += $item
                }
            }
            return $clean
        } else {
            return @($data)
        }
    } catch {
        Write-Bad "Could not read change log '$changesPath': $($_.Exception.Message)"
        return @()
    }
}

function Set-ChangesForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath,

        [AllowEmptyCollection()]
        [object[]]$Changes
    )

    if (-not $Changes) {
        $Changes = @()
    }

    $changesPath = Get-ChangeLogPathForXml -XmlPath $XmlPath
    $changesDir  = Split-Path $changesPath -Parent

    if ($Changes.Count -eq 0) {
        try {
            if (Test-Path $changesPath) {
                Remove-Item -Path $changesPath -Force
                Write-Info "Removed change log '$changesPath' (no saved edits left)."
            } else {
                Write-Info "No change log to keep for '$XmlPath' (no saved edits left)."
            }
        } catch {
            Write-Bad "Failed to remove change log '$changesPath': $($_.Exception.Message)"
        }
        return
    }

    $normalized = @()
    foreach ($c in $Changes) {
        if ($null -eq $c) { continue }

        $statusProp = $c.PSObject.Properties['Status']
        if (-not $statusProp) {
            $c | Add-Member -NotePropertyName 'Status' -NotePropertyValue 'Pending' -Force
        } else {
            $val = [string]$statusProp.Value
            if ([string]::IsNullOrWhiteSpace($val)) {
                $statusProp.Value = 'Pending'
            }
        }

        $normalized += $c
    }

    $Changes = $normalized

    if (-not (Test-Path $changesDir)) {
        New-Item -ItemType Directory -Path $changesDir -Force | Out-Null
    }

    try {
        $Changes | ConvertTo-Json -Depth 8 | Set-Content -Path $changesPath -Encoding UTF8
        Write-Info "Updated change log '$changesPath'."
    } catch {
        Write-Bad "Failed to update change log '$changesPath': $($_.Exception.Message)"
    }
}

function Mark-AllPendingChangesCommittedForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $changes = @(Load-ChangesForFile -XmlPath $XmlPath)
    if ($changes.Count -eq 0) {
        return
    }

    $updated = @()

    foreach ($c in $changes) {
        if ($null -eq $c) { continue }

        $statusProp = $c.PSObject.Properties['Status']
        if (-not $statusProp) {
            $c | Add-Member -NotePropertyName 'Status' -NotePropertyValue 'Committed' -Force
        } else {
            $val = [string]$statusProp.Value
            if ([string]::IsNullOrWhiteSpace($val) -or $val -eq 'Pending') {
                $statusProp.Value = 'Committed'
            }
        }

        $updated += $c
    }

    Set-ChangesForFile -XmlPath $XmlPath -Changes $updated
}

function Remove-PendingChangesForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $changes = @(Load-ChangesForFile -XmlPath $XmlPath)
    if ($changes.Count -eq 0) {
        Set-ChangesForFile -XmlPath $XmlPath -Changes @()
        return
    }

    $remaining = @(
        $changes | Where-Object {
            $p = $_.PSObject.Properties['Status']
            if (-not $p) { return $false }
            [string]$p.Value -eq 'Committed'
        }
    )

    Set-ChangesForFile -XmlPath $XmlPath -Changes $remaining
}

function Get-PendingChangesSummaryForAllFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    $items = @()

    foreach ($file in $XmlFiles) {
        $xmlPath = $file.FullName
        $changes = @(Load-ChangesForFile -XmlPath $xmlPath)
        if ($changes.Count -eq 0) { continue }

        $pending = @(
            $changes | Where-Object {
                $p = $_.PSObject.Properties['Status']
                if (-not $p) { return $true }
                $val = [string]$p.Value
                if ([string]::IsNullOrWhiteSpace($val)) { return $true }
                return ($val -ne 'Committed')
            }
        )

        if ($pending.Count -eq 0) { continue }

        $items += [pscustomobject]@{
            File         = $file
            XmlPath      = $xmlPath
            Pending      = $pending
            PendingCount = $pending.Count
        }
    }

    return $items
}

function Apply-AllPendingChangesForAllFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    $anyApplied = $false

    foreach ($file in $XmlFiles) {
        $xmlPath = $file.FullName
        $changes = @(Load-ChangesForFile -XmlPath $xmlPath)
        if ($changes.Count -eq 0) { continue }

        $pending = @(
            $changes | Where-Object {
                $p = $_.PSObject.Properties['Status']
                if (-not $p) { return $true }
                $val = [string]$p.Value
                if ([string]::IsNullOrWhiteSpace($val)) { return $true }
                return ($val -ne 'Committed')
            }
        )

        if ($pending.Count -eq 0) { continue }

        try {
            [xml]$xmlDoc = Load-XmlDocument -Path $xmlPath
        } catch {
            Write-Bad ("Could not read XML '{0}': {1}" -f $xmlPath, $_.Exception.Message)
            continue
        }

        $appliedCount = 0
        foreach ($c in $pending) {
            $ok = $false
            try {
                $ok = Apply-SingleSavedChange -XmlRoot $xmlDoc -Change $c -XmlPath $xmlPath
            } catch {
                $ok = $false
                Write-Bad ("Error applying pending change in {0}: {1}" -f $file.Name, $_.Exception.Message)
            }
            if ($ok) { $appliedCount++ }
        }

        if ($appliedCount -gt 0) {
            Backup-XmlFile -Path $xmlPath
            Save-XmlDocument -Xml $xmlDoc -Path $xmlPath
            Mark-AllPendingChangesCommittedForFile -XmlPath $xmlPath

            Write-Good ("Applied {0} pending change{1} and saved '{2}'" -f `
                $appliedCount,
                $(if ($appliedCount -eq 1) { "" } else { "s" }),
                $file.Name)

            $anyApplied = $true
        }
    }

    if (-not $anyApplied) {
        Write-Info "There were no pending changes to apply."
    }
}

function Get-SharedConfigsFolderForXmlFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    if (-not $XmlFiles -or $XmlFiles.Count -eq 0) {
        return $null
    }

    $rootDir    = Split-Path $XmlFiles[0].FullName -Parent
    $helperRoot = Join-Path $rootDir "LSR-XML-Helper"
    if (-not (Test-Path $helperRoot)) {
        New-Item -ItemType Directory -Path $helperRoot -Force | Out-Null
    }

    $sharedDir = Join-Path $helperRoot "Shared-Configs"

    if (-not (Test-Path $sharedDir)) {
        New-Item -ItemType Directory -Path $sharedDir -Force | Out-Null
    }

    return $sharedDir
}

function Add-ChangeRecordForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Change
    )

    $changesPath = Get-ChangeLogPathForXml -XmlPath $XmlPath
    $changesDir  = Split-Path $changesPath -Parent

    if (-not (Test-Path $changesDir)) {
        New-Item -ItemType Directory -Path $changesDir -Force | Out-Null
    }

    $existing = @()
    if (Test-Path $changesPath) {
        try {
            $json = Get-Content -Path $changesPath -Raw
            if (-not [string]::IsNullOrWhiteSpace($json)) {
                $data = $json | ConvertFrom-Json
                if ($data -is [System.Collections.IEnumerable]) {
                    $existing = @($data)
                } else {
                    $existing = @($data)
                }
            }
        } catch {
            Write-Bad "Could not read existing change log '$changesPath': $($_.Exception.Message)"
        }
    }

    if (-not ($Change.PSObject.Properties.Name -contains 'Status')) {
        $Change | Add-Member -NotePropertyName 'Status' -NotePropertyValue 'Pending' -Force
    } elseif ([string]::IsNullOrWhiteSpace([string]$Change.Status)) {
        $Change.Status = 'Pending'
    }

    $all = @($existing + $Change)
    try {
        $all | ConvertTo-Json -Depth 8 | Set-Content -Path $changesPath -Encoding UTF8
        Write-Info "Recorded change in '$changesPath'."
    } catch {
        Write-Bad "Failed to write change log '$changesPath': $($_.Exception.Message)"
    }
}

function Export-SharedConfigPack {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    $items = @()
    foreach ($file in $XmlFiles) {
        $xmlPath = $file.FullName
        $changes = @(Load-ChangesForFile -XmlPath $xmlPath)
        if ($changes.Count -gt 0) {
            $items += [pscustomobject]@{
                File    = $file
                XmlPath = $xmlPath
                Changes = $changes
                Count   = $changes.Count
            }
        }
    }

    if ($items.Count -eq 0) {
        Write-Bad "There are no saved edits to export in this folder."
        Start-Sleep -Seconds 2
        return
    }

    Clear-Screen
    Write-Title "Export saved edits as a shared config pack"

    Write-Host "The following XML files have saved edits:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $items.Count; $i++) {
        $idx  = $i + 1
        $item = $items[$i]
        Write-Host ("[{0}] {1} ({2} change{3})" -f `
            $idx,
            $item.File.Name,
            $item.Count,
            $(if ($item.Count -eq 1) { "" } else { "s" }))
    }
    Write-Host ""

    Write-Host "Export options:"
    Write-Host "[1] Export ALL XML files listed above"
    Write-Host "[2] Choose which XML files to include"
    Write-Host "[Q] Cancel"
    Write-Host ""

    $mode = Read-Host "Your choice"
    if ($mode -match '^[Qq]$') {
        Write-Info "Export cancelled."
        Start-Sleep -Seconds 1
        return
    }

    $selectedItems = $items

    if ($mode -eq '2') {
        $nums = Read-Host "Enter the numbers of the XML files to include (example: 1,3,5)"
        if ([string]::IsNullOrWhiteSpace($nums)) {
            Write-Bad "No selection entered. Export cancelled."
            Start-Sleep -Seconds 2
            return
        }

        $indices = @()
        foreach ($part in ($nums -split ',')) {
            $trim = $part.Trim()
            if ($trim -as [int]) {
                $n = [int]$trim
                if ($n -ge 1 -and $n -le $items.Count) {
                    $indices += $n
                }
            }
        }

        if ($indices.Count -eq 0) {
            Write-Bad "No valid XML numbers found in your input. Export cancelled."
            Start-Sleep -Seconds 2
            return
        }

        $selectedItems = @()
        foreach ($n in $indices | Sort-Object -Unique) {
            $selectedItems += $items[$n - 1]
        }
    }
    elseif ($mode -ne '1') {
        Write-Bad "Not a valid choice. Export cancelled."
        Start-Sleep -Seconds 2
        return
    }

    $sharedDir = Get-SharedConfigsFolderForXmlFiles -XmlFiles $XmlFiles
    if (-not $sharedDir) {
        Write-Bad "Could not determine Shared-Configs folder. Export cancelled."
        Start-Sleep -Seconds 2
        return
    }

    Write-Host ""
    Write-Info ("Shared configs will be saved under: {0}" -f $sharedDir)

    $description = Read-Host "Optional description for this config pack (leave blank for none)"
    $nameInput   = Read-Host "File name for this config pack (leave blank for automatic name)"

    $timeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    if ([string]::IsNullOrWhiteSpace($nameInput)) {
        $fileName = "ConfigPack_{0}.json" -f $timeStamp
    } else {
        $safeName = ($nameInput -replace '[^\w\-.]+', '_').Trim()
        if ([string]::IsNullOrWhiteSpace($safeName)) {
            $safeName = "ConfigPack_{0}" -f $timeStamp
        }
        if (-not $safeName.ToLower().EndsWith(".json")) {
            $safeName += ".json"
        }
        $fileName = $safeName
    }

    $outPath = Join-Path $sharedDir $fileName

    try {
        $pack = [pscustomobject]@{
            Tool          = "LSR-XML-Helper"
            FormatVersion = 1
            CreatedUtc    = (Get-Date).ToUniversalTime().ToString("o")
            Description   = $description
            XmlConfigs    = @()
        }

        foreach ($item in $selectedItems) {
            $pack.XmlConfigs += [pscustomobject]@{
                FileName = $item.File.Name
                Changes  = $item.Changes
            }
        }

        $json = $pack | ConvertTo-Json -Depth 10

        if (-not (Test-Path $sharedDir)) {
            New-Item -ItemType Directory -Path $sharedDir -Force | Out-Null
        }

        $json | Set-Content -Path $outPath -Encoding UTF8
        Write-Good "Exported config pack to: $outPath"
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Bad "Failed to export config pack: $($_.Exception.Message)"
        Start-Sleep -Seconds 3
    }
}

function Export-SavedEditsSummary {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    function Get-PropValue {
        param(
            [Parameter(Mandatory = $true)] $Obj,
            [Parameter(Mandatory = $true)] [string] $Name
        )

        $p = $Obj.PSObject.Properties[$Name]
        if ($p -and $null -ne $p.Value) {
            return [string]$p.Value
        }
        return ""
    }

    $groups = @()

    foreach ($file in $XmlFiles) {
        $xmlPath = $file.FullName
        $changes = @(Load-ChangesForFile -XmlPath $xmlPath)
        if ($changes.Count -eq 0) { continue }

        $groups += [pscustomobject]@{
            File    = $file
            XmlPath = $xmlPath
            Changes = $changes
        }
    }

    if ($groups.Count -eq 0) {
        Write-Bad "No saved edits found to export a summary."
        Start-Sleep -Seconds 2
        return
    }

    $sharedDir = Get-SharedConfigsFolderForXmlFiles -XmlFiles $XmlFiles
    if (-not $sharedDir) {
        Write-Bad "Could not determine Shared-Configs folder."
        Start-Sleep -Seconds 2
        return
    }

    Write-Host ""
    Write-Info ("Summaries will be saved under: {0}" -f $sharedDir)

    $description = Read-Host "Optional description for this summary (leave blank for none)"
    $nameInput   = Read-Host "File name for this summary (leave blank for automatic name)"

    $timeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    if ([string]::IsNullOrWhiteSpace($nameInput)) {
        $fileName = "SavedEdits_Summary_{0}.txt" -f $timeStamp
    } else {
        $safeName = ($nameInput -replace '[^\w\-.]+', '_').Trim()
        if ([string]::IsNullOrWhiteSpace($safeName)) {
            $safeName = "SavedEdits_Summary_{0}" -f $timeStamp
        }
        if (-not $safeName.ToLower().EndsWith(".txt")) {
            $safeName += ".txt"
        }
        $fileName = $safeName
    }

    $outPath = Join-Path $sharedDir $fileName

    $lines = @()
    $lines += "LSR-XML-Helper Saved Edits Summary"
    $lines += ("Created: {0}" -f (Get-Date))

    if (-not [string]::IsNullOrWhiteSpace($description)) {
        $lines += ("Description: {0}" -f $description)
    }

    $lines += ""

    foreach ($g in ($groups | Sort-Object { $_.File.Name })) {

        $byStatus = @{
            Pending   = @()
            Committed = @()
        }

        foreach ($c in $g.Changes) {
            if ($null -eq $c) { continue }

            $status = 'Pending'
            $sv = Get-PropValue -Obj $c -Name 'Status'
            if (-not [string]::IsNullOrWhiteSpace($sv) -and $sv -ieq 'Committed') {
                $status = 'Committed'
            }

            $byStatus[$status] += $c
        }

        foreach ($statusKey in @('Pending','Committed')) {
            $set = @($byStatus[$statusKey])
            if ($set.Count -eq 0) { continue }

            $lines += ("=" * 60)
            $lines += ("File: {0}" -f $g.File.Name)
            $lines += ("Status: {0}" -f $statusKey)
            $lines += ("-" * 60)
            $lines += ""

            $idx = 0
            foreach ($c in $set) {
                $idx++

                $type = Get-PropValue -Obj $c -Name 'Type'

                if ($type -eq 'EditField') {
                    $lines += ("[{0}] EditField" -f $idx)
                    $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
                    $lines += ("    Entry  : {0}" -f (Get-PropValue -Obj $c -Name 'EntryIndex'))
                    $lines += ("    Field  : {0}" -f (Get-PropValue -Obj $c -Name 'FieldPath'))
                    $lines += ("    Old    : {0}" -f (Get-PropValue -Obj $c -Name 'OldValue'))
                    $lines += ("    New    : {0}" -f (Get-PropValue -Obj $c -Name 'NewValue'))
                    $lines += ""
                    continue
                }

                if ($type -eq 'AddEntry') {
                    $lines += ("[{0}] AddEntry" -f $idx)
                    $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
                    $lines += ("    Entry  : (new)")
                    $lines += ("    Field  :")
                    $lines += ("    Old    :")
                    $lines += ("    New    :")
                    $lines += ""
                    continue
                }

                $lines += ("[{0}] {1}" -f $idx, $type)
                $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
                $lines += ("    Entry  : {0}" -f (Get-PropValue -Obj $c -Name 'EntryIndex'))
                $lines += ("    Field  : {0}" -f (Get-PropValue -Obj $c -Name 'FieldPath'))
                $lines += ("    Old    : {0}" -f (Get-PropValue -Obj $c -Name 'OldValue'))
                $lines += ("    New    : {0}" -f (Get-PropValue -Obj $c -Name 'NewValue'))
                $lines += ""
            }

            $lines += ""
        }
    }

    try {
        $lines | Set-Content -Path $outPath -Encoding UTF8
        Write-Good "Summary exported to: $outPath"
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Bad "Failed to export summary: $($_.Exception.Message)"
        Write-LogError $_
        Start-Sleep -Seconds 3
    }
}

function Export-SavedEditsSummaryForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath,

        [Parameter(Mandatory = $true)]
        [object[]]$SavedChanges
    )

    function Get-PropValue {
        param(
            [Parameter(Mandatory = $true)] $Obj,
            [Parameter(Mandatory = $true)] [string] $Name
        )

        $p = $Obj.PSObject.Properties[$Name]
        if ($p -and $null -ne $p.Value) {
            return [string]$p.Value
        }
        return ""
    }

    if (-not (Test-Path $XmlPath)) {
        Write-Bad "XML not found: $XmlPath"
        Start-Sleep -Seconds 2
        return
    }

    $changes = @()
    if ($SavedChanges) {
        $changes = @($SavedChanges | Where-Object { $_ -ne $null })
    }

    if ($changes.Count -eq 0) {
        Write-Bad "No saved edits found to export a summary for this XML."
        Start-Sleep -Seconds 2
        return
    }

    $helperRoot = Get-HelperRootForXmlPath -XmlPath $XmlPath
    $sharedDir  = Join-Path $helperRoot "Shared-Configs"
    if (-not (Test-Path $sharedDir)) {
        New-Item -ItemType Directory -Path $sharedDir -Force | Out-Null
    }

    Write-Host ""
    Write-Info ("Summaries will be saved under: {0}" -f $sharedDir)

    $description = Read-Host "Optional description for this summary (leave blank for none)"
    $nameInput   = Read-Host "File name for this summary (leave blank for automatic name)"

    $timeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    if ([string]::IsNullOrWhiteSpace($nameInput)) {
        $fileName = "SavedEdits_Summary_{0}.txt" -f $timeStamp
    } else {
        $safeName = ($nameInput -replace '[^\w\-.]+', '_').Trim()
        if ([string]::IsNullOrWhiteSpace($safeName)) {
            $safeName = "SavedEdits_Summary_{0}" -f $timeStamp
        }
        if (-not $safeName.ToLower().EndsWith(".txt")) {
            $safeName += ".txt"
        }
        $fileName = $safeName
    }

    $outPath = Join-Path $sharedDir $fileName

    $byStatus = @{
        Pending   = @()
        Committed = @()
    }

    foreach ($c in $changes) {
        if ($null -eq $c) { continue }

        $status = 'Pending'
        $sv = Get-PropValue -Obj $c -Name 'Status'
        if (-not [string]::IsNullOrWhiteSpace($sv) -and $sv -ieq 'Committed') {
            $status = 'Committed'
        }

        $byStatus[$status] += $c
    }

    $lines = @()
    $lines += "LSR-XML-Helper Saved Edits Summary (Single XML)"
    $lines += ("Created: {0}" -f (Get-Date))
    $lines += ("XML: {0}" -f (Split-Path $XmlPath -Leaf))

    if (-not [string]::IsNullOrWhiteSpace($description)) {
        $lines += ("Description: {0}" -f $description)
    }

    $lines += ""

    foreach ($statusKey in @('Pending','Committed')) {
        $set = @($byStatus[$statusKey])
        if ($set.Count -eq 0) { continue }

        $lines += ("=" * 60)
        $lines += ("File: {0}" -f (Split-Path $XmlPath -Leaf))
        $lines += ("Status: {0}" -f $statusKey)
        $lines += ("-" * 60)
        $lines += ""

        $idx = 0
        foreach ($c in $set) {
            $idx++

            $type = Get-PropValue -Obj $c -Name 'Type'

            if ($type -eq 'EditField') {
                $lines += ("[{0}] EditField" -f $idx)
                $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
                $lines += ("    Entry  : {0}" -f (Get-PropValue -Obj $c -Name 'EntryIndex'))
                $lines += ("    Field  : {0}" -f (Get-PropValue -Obj $c -Name 'FieldPath'))
                $lines += ("    Old    : {0}" -f (Get-PropValue -Obj $c -Name 'OldValue'))
                $lines += ("    New    : {0}" -f (Get-PropValue -Obj $c -Name 'NewValue'))
                $lines += ""
                continue
            }

            if ($type -eq 'AddEntry') {
                $lines += ("[{0}] AddEntry" -f $idx)
                $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
                $lines += ("    Entry  : (new)")
                $lines += ("    Field  :")
                $lines += ("    Old    :")
                $lines += ("    New    :")
                $lines += ""
                continue
            }

            $lines += ("[{0}] {1}" -f $idx, $type)
            $lines += ("    Type   : {0}" -f (Get-PropValue -Obj $c -Name 'TypeName'))
            $lines += ("    Entry  : {0}" -f (Get-PropValue -Obj $c -Name 'EntryIndex'))
            $lines += ("    Field  : {0}" -f (Get-PropValue -Obj $c -Name 'FieldPath'))
            $lines += ("    Old    : {0}" -f (Get-PropValue -Obj $c -Name 'OldValue'))
            $lines += ("    New    : {0}" -f (Get-PropValue -Obj $c -Name 'NewValue'))
            $lines += ""
        }

        $lines += ""
    }

    try {
        $lines | Set-Content -Path $outPath -Encoding UTF8
        Write-Good "Summary exported to: $outPath"
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Bad "Failed to export summary: $($_.Exception.Message)"
        Write-LogError $_
        Start-Sleep -Seconds 3
    }
}

function Import-SharedConfigPack {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    Clear-Screen
    Write-Title "Import a shared config pack"

    $defaultDir = Get-SharedConfigsFolderForXmlFiles -XmlFiles $XmlFiles
    $inPath = $null

    if ($defaultDir -and (Test-Path $defaultDir)) {
        $packs = Get-ChildItem -Path $defaultDir -Filter "*.json" -File | Sort-Object LastWriteTime -Descending
        if ($packs.Count -gt 0) {
            Write-Host "Shared config packs found in:" -ForegroundColor Cyan
            Write-Host "  $defaultDir"
            Write-Host ""
            for ($i = 0; $i -lt $packs.Count; $i++) {
                $idx = $i + 1
                Write-Host ("[{0}] {1} (last modified: {2})" -f `
                    $idx,
                    $packs[$i].Name,
                    $packs[$i].LastWriteTime)
            }
            Write-Host ""
            Write-Host "[1] Pick a pack from the list above"
            Write-Host "[2] Enter a custom path manually"
            Write-Host "[Q] Cancel"
            Write-Host ""

            $pickMode = Read-Host "Your choice"
            if ($pickMode -match '^[Qq]$') {
                Write-Info "Import cancelled."
                Start-Sleep -Seconds 1
                return
            }
            elseif ($pickMode -eq '1') {
                $nStr = Read-Host "Enter the number of the pack to use"
                if ($nStr -as [int]) {
                    $n = [int]$nStr
                    if ($n -ge 1 -and $n -le $packs.Count) {
                        $inPath = $packs[$n - 1].FullName
                    }
                }
                if (-not $inPath) {
                    Write-Bad "Invalid selection. Import cancelled."
                    Start-Sleep -Seconds 2
                    return
                }
            }
            elseif ($pickMode -eq '2') {
                $inPath = Read-Host "Enter full path to a shared config pack (.json)"
            }
            else {
                Write-Bad "Not a valid choice. Import cancelled."
                Start-Sleep -Seconds 2
                return
            }
        }
        else {
            $inPath = Read-Host "Enter full path to a shared config pack (.json)"
        }
    }
    else {
        $inPath = Read-Host "Enter full path to a shared config pack (.json)"
    }

    if ([string]::IsNullOrWhiteSpace($inPath)) {
        Write-Bad "No file path entered. Import cancelled."
        Start-Sleep -Seconds 2
        return
    }

    if (-not (Test-Path $inPath)) {
        Write-Bad "File not found: $inPath"
        Start-Sleep -Seconds 2
        return
    }

    try {
        $json = Get-Content -Path $inPath -Raw
        $pack = $json | ConvertFrom-Json
    }
    catch {
        Write-Bad "Could not read config pack: $($_.Exception.Message)"
        Start-Sleep -Seconds 3
        return
    }

    if (-not $pack -or -not $pack.XmlConfigs) {
        Write-Bad "This file does not look like a valid config pack."
        Start-Sleep -Seconds 3
        return
    }

    Write-Host ""
    if ($pack.Tool) {
        Write-Info ("Config pack tool: {0}" -f $pack.Tool)
    }
    if ($pack.FormatVersion) {
        Write-Info ("Format version   : {0}" -f $pack.FormatVersion)
    }
    if ($pack.Description) {
        Write-Info ("Description      : {0}" -f $pack.Description)
    }
    Write-Host ""

    $cfgs = @($pack.XmlConfigs)

    Write-Host "Import options:"
    Write-Host "[1] Import ALL sections in this config pack"
    Write-Host "[2] Choose which sections (XML files) to import"
    Write-Host "[Q] Cancel"
    Write-Host ""

    $mode = Read-Host "Your choice"
    if ($mode -match '^[Qq]$') {
        Write-Info "Import cancelled."
        Start-Sleep -Seconds 2
        return
    }

    $cfgsToImport = $cfgs

    if ($mode -eq '2') {
        Write-Host ""
        Write-Host "Sections in this pack:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $cfgs.Count; $i++) {
            $idx = $i + 1
            Write-Host ("[{0}] {1}" -f $idx, $cfgs[$i].FileName)
        }
        Write-Host ""
        $nums = Read-Host "Enter the numbers of the sections to import (example: 1,3,4)"
        if ([string]::IsNullOrWhiteSpace($nums)) {
            Write-Bad "No sections selected. Import cancelled."
            Start-Sleep -Seconds 2
            return
        }

        $indices = @()
        foreach ($part in ($nums -split ',')) {
            $trim = $part.Trim()
            if ($trim -as [int]) {
                $n = [int]$trim
                if ($n -ge 1 -and $n -le $cfgs.Count) {
                    $indices += $n
                }
            }
        }

        if ($indices.Count -eq 0) {
            Write-Bad "No valid section numbers found. Import cancelled."
            Start-Sleep -Seconds 2
            return
        }

        $cfgsToImport = @()
        foreach ($n in $indices | Sort-Object -Unique) {
            $cfgsToImport += $cfgs[$n - 1]
        }
    }
    elseif ($mode -ne '1') {
        Write-Bad "Not a valid choice. Import cancelled."
        Start-Sleep -Seconds 2
        return
    }

    Write-Host ""
    Write-Host "[1] Yes, add these edits into my local LSR-Changes logs"
    Write-Host "[2] No, cancel"
    Write-Host ""
    $apply = Read-Host "Your choice"
    if ($apply -ne '1') {
        Write-Info "Import cancelled."
        Start-Sleep -Seconds 2
        return
    }

    foreach ($cfg in $cfgsToImport) {
        $fileName = $cfg.FileName
        $target   = $XmlFiles | Where-Object { $_.Name -eq $fileName } | Select-Object -First 1
        if (-not $target) {
            Write-Bad ("Skipping '{0}': no matching XML file in this folder." -f $fileName)
            continue
        }

        $xmlPath  = $target.FullName
        $existing = @(Load-ChangesForFile -XmlPath $xmlPath)
        $incoming = @($cfg.Changes)

        if ($incoming.Count -eq 0) {
            continue
        }

        $merged = @($existing + $incoming)
        try {
            Set-ChangesForFile -XmlPath $xmlPath -Changes $merged
            Write-Good ("Imported {0} change{1} into {2}" -f `
                $incoming.Count,
                $(if ($incoming.Count -eq 1) { "" } else { "s" }),
                $fileName)
        }
        catch {
            Write-Bad ("Failed to update changes for {0}: {1}" -f $fileName, $_.Exception.Message)
        }
    }

    Write-Host ""
    Write-Info "You can now use 'Review saved edits' to inspect and apply the imported changes."
    Start-Sleep -Seconds 3
}

function Backup-XmlFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        Write-Bad "File not found, cannot backup: $Path"
        Start-Sleep -Seconds 1
        return
    }

    $helperRoot = Get-HelperRootForXmlPath -XmlPath $Path
    $backupDir  = Join-Path $helperRoot "BackupXMLs"

    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

    $name     = Split-Path $Path -Leaf
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($name)

    $culture   = [System.Globalization.CultureInfo]::CurrentCulture
    $now       = Get-Date
    $stampRaw  = $now.ToString("g", $culture)
    $stampSafe = ($stampRaw -replace '[^0-9A-Za-z]+','-').Trim('-')

    $backupName = "{0}_{1}.xml" -f $baseName, $stampSafe
    $backupPath = Join-Path $backupDir $backupName
    $idx = 1
    while (Test-Path $backupPath) {
        $backupName = "{0}_{1}_{2}.xml" -f $baseName, $stampSafe, $idx
        $backupPath = Join-Path $backupDir $backupName
        $idx++
    }

    Copy-Item -Path $Path -Destination $backupPath -Force
    Write-Good "Backup created: $backupPath"
    Start-Sleep -Seconds 2
}

function Get-BackupFilesForXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $helperRoot = Get-HelperRootForXmlPath -XmlPath $XmlPath
    $backupDir  = Join-Path $helperRoot "BackupXMLs"

    if (-not (Test-Path $backupDir)) {
        return @()
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($XmlPath)

    $files = Get-ChildItem -Path $backupDir -Filter "*.xml" -File -ErrorAction SilentlyContinue

    $filtered = @(
        $files | Where-Object {
            $_.BaseName -eq $baseName -or $_.BaseName -like ($baseName + "_*")
        } | Sort-Object LastWriteTime -Descending
    )

    return $filtered
}

function Restore-BackupToXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath,

        [Parameter(Mandatory = $true)]
        [string]$BackupPath
    )

    if (-not (Test-Path $XmlPath)) {
        Write-Bad "Target XML not found: $XmlPath"
        Start-Sleep -Seconds 2
        return
    }

    if (-not (Test-Path $BackupPath)) {
        Write-Bad "Backup file not found: $BackupPath"
        Start-Sleep -Seconds 2
        return
    }

    $confirm = Read-Host "This will OVERWRITE the XML with the selected backup. Type RESTORE to continue"
    if ($confirm -ne "RESTORE") {
        Write-Info "Restore cancelled."
        Start-Sleep -Seconds 1
        return
    }

    Backup-XmlFile -Path $XmlPath

    try {
        Copy-Item -Path $BackupPath -Destination $XmlPath -Force
        Write-Good "Restored backup into: $XmlPath"
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Bad "Restore failed: $($_.Exception.Message)"
        Write-LogError $_
        Start-Sleep -Seconds 3
    }
}

function Get-AllBackupItemsForXmlFiles {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    $items = @()

    foreach ($f in $XmlFiles) {
        $xmlPath = $f.FullName
        $backups = @(Get-BackupFilesForXml -XmlPath $xmlPath)

        foreach ($b in $backups) {
            $items += [pscustomobject]@{
                XmlFileName = $f.Name
                XmlPath     = $xmlPath
                BackupFile  = $b
            }
        }
    }

    return @($items | Sort-Object { $_.BackupFile.LastWriteTime } -Descending)
}

function Show-BackupBrowser {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    while ($true) {
        Clear-Screen
        Write-Title "Restore Backups"

        Write-Host "[A] View ALL backups (across all XMLs)" -ForegroundColor Cyan

        for ($i = 0; $i -lt $XmlFiles.Count; $i++) {
            $idx = $i + 1
            Write-Host ("[{0}] {1}" -f $idx, $XmlFiles[$i].Name) -ForegroundColor Cyan
        }

        Write-Host ""
        Write-Host "Pick A for all backups, or a number to view backups for one XML."
        Write-Host "Type Q to go back."
        Write-Host ""

        $pick = Read-Host "Your choice"
        if ($pick -match '^[Qq]$') { return }

        if ($pick -match '^[Aa]$') {

            while ($true) {
                $items = @(Get-AllBackupItemsForXmlFiles -XmlFiles $XmlFiles)

                Clear-Screen
                Write-Title "All Backups"

                if ($items.Count -eq 0) {
                    Write-Bad "No backups found."
                    Read-Host "Press Enter to go back" | Out-Null
                    break
                }

                for ($i = 0; $i -lt $items.Count; $i++) {
                    $idx    = $i + 1
                    $bkName = $items[$i].BackupFile.Name
                    $bkTime = $items[$i].BackupFile.LastWriteTime
                    $xml    = $items[$i].XmlFileName

                    Write-Host ("[{0}] {1}  ({2})" -f $idx, $bkName, $bkTime) -ForegroundColor Yellow
                    Write-Host ("     -> {0}" -f $xml) -ForegroundColor Cyan
                }

                Write-Host ""
                Write-Host "Pick a backup number to restore."
                Write-Host "Type Q to go back."
                Write-Host ""

                $bPick = Read-Host "Your choice"
                if ($bPick -match '^[Qq]$') { break }

                if (-not ($bPick -as [int])) {
                    Write-Bad "Not a valid number."
                    Start-Sleep -Seconds 1
                    continue
                }

                $bn = [int]$bPick
                if ($bn -lt 1 -or $bn -gt $items.Count) {
                    Write-Bad "Out of range."
                    Start-Sleep -Seconds 1
                    continue
                }

                $chosen = $items[$bn - 1]
                Restore-BackupToXml -XmlPath $chosen.XmlPath -BackupPath $chosen.BackupFile.FullName
            }

            continue
        }

        if (-not ($pick -as [int])) {
            Write-Bad "Not a valid option."
            Start-Sleep -Seconds 1
            continue
        }

        $n = [int]$pick
        if ($n -lt 1 -or $n -gt $XmlFiles.Count) {
            Write-Bad "Out of range."
            Start-Sleep -Seconds 1
            continue
        }

        $xmlPath = $XmlFiles[$n - 1].FullName
        $backups = @(Get-BackupFilesForXml -XmlPath $xmlPath)

        while ($true) {
            Clear-Screen
            Write-Title ("Backups for: {0}" -f ([System.IO.Path]::GetFileName($xmlPath)))

            if ($backups.Count -eq 0) {
                Write-Bad "No backups found for this XML."
                Read-Host "Press Enter to go back" | Out-Null
                break
            }

            for ($i = 0; $i -lt $backups.Count; $i++) {
                $idx = $i + 1
                Write-Host ("[{0}] {1}  ({2})" -f $idx, $backups[$i].Name, $backups[$i].LastWriteTime) -ForegroundColor Yellow
            }

            Write-Host ""
            Write-Host "Pick a backup number to restore it into the XML."
            Write-Host "Type Q to go back."
            Write-Host ""

            $bPick = Read-Host "Your choice"
            if ($bPick -match '^[Qq]$') { break }

            if (-not ($bPick -as [int])) {
                Write-Bad "Not a valid number."
                Start-Sleep -Seconds 1
                continue
            }

            $bn = [int]$bPick
            if ($bn -lt 1 -or $bn -gt $backups.Count) {
                Write-Bad "Out of range."
                Start-Sleep -Seconds 1
                continue
            }

            Restore-BackupToXml -XmlPath $xmlPath -BackupPath $backups[$bn - 1].FullName
        }
    }
}

function Get-XmlFiles {
    param(
        [Parameter(Mandatory = $true)][string]$Folder
    )

    if (-not (Test-Path $Folder)) {
        throw "Root folder does not exist: $Folder"
    }

    Get-ChildItem -Path $Folder -Filter "*.xml" -File | Sort-Object Name
}

function Select-FileFromList {
    param(
        [Parameter(Mandatory = $true)][System.IO.FileInfo[]]$Files
    )

    if ($Files.Count -eq 0) {
        Write-Bad "No XML files found in that folder."
        return $null
    }

    while ($true) {
        Clear-Screen
        Write-Title "Pick an XML file to work with"

        for ($i = 0; $i -lt $Files.Count; $i++) {
            $index = $i + 1
            Write-Host "[$index] $($Files[$i].Name)"
        }

        Write-Host ""
        Write-Host "Type the number of the file you want. Example: 3"
        Write-Host "Or type Q to cancel."

        $choice = Read-Host "Your choice"
        if ($choice -match '^[Qq]$') { return $null }

        if ($choice -as [int]) {
            $num = [int]$choice
            if ($num -ge 1 -and $num -le $Files.Count) {
                return $Files[$num - 1]
            }
        }

        Write-Bad "Not a valid number. Try again."
        Start-Sleep -Seconds 1
    }
}

function Load-XmlDocument {
    param(
        [Parameter(Mandatory = $true)][string]$Path
    )
    try {
        [xml]$xml = Get-Content -Path $Path -Raw
        return $xml
    } catch {
        throw ("Could not read XML from ${Path}: $($_.Exception.Message)")
    }
}

function Save-XmlDocument {
    param(
        [Parameter(Mandatory = $true)][xml]$Xml,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $Xml.Save($Path)
    Write-Good "Saved changes to: $Path"
}

function Get-EntryTypes {
    param(
        [Parameter(Mandatory = $true)][xml]$Xml
    )

    $root = $Xml.DocumentElement
    if (-not $root) { return @() }

    $nodes = $root.SelectNodes("*") |
             Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }

    $groups = $nodes | Group-Object LocalName | Sort-Object Name
    return $groups
}

function Get-EntriesOfType {
    param(
        [Parameter(Mandatory = $true)][xml]$Xml,
        [Parameter(Mandatory = $true)][string]$ElementName
    )

    $root = $Xml.DocumentElement
    if (-not $root) { return @() }

    $rootChildren = @(
        $root.ChildNodes |
        Where-Object {
            $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and
            $_.LocalName -eq $ElementName
        }
    )

    function Get-EntriesFromContainer {
        param([System.Xml.XmlElement]$container)
        $elementChildren = @(
            $container.ChildNodes |
            Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }
        )

        if ($elementChildren.Count -gt 0) {
            $uniqueNames = $elementChildren |
                Select-Object -ExpandProperty LocalName -Unique

            if ($uniqueNames.Count -eq 1) {
                return $elementChildren
            } else {
                return @($container)
            }
        }

        $hasAttributes = ($container.Attributes -and $container.Attributes.Count -gt 0)
        $leafNodes = @(
            $container.SelectNodes(".//*[not(*)]") |
            Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }
        )

        if ($hasAttributes -or $leafNodes.Count -gt 0) {
            return @($container)
        }

        return @()
    }

    if ($rootChildren.Count -gt 1) {
        return $rootChildren
    }

    if ($rootChildren.Count -eq 1) {
        return Get-EntriesFromContainer -container $rootChildren[0]
    }

    $container = $root.SelectSingleNode($ElementName)
    if (-not $container) { return @() }

    return Get-EntriesFromContainer -container $container
}

function Get-EntryFieldsList {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlElement]$Entry
    )

    $fields = @()

    if ($Entry.Attributes) {
        foreach ($attr in $Entry.Attributes) {
            $fields += [pscustomobject]@{
                Kind    = 'Attribute'
                Path    = $attr.Name
                Display = "[Attr] $($attr.Name) = '$($attr.Value)'"
                Node    = $attr
            }
        }
    }

    $leafNodes = New-Object System.Collections.ArrayList

    $hasElementChild = $false
    foreach ($c in $Entry.ChildNodes) {
        if ($c.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $hasElementChild = $true
            break
        }
    }

    if (-not $hasElementChild -and -not [string]::IsNullOrWhiteSpace($Entry.InnerText)) {
        [void]$leafNodes.Add($Entry)
    }

    $descLeaves = $Entry.SelectNodes(".//*[not(*)]")
    foreach ($n in $descLeaves) {
        [void]$leafNodes.Add($n)
    }

    foreach ($node in $leafNodes) {
        if (-not $node) { continue }
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }

        if ($node -eq $Entry) {
            $path = $Entry.LocalName
        } else {
            $names   = @()
            $current = $node
            while ($current -and $current -ne $Entry) {
                $names   = ,$current.LocalName + $names
                $current = $current.ParentNode
            }
            $path = $names -join "."
        }

        if ([string]::IsNullOrWhiteSpace($path)) { continue }

        $value = $node.InnerText

        $fields += [pscustomobject]@{
            Kind    = 'Element'
            Path    = $path
            Display = "$path = '$value'"
            Node    = $node
        }
    }

    $pathCounts = @{}
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $p = $fields[$i].Path
        if (-not $pathCounts.ContainsKey($p)) {
            $pathCounts[$p] = 0
        }
        $idx = $pathCounts[$p]
        $pathCounts[$p] = $idx + 1

        $fieldKey = "{0}#{1}" -f $p, $idx
        $fields[$i] | Add-Member -NotePropertyName 'FieldKey' -NotePropertyValue $fieldKey -Force
    }

    return $fields
}

function Get-EntrySummary {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Entry
    )

    $fields = @(Get-EntryFieldsList -Entry $Entry)

    if ($fields.Count -eq 0) {
        return "(no fields on this entry)"
    }

    $take  = [Math]::Min(3, $fields.Count)
    $parts = @()
    for ($i = 0; $i -lt $take; $i++) {
        $parts += $fields[$i].Display
    }

    return ($parts -join "; ")
}

function Apply-EditFieldChange {
    param(
        [xml]$Xml,
        [string]$XmlPath,
        [pscustomobject]$Change,
        [switch]$ValidateOnly
    )

    $typeName   = $Change.TypeName
    $entryIndex = [int]$Change.EntryIndex

    $entries = @(Get-EntriesOfType -Xml $Xml -ElementName $typeName)
    if ($entries.Count -eq 0) {
        Write-Bad "Cannot apply edit: type '$typeName' no longer exists in '$XmlPath'."
        return $false
    }

    if ($entryIndex -lt 1 -or $entryIndex -gt $entries.Count) {
        Write-Bad "Cannot apply edit: entry index $entryIndex for type '$typeName' is out of range in '$XmlPath'."
        return $false
    }

    $entry  = [System.Xml.XmlElement]$entries[$entryIndex - 1]
    $fields = Get-EntryFieldsList -Entry $entry
    $field  = $fields | Where-Object { $_.Path -eq $Change.FieldPath } | Select-Object -First 1

    if (-not $field) {
        Write-Bad "Cannot apply edit: field '$($Change.FieldPath)' not found in type '$typeName' entry #$entryIndex in '$XmlPath'."
        return $false
    }

    $currentValue = if ($field.Kind -eq 'Attribute') { $field.Node.Value } else { $field.Node.InnerText }

    if ($Change.PSObject.Properties.Match('OldValue') -and $Change.OldValue -ne $null) {
        if ($currentValue -ne $Change.OldValue) {
            Write-Bad "Skipping edit for '$typeName' entry #$entryIndex, field '$($Change.FieldPath)' in '$XmlPath': expected old value '$($Change.OldValue)', found '$currentValue'."
            return $false
        }
    }

    if ($ValidateOnly) {
        Write-Good "Would apply edit: $typeName entry #$entryIndex, $($Change.FieldPath) -> '$($Change.NewValue)'"
        return $true
    }

    if ($field.Kind -eq 'Attribute') {
        $field.Node.Value = $Change.NewValue
    } else {
        $field.Node.InnerText = $Change.NewValue
    }

    Write-Good "Applied edit: $typeName entry #$entryIndex, $($Change.FieldPath) set to '$($Change.NewValue)'"
    return $true
}

function Apply-AddEntryChange {
    param(
        [xml]$Xml,
        [string]$XmlPath,
        [pscustomobject]$Change,
        [switch]$ValidateOnly
    )

    $typeName = $Change.TypeName

    $entries = @(Get-EntriesOfType -Xml $Xml -ElementName $typeName)
    if ($entries.Count -eq 0) {
        Write-Bad "Cannot apply added entry: type '$typeName' no longer exists in '$XmlPath'."
        return $false
    }

    $firstEntry = $entries[0]
    $parent = $firstEntry.ParentNode
    if (-not $parent) {
        Write-Bad "Cannot apply added entry: could not find parent container for type '$typeName' in '$XmlPath'."
        return $false
    }

    if (-not $Change.EntryXml) {
        Write-Bad "Cannot apply added entry for '$typeName' in '$XmlPath': stored EntryXml is empty."
        return $false
    }

    try {
        $tmpDoc = New-Object System.Xml.XmlDocument
        $tmpDoc.LoadXml($Change.EntryXml)
        $newNode = $Xml.ImportNode($tmpDoc.DocumentElement, $true)
    } catch {
        Write-Bad "Cannot apply added entry for '$typeName' in '$XmlPath': stored EntryXml is not valid XML. $($_.Exception.Message)"
        return $false
    }

    if ($ValidateOnly) {
        Write-Good "Would append new '$typeName' entry to '$XmlPath'."
        return $true
    }

    [void]$parent.AppendChild($newNode)
    Write-Good "Applied added entry: new '$typeName' entry appended in '$XmlPath'."
    return $true
}

function Apply-ChangeRecord {
    param(
        [xml]$Xml,
        [string]$XmlPath,
        [pscustomobject]$Change,
        [switch]$ValidateOnly
    )

    switch ($Change.Type) {
        'EditField' {
            return Apply-EditFieldChange -Xml $Xml -XmlPath $XmlPath -Change $Change -ValidateOnly:$ValidateOnly
        }
        'AddEntry' {
            return Apply-AddEntryChange -Xml $Xml -XmlPath $XmlPath -Change $Change -ValidateOnly:$ValidateOnly
        }
        default {
            Write-Bad "Unknown change type '$($Change.Type)' in log for '$XmlPath'."
            return $false
        }
    }
}

function Select-ItemWithSearch {
    param(
        [Parameter(Mandatory = $true)][object[]]$Items,
        [Parameter(Mandatory = $true)][scriptblock]$GetLabel,
        [scriptblock]$IsChanged     = { param($x) $false },
        [scriptblock]$IsHighlighted = { param($x) $false },
        [string]$Title        = "Select item",
        [string]$EmptyMessage = "No items available.",
        [string]$ItemWord     = "item"
    )

    $all = @($Items)
    if ($all.Count -eq 0) {
        Write-Bad $EmptyMessage
        Start-Sleep -Seconds 1
        return $null
    }

    $current      = $all
    $defaultColor = [System.Console]::ForegroundColor

    while ($true) {
        Clear-Screen
        Write-Title $Title

        for ($i = 0; $i -lt $current.Count; $i++) {
            $index = $i + 1
            $item  = $current[$i]
            $label = & $GetLabel $item

            $changed     = $false
            $highlighted = $false

            if ($IsChanged) {
                $changed = & $IsChanged $item
            }
            if ($IsHighlighted) {
                $highlighted = & $IsHighlighted $item
            }

            $labelLines = $label -split "`n"
            $firstLine  = "[{0}] {1}" -f $index, $labelLines[0]

            if ($highlighted) {
                Write-Host $firstLine -ForegroundColor Yellow
            } elseif ($changed) {
                Write-Host $firstLine -ForegroundColor Green
            } else {
                Write-Host $firstLine -ForegroundColor $defaultColor
            }

            if ($labelLines.Length -gt 1) {
                for ($k = 1; $k -lt $labelLines.Length; $k++) {
                    $extra = "    " + $labelLines[$k]

                    if ($highlighted -and $labelLines[$k].StartsWith("   * ")) {
                        Write-Host $extra -ForegroundColor Yellow
                    } elseif ($changed) {
                        Write-Host $extra -ForegroundColor Green
                    } else {
                        Write-Host $extra -ForegroundColor $defaultColor
                    }
                }
            }
        }

        [System.Console]::ForegroundColor = $defaultColor

        Write-Host ""
        Write-Host "Type a number to pick a $ItemWord."
        Write-Host "Type S to search $ItemWord`s."
        Write-Host "Type R to reset search."
        Write-Host "Type Q to go back."
        Write-Host ""

        $choice = Read-Host "Your choice"

        if ($choice -match '^[Qq]$') {
            return $null
        }

        if ($choice -match '^[Ss]$') {
            $term = Read-Host "Enter text to search for"
            if ([string]::IsNullOrWhiteSpace($term)) {
                Write-Bad "Search text empty. Showing full list again."
                Start-Sleep -Seconds 1
                $current = $all
                continue
            }

            $needle   = $term.ToLower()
            $filtered = @(
                $all | Where-Object {
                    ((& $GetLabel $_).ToLower()) -like "*$needle*"
                }
            )

            if ($filtered.Count -eq 0) {
                Write-Bad "No $ItemWord`s matched '$term'."
                Start-Sleep -Seconds 1
                $current = $all
            } else {
                Write-Good "$($filtered.Count) $ItemWord`s matched '$term'."
                Start-Sleep -Seconds 1
                $current = $filtered
            }

            continue
        }

        if ($choice -match '^[Rr]$') {
            $current = $all
            continue
        }

        if ($choice -as [int]) {
            $num = [int]$choice
            if ($num -ge 1 -and $num -le $current.Count) {
                return $current[$num - 1]
            }
        }

        Write-Bad "Not a valid choice. Try again."
        Start-Sleep -Seconds 1
    }
}

function Select-EntryType {
    param(
        [Parameter(Mandatory = $true)][xml]$Xml,
        [string]$HighlightTypeName
    )

    $groups = Get-EntryTypes -Xml $Xml

    $selected = Select-ItemWithSearch `
        -Items $groups `
        -GetLabel {
            param($g)
            $typeName = $g.Name
            $entries  = Get-EntriesOfType -Xml $Xml -ElementName $typeName
            $count    = @($entries).Count
            "{0} (entries: {1})" -f $typeName, $count
        } `
        -IsChanged { param($x) $false } `
        -IsHighlighted {
            param($g)
            if (-not $HighlightTypeName) { return $false }
            return ($g.Name -eq $HighlightTypeName)
        } `
        -Title "What type of entry do you want to work with?" `
        -EmptyMessage "No entry types found in this file." `
        -ItemWord "type"

    if (-not $selected) { return $null }
    return $selected.Name
}

function Select-Entry {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Entries,

        [int]$HighlightIndex = 0,
        [object[]]$HighlightMatches
    )

    $entriesArray  = @($Entries)
    $highlightKeys = @{}
    if ($HighlightMatches) {
        foreach ($hm in $HighlightMatches) {
            if ($hm.FieldPath -and $hm.ValueRaw -ne $null) {
                $key = if ($hm.PSObject.Properties.Name -contains 'FieldKey' -and $hm.FieldKey) {
                    $hm.FieldKey
                } else {
                    "{0}::{1}" -f $hm.FieldPath, $hm.ValueRaw
                }
                $highlightKeys[$key] = $true
            }
        }
    }

    $selected = Select-ItemWithSearch `
        -Items $entriesArray `
        -GetLabel {
            param($e)

            $idx          = [array]::IndexOf($entriesArray, $e)
            $isJumpTarget = ($HighlightIndex -gt 0 -and $idx -ge 0 -and ($idx + 1) -eq $HighlightIndex)

            if ($isJumpTarget -and $HighlightMatches -and $HighlightMatches.Count -gt 0) {
                $lines = @()

                $summary = Get-EntrySummary -Entry $e
                if (-not [string]::IsNullOrWhiteSpace($summary)) {
                    $lines += $summary
                }

                $fields = Get-EntryFieldsList -Entry $e
                foreach ($field in $fields) {
                    $value = if ($field.Kind -eq 'Attribute') {
                        $field.Node.Value
                    } else {
                        $field.Node.InnerText
                    }

                    $fieldKey = if ($field.PSObject.Properties.Name -contains 'FieldKey') {
                        $field.FieldKey
                    } else {
                        "{0}::{1}" -f $field.Path, $value
                    }

                    $prefix = if ($highlightKeys.ContainsKey($fieldKey)) { "   * " } else { "     " }
                    $lines += ($prefix + $field.Display)
                }

                return ($lines -join "`n")
            }
            else {
                return (Get-EntrySummary -Entry $e)
            }
        } `
        -IsChanged { param($x) $false } `
        -IsHighlighted {
            param($x)
            if ($HighlightIndex -le 0) { return $false }
            $idx = [array]::IndexOf($entriesArray, $x)
            return ($idx -ge 0 -and ($idx + 1) -eq $HighlightIndex)
        } `
        -Title "Select an entry" `
        -EmptyMessage "No entries of this type." `
        -ItemWord "entry"

    if (-not $selected) { return $null }
    return [System.Xml.XmlElement]$selected
}

function Show-EntriesList {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Entries,
        [int]$HighlightIndex = 0,
        [object[]]$HighlightMatches
    )

    $list = @($Entries)
    if ($list.Count -eq 0) {
        Write-Bad "No entries of this type were found."
        return
    }

    $highlightKeys = @{}
    if ($HighlightMatches) {
        foreach ($hm in $HighlightMatches) {
            if ($hm.FieldPath -and $hm.ValueRaw -ne $null) {
                $key = if ($hm.PSObject.Properties.Name -contains 'FieldKey' -and $hm.FieldKey) {
                    $hm.FieldKey
                } else {
                    "{0}::{1}" -f $hm.FieldPath, $hm.ValueRaw
                }
                $highlightKeys[$key] = $true
            }
        }
    }

    Write-Title "Entries"

    for ($i = 0; $i -lt $list.Count; $i++) {
        $entry = $list[$i]
        $index = $i + 1

        if ($HighlightIndex -eq $index -and $HighlightMatches -and $HighlightMatches.Count -gt 0) {

            Write-Host ("[{0}] (highlighted entry)" -f $index) -ForegroundColor Yellow

            $fields = Get-EntryFieldsList -Entry $entry
            foreach ($field in $fields) {
                $value = if ($field.Kind -eq 'Attribute') {
                    $field.Node.Value
                } else {
                    $field.Node.InnerText
                }

                $fieldKey = if ($field.PSObject.Properties.Name -contains 'FieldKey') {
                    $field.FieldKey
                } else {
                    "{0}::{1}" -f $field.Path, $value
                }

                $line = "   {0}" -f $field.Display
                if ($highlightKeys.ContainsKey($fieldKey)) {
                    Write-Host $line -ForegroundColor Yellow
                } else {
                    Write-Host $line
                }
            }

            Write-Host ""
        }
        else {
            $summary = Get-EntrySummary -Entry $entry
            $line    = ("[{0}] {1}" -f $index, $summary)

            if ($HighlightIndex -eq $index) {
                Write-Host $line -ForegroundColor Yellow
            } else {
                Write-Host $line
            }
        }
    }
}

function Edit-EntryFields {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Entry,

        [object[]]$HighlightMatches,

        [string]$TypeName,
        [int]$EntryIndex,
        [string]$XmlPath
    )

    $highlightList = @()
    if ($HighlightMatches) {
        foreach ($hm in $HighlightMatches) {
            if ($hm.FieldPath -and $hm.ValueRaw -ne $null) {
                $key = "{0}::{1}" -f $hm.FieldPath, $hm.ValueRaw
                $highlightList += $key
            }
        }
    }

    $changedFieldKeys = @()
    $hadAnyFields     = $false
    $changedAny       = $false

    while ($true) {
        $fields = Get-EntryFieldsList -Entry $Entry

        if ($fields.Count -eq 0) {
            if (-not $hadAnyFields) {
                Write-Info "No editable fields found on this entry."
                Read-Host "Press Enter to go back"
                return $false
            } else {
                return $changedAny
            }
        }

        $hadAnyFields = $true

        $field = Select-ItemWithSearch `
            -Items $fields `
            -GetLabel { param($f) $f.Display } `
            -IsChanged {
                param($f)
                $key = if ($f.PSObject.Properties.Name -contains 'FieldKey' -and $f.FieldKey) {
                    $f.FieldKey
                } else {
                    $f.Path
                }
                return $changedFieldKeys -contains $key
            } `
            -IsHighlighted {
                param($f)

                if (-not $highlightList -or $highlightList.Count -eq 0) { return $false }

                $value = if ($f.Kind -eq 'Attribute') {
                    $f.Node.Value
                } else {
                    $f.Node.InnerText
                }

                $key = "{0}::{1}" -f $f.Path, $value
                return $highlightList -contains $key
            } `
            -Title "Pick a field to change" `
            -EmptyMessage "No editable fields." `
            -ItemWord "field"

        if (-not $field) {
            return $changedAny
        }

        $currentValue = if ($field.Kind -eq 'Attribute') {
            $field.Node.Value
        } else {
            $field.Node.InnerText
        }

        $prompt   = "New value for $($field.Path) (current: '$currentValue', leave blank to keep as it is)"
        $newValue = Read-Host $prompt

        if ([string]::IsNullOrWhiteSpace($newValue)) {
            Write-Info "Value left unchanged."
            Start-Sleep -Seconds 1
            continue
        }

        if ($field.Kind -eq 'Attribute') {
            $field.Node.Value = $newValue
        } else {
            $field.Node.InnerText = $newValue
        }

        $fieldKey = if ($field.PSObject.Properties.Name -contains 'FieldKey' -and $field.FieldKey) {
            $field.FieldKey
        } else {
            $field.Path
        }

        if ($changedFieldKeys -notcontains $fieldKey) {
            $changedFieldKeys += $fieldKey
        }
        $changedAny = $true

        Write-Good "Updated $($field.Path) to '$newValue'"

        if ($TypeName -and $EntryIndex -gt 0 -and $XmlPath) {
            $changeObj = [pscustomobject]@{
                Type       = 'EditField'
                TypeName   = $TypeName
                EntryIndex = $EntryIndex
                FieldPath  = $field.Path
                OldValue   = $currentValue
                NewValue   = $newValue
                TimeUtc    = (Get-Date).ToUniversalTime().ToString("o")
                Status     = 'Pending'
            }
            Add-ChangeRecordForFile -XmlPath $XmlPath -Change $changeObj
        }

        Start-Sleep -Seconds 1
    }
}

function Duplicate-Entry {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlElement]$Entry,
        [object[]]$HighlightMatches,

        [string]$TypeName,
        [int]$TemplateIndex,
        [string]$XmlPath
    )

    $parent = $Entry.ParentNode
    if (-not $parent) {
        Write-Bad "Cannot duplicate. Entry has no parent in the XML."
        return $null
    }

    $newEntry = $Entry.Clone()

    Write-Info "You are now editing the copy of the entry."
    $hadChanges = Edit-EntryFields `
        -Entry $newEntry `
        -HighlightMatches $HighlightMatches `
        -TypeName $TypeName `
        -EntryIndex 0 `
        -XmlPath $XmlPath

    if (-not $hadChanges) {
        Write-Info "No changes were made, copy was not added."
        Start-Sleep -Seconds 1
        return $null
    }

    $parent.AppendChild($newEntry) | Out-Null
    Write-Good "New entry added under the same section."
    Start-Sleep -Seconds 1

    if ($XmlPath -and $TypeName) {
        try {
            $entryXml  = $newEntry.OuterXml
            $changeObj = [pscustomobject]@{
                Type      = 'AddEntry'
                TypeName  = $TypeName
                EntryXml  = $entryXml
                TimeUtc   = (Get-Date).ToUniversalTime().ToString("o")
                Status    = 'Pending'
            }
            Add-ChangeRecordForFile -XmlPath $XmlPath -Change $changeObj
        } catch {
            Write-Bad "Could not record added entry change for '$XmlPath': $($_.Exception.Message)"
        }
    }

    return $newEntry
}

function Apply-SingleSavedChange {
    param(
        [Parameter(Mandatory = $true)][System.Xml.XmlDocument]$XmlRoot,
        [Parameter(Mandatory = $true)][object]$Change,
        [Parameter(Mandatory = $true)][string]$XmlPath,
        [switch]$ValidateOnly
    )

    if ($Change.PSObject.Properties.Name -contains 'Type') {
        $normalized = $Change
        Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized -ValidateOnly:$ValidateOnly
        return
    }

    if ($Change.PSObject.Properties.Name -contains 'Kind') {
        if ($Change.Kind -eq 'ADD') {
            $entryXml = $null

            if ($Change.PSObject.Properties.Name -contains 'EntryXml' -and $Change.EntryXml) {
                $entryXml = $Change.EntryXml
            } elseif ($Change.PSObject.Properties.Name -contains 'Fields') {
                $tmpDoc   = New-Object System.Xml.XmlDocument
                $rootName = $Change.EntryType
                $rootNode = $tmpDoc.CreateElement($rootName)
                $tmpDoc.AppendChild($rootNode) | Out-Null

                foreach ($field in $Change.Fields) {
                    $parts       = $field.Path.Split(".")
                    $currentNode = $rootNode
                    foreach ($p in $parts) {
                        $found = $currentNode.SelectSingleNode($p)
                        if (-not $found) {
                            $found = $tmpDoc.CreateElement($p)
                            $currentNode.AppendChild($found) | Out-Null
                        }
                        $currentNode = $found
                    }
                    $currentNode.InnerText = $field.NewValue
                }

                $entryXml = $rootNode.OuterXml
            }

            if ($entryXml) {
                $normalized = [pscustomobject]@{
                    Type     = 'AddEntry'
                    TypeName = $Change.EntryType
                    EntryXml = $entryXml
                }
                Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized -ValidateOnly:$ValidateOnly
            }

            return
        }

        if ($Change.Kind -eq 'EDIT') {
            $fields = @()
            if ($Change.PSObject.Properties.Name -contains 'Fields') {
                foreach ($f in $Change.Fields) {
                    $fields += [pscustomobject]@{
                        Path     = $f.Path
                        OldValue = $f.OldValue
                        NewValue = $f.NewValue
                    }
                }
            }

            foreach ($f in $fields) {
                $normalized = [pscustomobject]@{
                    Type       = 'EditField'
                    TypeName   = $Change.EntryType
                    EntryIndex = $Change.EntryIndex
                    FieldPath  = $f.Path
                    OldValue   = $f.OldValue
                    NewValue   = $f.NewValue
                }

                Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized -ValidateOnly:$ValidateOnly
            }

            return
        }
    }
}

function Review-AllSavedEdits {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$XmlFiles
    )

    while ($true) {
        Clear-Screen
        Write-Title "Review saved edits"

        $items = @()

        foreach ($file in $XmlFiles) {
            $xmlPath = $file.FullName
            $changes = @(Load-ChangesForFile -XmlPath $xmlPath)

            if ($changes.Count -gt 0) {
                $items += [pscustomobject]@{
                    File    = $file
                    XmlPath = $xmlPath
                    Count   = $changes.Count
                    Changes = $changes
                }
            }
        }

        if ($items.Count -eq 0) {
            Write-Bad "There are no saved edits in this folder yet."
            Write-Host ""
            Write-Host "Options:"
            Write-Host "[I] Import a shared config pack"
            Write-Host "[Q] Back"
            Write-Host ""

            $choice = Read-Host "Your choice"

            if ($choice -match '^[Qq]$') {
                return
            }

            if ($choice -match '^[Ii]$') {
                Import-SharedConfigPack -XmlFiles $XmlFiles
                continue
            }

            Write-Bad "Not a valid option."
            Start-Sleep -Seconds 1
            continue
        }

        for ($i = 0; $i -lt $items.Count; $i++) {
            $idx  = $i + 1
            $item = $items[$i]

            $suffix = if ($item.Count -eq 1) { "change" } else { "changes" }
            Write-Host ("[{0}] {1} ({2} {3})" -f `
                $idx,
                $item.File.Name,
                $item.Count,
                $suffix) -ForegroundColor Cyan
        }

        Write-Host ""
        Write-Host "Options:"
        Write-Host "[number] Review saved edits for that XML"
        Write-Host "[A] Apply ALL saved edits to ALL XML files"
        Write-Host "[E] Export saved edits as a shared config pack"
        Write-Host "[X] Export saved edits summary (text)"
        Write-Host "[I] Import a shared config pack"
        Write-Host "[Q] Back"
        Write-Host ""

        $choice = Read-Host "Your choice"

        if ($choice -match '^[Qq]$') {
            return
        }

        if ($choice -match '^[Aa]$') {
            Write-Host ""
            Write-Host "This will apply ALL saved edits to ALL XML files," -ForegroundColor Yellow
            Write-Host "create backups, save XMLs, and mark edits as committed." -ForegroundColor Yellow
            Write-Host ""

            $confirm = Read-Host "Type APPLY to continue"
            if ($confirm -ne "APPLY") {
                Write-Info "Bulk apply cancelled."
                Start-Sleep -Seconds 2
                continue
            }

            foreach ($item in $items) {
                Write-Host ""
                Write-Title ("Applying saved edits to {0}" -f $item.File.Name)

                try {
                    [xml]$xmlDoc = Load-XmlDocument -Path $item.XmlPath
                }
                catch {
                    Write-Bad ("Could not read XML '{0}': {1}" -f `
                        $item.XmlPath,
                        $_.Exception.Message)
                    Write-LogError $_
                    continue
                }

                $applied = 0

                foreach ($c in $item.Changes) {
                    try {
                        if (Apply-SingleSavedChange -XmlRoot $xmlDoc -Change $c -XmlPath $item.XmlPath) {
                            $applied++
                        }
                    }
                    catch {
                        Write-Bad ("Error applying change in {0}: {1}" -f `
                            $item.File.Name,
                            $_.Exception.Message)
                        Write-LogError $_
                    }
                }

                Backup-XmlFile -Path $item.XmlPath
                Save-XmlDocument -Xml $xmlDoc -Path $item.XmlPath
                Mark-AllPendingChangesCommittedForFile -XmlPath $item.XmlPath

                Write-Good ("Applied {0} change{1} to {2}" -f `
                    $applied,
                    $(if ($applied -eq 1) { "" } else { "s" }),
                    $item.File.Name)

                Start-Sleep -Seconds 1
            }

            Write-Host ""
            Write-Good "Bulk apply completed."
            Read-Host "Press Enter to continue" | Out-Null
            continue
        }

        if ($choice -match '^[Ee]$') {
            Export-SharedConfigPack -XmlFiles $XmlFiles
            continue
        }

        if ($choice -match '^[Xx]$') {
            Export-SavedEditsSummary -XmlFiles $XmlFiles
            continue
        }

        if ($choice -match '^[Ii]$') {
            Import-SharedConfigPack -XmlFiles $XmlFiles
            continue
        }

        if ($choice -as [int]) {
            $num = [int]$choice
            if ($num -ge 1 -and $num -le $items.Count) {
                $item    = $items[$num - 1]
                $changes = @(Load-ChangesForFile -XmlPath $item.XmlPath)
                Review-SavedEditsForFile -XmlPath $item.XmlPath -SavedChanges $changes
                continue
            }

            Write-Bad "Not a valid number."
            Start-Sleep -Seconds 1
            continue
        }

        Write-Bad "Not a valid option."
        Start-Sleep -Seconds 1
    }
}

function Review-SavedEditsForFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath,

        [Parameter(Mandatory = $true)]
        [object[]]$SavedChanges
    )

    if (-not (Test-Path $XmlPath)) {
        Write-Bad "XML not found on disk: $XmlPath"
        Start-Sleep -Seconds 2
        return
    }

    try {
        [xml]$Xml = Load-XmlDocument -Path $XmlPath
    } catch {
        Write-Bad "Could not open XML for review: $XmlPath"
        Start-Sleep -Seconds 2
        return
    }

    $changesList = [System.Collections.ArrayList]::new()
    if ($SavedChanges) {
        $SavedChanges = $SavedChanges | Where-Object { $_ -ne $null }
        [void]$changesList.AddRange($SavedChanges)
    }

    $appliedAllInMemory = $false
    $currentFilter      = 'All'

    while ($true) {
        Clear-Screen
        Write-Title "Saved edits for:"
        Write-Host $XmlPath
        Write-Host ""

        if ($changesList.Count -eq 0) {
            Write-Bad "Nothing tracked for this XML anymore."
            Read-Host "Press Enter to return" | Out-Null
            return
        }

        Write-Host -NoNewline "Current filter: " -ForegroundColor Cyan
        switch ($currentFilter) {
            'Pending'   { Write-Host "Pending"   -ForegroundColor Yellow }
            'Committed' { Write-Host "Committed" -ForegroundColor Green }
            default     { Write-Host "All"       -ForegroundColor White }
        }
        Write-Host ""

        $allChanges = @($changesList)
        switch ($currentFilter) {
            'Pending' {
                $changes = @(
                    $allChanges | Where-Object {
                        $p = $_.PSObject.Properties['Status']
                        if (-not $p) { return $true }
                        $val = [string]$p.Value
                        if ([string]::IsNullOrWhiteSpace($val)) { return $true }
                        return ($val -ne 'Committed')
                    }
                )
            }
            'Committed' {
                $changes = @(
                    $allChanges | Where-Object {
                        $p = $_.PSObject.Properties['Status']
                        if (-not $p) { return $false }
                        $val = [string]$p.Value
                        return ($val -eq 'Committed')
                    }
                )
            }
            default {
                $changes = $allChanges
            }
        }

        if ($changes.Count -eq 0) {
            Write-Bad "No changes match the current filter."
            Write-Host ""
        }

        $defaultColor = [System.Console]::ForegroundColor
        $index = 0
        foreach ($change in $changes) {
            $index++

            $statusLabel = $null
            $statusColor = $null

            if ($change.PSObject.Properties.Name -contains 'Status') {
                $rawStatus = [string]$change.Status
                if ([string]::IsNullOrWhiteSpace($rawStatus)) {
                    $rawStatus = 'Pending'
                }

                if ($rawStatus -ieq 'Committed') {
                    $statusLabel = 'Committed'
                    $statusColor = 'Green'
                } else {
                    $statusLabel = 'Pending'
                    $statusColor = 'Yellow'
                }
            }

            function Write-ChangeLine {
                param(
                    [int]$Idx,
                    [string]$MainText,
                    [string]$MainColor,
                    [string]$StatusText,
                    [string]$StatusColor
                )

                Write-Host "[" -NoNewline -ForegroundColor White
                Write-Host $Idx -NoNewline -ForegroundColor $MainColor
                Write-Host "] " -NoNewline -ForegroundColor White

                Write-Host $MainText -NoNewline -ForegroundColor $MainColor

                if ($StatusText) {
                    Write-Host " [" -NoNewline -ForegroundColor White
                    Write-Host $StatusText -NoNewline -ForegroundColor $StatusColor
                    Write-Host "]" -NoNewline -ForegroundColor White
                }

                Write-Host ""
            }

            if ($change.Type -eq 'EditField') {
                $mainText = ("EDIT Type='{0}', EntryIndex={1}, Field='{2}'" -f `
                    $change.TypeName,
                    $change.EntryIndex,
                    $change.FieldPath)

                Write-ChangeLine -Idx $index `
                                 -MainText $mainText `
                                 -MainColor 'Cyan' `
                                 -StatusText $statusLabel `
                                 -StatusColor $(if ($statusColor) { $statusColor } else { 'Yellow' })

                $valueColor = if ($statusColor) { $statusColor } else { 'Green' }
                Write-Host ("     {0} -> {1}" -f $change.OldValue, $change.NewValue) -ForegroundColor $valueColor
                Write-Host ""
            }
            elseif ($change.Type -eq 'AddEntry') {
                $mainText = ("ADD Type='{0}' (new entry)" -f $change.TypeName)

                Write-ChangeLine -Idx $index `
                                 -MainText $mainText `
                                 -MainColor 'Yellow' `
                                 -StatusText $statusLabel `
                                 -StatusColor $(if ($statusColor) { $statusColor } else { 'Yellow' })

                $fieldColor = if ($statusColor) { $statusColor } else { 'Green' }

                try {
                    if ($change.EntryXml) {
                        $tmpDoc   = New-Object System.Xml.XmlDocument
                        $tmpDoc.LoadXml($change.EntryXml)
                        $tmpEntry      = [System.Xml.XmlElement]$tmpDoc.DocumentElement
                        $summaryFields = Get-EntryFieldsList -Entry $tmpEntry
                        foreach ($sf in $summaryFields) {
                            Write-Host ("     {0}" -f $sf.Display) -ForegroundColor $fieldColor
                        }
                    }
                } catch {
                    Write-Bad "     (Could not read stored EntryXml)"
                }

                Write-Host ""
            }
            else {
                $mainText = ("Unknown change type '{0}'" -f $change.Type)

                Write-ChangeLine -Idx $index `
                                 -MainText $mainText `
                                 -MainColor 'Yellow' `
                                 -StatusText $statusLabel `
                                 -StatusColor $(if ($statusColor) { $statusColor } else { 'Yellow' })

                Write-Host ""
            }
        }

        Write-Host "--------------------------------------------------------------"
        Write-Host "Options:"
        Write-Host "  [1] Apply ALL valid changes (in memory only)"
        Write-Host "  [2] Apply ONE change by number"
        Write-Host "  [3] Delete ONE change by number"
        Write-Host "  [4] Delete ALL saved changes for this XML"
        Write-Host "  [5] Save XML now (backup + file write, mark changes committed)"
        Write-Host "  [6] Reload XML from disk (discard in-memory)"
        Write-Host "  [7] Test apply ALL changes (no modifications)"
        Write-Host "  [8] Change filter (All / Pending / Committed)"
        Write-Host "  [9] Export saved edits summary for this XML (text)"
        Write-Host "  [Q] Go back"
        Write-Host ""

        $choice = Read-Host "Your choice"

        if ($choice -match '^[Qq]$') {
            return
        }

        switch ($choice) {

            '4' {
                $confirm = Read-Host "This will delete ALL saved edits for this XML from the history log. It does NOT modify the XML file itself. Are you sure? (Y/N)"
                if ($confirm -notmatch '^[Yy]$') {
                    Write-Info "Delete cancelled. Saved edits are unchanged."
                    Start-Sleep -Seconds 1
                    continue
                }

                $changesList.Clear()
                Set-ChangesForFile -XmlPath $XmlPath -Changes ([object[]]$changesList)
                Write-Good "Cleared all saved edits for this XML."
                Read-Host "Press Enter to return" | Out-Null
                return
            }

            '3' {
                if ($changes.Count -eq 0) {
                    Write-Bad "No changes to delete for this filter."
                    Read-Host "Press Enter to continue" | Out-Null
                    continue
                }

                $target = Read-Host "Delete which one? (number)"
                if ($target -as [int]) {
                    $n = [int]$target
                    if ($n -ge 1 -and $n -le $changes.Count) {
                        $changeToDelete = $changes[$n - 1]
                        $confirm = Read-Host "This will remove change #$n from the saved edits history for this XML only. It does NOT modify the XML file. Are you sure? (Y/N)"
                        if ($confirm -notmatch '^[Yy]$') {
                            Write-Info "Delete cancelled. Saved edits are unchanged."
                            Start-Sleep -Seconds 1
                            continue
                        }

                        $idxInList = $changesList.IndexOf($changeToDelete)
                        if ($idxInList -ge 0) {
                            $changesList.RemoveAt($idxInList)
                        }

                        Set-ChangesForFile -XmlPath $XmlPath -Changes ([object[]]$changesList)
                        Write-Good "Removed change #$n"
                    } else {
                        Write-Bad "Invalid number"
                    }
                }
                Read-Host "Press Enter to continue" | Out-Null
                continue
            }

            '2' {
                if ($changes.Count -eq 0) {
                    Write-Bad "No changes to apply for this filter."
                    Read-Host "Press Enter to continue" | Out-Null
                    continue
                }

                $target = Read-Host "Apply which change? (number)"
                if ($target -as [int]) {
                    $n = [int]$target
                    if ($n -ge 1 -and $n -le $changes.Count) {
                        $changeToApply = $changes[$n - 1]
                        Apply-SingleSavedChange -XmlRoot $Xml -Change $changeToApply -XmlPath $XmlPath
                        Write-Good "Applied change #$n"
                        Write-Info "This edit is only applied in memory. Use 'Save XML now' to write it to disk."
                        $appliedAllInMemory = $false
                    } else {
                        Write-Bad "Invalid number"
                    }
                }
                Read-Host "Press Enter to continue" | Out-Null
                continue
            }

            '1' {
                foreach ($c in $changesList) {
                    Apply-SingleSavedChange -XmlRoot $Xml -Change $c -XmlPath $XmlPath
                }
                $appliedAllInMemory = $true
                Write-Good "Applied ALL changes (in memory only)"
                Write-Info "Use 'Save XML now' if you want these edits written to disk."
                Read-Host "Press Enter to continue" | Out-Null
                continue
            }

            '5' {
                if (-not $appliedAllInMemory) {
                    $pendingToApply = @(
                        $changesList | Where-Object {
                            $p = $_.PSObject.Properties['Status']
                            if (-not $p) { return $true }
                            $val = [string]$p.Value
                            if ([string]::IsNullOrWhiteSpace($val)) { return $true }
                            return ($val -ne 'Committed')
                        }
                    )

                    if ($pendingToApply.Count -gt 0) {
                        $appliedCount = 0
                        foreach ($c in $pendingToApply) {
                            $ok = $false
                            try {
                                $ok = Apply-SingleSavedChange -XmlRoot $Xml -Change $c -XmlPath $XmlPath
                            } catch {
                                $ok = $false
                                Write-Bad ("Error applying pending change: {0}" -f $_.Exception.Message)
                            }
                            if ($ok) { $appliedCount++ }
                        }

                        Write-Good ("Applied {0} pending change{1} before saving." -f `
                            $appliedCount,
                            $(if ($appliedCount -eq 1) { "" } else { "s" }))
                    } else {
                        Write-Info "No pending changes to apply before saving."
                    }
                } else {
                    Write-Info "All changes were already applied in memory earlier; saving without reapplying."
                }

                Backup-XmlFile -Path $XmlPath
                Save-XmlDocument -Xml $Xml -Path $XmlPath
                Mark-AllPendingChangesCommittedForFile -XmlPath $XmlPath

                $changePath = Get-ChangeLogPathForXml -XmlPath $XmlPath

                Write-Host ""
                Write-Good "XML written to disk and backed up"

                if ($changePath) {
                    Write-Host ""
                    Write-Info "Changes log is stored at: $changePath"
                }

                $updated = @(Load-ChangesForFile -XmlPath $XmlPath)
                $changesList.Clear()
                if ($updated) {
                    [void]$changesList.AddRange($updated)
                }

                $appliedAllInMemory = $false

                Write-Host ""
                Read-Host "Press Enter to continue" | Out-Null
                continue
            }

            '6' {
                Write-Bad "Reloading XML from disk..."
                $Xml = Load-XmlDocument -Path $XmlPath
                $appliedAllInMemory = $false
                Start-Sleep 1
            }

            '7' {
                Write-Host ""
                Write-Info "Testing all changes in dry-run mode. XML will not be modified or saved."

                [xml]$xmlClone = $Xml.OuterXml

                foreach ($c in $changesList) {
                    Apply-SingleSavedChange -XmlRoot $xmlClone -Change $c -XmlPath $XmlPath -ValidateOnly
                }
                Write-Host ""
                Read-Host "Dry-run completed. Press Enter to continue" | Out-Null
                continue
            }

            '8' {
                $changing = $true
                while ($changing) {
                    Clear-Screen
                    Write-Title "Change filter"
                    Write-Host -NoNewline "Current filter: " -ForegroundColor Cyan
                    switch ($currentFilter) {
                        'Pending'   { Write-Host "Pending"   -ForegroundColor Yellow }
                        'Committed' { Write-Host "Committed" -ForegroundColor Green }
                        default     { Write-Host "All"       -ForegroundColor White }
                    }
                    Write-Host ""
                    Write-Host "Select filter:"
                    Write-Host "  [1] All changes"
                    Write-Host "  [2] Pending only"
                    Write-Host "  [3] Committed only"
                    Write-Host "  [Q] Cancel"
                    Write-Host ""

                    $fChoice = Read-Host "Your choice"

                    switch ($fChoice) {
                        '1' { $currentFilter = 'All';       $changing = $false }
                        '2' { $currentFilter = 'Pending';   $changing = $false }
                        '3' { $currentFilter = 'Committed'; $changing = $false }
                        default {
                            if ($fChoice -match '^[Qq]$') {
                                $changing = $false
                            } else {
                                Write-Bad "Not a valid option."
                                Start-Sleep -Seconds 1
                            }
                        }
                    }
                }
            }

            '9' {
                Export-SavedEditsSummaryForFile -XmlPath $XmlPath -SavedChanges ([object[]]@($changesList))
                continue
            }

            default {
                Write-Bad "Invalid input"
                Start-Sleep 1
            }
        }
    }
}

function Show-FileMenu {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    Clear-Screen
    Write-Title "Working with file:"
    Write-Host $FilePath
    Write-Host ""

    if (-not $Script:ShownBackupInfoShown) {
        Write-Host "[i] A backup is automatically made when you save changes." -ForegroundColor Yellow
        Write-Host "[i] The backup will be saved in a folder called 'BackupXMLs' next to your XML file." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to continue..."
        $Script:ShownBackupInfoShown = $true

        Clear-Screen
        Write-Title "Working with file:"
        Write-Host $FilePath
        Write-Host ""
    }

    Write-Host "[1] Look at entry types in this file (no changes will be made)"
    Write-Host "[2] Add a new entry by copying an existing one"
    Write-Host "[3] Edit an existing entry"
    Write-Host "[4] Save changes to XML (backup, write to disk + mark changes committed), then return"
    Write-Host "[5] Discard in-memory changes for this file and return"
    Write-Host "[6] Review saved edits for this file"
    Write-Host "[Q] Go back"
}

function Show-InfoMenu {
    param(
        [Parameter(Mandatory = $true)][string]$RootFolder
    )

    while ($true) {
        Clear-Screen
        Write-Title "Tool information & folders"

        $version          = Get-LocalVersion
        $helperRoot       = Join-Path $RootFolder "LSR-XML-Helper"
        $xmlEditsDir      = Join-Path $helperRoot "XML-Edits"
        $backupDir        = Join-Path $helperRoot "BackupXMLs"
        $sharedConfigsDir = Join-Path $helperRoot "Shared-Configs"
        $logsDir          = Join-Path $helperRoot "Logs"

        if ($Script:SkipUpdateCheck) {
            $versionColor = "Green"
        } else {
            $versionColor = "White"
            if ($Script:VersionStatus -eq "Latest") {
                $versionColor = "Green"
            } elseif ($Script:VersionStatus -eq "Outdated") {
                $versionColor = "Yellow"
            }
        }

        Write-Host -NoNewline "Current version        : " -ForegroundColor White
        Write-Host $version -ForegroundColor $versionColor

        Write-Host -NoNewline "Root XML folder        : " -ForegroundColor White
        Write-Host $RootFolder -ForegroundColor Cyan

        Write-Host -NoNewline "AppData config folder  : " -ForegroundColor White
        Write-Host $Script:AppDataDir -ForegroundColor Cyan

        Write-Host -NoNewline "Helper root under XMLs : " -ForegroundColor White
        Write-Host $helperRoot -ForegroundColor Cyan

        Write-Host ""
        Write-Host "Subfolders (created when needed):" -ForegroundColor White

        Write-Host -NoNewline "  XML-Edits        : " -ForegroundColor White
        Write-Host $xmlEditsDir -ForegroundColor Cyan

        Write-Host -NoNewline "  BackupXMLs       : " -ForegroundColor White
        Write-Host $backupDir -ForegroundColor Cyan

        Write-Host -NoNewline "  Shared-Configs   : " -ForegroundColor White
        Write-Host $sharedConfigsDir -ForegroundColor Cyan

        Write-Host -NoNewline "  Logs             : " -ForegroundColor White
        Write-Host $logsDir -ForegroundColor Cyan

        Write-Host ""

        $updateStatusText  = if ($Script:SkipUpdateCheck) { "OFF" } else { "ON" }
        $updateStatusColor = if ($Script:SkipUpdateCheck) { "Yellow" } else { "Green" }

        Write-Host -NoNewline "Automatic update check    : " -ForegroundColor White
        Write-Host $updateStatusText -ForegroundColor $updateStatusColor

        $autoUseText  = if ($Script:AutoUseLastFolder) { "ON" } else { "OFF" }
        $autoUseColor = if ($Script:AutoUseLastFolder) { "Green" } else { "Yellow" }

        Write-Host -NoNewline "Auto-use last XML folder  : " -ForegroundColor White
        Write-Host $autoUseText -ForegroundColor $autoUseColor

        Write-Host ""
        Write-Host "[1] Open main XML folder"
        Write-Host "[2] Open helper root (LSR-XML-Helper)"
        Write-Host "[3] Open XML-Edits"
        Write-Host "[4] Open BackupXMLs"
        Write-Host "[5] Open Shared-Configs"
        Write-Host "[6] Open Logs"
        Write-Host "[7] Toggle automatic update check"
        Write-Host "[8] Toggle auto-use last XML folder"
        Write-Host "[Q] Go back"
        Write-Host ""

        $choice = Read-Host "Your choice"

        switch ($choice) {
            '1' { Open-Folder -Path $RootFolder }
            '2' { Open-Folder -Path $helperRoot }
            '3' { Open-Folder -Path $xmlEditsDir }
            '4' { Open-Folder -Path $backupDir }
            '5' { Open-Folder -Path $sharedConfigsDir }
            '6' { Open-Folder -Path $logsDir }
            '7' {
                $Script:SkipUpdateCheck = -not $Script:SkipUpdateCheck
                $configDir  = Join-Path $env:LOCALAPPDATA "LSR-XML-Helper"
                if (-not (Test-Path $configDir)) {
                    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
                }
                $configPath = Join-Path $configDir "config.json"
                $cfgObj = [pscustomobject]@{
                    RootFolder        = $RootFolder
                    SkipUpdateCheck   = $Script:SkipUpdateCheck
                    AutoUseLastFolder = $Script:AutoUseLastFolder
                }
                try {
                    $cfgObj | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
                    $stateText = if ($Script:SkipUpdateCheck) { "OFF" } else { "ON" }
                    Write-Good ("Automatic update check is now {0}" -f $stateText)
                } catch {
                    Write-Bad "Could not update configuration."
                }
                Start-Sleep -Seconds 2
            }
            '8' {
                $Script:AutoUseLastFolder = -not $Script:AutoUseLastFolder
                $configDir  = Join-Path $env:LOCALAPPDATA "LSR-XML-Helper"
                if (-not (Test-Path $configDir)) {
                    New-Item -ItemType Directory -Path $configDir -Force | Out-Null
                }
                $configPath = Join-Path $configDir "config.json"
                $cfgObj = [pscustomobject]@{
                    RootFolder        = $RootFolder
                    SkipUpdateCheck   = $Script:SkipUpdateCheck
                    AutoUseLastFolder = $Script:AutoUseLastFolder
                }
                try {
                    $cfgObj | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
                    $stateText = if ($Script:AutoUseLastFolder) { "ON" } else { "OFF" }
                    Write-Good ("Auto-use last XML folder is now {0}" -f $stateText)
                } catch {
                    Write-Bad "Could not update configuration."
                }
                Start-Sleep -Seconds 2
            }
            default {
                if ($choice -match '^[Qq]$') {
                    return
                }
            }
        }
    }
}

function Get-RootFolderInteractive {
    param(
        [string]$InitialValue
    )

    if ($InitialValue -and (Test-Path $InitialValue)) {
        return $InitialValue
    }

    $configDir = Join-Path $env:LOCALAPPDATA "LSR-XML-Helper"
    if (-not (Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }
    $configPath = Join-Path $configDir "config.json"

    $remembered = $null
    if (Test-Path $configPath) {
        try {
            $cfgText = Get-Content -Path $configPath -Raw
            if (-not [string]::IsNullOrWhiteSpace($cfgText)) {
                $cfg = $cfgText | ConvertFrom-Json
                if ($cfg.RootFolder -and (Test-Path $cfg.RootFolder)) {
                    $remembered = $cfg.RootFolder
                }
                if ($cfg.SkipUpdateCheck -eq $true) {
                    $Script:SkipUpdateCheck = $true
                }
                if ($cfg.AutoUseLastFolder -eq $true) {
                    $Script:AutoUseLastFolder = $true
                }
            }
        } catch {
            Write-Bad "Could not read saved folder. Will ask again."
        }
    }

    if ($remembered -and $Script:AutoUseLastFolder) {
        return $remembered
    }

    if ($remembered) {
        Clear-Screen
        Write-Title "Last used folder found"
        Write-Info "Last folder: $remembered"
        $useLast = Read-Host "Use this folder? (Y/N)"
        if ($useLast -match '^[Yy]$') {
            return $remembered
        }
    }

    Clear-Screen
    Write-Title "Pick the folder with your LSR XML files"
    $pickedFolder = $null

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dialog                     = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description         = "Pick the folder that contains your Los Santos RED XML files"
        $dialog.ShowNewFolderButton = $false

        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and
            -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
            $pickedFolder = $dialog.SelectedPath
        }
    } catch {
        Write-Bad "Could not open folder picker window. Type the folder instead."
        $pickedFolder = Read-Host "Folder with your LSR XML files"
    }

    if ([string]::IsNullOrWhiteSpace($pickedFolder)) {
        throw "No folder was selected. Please run the tool again and pick the folder that contains your LSR XML files."
    }

    if (-not (Test-Path $pickedFolder)) {
        throw "Folder does not exist: $pickedFolder"
    }

    try {
        $cfgObj = [pscustomobject]@{
            RootFolder        = $pickedFolder
            SkipUpdateCheck   = $Script:SkipUpdateCheck
            AutoUseLastFolder = $Script:AutoUseLastFolder
        }
        $cfgObj | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8
        Write-Good "Folder saved for next time."
        Start-Sleep -Seconds 1
    } catch {
        Write-Bad "Could not save folder setting, but you can still use the tool."
    }

    return $pickedFolder
}

function Edit-XmlFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [string]$JumpTypeName,
        [int]$JumpEntryIndex,
        [object[]]$HighlightMatches
    )

    $global:currentPath  = $File.FullName
    $global:originalText = Get-Content -Path $global:currentPath -Raw
    $global:currentXml   = Load-XmlDocument -Path $global:currentPath

    $localJumpTypeName   = $JumpTypeName
    $localJumpEntryIndex = $JumpEntryIndex
    $localHighlight      = $HighlightMatches

    $editing = $true
    while ($editing) {
        Show-FileMenu -FilePath $global:currentPath
        $choice = Read-Host "Your choice"

        $hasJump = ($localJumpTypeName -and $localJumpEntryIndex -gt 0)

        switch ($choice) {

            '1' {
                $viewing = $true
                while ($viewing) {
                    $typeName = Select-EntryType -Xml $global:currentXml -HighlightTypeName $localJumpTypeName
                    if (-not $typeName) {
                        $viewing = $false
                        break
                    }

                    $entries = Get-EntriesOfType -Xml $global:currentXml -ElementName $typeName
                    if (-not $entries -or @($entries).Count -eq 0) {
                        Write-Bad "No entries of that type."
                        Start-Sleep -Seconds 1
                        continue
                    }

                    $entriesArray   = @($entries)
                    $highlightIndex = 0
                    if ($hasJump -and $typeName -eq $localJumpTypeName -and
                        $localJumpEntryIndex -gt 0 -and
                        $localJumpEntryIndex -le $entriesArray.Count) {
                        $highlightIndex = $localJumpEntryIndex
                    }

                    Clear-Screen
                    Show-EntriesList -Entries $entriesArray -HighlightIndex $highlightIndex -HighlightMatches $localHighlight
                    Write-Host ""
                    Write-Host "Press Enter to go back."
                    Read-Host | Out-Null
                }
            }

            '2' {
                $adding = $true
                while ($adding) {
                    $typeName = Select-EntryType -Xml $global:currentXml -HighlightTypeName $localJumpTypeName
                    if (-not $typeName) {
                        $adding = $false
                        break
                    }

                    $entries = Get-EntriesOfType -Xml $global:currentXml -ElementName $typeName
                    $entriesArray = @($entries)
                    if ($entriesArray.Count -eq 0) {
                        Write-Bad "No entries of that type."
                        Start-Sleep -Seconds 1
                        continue
                    }

                    $highlightIndex = 0
                    if ($hasJump -and $typeName -eq $localJumpTypeName -and
                        $localJumpEntryIndex -gt 0 -and
                        $localJumpEntryIndex -le $entriesArray.Count) {
                        $highlightIndex = $localJumpEntryIndex
                    }

                    $pickTemplate = $true
                    while ($pickTemplate) {
                        $template = Select-Entry -Entries $entriesArray -HighlightIndex $highlightIndex -HighlightMatches $localHighlight
                        if (-not $template) {
                            $pickTemplate = $false
                            break
                        }

                        $templateMatches = $null
                        if ($hasJump -and $typeName -eq $localJumpTypeName) {
                            $idx = [Array]::IndexOf($entriesArray, $template)
                            if ($idx -ge 0 -and ($idx + 1) -eq $localJumpEntryIndex) {
                                $templateMatches = $localHighlight
                            }
                        }

                        $templateIndex = [Array]::IndexOf($entriesArray, $template)
                        if ($templateIndex -lt 0) { $templateIndex = 0 }

                        Duplicate-Entry `
                            -Entry $template `
                            -HighlightMatches $templateMatches `
                            -TypeName $typeName `
                            -TemplateIndex ($templateIndex + 1) `
                            -XmlPath $global:currentPath | Out-Null
                    }
                }
            }

            '3' {
                $editingTypes = $true
                while ($editingTypes) {
                    $typeName = Select-EntryType -Xml $global:currentXml -HighlightTypeName $localJumpTypeName
                    if (-not $typeName) {
                        $editingTypes = $false
                        break
                    }

                    $entries = Get-EntriesOfType -Xml $global:currentXml -ElementName $typeName
                    $entriesArray = @($entries)
                    if ($entriesArray.Count -eq 0) {
                        Write-Bad "No entries of that type."
                        Start-Sleep -Seconds 1
                        continue
                    }

                    $highlightIndex = 0
                    if ($hasJump -and $typeName -eq $localJumpTypeName -and
                        $localJumpEntryIndex -gt 0 -and
                        $localJumpEntryIndex -le $entriesArray.Count) {
                        $highlightIndex = $localJumpEntryIndex
                    }

                    $editingEntries = $true
                    while ($editingEntries) {
                        $entry = Select-Entry -Entries $entriesArray -HighlightIndex $highlightIndex -HighlightMatches $localHighlight
                        if (-not $entry) {
                            $editingEntries = $false
                            break
                        }

                        $entryMatches = $null
                        if ($hasJump -and $typeName -eq $localJumpTypeName) {
                            $idx = [Array]::IndexOf($entriesArray, $entry)
                            if ($idx -ge 0 -and ($idx + 1) -eq $localJumpEntryIndex) {
                                $entryMatches = $localHighlight
                            }
                        }

                        $entryIndex = [Array]::IndexOf($entriesArray, $entry)
                        if ($entryIndex -lt 0) { $entryIndex = 0 }

                        Edit-EntryFields `
                            -Entry $entry `
                            -HighlightMatches $entryMatches `
                            -TypeName $typeName `
                            -EntryIndex ($entryIndex + 1) `
                            -XmlPath $global:currentPath | Out-Null
                    }
                }
            }

            '4' {
                Backup-XmlFile -Path $global:currentPath
                Save-XmlDocument -Xml $global:currentXml -Path $global:currentPath
                Mark-AllPendingChangesCommittedForFile -XmlPath $global:currentPath

                Write-Host ""
                Write-Host "[i] A backup was made when you saved changes." -ForegroundColor Cyan
                Write-Host "[i] Backup files are located inside the 'BackupXMLs' folder next to your XML files." -ForegroundColor Cyan

                $changePath = Get-ChangeLogPathForXml -XmlPath $global:currentPath
                if (Test-Path $changePath) {
                    Write-Host ""
                    Write-Info "Changes log is stored at: $changePath"
                }

                Write-Host ""
                Read-Host "Press Enter to return to the XML file list" | Out-Null

                $editing = $false
            }

            '5' {
                $answer = Read-Host "Are you sure you want to discard all in-memory changes for this file and reload from disk? (Y/N)"
                if ($answer -match '^[Yy]$') {
                    Write-Bad "Throwing away in-memory changes and reloading from disk."
                    $global:currentXml = Load-XmlDocument -Path $global:currentPath
                    $editing = $false
                } else {
                    Write-Info "Discard cancelled, in-memory changes are still present."
                    Start-Sleep -Seconds 1
                }
            }

            '6' {
                Review-SavedEditsForFile -XmlPath $global:currentPath -SavedChanges (Load-ChangesForFile -XmlPath $global:currentPath)
            }

            'Q' {
                $editing = $false
            }

            default {
                Write-Bad "Not a valid option."
                Start-Sleep -Seconds 2
            }
        }
    }
}

function Get-EntryKeywordMatches {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [string[]]$Words,
        [string[]]$Phrases
    )

    $results = @()

    $xml = $null
    try {
        $xml = Load-XmlDocument -Path $File.FullName
    } catch {
        return @()
    }

    $wordsLower   = @()
    $phrasesLower = @()

    if ($Words) {
        foreach ($w in $Words) {
            $t = $w.Trim()
            if (-not [string]::IsNullOrWhiteSpace($t)) {
                $wordsLower += $t.ToLower()
            }
        }
    }

    if ($Phrases) {
        foreach ($p in $Phrases) {
            $t = $p.Trim()
            if (-not [string]::IsNullOrWhiteSpace($t)) {
                $phrasesLower += $t.ToLower()
            }
        }
    }

    if ($wordsLower.Count -eq 0 -and $phrasesLower.Count -eq 0) {
        return @()
    }

    $groups = Get-EntryTypes -Xml $xml
    foreach ($g in $groups) {
        $typeName = $g.Name
        $entries  = Get-EntriesOfType -Xml $xml -ElementName $typeName
        if (-not $entries) { continue }

        $entryIndex = 1
        foreach ($entry in $entries) {

            $fields       = Get-EntryFieldsList -Entry $entry
            $valueMatches = @()

            foreach ($field in $fields) {
                $pathLower = $field.Path.ToLower()

                $value = if ($field.Kind -eq 'Attribute') {
                    $field.Node.Value
                } else {
                    $field.Node.InnerText
                }
                $valueStr   = [string]$value
                $valueLower = $valueStr.ToLower()

                $displayText  = [string]$field.Display
                $displayLower = $displayText.ToLower()

                $combined = $pathLower + "`n" + $valueLower + "`n" + $displayLower

                $hit = $true

                foreach ($w in $wordsLower) {
                    if (-not ($pathLower.Contains($w) -or $valueLower.Contains($w) -or $displayLower.Contains($w))) {
                        $hit = $false
                        break
                    }
                }

                if ($hit) {
                    foreach ($p in $phrasesLower) {
                        if (-not $combined.Contains($p)) {
                            $hit = $false
                            break
                        }
                    }
                }

                if ($hit) {
                    $valueMatches += [pscustomobject]@{
                        FieldPath = $field.Path
                        Display   = $field.Display
                        ValueRaw  = $valueStr
                    }
                }
            }

            if ($valueMatches.Count -gt 0) {
                $results += [pscustomobject]@{
                    TypeName     = $typeName
                    EntryIndex   = $entryIndex
                    ValueMatches = $valueMatches
                }
            }

            $entryIndex++
        }
    }

    return $results
}

function Search-XmlFilesByKeyword {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    $searchScope = $Files

    while ($true) {

        Clear-Screen
        Write-Title "Search XML files by keyword"

        $keyword = Read-Host "Enter keyword(s) to search for in $($searchScope.Count) XML file(s) (blank = go back)"
        if ([string]::IsNullOrWhiteSpace($keyword)) {
            return
        }

        $phrases = @()
        $wordSource = $keyword

        $matches = [regex]::Matches($keyword, '"([^"]+)"')
        foreach ($m in $matches) {
            $phrases += $m.Groups[1].Value
        }
        if ($matches.Count -gt 0) {
            $wordSource = [regex]::Replace($keyword, '"[^"]+"', ' ')
        }

        $words = @()
        $split = $wordSource.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        foreach ($w in $split) {
            $words += $w
        }

        if ($words.Count -eq 0 -and $phrases.Count -eq 0) {
            Write-Bad "No usable words or phrases in your search."
            Start-Sleep -Seconds 1
            continue
        }

        Write-Host ""
        Write-Host ("[i] Searching for '{0}' in {1} XML file{2}. This may take a moment..." -f `
            $keyword,
            $searchScope.Count,
            ($(if ($searchScope.Count -eq 1) { "" } else { "s" }))) -ForegroundColor DarkGray
        Write-Host ""

        $results = @()

        foreach ($file in $searchScope) {
            $entryMatches = Get-EntryKeywordMatches -File $file -Words $words -Phrases $phrases
            if ($entryMatches -and $entryMatches.Count -gt 0) {

                $totalValueMatches = 0
                foreach ($em in $entryMatches) {
                    $totalValueMatches += $em.ValueMatches.Count
                }

                $results += [pscustomobject]@{
                    File         = $file
                    TotalMatches = $totalValueMatches
                    EntryMatches = $entryMatches
                }
            }
        }

        if ($results.Count -eq 0) {
            Write-Bad "No XML files in this scope contained '$keyword'."
            Start-Sleep -Seconds 1
            continue
        }

        while ($true) {
            Clear-Screen
            Write-Title "Files containing '$keyword'"

            for ($i = 0; $i -lt $results.Count; $i++) {
                $idx = $i + 1
                $r   = $results[$i]

                Write-Host ("[{0}] {1} (total matches: {2})" -f `
                    $idx,
                    $r.File.Name,
                    $r.TotalMatches) -ForegroundColor Cyan

                foreach ($em in $r.EntryMatches) {
                    $valueCount = $em.ValueMatches.Count
                    $valueWord  = if ($valueCount -eq 1) { "match" } else { "matches" }

                    Write-Host ("   {0} entry #{1}: {2} {3}" -f `
                        $em.TypeName,
                        $em.EntryIndex,
                        $valueCount,
                        $valueWord) -ForegroundColor Yellow

                    foreach ($s in $em.ValueMatches) {
                        Write-Host ("      • {0}" -f $s.Display)
                    }

                    Write-Host ""
                }

                Write-Host ""
            }

            Write-Host "Type a number to pick an XML file."
            Write-Host "Type S to search again (only inside these result XML files)."
            Write-Host "Type R to start a new search on ALL XML files."
            Write-Host "Type Q to go back to the main menu."
            Write-Host ""

            $choice = Read-Host "Your choice"

            if ($choice -match '^[Qq]$') {
                return
            }

            if ($choice -match '^[Rr]$') {
                $searchScope = $Files
                break
            }

            if ($choice -match '^[Ss]$') {
                $searchScope = $results.File
                break
            }

            if ($choice -as [int]) {
                $num = [int]$choice
                if ($num -ge 1 -and $num -le $results.Count) {
                    $picked = $results[$num - 1]

                    if ($picked.EntryMatches -and $picked.EntryMatches.Count -gt 0) {
                        while ($true) {
                            Clear-Screen
                            Write-Title ("Matches in {0} for '{1}'" -f $picked.File.Name, $keyword)
                            Write-Host ""

                            for ($j = 0; $j -lt $picked.EntryMatches.Count; $j++) {
                                $em         = $picked.EntryMatches[$j]
                                $entryIdx   = $j + 1
                                $valueCount = $em.ValueMatches.Count
                                $valueWord  = if ($valueCount -eq 1) { "match" } else { "matches" }

                                Write-Host ("[{0}] {1} entry #{2} ({3} {4})" -f `
                                    $entryIdx,
                                    $em.TypeName,
                                    $em.EntryIndex,
                                    $valueCount,
                                    $valueWord) -ForegroundColor Yellow
                            }

                            Write-Host ""
                            Write-Host "Type a number to jump straight to that entry."
                            Write-Host "Press Enter with nothing typed to open the file normally."
                            Write-Host "Type Q to go back to the file list."
                            Write-Host ""

                            $subChoice = Read-Host "Your choice"

                            if ([string]::IsNullOrWhiteSpace($subChoice)) {
                                Edit-XmlFile -File $picked.File
                                return
                            }

                            if ($subChoice -match '^[Qq]$') {
                                break
                            }

                            if ($subChoice -as [int]) {
                                $subNum = [int]$subChoice
                                if ($subNum -ge 1 -and $subNum -le $picked.EntryMatches.Count) {
                                    $em = $picked.EntryMatches[$subNum - 1]

                                    Edit-XmlFile -File $picked.File `
                                        -JumpTypeName     $em.TypeName `
                                        -JumpEntryIndex   $em.EntryIndex `
                                        -HighlightMatches $em.ValueMatches
                                    return
                                }
                            }

                            Write-Bad "Not a valid choice. Try again."
                            Start-Sleep -Seconds 1
                        }

                        continue
                    }

                    Edit-XmlFile -File $picked.File
                    return
                }
            }

            Write-Bad "Not a valid choice. Try again."
            Start-Sleep -Seconds 1
        }
    }
}

try {
    Clear-Screen
    Write-Title "LSR XML Helper"

    $RootFolder                 = Get-RootFolderInteractive -InitialValue $RootFolder
    $Script:RootFolderForLogs   = $RootFolder
    $xmlFiles                   = Get-XmlFiles -Folder $RootFolder

    if (-not $Script:SkipUpdateCheck) {
        Check-ForUpdate
    }

    if ($xmlFiles.Count -eq 0) {
        Write-Bad "No XML files found in that folder. Nothing to do."
        return
    }

    $global:currentXml   = $null
    $global:currentPath  = $null
    $global:originalText = $null

    while ($true) {
    Clear-Screen
    Write-Title "Main menu"

    Write-Host "[1] Pick an XML file to edit"
    Write-Host "[2] Search all XML files for keyword(s)"
    Write-Host "[3] Review saved edits"
    Write-Host "[4] Restore Backups"
    Write-Host "[5] Refresh XML file list"
    Write-Host "[6] Settings & Info"
    Write-Host "[Q] Quit"

    $mainChoice = Read-Host "Your choice"

    switch ($mainChoice.ToUpper()) {

        '1' {
            $file = Select-FileFromList -Files $xmlFiles
            if (-not $file) { continue }
            Edit-XmlFile -File $file
        }

        '2' {
            Search-XmlFilesByKeyword -Files $xmlFiles
        }

        '3' {
            Review-AllSavedEdits -XmlFiles $xmlFiles
        }

        '4' {
            Show-BackupBrowser -XmlFiles $xmlFiles
        }

        '5' {
            try {
                $xmlFiles = Get-XmlFiles -Folder $RootFolder
                if ($xmlFiles.Count -eq 0) {
                    Write-Bad "No XML files found in that folder after refresh."
                    Start-Sleep -Seconds 2
                }
                else {
                    Write-Good ("Refreshed XML file list. {0} file{1} found." -f `
                        $xmlFiles.Count,
                        $(if ($xmlFiles.Count -eq 1) { "" } else { "s" }))
                    Start-Sleep -Seconds 1
                }
            }
            catch {
                Write-Bad ("Could not refresh XML file list: {0}" -f $_.Exception.Message)
                Write-LogError $_
                Start-Sleep -Seconds 2
            }
        }

        '6' {
            Show-InfoMenu -RootFolder $RootFolder
        }

        'Q' {
            $pendingInfo = Get-PendingChangesSummaryForAllFiles -XmlFiles $xmlFiles

            if (-not $pendingInfo -or $pendingInfo.Count -eq 0) {
                Write-Host "[i] Closing LSR XML Helper." -ForegroundColor DarkYellow
                Start-Sleep -Seconds 2
                $Script:SkipExitPause = $true
                return
            }

            $exitQuitMenu = $false
            while (-not $exitQuitMenu) {
                Clear-Screen
                Write-Title "Pending saved edits detected"

                foreach ($item in $pendingInfo) {
                    $fileName = $item.File.Name
                    $count    = $item.PendingCount
                    $suffix   = if ($count -eq 1) { "change" } else { "changes" }
                    Write-Host ("{0} : {1} pending {2}" -f $fileName, $count, $suffix) -ForegroundColor Yellow
                }

                Write-Host ""
                Write-Host "[1] Apply all pending edits, then quit"
                Write-Host "[2] Discard all pending edits, then quit"
                Write-Host "[3] Return to the main menu"
                Write-Host "[Q] Quit and leave pending edits untouched"
                Write-Host ""

                $answer = Read-Host "Select an option"

                switch ($answer.ToUpper()) {

                    '1' {
                        Apply-AllPendingChangesForAllFiles -XmlFiles $xmlFiles
                        Write-Host ""
                        Write-Host "[i] Closing LSR XML Helper." -ForegroundColor DarkYellow
                        Start-Sleep -Seconds 2
                        $Script:SkipExitPause = $true
                        return
                    }

                    '2' {
                        foreach ($item in $pendingInfo) {
                            Remove-PendingChangesForFile -XmlPath $item.XmlPath
                        }
                        Write-Host ""
                        Write-Host "[i] All pending saved edits discarded." -ForegroundColor Yellow
                        Start-Sleep -Seconds 2
                        $Script:SkipExitPause = $true
                        return
                    }

                    '3' {
                        $exitQuitMenu = $true
                    }

                    'Q' {
                        Write-Host ""
                        Write-Host "[i] Leaving pending edits in place and closing." -ForegroundColor Yellow
                        Start-Sleep -Seconds 2
                        $Script:SkipExitPause = $true
                        return
                    }

                    default {
                        Write-Bad "Not a valid option."
                        Start-Sleep -Seconds 1
                    }
                }
            }
        }

        default {
            Write-Bad "Not a valid option."
            Start-Sleep -Seconds 1

           }
        }
    }
}

catch {
    Write-Host ""
    Write-Host "================ ERROR ================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-LogError $_
}
finally {
    if (-not $Script:SkipExitPause) {
        Write-Host ""
        Read-Host "Press Enter to close LSR XML Helper"
    }
}