[CmdletBinding()]
param(
    [string]$RootFolder
)

$ErrorActionPreference           = "Stop"
$Script:ShownBackupInfoShown     = $false
$Script:SkipExitPause            = $false
$Script:PendingJumpTypeName      = $null
$Script:PendingJumpEntryIndex    = 0
$Script:PendingHighlightMatches  = $null
$Script:ScriptPath               = $MyInvocation.MyCommand.Path
$Script:AppDataDir               = Join-Path $env:LOCALAPPDATA "LSR-XML-Helper"
$Script:LocalVersionFile         = Join-Path $Script:AppDataDir "version.txt"
$Script:RemoteVersionUrl         = "https://pastebin.com/raw/56yTg6aw"
$Script:RemoteScriptUrl          = "https://drive.usercontent.google.com/download?id=1uunlxT5bV5sXCDO-OGaIOT4QFuuIF4G2&export=download&confirm=t"

if (-not (Test-Path $Script:AppDataDir)) {
    New-Item -ItemType Directory -Path $Script:AppDataDir -Force | Out-Null
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

function Get-ChangeLogPathForXml {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XmlPath
    )

    $dir      = Split-Path $XmlPath -Parent
    $fileName = Split-Path $XmlPath -Leaf
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)

    $changesDir = Join-Path $dir "LSR-Changes"
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
        if ($data -is [System.Collections.IEnumerable]) {
            return @($data)
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

    $all = @($existing + $Change)
    try {
        $all | ConvertTo-Json -Depth 8 | Set-Content -Path $changesPath -Encoding UTF8
        Write-Info "Recorded change in '$changesPath'."
    } catch {
        Write-Bad "Failed to write change log '$changesPath': $($_.Exception.Message)"
    }
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

    $dir       = Split-Path $Path -Parent
    $name      = Split-Path $Path -Leaf
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $backupDir = Join-Path $dir "BackupXMLs"
    if (-not (Test-Path $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

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
        Write-Host "Type S to search $ItemWord(s)."
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

    $entriesArray = @($Entries)
    $highlightKeys = @{}
    if ($HighlightMatches) {
        foreach ($hm in $HighlightMatches) {
            if ($hm.FieldPath -and $hm.ValueRaw -ne $null) {
                $key = "{0}::{1}" -f $hm.FieldPath, $hm.ValueRaw
                $highlightKeys[$key] = $true
            }
        }
    }

    $selected = Select-ItemWithSearch `
        -Items $entriesArray `
        -GetLabel {
            param($e)

            $idx = [array]::IndexOf($entriesArray, $e)
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

                    $key    = "{0}::{1}" -f $field.Path, $value
                    $prefix = if ($highlightKeys.ContainsKey($key)) { "   * " } else { "     " }

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
                $key = "{0}::{1}" -f $hm.FieldPath, $hm.ValueRaw
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

                $key  = "{0}::{1}" -f $field.Path, $value
                $line = "   {0}" -f $field.Display

                if ($highlightKeys.ContainsKey($key)) {
                    Write-Host $line -ForegroundColor Yellow
                } else {
                    Write-Host $line
                }
            }

            Write-Host ""

        } else {

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

    $changedFieldPaths = @()
    $hadAnyFields      = $false
    $changedAny        = $false

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
            -IsChanged { param($f) $changedFieldPaths -contains $f.Path } `
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

        if ($changedFieldPaths -notcontains $field.Path) {
            $changedFieldPaths += $field.Path
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
        [Parameter(Mandatory = $true)][string]$XmlPath
    )

    if ($Change.PSObject.Properties.Name -contains 'Type') {
        $normalized = $Change
        Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized
        return
    }

    if ($Change.PSObject.Properties.Name -contains 'Kind') {
        if ($Change.Kind -eq 'ADD') {
            $entryXml = $null

            if ($Change.PSObject.Properties.Name -contains 'EntryXml' -and $Change.EntryXml) {
                $entryXml = $Change.EntryXml
            } elseif ($Change.PSObject.Properties.Name -contains 'Fields') {
                $tmpDoc = New-Object System.Xml.XmlDocument
                $rootName = $Change.EntryType
                $rootNode = $tmpDoc.CreateElement($rootName)
                $tmpDoc.AppendChild($rootNode) | Out-Null

                foreach ($field in $Change.Fields) {
                    $parts = $field.Path.Split(".")
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
                Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized
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

                Apply-ChangeRecord -Xml $XmlRoot -XmlPath $XmlPath -Change $normalized
            }

            return
        }
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

    while ($true) {
        Clear-Screen
        Write-Title "Saved edits for:"
        Write-Host $XmlPath
        Write-Host ""

        if (-not $SavedChanges -or $SavedChanges.Count -eq 0) {
            Write-Bad "Nothing tracked for this XML anymore."
            Read-Host "Press Enter to return"
            return
        }

        $changes = @($SavedChanges)

        $index = 0
        foreach ($change in $changes) {
            $index++

            if ($change.Type -eq 'EditField') {
                Write-Host ("[{0}] EDIT Type='{1}', EntryIndex={2}, Field='{3}'" -f `
                    $index, $change.TypeName, $change.EntryIndex, $change.FieldPath) -ForegroundColor Cyan
                Write-Host ("     {0} -> {1}" -f $change.OldValue, $change.NewValue) -ForegroundColor Green
                Write-Host ""
            }
            elseif ($change.Type -eq 'AddEntry') {
                Write-Host ("[{0}] ADD Type='{1}' (new entry)" -f $index, $change.TypeName) -ForegroundColor Yellow

                try {
                    if ($change.EntryXml) {
                        $tmpDoc  = New-Object System.Xml.XmlDocument
                        $tmpDoc.LoadXml($change.EntryXml)
                        $tmpEntry = [System.Xml.XmlElement]$tmpDoc.DocumentElement
                        $summaryFields = Get-EntryFieldsList -Entry $tmpEntry
                        foreach ($sf in $summaryFields) {
                            Write-Host ("     {0}" -f $sf.Display) -ForegroundColor Green
                        }
                    }
                } catch {
                    Write-Bad "     (Could not read stored EntryXml)"
                }

                Write-Host ""
            }
            else {
                Write-Host ("[{0}] Unknown change type '{1}'" -f $index, $change.Type) -ForegroundColor Yellow
                Write-Host ""
            }
        }

        Write-Host "--------------------------------------------------------------"
        Write-Host "Options:"
        Write-Host "  [1] Apply ALL valid changes (in memory only)"
        Write-Host "  [2] Apply ONE change by number"
        Write-Host "  [3] Delete ONE change by number"
        Write-Host "  [4] Delete ALL saved changes for this XML"
        Write-Host "  [5] Save XML now (backup + file write)"
        Write-Host "  [6] Reload XML from disk (discard in-memory)"
        Write-Host "  [7] Go back"
        Write-Host ""

        $choice = Read-Host "Your choice"

        switch ($choice) {

            '7' {
                return
            }

            '4' {
                $SavedChanges.Clear()
                Write-Good "Cleared all saved edits"
                Start-Sleep 1
                continue
            }

            '3' {
                $target = Read-Host "Delete which one? (number)"
                if ($target -as [int]) {
                    $n = [int]$target
                    if ($n -ge 1 -and $n -le $SavedChanges.Count) {
                        $SavedChanges.RemoveAt($n - 1)
                        Write-Good "Removed change #$n"
                    } else {
                        Write-Bad "Invalid number"
                    }
                }
                Start-Sleep 1
                continue
            }

            '2' {
                $target = Read-Host "Apply which change? (number)"
                if ($target -as [int]) {
                    $n = [int]$target
                    if ($n -ge 1 -and $n -le $SavedChanges.Count) {
                        Apply-SingleSavedChange -XmlRoot $Xml -Change $SavedChanges[$n - 1] -XmlPath $XmlPath
                        Write-Good "Applied change #$n"
                    } else {
                        Write-Bad "Invalid number"
                    }
                }
                Start-Sleep 1
                continue
            }

            '1' {
                foreach ($c in $SavedChanges) {
                    Apply-SingleSavedChange -XmlRoot $Xml -Change $c -XmlPath $XmlPath
                }
                Write-Good "Applied ALL changes (in memory only)"
                Start-Sleep 1
                continue
            }

            '5' {
                Backup-XmlFile -Path $XmlPath
                Save-XmlDocument -Xml $Xml -Path $XmlPath

                $changePath = Get-ChangeLogPathForXml -XmlPath $XmlPath

                Write-Host ""
                Write-Good "XML written to disk and backed up"

                if ($changePath) {
                    Write-Host ""
                    Write-Info "Changes log is stored at: $changePath"
                }

                Write-Host ""
                Read-Host "Press Enter to continue"
                continue
            }

            '6' {
                Write-Bad "Reloading XML from disk..."
                $Xml = Load-XmlDocument -Path $XmlPath
                Start-Sleep 1
            }

            default {
                Write-Bad "Invalid input"
                Start-Sleep 1
            }
        }
    }
}


function Review-AllSavedEdits {
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
                File      = $file
                XmlPath   = $xmlPath
                Changes   = $changes
                Count     = $changes.Count
            }
        }
    }

    if ($items.Count -eq 0) {
        Write-Info "No saved edits found for any XML in this folder."
        Read-Host "Press Enter to go back to the main menu" | Out-Null
        return
    }

    while ($true) {
        Clear-Screen
        Write-Title "Review saved edits (all XML files)"

        for ($i = 0; $i -lt $items.Count; $i++) {
            $idx   = $i + 1
            $item  = $items[$i]
            $fname = $item.File.Name

            Write-Host ("[{0}] {1} ({2} change{3})" -f `
                $idx, $fname, $item.Count, $(if ($item.Count -eq 1) { "" } else { "s" })) -ForegroundColor Green
            Write-Host ("     {0}" -f $item.XmlPath) -ForegroundColor DarkGray
            Write-Host ""
        }

        Write-Host "Options:"
        Write-Host "  A = Show ALL edits for ALL XML files (read-only list)"
        Write-Host "  Number = Review edits for a single XML (apply / delete etc.)"
        Write-Host "  Q = Go back to main menu"
        Write-Host ""

        $choice = Read-Host "Your choice"

        if ($choice -match '^[Qq]$') {
            return
        }

        if ($choice -match '^[Aa]$') {
            Clear-Screen
            Write-Title "All saved edits for all XML files"
            foreach ($item in $items) {
                Write-Host ""
                Write-Host $item.File.Name -ForegroundColor Green
                Write-Host $item.XmlPath -ForegroundColor DarkGray
                Write-Host ""

                $dummyXml = $null
                try {
                    $dummyXml = Load-XmlDocument -Path $item.XmlPath
                } catch {
                    Write-Bad "Could not load XML for preview: $($item.XmlPath): $($_.Exception.Message)"
                }

                $changes = $item.Changes
                for ($i = 0; $i -lt $changes.Count; $i++) {
                    $idx    = $i + 1
                    $change = $changes[$i]

                    if ($change.Type -eq 'EditField') {
                        Write-Host (" [{0}] EDIT  Type='{1}', EntryIndex={2}, Field='{3}'" -f `
                            $idx, $change.TypeName, $change.EntryIndex, $change.FieldPath) -ForegroundColor Cyan
                        Write-Host ("      {0} -> {1}" -f $change.OldValue, $change.NewValue)
                        Write-Host ""
                    }
                    elseif ($change.Type -eq 'AddEntry') {
                        Write-Host (" [{0}] ADD   Type='{1}' (new entry)" -f $idx, $change.TypeName) -ForegroundColor Yellow

                        try {
                            if ($change.EntryXml) {
                                $tmpDoc = New-Object System.Xml.XmlDocument
                                $tmpDoc.LoadXml($change.EntryXml)
                                $tmpEntry = [System.Xml.XmlElement]$tmpDoc.DocumentElement
                                $summaryFields = Get-EntryFieldsList -Entry $tmpEntry
                                foreach ($sf in $summaryFields) {
                                    Write-Host ("      {0}" -f $sf.Display)
                                }
                            }
                        } catch {
                            Write-Bad "      (Could not parse stored EntryXml: $($_.Exception.Message))"
                        }

                        Write-Host ""
                    }
                    else {
                        Write-Host (" [{0}] Unknown change type '{1}'" -f $idx, $change.Type) -ForegroundColor Yellow
                        Write-Host ""
                    }
                }
                Write-Host "--------------------------------------------------------"
            }

            Write-Host ""
            Read-Host "Press Enter to go back to the review menu" | Out-Null
            continue
        }

        if ($choice -as [int]) {
            $num = [int]$choice
            if ($num -ge 1 -and $num -le $items.Count) {
                $item   = $items[$num - 1]
                $xml    = Load-XmlDocument -Path $item.XmlPath
                Review-SavedEditsForFile -XmlPath $item.XmlPath -SavedChanges $item.Changes
                return
            }
        }

        Write-Bad "Not a valid choice."
        Start-Sleep -Seconds 1
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
    Write-Host "[4] Save changes and return to XML file list"
    Write-Host "[5] Discard changes and return to XML file list"
    Write-Host "[6] Review saved edits for this file"
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
            }
        } catch {
            Write-Bad "Could not read saved folder. Will ask again."
        }
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
            RootFolder = $pickedFolder
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
                Write-Bad "Throwing away in memory changes and reloading from disk."
                $global:currentXml = Load-XmlDocument -Path $global:currentPath
                $editing = $false
            }

            '6' {
                Review-SavedEditsForFile -XmlPath $global:currentPath -SavedChanges (Load-ChangesForFile -XmlPath $global:currentPath)
            }

            default {
                Write-Bad "Not a valid option."
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Get-EntryKeywordMatches {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$Keyword
    )

    $results = @()

    $xml = $null
    try {
        $xml = Load-XmlDocument -Path $File.FullName
    } catch {
        return @()
    }

    $keywordLower = $Keyword.ToLower()

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
                $value = if ($field.Kind -eq 'Attribute') {
                    $field.Node.Value
                } else {
                    $field.Node.InnerText
                }

                if (-not [string]::IsNullOrWhiteSpace($value)) {
                    $valLower = $value.ToString().ToLower()
                    if ($valLower.Contains($keywordLower)) {
                        $valueMatches += [pscustomobject]@{
                            FieldPath = $field.Path
                            Display   = $field.Display
                            ValueRaw  = $value
                        }
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

        Write-Host ""
        Write-Host ("[i] Searching for '{0}' in {1} XML file{2}. This may take a moment..." -f `
            $keyword,
            $searchScope.Count,
            ($(if ($searchScope.Count -eq 1) { "" } else { "s" }))) -ForegroundColor DarkGray
        Write-Host ""

        $results = @()

        foreach ($file in $searchScope) {
            $entryMatches = Get-EntryKeywordMatches -File $file -Keyword $keyword
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

                Write-Host ("[{0}] {1} (total value matches: {2})" -f `
                    $idx,
                    $r.File.Name,
                    $r.TotalMatches) -ForegroundColor Cyan

                foreach ($em in $r.EntryMatches) {
                    $valueCount = $em.ValueMatches.Count
                    $valueWord  = if ($valueCount -eq 1) { "value match" } else { "value matches" }

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
                                $valueWord  = if ($valueCount -eq 1) { "value match" } else { "value matches" }

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

    Check-ForUpdate

    $RootFolder = Get-RootFolderInteractive -InitialValue $RootFolder
    $xmlFiles   = Get-XmlFiles -Folder $RootFolder

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
        Write-Host "[2] Search all XML files for keyword(s) and list only the XML files that contain a match"
        Write-Host "[3] Review saved edits"
        Write-Host "[4] Quit"

        $mainChoice = Read-Host "Your choice"

        switch ($mainChoice) {
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
                Write-Host "[i] Closing LSR XML Helper." -ForegroundColor DarkYellow
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
catch {
    Write-Host ""
    Write-Host "================ ERROR ================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {
    if (-not $Script:SkipExitPause) {
        Write-Host ""
        Read-Host "Press Enter to close LSR XML Helper"
    }
}
