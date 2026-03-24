[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Win32Native {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
}
"@

$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:AppRoot "macros.json"
$script:Colors = @{
    Background = [System.Drawing.ColorTranslator]::FromHtml("#F4F7FB")
    Surface = [System.Drawing.Color]::White
    SurfaceAlt = [System.Drawing.ColorTranslator]::FromHtml("#EAF0F8")
    Primary = [System.Drawing.ColorTranslator]::FromHtml("#0F5FDB")
    PrimaryDark = [System.Drawing.ColorTranslator]::FromHtml("#0A4AAE")
    Accent = [System.Drawing.ColorTranslator]::FromHtml("#0A7A5A")
    Text = [System.Drawing.ColorTranslator]::FromHtml("#182433")
    Muted = [System.Drawing.ColorTranslator]::FromHtml("#5D6B7E")
    Border = [System.Drawing.ColorTranslator]::FromHtml("#D7DFEA")
    Warning = [System.Drawing.ColorTranslator]::FromHtml("#9A6700")
}
$script:LogDirectory = Join-Path $script:AppRoot "logs"
$script:LogPath = Join-Path $script:LogDirectory "macro-runs.log"

function Initialize-Logging {
    if (-not (Test-Path -LiteralPath $script:LogDirectory)) {
        [void](New-Item -ItemType Directory -Path $script:LogDirectory -Force)
    }

    if (-not (Test-Path -LiteralPath $script:LogPath)) {
        [System.IO.File]::WriteAllText($script:LogPath, "", [System.Text.Encoding]::UTF8)
    }
}

function Get-MacroLogPath {
    Initialize-Logging
    return $script:LogPath
}

function Write-MacroLog {
    param(
        [string]$MacroName,
        [string]$Status,
        [string]$Details
    )

    Initialize-Logging
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] {1} - Status: {2} - Details: {3}" -f $timestamp, $MacroName, $Status, $Details
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function New-AppFont {
    param(
        [float]$Size = 9,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )

    return New-Object System.Drawing.Font("Segoe UI", $Size, $Style)
}

function Resolve-AppRelativePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $Path
    }

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    return [System.IO.Path]::GetFullPath((Join-Path $script:AppRoot $Path))
}

function Get-MacroResolvedPath {
    param([hashtable]$Macro)

    $configuredPath = Resolve-AppRelativePath -Path ([string]$Macro.path)
    if (-not [string]::IsNullOrWhiteSpace($configuredPath) -and (Test-Path -LiteralPath $configuredPath)) {
        return $configuredPath
    }

    $fileName = [System.IO.Path]::GetFileName([string]$Macro.path)
    if (-not [string]::IsNullOrWhiteSpace($fileName)) {
        $macrosFolderPath = Join-Path (Join-Path $script:AppRoot "macros") $fileName
        if (Test-Path -LiteralPath $macrosFolderPath) {
            return $macrosFolderPath
        }
    }

    return $configuredPath
}

function New-DefaultConfig {
    $macros = [System.Collections.ArrayList]::new()
    [void]$macros.Add(@{
        name = "Rysunki do PDF i DXF v4"
        description = "Eksport rysunkow do PDF i DXF."
        path = ".\\macros\\Macro rysunków pod pdf i pod dxf v4.swp"
        module = "Macro_rysunkow_pod_pdf_i_"
        procedure = "main"
        category = "Eksport"
        launchMode = "attached_or_start"
        requiresOpenDoc = $false
        requirements = "Makro moze prosic o wskazanie folderu lub plikow po uruchomieniu."
    })
    [void]$macros.Add(@{
        name = "Zapis czesci v4"
        description = "Zapis czesci z numerowaniem i podzialem na foldery."
        path = ".\\macros\\MacroZapisuCzesciv4_dzialajacy.swp"
        module = "MacroZapisuCzesciv41"
        procedure = "main"
        category = "Zapis"
        launchMode = "attached_or_start"
        requiresOpenDoc = $true
        requiredDocType = "assembly"
        requirements = "Przed uruchomieniem otworz w SolidWorks zlozenie z wirtualnymi komponentami. To makro pracuje na aktualnie otwartym zlozeniu."
    })
    [void]$macros.Add(@{
        name = "Export do PDF i DXF z rozpoznaniem"
        description = "Starsza wersja eksportu z rozpoznaniem."
        path = ".\\macros\\V1 - export do pdf i dxf z rozpoznaniem.swp"
        module = "V1__export_do_pdf_i_dxf_"
        procedure = "main"
        category = "Eksport"
        launchMode = "attached_or_start"
        requiresOpenDoc = $true
        requirements = "Najlepiej uruchamiac przy juz otwartym odpowiednim pliku w SolidWorks."
    })

    return @{
        solidworks = @{
            visible = $true
        }
        macros = $macros
    }
}

function ConvertTo-MacroList {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Items
    )

    $list = [System.Collections.ArrayList]::new()
    foreach ($item in $Items) {
        if ($item -is [hashtable]) {
            [void]$list.Add($item)
        }
        else {
            [void]$list.Add([hashtable]$item)
        }
    }

    return $list
}

function ConvertTo-Hashtable {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $hash = @{}
        foreach ($key in $InputObject.Keys) {
            $hash[$key] = ConvertTo-Hashtable -InputObject $InputObject[$key]
        }
        return $hash
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $list = [System.Collections.ArrayList]::new()
        foreach ($item in $InputObject) {
            [void]$list.Add((ConvertTo-Hashtable -InputObject $item))
        }
        return $list
    }

    if ($InputObject -is [pscustomobject]) {
        $hash = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-Hashtable -InputObject $property.Value
        }
        return $hash
    }

    return $InputObject
}

function Normalize-Config {
    param([hashtable]$Config)

    if (-not $Config.ContainsKey("solidworks")) {
        $Config.solidworks = @{ visible = $true }
    }
    if (-not $Config.solidworks.ContainsKey("visible")) {
        $Config.solidworks.visible = $true
    }
    if (-not $Config.ContainsKey("macros")) {
        $Config.macros = [System.Collections.ArrayList]::new()
    }
    else {
        $Config.macros = ConvertTo-MacroList -Items $Config.macros
    }

    foreach ($macro in $Config.macros) {
        if (-not $macro.ContainsKey("launchMode") -or [string]::IsNullOrWhiteSpace([string]$macro.launchMode)) {
            $macro.launchMode = "attached_or_start"
        }
        elseif ([string]$macro.launchMode -eq "attached_only") {
            $macro.launchMode = "attached_or_start"
        }
        if (-not $macro.ContainsKey("requiresOpenDoc")) {
            $macro.requiresOpenDoc = $false
        }
        if (-not $macro.ContainsKey("requiredDocType")) {
            $macro.requiredDocType = ""
        }
        if (-not $macro.ContainsKey("requirements")) {
            $macro.requirements = ""
        }
        if (-not $macro.ContainsKey("steps")) {
            $macro.steps = ""
        }
        if (-not $macro.ContainsKey("origin")) {
            $macro.origin = ""
        }
        if (-not $macro.ContainsKey("sourceUrl")) {
            $macro.sourceUrl = ""
        }
    }

    return $Config
}

function Save-Config {
    param([hashtable]$Config)

    $json = $Config | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($script:ConfigPath, $json, [System.Text.Encoding]::UTF8)
}

function Import-Config {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        $defaultConfig = Normalize-Config -Config (New-DefaultConfig)
        Save-Config -Config $defaultConfig
        return $defaultConfig
    }

    $jsonObject = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $raw = ConvertTo-Hashtable -InputObject $jsonObject
    return Normalize-Config -Config $raw
}

function Get-MacroLabel {
    param([hashtable]$Macro)

    if ([string]::IsNullOrWhiteSpace([string]$Macro.category)) {
        return $Macro.name
    }

    return "[{0}] {1}" -f $Macro.category, $Macro.name
}

function Get-MacroGlyph {
    param([hashtable]$Macro)

    $category = ([string]$Macro.category).ToLowerInvariant()
    $name = ([string]$Macro.name).ToLowerInvariant()

    if ($name.Contains("moje: rysunki")) { return "PDF" }
    if ($name.Contains("moje: zapis")) { return "NUM" }
    if ($name.Contains("moje: export z rozpoznaniem")) { return "SEN" }
    if ($name.Contains("pobrane:")) { return "DL" }
    if ($name.Contains("bom")) { return "BOM" }
    if ($name.Contains("dxf")) { return "DXF" }
    if ($name.Contains("pdf")) { return "PDF" }
    if ($name.Contains("sheet")) { return "SHT" }
    if ($name.Contains("cut list") -or $name.Contains("cutlist")) { return "CUT" }
    if ($name.Contains("bounding")) { return "BOX" }
    if ($name.Contains("virtual")) { return "ASM" }

    switch ($category) {
        "eksport" { return "EXP" }
        "zapis" { return "SAV" }
        "rysunki" { return "DRW" }
        "czesci" { return "PRT" }
        "zlozenia" { return "ASM" }
        "github" { return "GIT" }
        default { return "MAC" }
    }
}

function Get-MacroAccentColor {
    param([hashtable]$Macro)

    $category = ([string]$Macro.category).ToLowerInvariant()
    $name = ([string]$Macro.name).ToLowerInvariant()

    if ($name.Contains("moje: rysunki")) { return [System.Drawing.ColorTranslator]::FromHtml("#0A7A5A") }
    if ($name.Contains("moje: zapis")) { return [System.Drawing.ColorTranslator]::FromHtml("#C26D00") }
    if ($name.Contains("moje: export z rozpoznaniem")) { return [System.Drawing.ColorTranslator]::FromHtml("#0F5FDB") }
    if ($name.Contains("pobrane:")) { return [System.Drawing.ColorTranslator]::FromHtml("#6B7280") }
    if ($name.Contains("bom")) { return [System.Drawing.ColorTranslator]::FromHtml("#C26D00") }
    if ($name.Contains("dxf")) { return [System.Drawing.ColorTranslator]::FromHtml("#0A7A5A") }
    if ($name.Contains("pdf")) { return [System.Drawing.ColorTranslator]::FromHtml("#C0392B") }
    if ($name.Contains("sheet")) { return [System.Drawing.ColorTranslator]::FromHtml("#00838F") }
    if ($name.Contains("bounding")) { return [System.Drawing.ColorTranslator]::FromHtml("#8E5A00") }
    if ($name.Contains("virtual")) { return [System.Drawing.ColorTranslator]::FromHtml("#7B1FA2") }

    switch ($category) {
        "eksport" { return [System.Drawing.ColorTranslator]::FromHtml("#0A7A5A") }
        "zapis" { return [System.Drawing.ColorTranslator]::FromHtml("#C26D00") }
        "rysunki" { return [System.Drawing.ColorTranslator]::FromHtml("#00838F") }
        "czesci" { return [System.Drawing.ColorTranslator]::FromHtml("#8E5A00") }
        "zlozenia" { return [System.Drawing.ColorTranslator]::FromHtml("#7B1FA2") }
        "github" { return [System.Drawing.ColorTranslator]::FromHtml("#24292F") }
        default { return $script:Colors.Primary }
    }
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = "SolidWorks Macro Center",
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    [void][System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    )
}

function Set-ButtonStyle {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BackColor = $script:Colors.SurfaceAlt,
        [System.Drawing.Color]$ForeColor = $script:Colors.Text
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.FlatAppearance.BorderColor = $script:Colors.Border
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.Font = New-AppFont -Size 9
    $Button.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
    $Button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
}

function Set-PrimaryButtonStyle {
    param([System.Windows.Forms.Button]$Button)

    Set-ButtonStyle -Button $Button -BackColor $script:Colors.Primary -ForeColor ([System.Drawing.Color]::White)
    $Button.FlatAppearance.BorderColor = $script:Colors.PrimaryDark
    $Button.Font = New-AppFont -Size 9.5 -Style Bold
}

function New-SystemBitmap {
    param([System.Drawing.Icon]$Icon)
    return $Icon.ToBitmap()
}

function Get-SolidWorksExecutablePath {
    $candidates = @(
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS 2023\SLDWORKS.exe",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS 2024\SLDWORKS.exe",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS 2025\SLDWORKS.exe",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS 2026\SLDWORKS.exe"
    )

    foreach ($path in $candidates) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    return $null
}

function Get-ActiveSolidWorksApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
    }
    catch {
        return $null
    }
}

function Get-SolidWorksProcess {
    return Get-Process -Name "SLDWORKS" -ErrorAction SilentlyContinue | Sort-Object StartTime | Select-Object -First 1
}

function Test-SolidWorksProcessRunning {
    return $null -ne (Get-SolidWorksProcess)
}

function Wait-ForSolidWorksApplication {
    param([int]$TimeoutSeconds = 90)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $app = Get-ActiveSolidWorksApplication
        if ($null -ne $app) {
            return $app
        }
        Start-Sleep -Milliseconds 1500
    } while ((Get-Date) -lt $deadline)

    return $null
}

function Start-SolidWorks {
    $exePath = Get-SolidWorksExecutablePath
    if ([string]::IsNullOrWhiteSpace($exePath)) {
        throw "Nie znaleziono pliku SLDWORKS.exe. Sprawdz instalacje SolidWorks."
    }

    Start-Process -FilePath $exePath | Out-Null
    $app = Wait-ForSolidWorksApplication -TimeoutSeconds 90
    if ($null -eq $app) {
        throw "Nie udalo sie nawiazac polaczenia z SolidWorks po uruchomieniu programu."
    }
    return $app
}

function Start-SolidWorksWithoutMacro {
    $exePath = Get-SolidWorksExecutablePath
    if ([string]::IsNullOrWhiteSpace($exePath)) {
        throw "Nie znaleziono pliku SLDWORKS.exe. Sprawdz instalacje SolidWorks."
    }

    Start-Process -FilePath $exePath | Out-Null
}

function Start-SolidWorksWithMacroArgument {
    param([string]$MacroPath)

    $exePath = Get-SolidWorksExecutablePath
    if ([string]::IsNullOrWhiteSpace($exePath)) {
        throw "Nie znaleziono pliku SLDWORKS.exe. Sprawdz instalacje SolidWorks."
    }

    Start-Process -FilePath $exePath -ArgumentList "/m `"$MacroPath`"" | Out-Null
}

function Get-OrWaitForSolidWorksApplication {
    param([int]$TimeoutSeconds = 10)

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $app = Get-ActiveSolidWorksApplication
        if ($null -ne $app) {
            return $app
        }

        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)

    return $null
}

function Get-SolidWorksMainWindowProcess {
    $candidates = @(Get-Process -Name "SLDWORKS" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($candidates.Count -eq 0) {
        return $null
    }

    return $candidates | Sort-Object StartTime -Descending | Select-Object -First 1
}

function Get-PreferredSolidWorksProcesses {
    param([string]$RequiredDocType = "")

    $processes = @(Get-Process -Name "SLDWORKS" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($processes.Count -eq 0) {
        return @()
    }

    $preferredPattern = switch ($RequiredDocType) {
        "assembly" { ".SLDASM" }
        "part" { ".SLDPRT" }
        "drawing" { ".SLDDRW" }
        default { "" }
    }

    $genericPatterns = @(".SLDASM", ".SLDPRT", ".SLDDRW")

    return @(
        $processes |
        Sort-Object `
            @{ Expression = {
                    $title = [string]$_.MainWindowTitle
                    if (-not [string]::IsNullOrWhiteSpace($preferredPattern) -and $title.ToUpperInvariant().Contains($preferredPattern)) {
                        0
                    }
                    elseif ($genericPatterns | Where-Object { $title.ToUpperInvariant().Contains($_) }) {
                        1
                    }
                    else {
                        2
                    }
                }
            }, `
            @{ Expression = { $_.StartTime }; Descending = $true }
    )
}

function Test-SolidWorksDocumentWindowOpen {
    $processes = @(Get-PreferredSolidWorksProcesses)
    if ($processes.Count -eq 0) {
        return $false
    }

    foreach ($proc in $processes) {
        $title = [string]$proc.MainWindowTitle
        $upper = $title.ToUpperInvariant()
        if ($upper.Contains(".SLDASM") -or $upper.Contains(".SLDPRT") -or $upper.Contains(".SLDDRW")) {
            return $true
        }
    }

    return $false
}

function Try-OpenMacroDialogViaCommand {
    param(
        $SwApp,
        [string]$MacroName
    )

    if ($null -eq $SwApp) {
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "RunCommand skipped because COM app object is null."
        return $false
    }

    try {
        $result = $SwApp.RunCommand(30, "")
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "ISldWorks.RunCommand(30) returned: $result"
        if ($result) {
            return $true
        }
    }
    catch {
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "ISldWorks.RunCommand(30) failed: $($_.Exception.Message)"
    }

    try {
        $activeDoc = $SwApp.ActiveDoc
        if ($null -eq $activeDoc) {
            Write-MacroLog -MacroName $MacroName -Status "trace" -Details "ModelDocExtension.RunCommand(30) skipped because ActiveDoc is null."
            return $false
        }

        $extension = $activeDoc.Extension
        if ($null -eq $extension) {
            Write-MacroLog -MacroName $MacroName -Status "trace" -Details "ModelDocExtension.RunCommand(30) skipped because Extension is null."
            return $false
        }

        $result = $extension.RunCommand(30, "")
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "IModelDocExtension.RunCommand(30) returned: $result"
        return [bool]$result
    }
    catch {
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "IModelDocExtension.RunCommand(30) failed: $($_.Exception.Message)"
        return $false
    }
}

function Activate-SolidWorksWindow {
    param([string]$RequiredDocType = "")

    $shell = $null
    try {
        $shell = New-Object -ComObject WScript.Shell
        $processes = @(Get-PreferredSolidWorksProcesses -RequiredDocType $RequiredDocType)
        if ($processes.Count -eq 0) {
            return $false
        }

        foreach ($proc in $processes) {
            try {
                if ($proc.MainWindowHandle -ne 0) {
                    $hWnd = [IntPtr]::new([int64]$proc.MainWindowHandle)
                    if ($hWnd -ne [IntPtr]::Zero) {
                        if ([Win32Native]::IsIconic($hWnd)) {
                            [void][Win32Native]::ShowWindowAsync($hWnd, 9)
                        }
                        else {
                            [void][Win32Native]::ShowWindowAsync($hWnd, 5)
                        }
                        Start-Sleep -Milliseconds 200
                        [void]$shell.AppActivate($proc.Id)
                        Start-Sleep -Milliseconds 200
                        [void][Win32Native]::SetForegroundWindow($hWnd)
                        Start-Sleep -Milliseconds 300
                        return $true
                    }
                }
            }
            catch {
            }
        }

        foreach ($proc in $processes) {
            try {
                if ($shell.AppActivate($proc.Id)) {
                    Start-Sleep -Milliseconds 300
                    return $true
                }
            }
            catch {
            }
        }

        foreach ($title in @("SOLIDWORKS", "SolidWorks", "DS SOLIDWORKS")) {
            try {
                if ($shell.AppActivate($title)) {
                    Start-Sleep -Milliseconds 300
                    return $true
                }
            }
            catch {
            }
        }

        return $false
    }
    catch {
        return $false
    }
}

function Invoke-SolidWorksRunDialogFallback {
    param(
        [string]$MacroPath,
        [string]$MacroName = "Launcher",
        [string]$RequiredDocType = "",
        $SwApp = $null
    )

    if (-not (Activate-SolidWorksWindow -RequiredDocType $RequiredDocType)) {
        throw "Nie udalo sie aktywowac glownego okna SolidWorks."
    }

    $shell = New-Object -ComObject WScript.Shell
    $dialogOpened = $false

    if ($null -ne $SwApp) {
        $dialogOpened = Try-OpenMacroDialogViaCommand -SwApp $SwApp -MacroName $MacroName
    }

    if (-not $dialogOpened) {
        Write-MacroLog -MacroName $MacroName -Status "trace" -Details "Fallback open-macro dialog through menu navigation for path: $MacroPath"
        $useLongToolsMenu = Test-SolidWorksDocumentWindowOpen
        Start-Sleep -Milliseconds 450
        $shell.SendKeys("{ESC}")
        Start-Sleep -Milliseconds 120
        $shell.SendKeys("{ESC}")
        Start-Sleep -Milliseconds 180
        $shell.SendKeys("%n")
        Start-Sleep -Milliseconds 260
        if ($useLongToolsMenu) {
            Write-MacroLog -MacroName $MacroName -Status "trace" -Details "Using long Tools menu navigation because a document window is open."
            $shell.SendKeys("{END}")
            Start-Sleep -Milliseconds 220
            for ($i = 0; $i -lt 7; $i++) {
                $shell.SendKeys("{UP}")
                Start-Sleep -Milliseconds 90
            }
        }
        else {
            Write-MacroLog -MacroName $MacroName -Status "trace" -Details "Using short Tools menu navigation."
            $shell.SendKeys("m")
            Start-Sleep -Milliseconds 220
        }
        $shell.SendKeys("{RIGHT}")
        Start-Sleep -Milliseconds 220
        $shell.SendKeys("{DOWN}")
        Start-Sleep -Milliseconds 90
        $shell.SendKeys("{DOWN}")
        Start-Sleep -Milliseconds 90
        $shell.SendKeys("{DOWN}")
        Start-Sleep -Milliseconds 180
        $shell.SendKeys("{ENTER}")
        Start-Sleep -Milliseconds 950
    }

    $shell.SendKeys("%n")
    Start-Sleep -Milliseconds 120
    $shell.SendKeys("^a")
    Start-Sleep -Milliseconds 120
    [System.Windows.Forms.Clipboard]::SetText($MacroPath)
    Start-Sleep -Milliseconds 120
    $shell.SendKeys("^v")
    Start-Sleep -Milliseconds 180
    $shell.SendKeys("{ENTER}")
}

function Get-SwDocumentTypeName {
    param($ActiveDoc)

    try {
        switch ([int]$ActiveDoc.GetType()) {
            1 { return "part" }
            2 { return "assembly" }
            3 { return "drawing" }
            default { return "unknown" }
        }
    }
    catch {
        return "unknown"
    }
}

function Assert-MacroRequirements {
    param(
        [hashtable]$Macro,
        $SwApp
    )

    if (-not [bool]$Macro.requiresOpenDoc) {
        return
    }

    $activeDoc = $SwApp.ActiveDoc
    if ($null -eq $activeDoc) {
        throw "To makro wymaga otwartego dokumentu w SolidWorks."
    }

    $requiredDocType = [string]$Macro.requiredDocType
    if (-not [string]::IsNullOrWhiteSpace($requiredDocType)) {
        $currentDocType = Get-SwDocumentTypeName -ActiveDoc $activeDoc
        if ($currentDocType -ne $requiredDocType) {
            switch ($requiredDocType) {
                "assembly" { throw "To makro wymaga otwartego zlozenia w SolidWorks." }
                "drawing" { throw "To makro wymaga otwartego rysunku w SolidWorks." }
                "part" { throw "To makro wymaga otwartej czesci w SolidWorks." }
                default { throw "Aktywny dokument ma zly typ dla tego makra." }
            }
        }
    }
}

function Invoke-SolidWorksMacro {
    param(
        [hashtable]$Macro,
        [hashtable]$Config
    )

    $macroPath = Get-MacroResolvedPath -Macro $Macro
    if (-not (Test-Path -LiteralPath $macroPath)) {
        throw "Nie znaleziono pliku makra:`n$macroPath"
    }

    $launchMode = [string]$Macro.launchMode
    if ([string]::IsNullOrWhiteSpace($launchMode)) {
        $launchMode = "attached_or_start"
    }

    $swProcessRunning = Test-SolidWorksProcessRunning
    $startedNow = $false
    $primaryError = $null

    if (-not $swProcessRunning) {
        if ([bool]$Macro.requiresOpenDoc) {
            Write-MacroLog -MacroName $Macro.name -Status "trace" -Details "No running SolidWorks process. Starting SolidWorks without macro because document is required."
            Start-SolidWorksWithoutMacro
            return @{ startedNow = $true; fallbackMode = "start-only" }
        }

        Write-MacroLog -MacroName $Macro.name -Status "trace" -Details "No running SolidWorks process. Starting with /m switch."
        Start-SolidWorksWithMacroArgument -MacroPath $macroPath
        return @{ startedNow = $true; fallbackMode = "startup-switch" }
    }

    $swApp = Get-OrWaitForSolidWorksApplication -TimeoutSeconds 10
    try {
        Write-MacroLog -MacroName $Macro.name -Status "trace" -Details "SolidWorks process detected. Using menu-dialog fallback only."
        Invoke-SolidWorksRunDialogFallback -MacroPath $macroPath -MacroName $Macro.name -RequiredDocType ([string]$Macro.requiredDocType) -SwApp $swApp
        return @{ startedNow = $startedNow; fallbackMode = "menu-dialog" }
    }
    catch {
        $primaryError = $_.Exception.Message
    }

    throw "SolidWorks nie uruchomil makra.`nPlik: $macroPath`nModul: $($Macro.module)`nProcedura: $($Macro.procedure)`nBlad: $primaryError"
}

function Show-MacroEditor {
    param([hashtable]$InitialMacro)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Makro"
    $form.Size = New-Object System.Drawing.Size(620, 560)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = $script:Colors.Background
    $form.Font = New-AppFont -Size 9

    $labels = @(
        @{ Text = "Nazwa"; Top = 20 },
        @{ Text = "Kategoria"; Top = 60 },
        @{ Text = "Opis"; Top = 100 },
        @{ Text = "Sciezka .swp"; Top = 160 },
        @{ Text = "Modul VBA"; Top = 210 },
        @{ Text = "Procedura startowa"; Top = 250 },
        @{ Text = "Tryb uruchamiania"; Top = 290 },
        @{ Text = "Wymagania"; Top = 330 },
        @{ Text = "Krok po kroku"; Top = 390 }
    )

    foreach ($item in $labels) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $item.Text
        $label.Left = 20
        $label.Top = $item.Top
        $label.Width = 130
        $form.Controls.Add($label)
    }

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Left = 160; $txtName.Top = 18; $txtName.Width = 430

    $txtCategory = New-Object System.Windows.Forms.TextBox
    $txtCategory.Left = 160; $txtCategory.Top = 58; $txtCategory.Width = 180

    $txtDescription = New-Object System.Windows.Forms.TextBox
    $txtDescription.Left = 160; $txtDescription.Top = 98; $txtDescription.Width = 430; $txtDescription.Height = 48; $txtDescription.Multiline = $true

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Left = 160; $txtPath.Top = 158; $txtPath.Width = 350

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Przegladaj..."; $btnBrowse.Left = 520; $btnBrowse.Top = 156; $btnBrowse.Width = 70

    $txtModule = New-Object System.Windows.Forms.TextBox
    $txtModule.Left = 160; $txtModule.Top = 208; $txtModule.Width = 180

    $txtProcedure = New-Object System.Windows.Forms.TextBox
    $txtProcedure.Left = 160; $txtProcedure.Top = 248; $txtProcedure.Width = 180

    $cmbLaunchMode = New-Object System.Windows.Forms.ComboBox
    $cmbLaunchMode.Left = 160; $cmbLaunchMode.Top = 288; $cmbLaunchMode.Width = 220
    $cmbLaunchMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbLaunchMode.Items.AddRange(@("attached_or_start", "attached_only", "startup_switch"))

    $chkRequiresOpenDoc = New-Object System.Windows.Forms.CheckBox
    $chkRequiresOpenDoc.Left = 390; $chkRequiresOpenDoc.Top = 290; $chkRequiresOpenDoc.Width = 200
    $chkRequiresOpenDoc.Text = "Wymaga otwartego dokumentu"

    $cmbRequiredDocType = New-Object System.Windows.Forms.ComboBox
    $cmbRequiredDocType.Left = 160; $cmbRequiredDocType.Top = 356; $cmbRequiredDocType.Width = 160
    $cmbRequiredDocType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cmbRequiredDocType.Items.AddRange(@("", "assembly", "part", "drawing"))

    $txtRequirements = New-Object System.Windows.Forms.TextBox
    $txtRequirements.Left = 330; $txtRequirements.Top = 328; $txtRequirements.Width = 260; $txtRequirements.Height = 48; $txtRequirements.Multiline = $true

    $txtSteps = New-Object System.Windows.Forms.TextBox
    $txtSteps.Left = 160; $txtSteps.Top = 388; $txtSteps.Width = 430; $txtSteps.Height = 92; $txtSteps.Multiline = $true; $txtSteps.ScrollBars = "Vertical"

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Zapisz"; $btnSave.Left = 430; $btnSave.Top = 492; $btnSave.Width = 75

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Anuluj"; $btnCancel.Left = 515; $btnCancel.Top = 492; $btnCancel.Width = 75

    Set-PrimaryButtonStyle -Button $btnSave
    Set-ButtonStyle -Button $btnCancel

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Makra SolidWorks (*.swp)|*.swp|Wszystkie pliki (*.*)|*.*"

    $btnBrowse.Add_Click({
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtPath.Text = $openFileDialog.FileName
        }
    })

    if ($null -ne $InitialMacro) {
        $txtName.Text = [string]$InitialMacro.name
        $txtCategory.Text = [string]$InitialMacro.category
        $txtDescription.Text = [string]$InitialMacro.description
        $txtPath.Text = [string]$InitialMacro.path
        $txtModule.Text = [string]$InitialMacro.module
        $txtProcedure.Text = [string]$InitialMacro.procedure
        $cmbLaunchMode.SelectedItem = [string]$InitialMacro.launchMode
        $chkRequiresOpenDoc.Checked = [bool]$InitialMacro.requiresOpenDoc
        $cmbRequiredDocType.SelectedItem = [string]$InitialMacro.requiredDocType
        $txtRequirements.Text = [string]$InitialMacro.requirements
        $txtSteps.Text = [string]$InitialMacro.steps
    }

    if ($cmbLaunchMode.SelectedIndex -lt 0) { $cmbLaunchMode.SelectedIndex = 0 }
    if ($cmbRequiredDocType.SelectedIndex -lt 0) { $cmbRequiredDocType.SelectedIndex = 0 }

    $btnSave.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            Show-Message -Text "Podaj nazwe makra." -Title "Brak danych" -Icon Warning
            return
        }

        if ([string]::IsNullOrWhiteSpace($txtPath.Text)) {
            Show-Message -Text "Podaj sciezke do pliku .swp." -Title "Brak danych" -Icon Warning
            return
        }

        $script:editorResult = @{
            name = $txtName.Text.Trim()
            category = $txtCategory.Text.Trim()
            description = $txtDescription.Text.Trim()
            path = $txtPath.Text.Trim()
            module = $txtModule.Text.Trim()
            procedure = $txtProcedure.Text.Trim()
            launchMode = [string]$cmbLaunchMode.SelectedItem
            requiresOpenDoc = $chkRequiresOpenDoc.Checked
            requiredDocType = [string]$cmbRequiredDocType.SelectedItem
            requirements = $txtRequirements.Text.Trim()
            steps = $txtSteps.Text.Trim()
        }

        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    foreach ($control in @($txtName, $txtCategory, $txtDescription, $txtPath, $btnBrowse, $txtModule, $txtProcedure, $cmbLaunchMode, $chkRequiresOpenDoc, $cmbRequiredDocType, $txtRequirements, $txtSteps, $btnSave, $btnCancel)) {
        $form.Controls.Add($control)
    }

    $script:editorResult = $null
    $dialogResult = $form.ShowDialog()
    $result = $null
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = $script:editorResult
    }

    $form.Dispose()
    return $result
}

try {
    $config = Import-Config

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "SolidWorks Macro Center"
    $form.Size = New-Object System.Drawing.Size(1060, 650)
    $form.MinimumSize = New-Object System.Drawing.Size(1060, 650)
    $form.StartPosition = "CenterScreen"
    $form.BackColor = $script:Colors.Background
    $form.Font = New-AppFont -Size 9

    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Left = 0; $headerPanel.Top = 0; $headerPanel.Width = 1060; $headerPanel.Height = 94
    $headerPanel.BackColor = $script:Colors.Primary; $headerPanel.Anchor = "Top,Left,Right"

    $headerIcon = New-Object System.Windows.Forms.PictureBox
    $headerIcon.Left = 24; $headerIcon.Top = 22; $headerIcon.Width = 48; $headerIcon.Height = 48
    $headerIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $headerIcon.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Application)

    $headerTitle = New-Object System.Windows.Forms.Label
    $headerTitle.Text = "SolidWorks Macro Center"
    $headerTitle.Left = 90; $headerTitle.Top = 20; $headerTitle.Width = 500; $headerTitle.Height = 30
    $headerTitle.ForeColor = [System.Drawing.Color]::White; $headerTitle.Font = New-AppFont -Size 18 -Style Bold

    $headerSubtitle = New-Object System.Windows.Forms.Label
    $headerSubtitle.Text = "Jedno miejsce do uruchamiania, porzadkowania i rozwijania Twoich makr."
    $headerSubtitle.Left = 92; $headerSubtitle.Top = 52; $headerSubtitle.Width = 620; $headerSubtitle.Height = 22
    $headerSubtitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml("#DCE8FF")
    $headerSubtitle.Font = New-AppFont -Size 9.5

    $leftPanel = New-Object System.Windows.Forms.Panel
    $leftPanel.Left = 20; $leftPanel.Top = 114; $leftPanel.Width = 340; $leftPanel.Height = 500
    $leftPanel.BackColor = $script:Colors.Surface; $leftPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $leftPanel.Anchor = "Top,Bottom,Left"

    $leftTitle = New-Object System.Windows.Forms.Label
    $leftTitle.Text = "Biblioteka makr"; $leftTitle.Left = 16; $leftTitle.Top = 14; $leftTitle.Width = 200
    $leftTitle.Font = New-AppFont -Size 11 -Style Bold

    $btnFilterAll = New-Object System.Windows.Forms.Button
    $btnFilterAll.Text = "Wszystkie"; $btnFilterAll.Left = 16; $btnFilterAll.Top = 44; $btnFilterAll.Width = 92; $btnFilterAll.Height = 28

    $btnFilterMine = New-Object System.Windows.Forms.Button
    $btnFilterMine.Text = "Moje"; $btnFilterMine.Left = 114; $btnFilterMine.Top = 44; $btnFilterMine.Width = 72; $btnFilterMine.Height = 28

    $btnFilterDownloaded = New-Object System.Windows.Forms.Button
    $btnFilterDownloaded.Text = "Pobrane"; $btnFilterDownloaded.Left = 192; $btnFilterDownloaded.Top = 44; $btnFilterDownloaded.Width = 88; $btnFilterDownloaded.Height = 28

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Left = 12; $listBox.Top = 82; $listBox.Width = 314; $listBox.Height = 314
    $listBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $listBox.BackColor = $script:Colors.Surface; $listBox.ForeColor = $script:Colors.Text
    $listBox.Font = New-AppFont -Size 10; $listBox.IntegralHeight = $false; $listBox.Anchor = "Top,Bottom,Left,Right"
    $listBox.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $listBox.ItemHeight = 82

    $rightPanel = New-Object System.Windows.Forms.Panel
    $rightPanel.Left = 380; $rightPanel.Top = 114; $rightPanel.Width = 658; $rightPanel.Height = 500
    $rightPanel.BackColor = $script:Colors.Surface; $rightPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $rightPanel.Anchor = "Top,Bottom,Left,Right"

    $detailsTitle = New-Object System.Windows.Forms.Label
    $detailsTitle.Text = "Szczegoly makra"; $detailsTitle.Left = 20; $detailsTitle.Top = 18; $detailsTitle.Width = 220
    $detailsTitle.Font = New-AppFont -Size 12 -Style Bold

    $detailsBox = New-Object System.Windows.Forms.TextBox
    $detailsBox.Left = 20; $detailsBox.Top = 54; $detailsBox.Width = 618; $detailsBox.Height = 270
    $detailsBox.Multiline = $true; $detailsBox.ReadOnly = $true; $detailsBox.ScrollBars = "Vertical"
    $detailsBox.BackColor = $script:Colors.SurfaceAlt; $detailsBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $detailsBox.Font = New-AppFont -Size 9.5; $detailsBox.Anchor = "Top,Bottom,Left,Right"

    $chkVisible = New-Object System.Windows.Forms.CheckBox
    $chkVisible.Left = 20; $chkVisible.Top = 338; $chkVisible.Width = 280
    $chkVisible.Text = "Pokazuj SolidWorks przy uruchamianiu"; $chkVisible.Checked = [bool]$config.solidworks.visible; $chkVisible.Anchor = "Left,Bottom"

    $hintLabel = New-Object System.Windows.Forms.Label
    $hintLabel.Left = 20; $hintLabel.Top = 366; $hintLabel.Width = 600; $hintLabel.Height = 52
    $hintLabel.Text = "Tryby: attached_only = tylko na juz otwartym SolidWorks, attached_or_start = podlacz lub uruchom, startup_switch = awaryjnie przez okno uruchamiania makra."
    $hintLabel.ForeColor = $script:Colors.Muted; $hintLabel.Font = New-AppFont -Size 8.5

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Left = 20; $statusLabel.Top = 596; $statusLabel.Width = 1018; $statusLabel.Height = 30
    $statusLabel.Text = "Gotowe."; $statusLabel.ForeColor = $script:Colors.Muted; $statusLabel.Font = New-AppFont -Size 9.5

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Uruchom w SolidWorks"; $btnRun.Left = 20; $btnRun.Top = 448; $btnRun.Width = 206; $btnRun.Height = 42
    $btnRun.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Shield); $btnRun.Anchor = "Left,Bottom"

    $btnOpenFile = New-Object System.Windows.Forms.Button
    $btnOpenFile.Text = "Pokaz plik"; $btnOpenFile.Left = 236; $btnOpenFile.Top = 448; $btnOpenFile.Width = 126; $btnOpenFile.Height = 42
    $btnOpenFile.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Information); $btnOpenFile.Anchor = "Left,Bottom"

    $btnOpenFolder = New-Object System.Windows.Forms.Button
    $btnOpenFolder.Text = "Otworz folder"; $btnOpenFolder.Left = 372; $btnOpenFolder.Top = 448; $btnOpenFolder.Width = 134; $btnOpenFolder.Height = 42
    $btnOpenFolder.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::WinLogo); $btnOpenFolder.Anchor = "Left,Bottom"

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Dodaj"; $btnAdd.Left = 16; $btnAdd.Top = 410; $btnAdd.Width = 94; $btnAdd.Height = 42
    $btnAdd.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Application); $btnAdd.Anchor = "Left,Bottom"

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = "Edytuj"; $btnEdit.Left = 120; $btnEdit.Top = 410; $btnEdit.Width = 94; $btnEdit.Height = 42
    $btnEdit.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Asterisk); $btnEdit.Anchor = "Left,Bottom"

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Usun"; $btnDelete.Left = 224; $btnDelete.Top = 410; $btnDelete.Width = 94; $btnDelete.Height = 42
    $btnDelete.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Error); $btnDelete.Anchor = "Left,Bottom"

    $btnSaveConfig = New-Object System.Windows.Forms.Button
    $btnSaveConfig.Text = "Zapisz konfiguracje"; $btnSaveConfig.Left = 460; $btnSaveConfig.Top = 416; $btnSaveConfig.Width = 178; $btnSaveConfig.Height = 42
    $btnSaveConfig.Image = New-SystemBitmap -Icon ([System.Drawing.SystemIcons]::Question); $btnSaveConfig.Anchor = "Right,Bottom"

    Set-PrimaryButtonStyle -Button $btnRun
    foreach ($button in @($btnOpenFile, $btnOpenFolder, $btnAdd, $btnEdit, $btnDelete, $btnSaveConfig, $btnFilterAll, $btnFilterMine, $btnFilterDownloaded)) {
        Set-ButtonStyle -Button $button
    }

    $script:macroFilter = "all"
    $script:displayedMacros = [System.Collections.ArrayList]::new()

    function Test-MacroMatchesFilter {
        param(
            [hashtable]$Macro,
            [string]$Filter
        )

        $name = ([string]$Macro.name).ToLowerInvariant()
        switch ($Filter) {
            "mine" { return $name.StartsWith("moje:") }
            "downloaded" { return $name.StartsWith("pobrane:") }
            default { return $true }
        }
    }

    function Update-FilterButtons {
        foreach ($btn in @($btnFilterAll, $btnFilterMine, $btnFilterDownloaded)) {
            Set-ButtonStyle -Button $btn
        }

        switch ($script:macroFilter) {
            "mine" { Set-PrimaryButtonStyle -Button $btnFilterMine }
            "downloaded" { Set-PrimaryButtonStyle -Button $btnFilterDownloaded }
            default { Set-PrimaryButtonStyle -Button $btnFilterAll }
        }
    }

    function Get-SelectedMacro {
        if ($listBox.SelectedIndex -lt 0) { return $null }
        return [hashtable]$script:displayedMacros[$listBox.SelectedIndex]
    }

    function Refresh-MacroList {
        $selectedName = $null
        if ($listBox.SelectedIndex -ge 0 -and $listBox.SelectedIndex -lt $script:displayedMacros.Count) {
            $selectedName = [string]([hashtable]$script:displayedMacros[$listBox.SelectedIndex]).name
        }

        $script:displayedMacros = [System.Collections.ArrayList]::new()
        $listBox.Items.Clear()
        foreach ($macro in $config.macros) {
            $macroHash = [hashtable]$macro
            if (Test-MacroMatchesFilter -Macro $macroHash -Filter $script:macroFilter) {
                [void]$script:displayedMacros.Add($macroHash)
                [void]$listBox.Items.Add($macroHash)
            }
        }

        if ($listBox.Items.Count -gt 0) {
            $newIndex = 0
            if (-not [string]::IsNullOrWhiteSpace($selectedName)) {
                for ($i = 0; $i -lt $script:displayedMacros.Count; $i++) {
                    if ([string]([hashtable]$script:displayedMacros[$i]).name -eq $selectedName) {
                        $newIndex = $i
                        break
                    }
                }
            }
            $listBox.SelectedIndex = $newIndex
        }
        else {
            $detailsBox.Text = "Brak makr w konfiguracji."
        }

        Update-FilterButtons
    }

    function Update-Details {
        $macro = Get-SelectedMacro
        if ($null -eq $macro) {
            $detailsBox.Text = "Wybierz makro z listy."
            return
        }

        $resolvedPath = Get-MacroResolvedPath -Macro $macro
        $docRequirementText = "Nie"
        if ([bool]$macro.requiresOpenDoc) {
            $docRequirementText = "Tak"
            if (-not [string]::IsNullOrWhiteSpace([string]$macro.requiredDocType)) {
                $docRequirementText += " ($($macro.requiredDocType))"
            }
        }

        $lines = @(
            "Nazwa: $($macro.name)",
            "Kategoria: $($macro.category)",
            "Tryb uruchamiania: $($macro.launchMode)",
            "Wymaga otwartego dokumentu: $docRequirementText",
            "",
            "Opis makra:",
            $macro.description,
            "",
            "Sciezka:",
            $resolvedPath
        )

        if (-not [string]::IsNullOrWhiteSpace([string]$macro.origin)) {
            $lines += ""
            $lines += "Pochodzenie:"
            $lines += $macro.origin
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$macro.sourceUrl)) {
            $lines += ""
            $lines += "Zrodlo:"
            $lines += [string]$macro.sourceUrl
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$macro.module) -or -not [string]::IsNullOrWhiteSpace([string]$macro.procedure)) {
            $lines += ""
            $lines += "Modul VBA: $($macro.module)"
            $lines += "Procedura startowa: $($macro.procedure)"
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$macro.requirements)) {
            $lines += ""
            $lines += "Wymagania:"
            $lines += $macro.requirements
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$macro.steps)) {
            $lines += ""
            $lines += "Krok po kroku:"
            foreach ($step in @([string]$macro.steps -split "\r?\n")) {
                if (-not [string]::IsNullOrWhiteSpace($step)) { $lines += $step }
            }
        }

        $detailsBox.Text = ($lines -join [Environment]::NewLine)
    }

    $listBox.Add_DrawItem({
        param($sender, $e)

        if ($e.Index -lt 0 -or $e.Index -ge $script:displayedMacros.Count) {
            return
        }

        $macro = [hashtable]$script:displayedMacros[$e.Index]
        $accent = Get-MacroAccentColor -Macro $macro
        $glyph = Get-MacroGlyph -Macro $macro
        $bounds = $e.Bounds

        $e.DrawBackground()
        $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

        $cardRect = New-Object System.Drawing.Rectangle ($bounds.Left + 4), ($bounds.Top + 4), ($bounds.Width - 8), ($bounds.Height - 8)
        $cardColor = if (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected) {
            [System.Drawing.ColorTranslator]::FromHtml("#E8F0FF")
        } else {
            [System.Drawing.Color]::White
        }

        $cardBrush = New-Object System.Drawing.SolidBrush($cardColor)
        $borderPen = New-Object System.Drawing.Pen($script:Colors.Border)
        $e.Graphics.FillRectangle($cardBrush, $cardRect)
        $e.Graphics.DrawRectangle($borderPen, $cardRect)

        $iconRect = New-Object System.Drawing.Rectangle ($cardRect.Left + 12), ($cardRect.Top + 14), 44, 44
        $iconBrush = New-Object System.Drawing.SolidBrush($accent)
        $e.Graphics.FillEllipse($iconBrush, $iconRect)

        $glyphFont = New-AppFont -Size 9 -Style Bold
        $glyphBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $glyphSize = $e.Graphics.MeasureString($glyph, $glyphFont)
        $glyphX = [single]($iconRect.Left + (($iconRect.Width - $glyphSize.Width) / 2))
        $glyphY = [single]($iconRect.Top + (($iconRect.Height - $glyphSize.Height) / 2))
        $e.Graphics.DrawString($glyph, $glyphFont, $glyphBrush, $glyphX, $glyphY)

        $titleFont = New-AppFont -Size 10.5 -Style Bold
        $bodyFont = New-AppFont -Size 8.6
        $metaFont = New-AppFont -Size 8 -Style Bold
        $textBrush = New-Object System.Drawing.SolidBrush($script:Colors.Text)
        $mutedBrush = New-Object System.Drawing.SolidBrush($script:Colors.Muted)

        $titleX = $iconRect.Right + 12
        $titleY = $cardRect.Top + 10
        $e.Graphics.DrawString([string]$macro.name, $titleFont, $textBrush, $titleX, $titleY)

        $chipText = if ([string]::IsNullOrWhiteSpace([string]$macro.category)) { "Makro" } else { [string]$macro.category }
        $chipSize = $e.Graphics.MeasureString($chipText, $metaFont)
        $chipRect = New-Object System.Drawing.RectangleF ($titleX), ($cardRect.Top + 32), ([Math]::Min(110, $chipSize.Width + 16)), 18
        $chipBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(20, $accent))
        $chipBorderPen = New-Object System.Drawing.Pen($accent)
        $e.Graphics.FillRectangle($chipBrush, $chipRect)
        $e.Graphics.DrawRectangle($chipBorderPen, [System.Drawing.Rectangle]::Round($chipRect))
        $e.Graphics.DrawString($chipText, $metaFont, (New-Object System.Drawing.SolidBrush($accent)), $chipRect.X + 8, $chipRect.Y + 2)

        $description = [string]$macro.description
        if ($description.Length -gt 58) { $description = $description.Substring(0, 55) + "..." }
        $e.Graphics.DrawString($description, $bodyFont, $mutedBrush, $titleX, ($cardRect.Top + 54))

        $cardBrush.Dispose()
        $borderPen.Dispose()
        $iconBrush.Dispose()
        $glyphBrush.Dispose()
        $textBrush.Dispose()
        $mutedBrush.Dispose()
        $chipBrush.Dispose()
        $chipBorderPen.Dispose()
        $glyphFont.Dispose()
        $titleFont.Dispose()
        $bodyFont.Dispose()
        $metaFont.Dispose()

        $e.DrawFocusRectangle()
    })

    $chkVisible.Add_CheckedChanged({ $config.solidworks.visible = $chkVisible.Checked })
    $listBox.Add_SelectedIndexChanged({ Update-Details })
    $btnFilterAll.Add_Click({ $script:macroFilter = "all"; Refresh-MacroList })
    $btnFilterMine.Add_Click({ $script:macroFilter = "mine"; Refresh-MacroList })
    $btnFilterDownloaded.Add_Click({ $script:macroFilter = "downloaded"; Refresh-MacroList })

    $btnRun.Add_Click({
        $macro = Get-SelectedMacro
        if ($null -eq $macro) {
            Show-Message -Text "Najpierw wybierz makro z listy." -Title "Brak wyboru" -Icon Warning
            return
        }

        try {
            $statusLabel.Text = "Uruchamianie makra: $($macro.name)"
            $runInfo = Invoke-SolidWorksMacro -Macro $macro -Config $config
            if ($runInfo.fallbackMode -eq "start-only") {
                $statusLabel.Text = "SolidWorks zostal uruchomiony. Otworz wymagany dokument i kliknij ponownie: $($macro.name)"
                Show-Message -Text "SolidWorks zostal uruchomiony, ale to makro wymaga juz otwartego dokumentu.`n`nOtworz teraz odpowiedni plik w SolidWorks i kliknij 'Uruchom w SolidWorks' jeszcze raz." -Title "Najpierw otworz dokument" -Icon Information
            }
            elseif ($runInfo.fallbackMode -eq "menu-dialog") {
                $statusLabel.Text = "Makro przekazane przez Narzedzia > Makro > Uruchom w otwartym SolidWorks: $($macro.name)"
            }
            elseif ($runInfo.startedNow) {
                $statusLabel.Text = "SolidWorks zostal uruchomiony i makro wystartowalo: $($macro.name)"
            }
            else {
                $statusLabel.Text = "Makro uruchomione w otwartym SolidWorks: $($macro.name)"
            }
        }
        catch {
            $statusLabel.Text = "Blad uruchamiania: $($_.Exception.Message)"
            Show-Message -Text $_.Exception.Message -Title "Nie udalo sie uruchomic makra" -Icon Error
        }
    })

    $btnOpenFile.Add_Click({
        $macro = Get-SelectedMacro
        if ($null -eq $macro) { return }
        $resolvedPath = Get-MacroResolvedPath -Macro $macro

        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            Show-Message -Text "Plik nie istnieje:`n$resolvedPath" -Title "Brak pliku" -Icon Warning
            return
        }

        Start-Process explorer.exe -ArgumentList "/select,`"$resolvedPath`""
    })

    $btnOpenFolder.Add_Click({
        $macro = Get-SelectedMacro
        if ($null -eq $macro) { return }
        $resolvedPath = Get-MacroResolvedPath -Macro $macro

        if (-not (Test-Path -LiteralPath $resolvedPath)) {
            Show-Message -Text "Plik nie istnieje:`n$resolvedPath" -Title "Brak pliku" -Icon Warning
            return
        }

        $folderPath = Split-Path -Parent $resolvedPath
        Start-Process explorer.exe -ArgumentList "`"$folderPath`""
    })

    $btnAdd.Add_Click({
        $newMacro = Show-MacroEditor
        if ($null -eq $newMacro) { return }
        [void]$config.macros.Add($newMacro)
        Refresh-MacroList
        $listBox.SelectedIndex = $config.macros.Count - 1
        Update-Details
    })

    $btnEdit.Add_Click({
        $selectedIndex = $listBox.SelectedIndex
        $macro = Get-SelectedMacro
        if ($null -eq $macro) {
            Show-Message -Text "Najpierw wybierz makro z listy." -Title "Brak wyboru" -Icon Warning
            return
        }

        $updatedMacro = Show-MacroEditor -InitialMacro $macro
        if ($null -eq $updatedMacro) { return }

        $config.macros[$selectedIndex] = $updatedMacro
        Refresh-MacroList
        $listBox.SelectedIndex = $selectedIndex
        Update-Details
    })

    $btnDelete.Add_Click({
        $selectedIndex = $listBox.SelectedIndex
        $macro = Get-SelectedMacro
        if ($null -eq $macro) {
            Show-Message -Text "Najpierw wybierz makro z listy." -Title "Brak wyboru" -Icon Warning
            return
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Usunac makro '$($macro.name)' z konfiguracji?",
            "Potwierdzenie",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            $config.macros.RemoveAt($selectedIndex)
            Refresh-MacroList
            Update-Details
        }
    })

    $btnSaveConfig.Add_Click({
        try {
            Save-Config -Config $config
            $statusLabel.Text = "Konfiguracja zapisana."
        }
        catch {
            Show-Message -Text $_.Exception.Message -Title "Nie udalo sie zapisac konfiguracji" -Icon Error
        }
    })

    foreach ($control in @($headerIcon, $headerTitle, $headerSubtitle)) { $headerPanel.Controls.Add($control) }
    foreach ($control in @($leftTitle, $listBox, $btnAdd, $btnEdit, $btnDelete)) { $leftPanel.Controls.Add($control) }
    foreach ($control in @($detailsTitle, $detailsBox, $chkVisible, $hintLabel, $btnRun, $btnOpenFile, $btnOpenFolder, $btnSaveConfig)) { $rightPanel.Controls.Add($control) }
    foreach ($control in @($headerPanel, $leftPanel, $rightPanel, $statusLabel)) { $form.Controls.Add($control) }

    Refresh-MacroList
    Update-Details

    [void]$form.ShowDialog()
}
catch {
    Show-Message -Text $_.Exception.Message -Title "Blad startu aplikacji" -Icon Error
    throw
}
