#Requires -Version 5.1
# Device Repair Tool PRO v4

# ============================================================
# 1. Verification admin + relance auto
# ============================================================
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $PSCommandPath) {
        # WPF pas encore charge ici — charger juste MessageBox
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show(
            "Lancez ce script depuis un fichier .ps1, pas depuis la console.",
            "Erreur de lancement", "OK", "Error"
        ) | Out-Null
        exit 1
    }
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# ============================================================
# 2. Masquage de la console
# ============================================================
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
$hwnd = [Win32]::GetConsoleWindow()
if ($hwnd -ne [IntPtr]::Zero) { [Win32]::ShowWindow($hwnd, 0) | Out-Null }

# ============================================================
# 3. Chargement des assemblies WPF
# ============================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# ============================================================
# 4. Variables globales
# ============================================================
$Script:APP_VERSION        = "4.1"
$Script:APP_NAME           = "Device Repair Tool PRO"
$Script:DESKTOP_DIR        = [Environment]::GetFolderPath("Desktop")
$Script:Logs               = [System.Collections.Generic.List[object]]::new()
$Script:ProblemDevices     = [System.Collections.Generic.List[object]]::new()
$Script:DriverStoreDrivers = [System.Collections.Generic.List[object]]::new()
$Script:WindowsUpdates     = [System.Collections.Generic.List[object]]::new()
$Script:EventEntries       = [System.Collections.Generic.List[object]]::new()
$Script:ScanDone           = $false

# ============================================================
# 5. Detection systeme
# ============================================================
try {
    $Script:OSInfo      = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $Script:OSVersion   = [version]$Script:OSInfo.Version
    $Script:PSVersion   = $PSVersionTable.PSVersion
    $Script:CPU         = (Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1).Name
    # TotalVisibleMemorySize est en KB -> diviser par 1024*1024 pour GB
    $Script:TotalRAM_GB = [math]::Round($Script:OSInfo.TotalVisibleMemorySize / 1024 / 1024, 2)
}
catch {
    Write-Error "Impossible de recuperer les informations systeme : $_"
    exit 1
}

# ============================================================
# 6. Fonctions utilitaires
# ============================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR")]
        [string]$Level = "INFO"
    )
    $logEntry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level     = $Level
        Message   = $Message
    }
    $Script:Logs.Add($logEntry)
}

function Get-FormattedLogs {
    return ($Script:Logs | ForEach-Object {
        "[{0}] [{1,-7}] {2}" -f $_.Timestamp, $_.Level, $_.Message
    }) -join "`r`n"
}

function Clear-Logs {
    $Script:Logs.Clear()
    Write-Log "Logs effaces" "SUCCESS"
}

function Set-Status {
    param($Label, [string]$Text)
    if ($Label) { $Label.Text = $Text }
}

function Hide-ConsoleWindow {
    $hwnd = [Win32]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { 
        [Win32]::ShowWindow($hwnd, 0) | Out-Null
    }
}

# Helper HTML escape (used by Export-Report)
function ConvertTo-HtmlEscaped {
    param([string]$s)
    if (-not $s) { return '' }
    $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

# Dispatcher yield non-bloquant (remplace DoEvents WinForms)
function Invoke-DispatcherYield {
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [System.Action]{},
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

function Get-SystemInfo {
    $uptime    = (Get-Date) - $Script:OSInfo.LastBootUpTime
    $uptimeStr = "{0}j {1}h {2}m" -f [math]::Floor($uptime.TotalDays), $uptime.Hours, $uptime.Minutes
    return [PSCustomObject]@{
        OS           = $Script:OSInfo.Caption
        Architecture = $env:PROCESSOR_ARCHITECTURE
        CPU          = $Script:CPU
        RAM_GB       = $Script:TotalRAM_GB
        PSVersion    = $Script:PSVersion
        Uptime       = $uptimeStr
        LastBoot     = $Script:OSInfo.LastBootUpTime.ToString("yyyy-MM-dd HH:mm:ss")
    }
}

# ============================================================
# 7. Fonction Scan peripheriques — ASYNC Runspace
# ============================================================
function Invoke-SystemScan {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [System.Windows.Controls.ProgressBar]$ProgressBar,
        [System.Windows.Controls.Button]$BtnStart,
        [System.Windows.Controls.Button]$BtnRepairSel,
        [System.Windows.Controls.Button]$BtnRepairAll,
        [System.Windows.Controls.TextBox]$TxtLog
    )

    try {
        Set-Status $StatusLabel "Scan en cours..."
        if ($ProgressBar) { $ProgressBar.Value = 0 }

        Write-Log "Debut du scan des peripheriques (async Runspace - no UI freeze)" "INFO"

        if (-not $Script:ProblemDevices) {
            $Script:ProblemDevices = New-Object System.Collections.ArrayList
        } else {
            $Script:ProblemDevices.Clear()
        }

        # Capture reference as local var — safe inside .NET event handler closure
        # ($Script: scope is unreliable inside DispatcherTimer.Tick event handlers)
        $devicesList = $Script:ProblemDevices

        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        $ps.AddScript({
            param($q)
            try {
                $allDrivers   = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue
                $devices      = @(Get-PnpDevice -ErrorAction SilentlyContinue)
                $totalDevices = $devices.Count
                $errorCount   = 0
                $idx          = 0

                # Hashtable de lookup pour eviter Where-Object repete sur chaque device
                $driverMap = @{}
                foreach ($d in $allDrivers) { if ($d.DeviceID) { $driverMap[$d.DeviceID] = $d } }

                $q.Enqueue("[INIT] $totalDevices")

                foreach ($device in $devices) {
                    $idx++

                    $driverInfo    = $driverMap[$device.InstanceId]
                    $driverDateStr = if ($driverInfo -and $driverInfo.DriverDate) { $driverInfo.DriverDate.ToString("yyyy-MM-dd") } else { "-" }
                    $errorCode     = if ($device.PSObject.Properties["ConfigManagerErrorCode"]) { $device.ConfigManagerErrorCode } else { "-" }

                    $deviceInfo = [PSCustomObject]@{
                        Selected      = $false
                        Name          = if ($device.FriendlyName) { $device.FriendlyName } else { "Peripherique inconnu" }
                        Class         = if ($device.Class) { $device.Class } else { "Inconnu" }
                        StatusCode    = $device.Status
                        ErrorCode     = $errorCode
                        InstanceId    = $device.InstanceId
                        DriverVersion = if ($driverInfo) { $driverInfo.DriverVersion } else { "-" }
                        DriverDate    = $driverDateStr
                        Provider      = if ($driverInfo) { $driverInfo.DriverProviderName } else { "-" }
                    }

                    $q.Enqueue("[DEVICE] $($deviceInfo | ConvertTo-Json -Compress)")
                    if ($device.Status -in @("Error", "Degraded")) { $errorCount++ }

                    if ($idx % 10 -eq 0 -and $totalDevices -gt 0) {
                        $q.Enqueue("[PROGRESS] $([math]::Round(($idx / $totalDevices) * 100))")
                    }
                }

                $q.Enqueue("[DONE] $errorCount|$totalDevices")
            }
            catch {
                $q.Enqueue("[ERR] $_")
                $q.Enqueue("[DONE] -1|0")
            }
        }).AddArgument($queue) | Out-Null

        $ps.BeginInvoke() | Out-Null

        $timer   = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(200)

        # Array-based counter: reference type, reliably mutable inside closures
        $spinCount     = @(0)
        $totalExpected = @(0)

        $timer.Add_Tick(({
            $ln = $null
            while ($queue.TryDequeue([ref]$ln)) {
                if ($ln -match '^\[INIT\] (.*)$') {
                    $totalExpected[0] = [int]$matches[1]
                }
                elseif ($ln -match '^\[DEVICE\] (.*)$') {
                    try {
                        $di = $matches[1] | ConvertFrom-Json
                        [void]$devicesList.Add($di)   # local captured ref, not $Script:
                    }
                    catch { }
                }
                elseif ($ln -match '^\[PROGRESS\] (.*)$') {
                    if ($ProgressBar) { $ProgressBar.Value = [int]$matches[1] }
                }
                elseif ($ln -match '^\[DONE\] (.*)$') {
                    $parts        = $matches[1] -split '\|'
                    $problemCount = [int]$parts[0]
                    $total        = if ($parts.Count -gt 1) { [int]$parts[1] } else { $totalExpected[0] }

                    $timer.Stop()
                    try { $ps.Dispose() } catch { }
                    try { $rs.Close()   } catch { }

                    if ($problemCount -ge 0) {
                        $unkCount = @($devicesList | Where-Object { $_.StatusCode -eq 'Unknown' }).Count
                        $okCount  = @($devicesList | Where-Object { $_.StatusCode -eq 'OK' }).Count
                        Write-Log "Scan termine: $total total | $problemCount erreur(s) | $unkCount inconnu(s) | $okCount OK" "SUCCESS"
                        Set-Status $StatusLabel "$total peripheriques : $problemCount erreur(s)  $unkCount inconnu(s)  $okCount OK"
                        if ($ProgressBar) { $ProgressBar.Value = 100 }
                        if ($DataGrid) {
                            $DataGrid.ItemsSource = $null
                            $DataGrid.ItemsSource = $devicesList
                        }
                        $Script:ScanDone = $true

                        $hasErrors = ($problemCount -gt 0)
                        if ($BtnRepairSel) { $BtnRepairSel.IsEnabled = $hasErrors }
                        if ($BtnRepairAll) { $BtnRepairAll.IsEnabled = $hasErrors }
                    } else {
                        Write-Log "Erreur lors du scan (async)" "ERROR"
                        Set-Status $StatusLabel "Erreur lors du scan"
                    }

                    if ($BtnStart) { $BtnStart.IsEnabled = $true }
                    if ($TxtLog)   { $TxtLog.Text = Get-FormattedLogs; $TxtLog.ScrollToEnd() }
                }
                elseif ($ln -match '^\[ERR\] (.*)$') {
                    Write-Log "Erreur scan async: $($matches[1])" "ERROR"
                }
            }

            if ($timer.IsEnabled) {
                $spinCount[0]++
                $dots = '.' * (($spinCount[0] % 4) + 1)
                Set-Status $StatusLabel "Scan en cours$dots"
            }
        }).GetNewClosure())

        $timer.Start()
    }
    catch {
        Write-Log "Erreur lors du scan: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
        if ($BtnStart) { $BtnStart.IsEnabled = $true }
    }
}

# ============================================================
# 8. Fonction Scan Driver Store
# ============================================================
function Invoke-DriverStoreScan {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    try {
        Set-Status $StatusLabel "Analyse du Driver Store..."
        Write-Log "Analyse du Driver Store en cours..." "INFO"
        $Script:DriverStoreDrivers.Clear()

        $output = pnputil /enum-drivers 2>&1 | Out-String
        if ([string]::IsNullOrWhiteSpace($output)) {
            throw "Aucune sortie de pnputil. Verifiez les droits administrateur."
        }

        $lines = $output -split "[\r\n]+"
        $currentDriver = $null

        # Parsing locale-independant : on detecte les blocs par la valeur oemXX.inf
        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -notmatch ":\s*.+") { continue }

            $colonIdx = $line.IndexOf(':')
            if ($colonIdx -lt 0) { continue }

            $label = $line.Substring(0, $colonIdx).Trim().ToLower()
            $value = $line.Substring($colonIdx + 1).Trim()

            # Detection debut de bloc : valeur de type oem0.inf ... oem999.inf
            if ($value -match "^oem\d+\.inf$") {
                if ($currentDriver) { $Script:DriverStoreDrivers.Add($currentDriver) }
                $currentDriver = [PSCustomObject]@{
                    Selected      = $false
                    PublishedName = $value
                    OriginalName  = "-"
                    Provider      = "-"
                    Version       = "-"
                    Date          = "-"
                    Class         = "-"
                }
                continue
            }

            if (-not $currentDriver) { continue }

            # Correspondance par contenu du label (compatible EN/FR et autres locales)
            if ($label -match "fournisseur|provider") {
                $currentDriver.Provider = $value
            }
            # FR: "version du pilote" contient date ET version sur la meme ligne
            elseif ($label -match "version") {
                $parts = $value -split '\s+', 2
                if ($parts.Count -eq 2) {
                    $currentDriver.Date    = $parts[0]
                    $currentDriver.Version = $parts[1]
                } else {
                    $currentDriver.Version = $value
                }
            }
            elseif ($label -match "classe|class") {
                $currentDriver.Class = $value
            }
            elseif ($label -match "origine|original") {
                $currentDriver.OriginalName = $value
            }
        }
        if ($currentDriver) { $Script:DriverStoreDrivers.Add($currentDriver) }

        Write-Log "$($Script:DriverStoreDrivers.Count) pilotes trouves dans le Driver Store" "SUCCESS"
        if ($DataGrid) {
            $DataGrid.ItemsSource = $null
            $DataGrid.ItemsSource = $Script:DriverStoreDrivers
        }
        Set-Status $StatusLabel "$($Script:DriverStoreDrivers.Count) pilotes"
    }
    catch {
        Write-Log "Erreur lors de l'analyse du Driver Store: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
    }
}

# ============================================================
# 8b. Fonction Analyse Evenements Pilotes — ciblee + intelligente
# ============================================================
function Invoke-EventScan {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$LblCritical,
        [System.Windows.Controls.TextBlock]$LblError,
        [System.Windows.Controls.TextBlock]$LblWarning,
        [System.Windows.Controls.TextBlock]$LblTotal,
        [int]$DaysBack = 7
    )
    try {
        Set-Status $StatusLabel "Analyse des evenements ($DaysBack j)..."
        Write-Log "Analyse des evenements pilotes (derniers $DaysBack jours)..." "INFO"
        $Script:EventEntries.Clear()

        $startTime = (Get-Date).AddDays(-$DaysBack)

        # ── Base de connaissances EventID ──────────────────────────────────
        $kb = @{
            9    = @{ Desc = "Timeout peripherique"            ; Advice = "Verifier branchement physique ou remplacer le peripherique" }
            11   = @{ Desc = "Erreur controleur disque"        ; Advice = "Tester S.M.A.R.T., verifier cables SATA/NVMe" }
            15   = @{ Desc = "Peripherique non pret"           ; Advice = "Reset USB ou reinstaller le pilote" }
            20   = @{ Desc = "Installation pilote echouee"     ; Advice = "Verifier compatibilite OS/architecture du pilote" }
            43   = @{ Desc = "Probleme USB"                    ; Advice = "Reset USB ou desactiver suspension selective USB" }
            51   = @{ Desc = "Erreur pagination disque"        ; Advice = "Secteurs defectueux possibles - tester avec chkdsk /r" }
            129  = @{ Desc = "Reinitialisation controleur"     ; Advice = "Mettre a jour firmware/pilote controleur stockage" }
            153  = @{ Desc = "Delai I/O disque"                ; Advice = "Disque lent ou mourant - cloner et remplacer" }
            157  = @{ Desc = "Disque ejection anormale"        ; Advice = "Verifier alimentation/cables du disque" }
            219  = @{ Desc = "Pilote non charge"               ; Advice = "Reinstaller ou MAJ le pilote - verifier la signature" }
            411  = @{ Desc = "Echec installation peripherique" ; Advice = "Verifier Driver Store - relancer scan peripheriques" }
            1001 = @{ Desc = "Rapport erreur Windows (WER)"    ; Advice = "Analyser le dump dans %localappdata%\CrashDumps" }
            6008 = @{ Desc = "Arret systeme inattendu"         ; Advice = "Verifier temperatures, RAM (memtest), alimentation" }
            7026 = @{ Desc = "Pilote demarrage en echec"       ; Advice = "Verifier services en echec - reconfigurer via msconfig" }
            7031 = @{ Desc = "Service redemarre (auto)"        ; Advice = "Instabilite service - analyser journal Application" }
            7034 = @{ Desc = "Service arrete (inattendu)"      ; Advice = "Verifier crashs, mettre a jour le composant" }
            7045 = @{ Desc = "Nouveau service installe"        ; Advice = "Verifier l'origine du service (malware possible)" }
        }

        # ── Sources ciblees pilotes/peripheriques ──────────────────────────
        $targetPattern = 'kernel-pnp|disk|volmgr|storage|partmgr|nvlddmkm|nvhda|nvcv|amdkmpfd|atikmdag|' +
                         'iaStorAV|ahci|usbhub|USBXHCI|usbccgp|usbstor|Netwtw|NetAdapterCx|bcmwl|iastora|' +
                         'mraid|nvraid|viasraid|ql2300|elxcli|hpsa|HidUsb|HidBth|BTHUSB|i8042prt|kbdclass|' +
                         'mouclass|HDAudBus|intelppm|pciide|atapi|cdrom|ftdisk|mountmgr|WUDFRd|WpdBusEnum|' +
                         'wacom|elan|synaptics|dptf|intelide|storahci|stornvme|spaceport|Classpnp'

        $forceIds  = @(9, 11, 15, 20, 43, 51, 129, 153, 157, 219, 411, 1001, 6008, 7026, 7031, 7034, 7045)
        $allEvents = [System.Collections.Generic.List[object]]::new()

        # Journal System
        try {
            $sysEvts = Get-WinEvent -FilterHashtable @{
                LogName   = 'System'
                Level     = @(1, 2, 3)
                StartTime = $startTime
            } -ErrorAction SilentlyContinue

            foreach ($e in $sysEvts) {
                if ($e.ProviderName -match $targetPattern -or $e.Id -in $forceIds) {
                    $allEvents.Add($e)
                }
            }
            Write-Log "  System log: $($allEvents.Count) evenements pertinents" "INFO"
        }
        catch { Write-Log "  Impossible de lire le journal System: $_" "WARNING" }

        # Kernel-PnP (channel dedie peripheriques)
        try {
            $pnpEvts = Get-WinEvent -FilterHashtable @{
                LogName   = 'Microsoft-Windows-Kernel-PnP/Configuration'
                Level     = @(1, 2, 3)
                StartTime = $startTime
            } -ErrorAction SilentlyContinue
            foreach ($e in $pnpEvts) { $allEvents.Add($e) }
            Write-Log "  Kernel-PnP/Configuration: $($pnpEvts.Count) evenements" "INFO"
        }
        catch { Write-Log "  Channel Kernel-PnP non accessible (normal sur certains OS)" "INFO" }

        # Application : WER seulement (crashs lies aux pilotes)
        try {
            $appEvts = Get-WinEvent -FilterHashtable @{
                LogName   = 'Application'
                Level     = @(1, 2)
                StartTime = $startTime
                Id        = @(1000, 1001, 1002)
            } -ErrorAction SilentlyContinue
            foreach ($e in $appEvts) {
                if ($e.ProviderName -match 'wer|fault|crash|windows error') { $allEvents.Add($e) }
            }
        }
        catch { }

        # Deduplication par (Id + TimeCreated + ProviderName) + tri + limite 500
        $seen   = [System.Collections.Generic.HashSet[string]]::new()
        $sorted = $allEvents |
            Sort-Object TimeCreated -Descending |
            Where-Object {
                $key = "$($_.Id)|$($_.TimeCreated.Ticks)|$($_.ProviderName)"
                $seen.Add($key)
            } |
            Select-Object -First 500

        $nCrit = 0; $nErr = 0; $nWarn = 0

        foreach ($evt in $sorted) {
            $levelStr = switch ($evt.Level) {
                1       { $nCrit++; "Critique"       }
                2       { $nErr++;  "Erreur"          }
                3       { $nWarn++; "Avertissement"   }
                default { "Info" }
            }

            # Message : 1ere ligne non vide, tronque a 130 chars
            $msgFull  = try { $evt.Message } catch { "" }
            $msgShort = if ($msgFull) {
                $first = ($msgFull -split "[\r\n]+" | Where-Object { $_.Trim() } | Select-Object -First 1).Trim()
                if ($first.Length -gt 130) { $first.Substring(0, 127) + "..." } else { $first }
            } else { "-" }

            # Correlation avec les peripheriques en erreur du scan onglet 1
            $deviceHint = "-"
            if ($Script:ProblemDevices -and $Script:ProblemDevices.Count -gt 0) {
                foreach ($pd in $Script:ProblemDevices) {
                    if ($pd.InstanceId) {
                        $frag = ($pd.InstanceId -split '\\')[-1]
                        if ($frag -and $msgFull -match [regex]::Escape($frag)) {
                            $deviceHint = $pd.Name; break
                        }
                    }
                    if ($pd.Name -and $pd.Name -ne "Peripherique inconnu" -and $pd.Name.Length -gt 4) {
                        if ($msgFull -match [regex]::Escape($pd.Name)) {
                            $deviceHint = $pd.Name; break
                        }
                    }
                }
            }

            $kbEntry = if ($kb.ContainsKey([int]$evt.Id)) { $kb[[int]$evt.Id] } else { $null }
            $desc    = if ($kbEntry) { $kbEntry.Desc   } else { $msgShort }
            $advice  = if ($kbEntry) { $kbEntry.Advice } else { "-" }

            $Script:EventEntries.Add([PSCustomObject]@{
                Level        = $levelStr
                TimeCreated  = $evt.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                EventId      = $evt.Id
                ProviderName = $evt.ProviderName
                DeviceHint   = $deviceHint
                Message      = $desc
                FullMessage  = $msgFull
                Advice       = $advice
            })
        }

        if ($LblCritical) { $LblCritical.Text = "Critiques : $nCrit" }
        if ($LblError)    { $LblError.Text    = "Erreurs : $nErr"    }
        if ($LblWarning)  { $LblWarning.Text  = "Avertissements : $nWarn" }
        if ($LblTotal)    { $LblTotal.Text    = "Total : $($Script:EventEntries.Count)" }

        if ($DataGrid) {
            $DataGrid.ItemsSource = $null
            $DataGrid.ItemsSource = $Script:EventEntries
        }

        $corr    = @($Script:EventEntries | Where-Object { $_.DeviceHint -ne "-" }).Count
        $corrMsg = if ($corr -gt 0) { " | $corr correle(s) avec scan" } else {
            if (-not $Script:ScanDone) { " | Lancez le scan onglet 1 pour la correlation" } else { "" }
        }

        Write-Log "Evenements: $nCrit critiques, $nErr erreurs, $nWarn avertissements$corrMsg" "SUCCESS"
        Set-Status $StatusLabel "$($Script:EventEntries.Count) evenements ($nCrit crit. / $nErr err. / $nWarn avert.)$corrMsg"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
    }
    catch {
        Write-Log "Erreur analyse evenements: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
    }
}

# ============================================================
# 9. Fonction Suppression pilotes Driver Store
# ============================================================
function Invoke-RemoveDrivers {
    param(
        [System.Collections.IList]$Drivers,
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    if ($Drivers.Count -eq 0) {
        Write-Log "Aucun pilote selectionne" "WARNING"
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Supprimer $($Drivers.Count) pilote(s) du Driver Store ?`nCette operation est irreversible !",
        "Confirmation", "YesNo", "Warning"
    )
    if ($confirm -ne "Yes") {
        Write-Log "Suppression annulee par l'utilisateur" "INFO"
        return
    }

    $success = 0
    $failed  = 0

    foreach ($driver in $Drivers) {
        # Validation du nom avant suppression
        if ($driver.PublishedName -notmatch "^oem\d+\.inf$") {
            Write-Log "Nom de pilote invalide ignore: '$($driver.PublishedName)'" "WARNING"
            $failed++
            continue
        }

        Write-Log "Suppression du pilote: $($driver.PublishedName)" "INFO"
        try {
            $result = pnputil /delete-driver $driver.PublishedName /force 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Log "  -> Supprime avec succes" "SUCCESS"
                $success++
            }
            else {
                Write-Log "  -> Echec (code $LASTEXITCODE): $result" "WARNING"
                $failed++
            }
        }
        catch {
            Write-Log "  -> Erreur: $($_.Exception.Message)" "ERROR"
            $failed++
        }
        if ($LogBox) {
            $LogBox.Text = Get-FormattedLogs
            $LogBox.ScrollToEnd()
            Invoke-DispatcherYield
        }
    }

    Write-Log "Suppression terminee: $success succes, $failed echecs" "INFO"
    Set-Status $StatusLabel "Suppression terminee: $success/$($Drivers.Count) pilotes"
}

# ============================================================
# 10. Fonction Reparation peripheriques — asynchrone WPF
# ============================================================
function Invoke-RepairDevices {
    param(
        [System.Collections.IList]$Devices,
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )

    if ($Devices.Count -eq 0) {
        Write-Log "Aucun peripherique selectionne" "WARNING"
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Voulez-vous vraiment reparer les $($Devices.Count) peripherique(s) selectionne(s) ?`n(Cela peut reinitialiser ou reinstaller les drivers, certains peripheriques peuvent devenir temporairement inutilisables)",
        "Confirmation reparation",
        "YesNo",
        "Warning"
    )
    if ($confirm -ne "Yes") { return }

    $success = 0
    $total   = $Devices.Count
    $index   = 0

    $timer          = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(500)

    $timer.Add_Tick({
        if ($index -ge $total) {
            $timer.Stop()
            Write-Log "Reparation terminee : $success/$total reussies" "INFO"
            Set-Status $StatusLabel "Reparation terminee : $success/$total"
            if ($LogBox) {
                $LogBox.Text = Get-FormattedLogs
                $LogBox.ScrollToEnd()
            }
            return
        }

        $device = $Devices[$index]
        $index++

        Write-Log "=== Debut reparation de $($device.Name) (InstanceID: $($device.InstanceId)) ===" "INFO"
        Set-Status $StatusLabel "Reparation : $($device.Name)"

        try {
            # 1. Reinitialisation du peripherique
            Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Milliseconds 500
            Enable-PnpDevice  -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Log "  -> Peripherique reinitialise" "SUCCESS"

            # 2. Reinstallation du driver si InfFile disponible
            if ($device.InfFile -and $device.InfFile -ne "-") {
                Write-Log "  -> Reinstallation du driver $($device.InfFile)..." "INFO"
                Start-Process -FilePath "pnputil.exe" -ArgumentList "/delete-driver $($device.InfFile) /force" -Wait -NoNewWindow
                Start-Process -FilePath "pnputil.exe" -ArgumentList "/add-driver $($device.InfFile) /install" -Wait -NoNewWindow
                Write-Log "  -> Driver reinstalle" "SUCCESS"
            }

            $success++
        }
        catch {
            Write-Log "  -> echec : $($_.Exception.Message)" "ERROR"
        }

        if ($LogBox) {
            $LogBox.Text = Get-FormattedLogs
            $LogBox.ScrollToEnd()
            Invoke-DispatcherYield
        }

        Write-Log "=== Fin reparation de $($device.Name) ===" "INFO"
    }.GetNewClosure())

    $timer.Start()
}

# ============================================================
# 9. Fonction Windows Update Search — ASYNC MTA Runspace
# ============================================================
function Invoke-WindowsUpdateSearch {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [System.Windows.Controls.TextBox]$LogBox
    )

    try {
        Set-Status $StatusLabel "Recherche des mises a jour..."
        Write-Log "Recherche Windows Update (async MTA - non-blocking)..." "INFO"

        if (-not $Script:WindowsUpdates) {
            $Script:WindowsUpdates = New-Object System.Collections.ArrayList
        } else {
            $Script:WindowsUpdates.Clear()
        }

        # Capture reference as local var — safe in .NET event handler closure
        $updatesList = $Script:WindowsUpdates

        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

        # MTA required for COM objects (Windows Update Agent)
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::MTA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        $ps.AddScript({
            param($q)
            try {
                $q.Enqueue("[INIT]")

                $session  = New-Object -ComObject "Microsoft.Update.Session"
                $searcher = $session.CreateUpdateSearcher()
                $result   = $searcher.Search("IsInstalled=0")

                $count = 0
                foreach ($update in $result.Updates) {
                    if ($update.Type -ne 2) { continue }  # pilotes uniquement

                    $kbId = if ($update.KBArticleIDs.Count -gt 0) { $update.KBArticleIDs[0] } else { "-" }
                    $updateInfo = [PSCustomObject]@{
                        Title       = $update.Title
                        Description = $update.Description
                        Size_MB     = [math]::Round($update.MaxDownloadSize / 1MB, 1)
                        KBArticle   = $kbId
                    }

                    $q.Enqueue("[UPDATE] $($updateInfo | ConvertTo-Json -Compress)")
                    $count++
                }

                $q.Enqueue("[DONE] $count")
            }
            catch {
                $q.Enqueue("[ERR] $_")
                $q.Enqueue("[DONE] -1")
            }
        }).AddArgument($queue) | Out-Null

        $ps.BeginInvoke() | Out-Null

        $timer          = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(200)
        $spinCount      = @(0)

        $timer.Add_Tick(({
            $ln = $null
            while ($queue.TryDequeue([ref]$ln)) {
                if ($ln -match '^\[UPDATE\] (.*)$') {
                    try {
                        $ui = $matches[1] | ConvertFrom-Json
                        [void]$updatesList.Add($ui)   # local captured ref
                    }
                    catch { }
                }
                elseif ($ln -match '^\[DONE\] (.*)$') {
                    $final = [int]$matches[1]

                    $timer.Stop()
                    try { $ps.Dispose() } catch { }
                    try { $rs.Close()   } catch { }

                    if ($final -ge 0) {
                        Write-Log "Recherche Windows Update terminee: $final mise(s) a jour disponible(s)" "SUCCESS"
                        Set-Status $StatusLabel "$final mise(s) a jour disponible(s)"
                    } else {
                        Write-Log "Erreur lors de la recherche Windows Update" "ERROR"
                        Set-Status $StatusLabel "Erreur lors de la recherche"
                    }

                    if ($DataGrid) {
                        $DataGrid.ItemsSource = $null
                        $DataGrid.ItemsSource = $updatesList   # local ref, never null
                    }

                    if ($LogBox) {
                        $LogBox.Text = Get-FormattedLogs
                        $LogBox.ScrollToEnd()
                    }
                }
                elseif ($ln -match '^\[ERR\] (.*)$') {
                    Write-Log "Erreur WU async: $($matches[1])" "ERROR"
                }
            }

            if ($timer.IsEnabled) {
                $spinCount[0]++
                $dots = '.' * (($spinCount[0] % 4) + 1)
                Set-Status $StatusLabel "Recherche en cours$dots"
            }
        }).GetNewClosure())

        $timer.Start()
    }
    catch {
        Write-Log "Erreur Windows Update: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
    }
}

# ============================================================
# 12. Start-BackgroundProcess pour SFC/DISM
#     Pattern : Runspace + ConcurrentQueue + DispatcherTimer
#     Evite les problemes de closure des events .NET en PowerShell
# ============================================================
function Start-BackgroundProcess {
    param(
        [string]$Exe,
        [string]$ExeArgs,
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [string]$Label
    )

    try {
        # File thread-safe partagee entre le Runspace et le timer UI
        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        $ps.AddScript({
            param($exe, $exeArgs, $q)
            try {
                $isSFC = ($exe -match 'sfc' -or $exeArgs -match 'scannow')
                $proc  = $null

                if ($isSFC) {
                    # SFC necessite un environnement console. On lance via un
                    # processus PowerShell cache pour recuperer le code retour reel.
                    $psi = [System.Diagnostics.ProcessStartInfo]::new()
                    $psi.FileName  = "powershell.exe"
                    $psi.Arguments = "-NoProfile -WindowStyle Hidden -Command `"sfc.exe /scannow; exit `$LASTEXITCODE`""
                    $psi.UseShellExecute = $false
                    $psi.CreateNoWindow  = $true

                    $proc = [System.Diagnostics.Process]::Start($psi)
                    $q.Enqueue("[INFO] Lancement de SFC dans un processus PowerShell cache (PID: $($proc.Id))...")
                    $proc.WaitForExit()

                    # Lecture du log CBS.log apres la fin de SFC
                    $cbsLog = "$env:windir\Logs\CBS\CBS.log"
                    if (Test-Path $cbsLog) {
                        $lines    = Get-Content $cbsLog -Encoding Unicode -ErrorAction SilentlyContinue
                        $sfcLines = $lines | Where-Object { $_ -match '\[SR\]' } | Select-Object -Last 100

                        if ($sfcLines) {
                            foreach ($l in $sfcLines) {
                                if ($l.Trim()) { $q.Enqueue("[OUT] $l") }
                            }
                        } else {
                            $q.Enqueue("[OUT] Aucune ligne [SR] trouvee dans CBS.log. Le log peut etre vide ou SFC n'a rien rapporte.")
                        }
                    } else {
                        $q.Enqueue("[OUT] CBS.log introuvable : $cbsLog")
                    }

                    # Interpretation du code retour SFC
                    switch ($proc.ExitCode) {
                        0 { $q.Enqueue("[SFC_OK] Aucune violation d'integrite detectee.") }
                        1 { $q.Enqueue("[SFC_FIXED] Fichiers corrompus detectes et repares avec succes.") }
                        2 { $q.Enqueue("[SFC_FAIL] Fichiers corrompus detectes mais NON repares. Lancez DISM puis relancez SFC.") }
                        default { $q.Enqueue("[SFC_UNK] Code de sortie inconnu : $($proc.ExitCode)") }
                    }
                } else {
                    # DISM : lecture en temps reel de stdout
                    $psi = [System.Diagnostics.ProcessStartInfo]::new()
                    $psi.FileName               = $exe
                    $psi.Arguments              = $exeArgs
                    $psi.UseShellExecute        = $false
                    $psi.CreateNoWindow         = $true
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError  = $true

                    $proc = [System.Diagnostics.Process]::Start($psi)
                    $q.Enqueue("[INFO] PID $($proc.Id) lance")

                    while (-not $proc.StandardOutput.EndOfStream) {
                        $line = $proc.StandardOutput.ReadLine()
                        if ($line) { $q.Enqueue("[OUT] $line") }
                    }
                    $err = $proc.StandardError.ReadToEnd()
                    if ($err) {
                        foreach ($el in ($err -split "`n")) {
                            if ($el.Trim()) { $q.Enqueue("[ERR] $el") }
                        }
                    }
                    $proc.WaitForExit()
                }

                $q.Enqueue("[DONE] $($proc.ExitCode)")
            }
            catch {
                $q.Enqueue("[ERR] $_")
                $q.Enqueue("[DONE] -1")
            }
        }).AddArgument($Exe).AddArgument($ExeArgs).AddArgument($queue) | Out-Null

        # Lance le Runspace en arriere-plan
        $ps.BeginInvoke() | Out-Null

        Write-Log "$Label lance : $Exe $ExeArgs" "INFO"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
        Set-Status $StatusLabel "$Label en cours..."

        # DispatcherTimer sur le thread UI qui draine la queue
        $timer          = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromMilliseconds(500)

        $Script:_bgSpinCount = 0
        $timer.Add_Tick({
            # Spinner pendant l'execution
            $Script:_bgSpinCount++
            $dots = '.' * (($Script:_bgSpinCount % 4) + 1)
            $currentStatus = $StatusLabel.Text
            if ($currentStatus -notmatch 'termine') {
                Set-Status $StatusLabel "$Label en cours$dots"
            }

            $line = $null
            while ($queue.TryDequeue([ref]$line)) {
                if ($line -match '^\[DONE\] (.*)$') {
                    $code = $matches[1]
                    $timer.Stop()
                    $ps.Dispose()
                    $rs.Close()
                    $lvl = if ($code -eq '0') { "SUCCESS" } else { "WARNING" }
                    Write-Log "$Label termine (code : $code)" $lvl
                    Set-Status $StatusLabel "$Label termine (code $code)"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }

                    $msg = switch ($Label) {
                        "SFC" {
                            switch ($code) {
                                '0' { "SFC termine.`n`nAucune violation d'integrite detectee." }
                                '1' { "SFC termine.`n`nFichiers corrompus detectes et REPARES avec succes." }
                                '2' { "SFC termine.`n`nFichiers corrompus NON repares.`nLancez DISM puis relancez SFC." }
                                default { "SFC termine (code $code)." }
                            }
                        }
                        "DISM" {
                            if ($code -eq '0') { "DISM termine.`n`nImage systeme reparee avec succes." }
                            else               { "DISM termine avec le code $code.`nConsultez le log pour les details." }
                        }
                        default { "$Label termine (code $code)." }
                    }
                    $icon = if ($code -eq '0' -or $code -eq '1') { "Information" } else { "Warning" }
                    [System.Windows.MessageBox]::Show($msg, "$Label - Resultat", "OK", $icon) | Out-Null
                }
                elseif ($line -match '^\[SFC_OK\] (.*)$') {
                    Write-Log $matches[1] "SUCCESS"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line -match '^\[SFC_FIXED\] (.*)$') {
                    Write-Log $matches[1] "SUCCESS"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line -match '^\[SFC_FAIL\] (.*)$') {
                    Write-Log $matches[1] "ERROR"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line -match '^\[SFC_UNK\] (.*)$') {
                    Write-Log $matches[1] "WARNING"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line -match '^\[ERR\] (.*)$') {
                    Write-Log $matches[1] "WARNING"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line -match '^\[OUT\] (.*)$') {
                    Write-Log $matches[1] "INFO"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
                elseif ($line) {
                    Write-Log $line "INFO"
                    if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
                }
            }
        }.GetNewClosure())

        $timer.Start()
    }
    catch {
        $msgErr = "Erreur au lancement de $Label : $($_.Exception.Message)"
        Write-Log $msgErr "ERROR"
        Set-Status $StatusLabel $msgErr
        if ($LogBox) { $LogBox.Text += "$msgErr`r`n"; $LogBox.ScrollToEnd() }
    }
}

# ============================================================
# Invoke-SFCScan / Invoke-DISMRepair
# ============================================================
function Invoke-SFCScan {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    Start-BackgroundProcess -Exe "sfc.exe" -ExeArgs "/scannow" -LogBox $LogBox -StatusLabel $StatusLabel -Label "SFC"
}

function Invoke-DISMRepair {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    Start-BackgroundProcess -Exe "dism.exe" -ExeArgs "/Online /Cleanup-Image /RestoreHealth" -LogBox $LogBox -StatusLabel $StatusLabel -Label "DISM"
}

# ============================================================
# 13. Fonction Reset USB
# ============================================================
function Invoke-ResetUSBDevices {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    try {
        Set-Status $StatusLabel "Reset des peripheriques USB..."
        Write-Log "Reset des peripheriques USB en cours..." "INFO"

        # Filter by InstanceId prefix — far more reliable than -Class USB
        # Exclude HID/Keyboard/Mouse/System to avoid input loss
        $skipClasses = @('HIDClass','Keyboard','Mouse','System','SecurityDevices','Biometric')
        $usbDevices  = @(Get-PnpDevice -ErrorAction SilentlyContinue | Where-Object {
            $_.InstanceId -match '^USB\\' -and $_.Class -notin $skipClasses
        })

        if ($usbDevices.Count -eq 0) {
            Write-Log "Aucun peripherique USB non-HID trouve" "WARNING"
            Set-Status $StatusLabel "Aucun peripherique USB a resetter"
            if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
            return
        }

        Write-Log "$($usbDevices.Count) peripherique(s) USB cible(s) (HID exclus)" "INFO"
        $count = 0

        foreach ($device in $usbDevices) {
            $name = if ($device.FriendlyName) { $device.FriendlyName } else { $device.InstanceId }
            try {
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop | Out-Null
                Start-Sleep -Milliseconds 300
                Enable-PnpDevice  -InstanceId $device.InstanceId -Confirm:$false -ErrorAction Stop | Out-Null
                Write-Log "  -> Reset OK : $name" "SUCCESS"
                $count++
            }
            catch {
                Write-Log "  -> Ignore   : $name ($($_.Exception.Message))" "WARNING"
            }
            if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
            Set-Status $StatusLabel "Reset USB... ($count/$($usbDevices.Count))"
            Invoke-DispatcherYield
        }

        Write-Log "Reset USB termine: $count/$($usbDevices.Count) traite(s)" "SUCCESS"
        Set-Status $StatusLabel "Reset USB termine ($count/$($usbDevices.Count))"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
    }
    catch {
        Write-Log "Erreur lors du reset USB: $($_.Exception.Message)" "ERROR"
        Set-Status $StatusLabel "Erreur reset USB"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
    }
}

# ============================================================
# 14. Fonction Export Rapport — HTML dark theme
# ============================================================
function Export-Report {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    try {
        $now        = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $reportPath = "$Script:DESKTOP_DIR\DeviceRepairTool_Report_$stamp.html"
        $sysInfo    = Get-SystemInfo

        # Pre-compute counters and card styles
        $devCount   = $Script:ProblemDevices.Count
        $drvCount   = $Script:DriverStoreDrivers.Count
        $evtCount   = $Script:EventEntries.Count
        $critEvt    = @($Script:EventEntries | Where-Object { $_.Level -in @('Critique','Erreur') }).Count
        $cardDev    = if ($devCount  -gt 0) {'err'}  else {'ok'}
        $cardEvt    = if ($critEvt   -gt 0) {'warn'} else {'ok'}
        $badgeDev   = if ($devCount  -gt 0) {'err'}  else {''}
        $badgeEvt   = if ($critEvt   -gt 0) {'warn'} else {''}

        $sb = [System.Text.StringBuilder]::new()

        # ── HEAD ──────────────────────────────────────────────
        $null = $sb.AppendLine('<!DOCTYPE html>')
        $null = $sb.AppendLine('<html lang="fr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">')
        $null = $sb.AppendLine("<title>DRT PRO v$($Script:APP_VERSION) &mdash; $now</title>")

        # ── CSS ───────────────────────────────────────────────
        $null = $sb.AppendLine('<style>')
        $null = $sb.AppendLine('*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}')
        $null = $sb.AppendLine('html{scroll-behavior:smooth}')
        $null = $sb.AppendLine('body{background:#121212;color:#D4D4D4;font-family:"Segoe UI",system-ui,sans-serif;font-size:13px;display:flex;flex-direction:column;min-height:100vh}')
        # layout
        $null = $sb.AppendLine('.layout{display:flex;flex:1}')
        # sidebar
        $null = $sb.AppendLine('.sidebar{width:215px;min-width:215px;background:#181818;border-right:1px solid #252526;display:flex;flex-direction:column;position:sticky;top:0;height:100vh;overflow-y:auto;flex-shrink:0}')
        $null = $sb.AppendLine('.sb-logo{padding:20px 16px 16px;border-bottom:1px solid #252526}')
        $null = $sb.AppendLine('.sb-logo .app{color:#4EC9B0;font-size:13px;font-weight:700;line-height:1.45}')
        $null = $sb.AppendLine('.sb-logo .ver{color:#444;font-size:10px;margin-top:4px}')
        $null = $sb.AppendLine('.sb-logo .ts{color:#555;font-size:10px;margin-top:2px}')
        $null = $sb.AppendLine('.nav{padding:10px 0}')
        $null = $sb.AppendLine('.nav a{display:flex;align-items:center;gap:9px;padding:9px 16px;color:#888;font-size:12px;text-decoration:none;border-left:3px solid transparent;transition:background .12s,color .12s,border-color .12s}')
        $null = $sb.AppendLine('.nav a:hover{color:#D4D4D4;background:#252526;border-left-color:#4EC9B085}')
        $null = $sb.AppendLine('.nb{margin-left:auto;padding:1px 8px;border-radius:9px;font-size:10px;font-weight:700;background:#2D2D30;color:#9CDCFE}')
        $null = $sb.AppendLine('.nb.err{background:#4D1A1A;color:#F44747}.nb.warn{background:#3D3200;color:#DCDCAA}.nb.ok{background:#1A3A2A;color:#4EC9B0}')
        $null = $sb.AppendLine('.sb-foot{margin-top:auto;padding:14px 16px;border-top:1px solid #252526;color:#444;font-size:10px}')
        # main
        $null = $sb.AppendLine('.main{flex:1;padding:28px 32px;overflow-x:hidden;min-width:0}')
        # page header
        $null = $sb.AppendLine('.ph{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:28px;padding-bottom:18px;border-bottom:1px solid #252526;flex-wrap:wrap}')
        $null = $sb.AppendLine('.ph h1{color:#4EC9B0;font-size:19px;font-weight:700;line-height:1.3}')
        $null = $sb.AppendLine('.ph .sub{color:#555;font-size:11px;margin-top:4px}')
        $null = $sb.AppendLine('.ph-badge{display:flex;gap:8px;flex-wrap:wrap;align-items:center}')
        $null = $sb.AppendLine('.tag{padding:3px 10px;border-radius:4px;font-size:11px;font-weight:600;background:#2D2D30;color:#9CDCFE}')
        $null = $sb.AppendLine('.tag.err{background:#4D1A1A;color:#F44747}.tag.warn{background:#3D3200;color:#DCDCAA}.tag.ok{background:#1A3A2A;color:#4EC9B0}')
        # stat row
        $null = $sb.AppendLine('.stat-row{display:grid;grid-template-columns:repeat(auto-fill,minmax(175px,1fr));gap:14px;margin-bottom:32px}')
        $null = $sb.AppendLine('.sc{background:#1E1E1E;border:1px solid #2A2A2D;border-radius:8px;padding:18px 16px;border-top:3px solid #2D2D30;transition:border-color .15s}')
        $null = $sb.AppendLine('.sc.ok{border-top-color:#4EC9B0}.sc.err{border-top-color:#F44747}.sc.warn{border-top-color:#DCDCAA}.sc.info{border-top-color:#9CDCFE}')
        $null = $sb.AppendLine('.sc-lbl{color:#555;font-size:10px;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px}')
        $null = $sb.AppendLine('.sc-val{font-size:32px;font-weight:700;line-height:1;color:#D4D4D4}')
        $null = $sb.AppendLine('.sc.err .sc-val{color:#F44747}.sc.warn .sc-val{color:#DCDCAA}.sc.ok .sc-val{color:#4EC9B0}.sc.info .sc-val{color:#9CDCFE}')
        $null = $sb.AppendLine('.sc-sub{color:#555;font-size:11px;margin-top:8px}')
        # sections
        $null = $sb.AppendLine('.section{margin-bottom:38px}')
        $null = $sb.AppendLine('.sec-hdr{display:flex;align-items:center;gap:10px;margin-bottom:10px}')
        $null = $sb.AppendLine('.sec-hdr h2{color:#4EC9B0;font-size:13px;font-weight:700}')
        $null = $sb.AppendLine('.badge{padding:2px 9px;border-radius:10px;font-size:11px;font-weight:700;background:#2D2D30;color:#9CDCFE}')
        $null = $sb.AppendLine('.badge.err{background:#4D1A1A;color:#F44747}.badge.warn{background:#3D3200;color:#DCDCAA}.badge.ok{background:#1A3A2A;color:#4EC9B0}')
        # sysbox
        $null = $sb.AppendLine('.sysbox{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px;background:#1E1E1E;border:1px solid #2A2A2D;border-radius:6px;padding:16px}')
        $null = $sb.AppendLine('.sl{color:#555;font-size:10px;text-transform:uppercase;letter-spacing:.04em}.sv{color:#9CDCFE;font-weight:600;margin-top:3px;font-size:12px;word-break:break-word}')
        # tables
        $null = $sb.AppendLine('.tbl-wrap{border:1px solid #2A2A2D;border-radius:6px;overflow:hidden}')
        $null = $sb.AppendLine('.tbl-scroll{overflow-x:auto;-webkit-overflow-scrolling:touch}')
        $null = $sb.AppendLine('table{width:100%;border-collapse:collapse;min-width:600px}')
        $null = $sb.AppendLine('thead{position:sticky;top:0;z-index:2}')
        $null = $sb.AppendLine('th{background:#252526;color:#4EC9B0;text-align:left;padding:8px 12px;font-size:11px;font-weight:600;white-space:nowrap;border-bottom:2px solid #3E3E42}')
        $null = $sb.AppendLine('td{padding:6px 12px;border-bottom:1px solid #1E1E1E;vertical-align:top;font-size:12px;word-break:break-word}')
        $null = $sb.AppendLine('tr:nth-child(odd) td{background:#222222}tr:nth-child(even) td{background:#272727}tr:hover td{background:#333!important}')
        $null = $sb.AppendLine('.re td{background:#2D1A0E!important}.rk td{background:#25230A!important}.rd td{background:#3A1515!important}')
        $null = $sb.AppendLine('.rc td{background:#3A1515!important}.rr td{background:#2D1A0E!important}.ra td{background:#25230A!important}')
        # chips
        $null = $sb.AppendLine('.chip{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;white-space:nowrap}')
        $null = $sb.AppendLine('.c-err{background:#4D1A1A;color:#F44747}.c-warn{background:#3D3200;color:#DCDCAA}.c-crit{background:#3A1515;color:#F47070;text-transform:uppercase}.c-ok{background:#1A3A2A;color:#4EC9B0}')
        $null = $sb.AppendLine('.mono{font-family:Consolas,monospace;font-size:11px;color:#9CDCFE;word-break:break-all}')
        # empty states
        $null = $sb.AppendLine('.empty{color:#444;font-size:12px;padding:20px 16px;font-style:italic;background:#1A1A1A;border:1px solid #2A2A2D;border-radius:6px}')
        $null = $sb.AppendLine('.empty-ok{color:#4EC9B0;font-size:12px;padding:14px 16px;background:#1A2A24;border:1px solid #1E3A2E;border-radius:6px}')
        # log
        $null = $sb.AppendLine('.log-wrap{background:#0B0B0B;border:1px solid #2A2A2D;border-radius:6px;overflow:hidden}')
        $null = $sb.AppendLine('pre{color:#9CDCFE;font-family:Consolas,monospace;font-size:11px;padding:16px;white-space:pre-wrap;word-break:break-all;line-height:1.65;max-height:540px;overflow-y:auto}')
        # footer
        $null = $sb.AppendLine('.foot{color:#333;font-size:11px;text-align:center;border-top:1px solid #1E1E1E;padding:14px 32px;background:#0D0D0D}')
        # responsive
        $null = $sb.AppendLine('@media(max-width:960px){.sidebar{display:none}.main{padding:16px 14px}}')
        $null = $sb.AppendLine('@media(max-width:640px){.stat-row{grid-template-columns:1fr 1fr}.sysbox{grid-template-columns:1fr 1fr}}')
        $null = $sb.AppendLine('@media(max-width:420px){.stat-row{grid-template-columns:1fr}.sysbox{grid-template-columns:1fr}}')
        $null = $sb.AppendLine('</style></head><body>')

        # ── Layout ────────────────────────────────────────────
        $null = $sb.AppendLine('<div class="layout">')

        # ── Sidebar ───────────────────────────────────────────
        $null = $sb.AppendLine('<nav class="sidebar">')
        $null = $sb.AppendLine('<div class="sb-logo">')
        $null = $sb.AppendLine("<div class='app'>Device Repair<br>Tool PRO</div>")
        $null = $sb.AppendLine("<div class='ver'>v$($Script:APP_VERSION)</div>")
        $null = $sb.AppendLine("<div class='ts'>$now</div>")
        $null = $sb.AppendLine('</div>')
        $null = $sb.AppendLine('<div class="nav">')
        $null = $sb.AppendLine('<a href="#sys">&#x1F4BB;&nbsp; Syst&egrave;me</a>')
        $null = $sb.AppendLine("<a href='#devices'>&#x26A0;&nbsp; P&eacute;riph&eacute;riques <span class='nb $badgeDev'>$devCount</span></a>")
        $null = $sb.AppendLine("<a href='#drivers'>&#x1F4C2;&nbsp; Driver Store <span class='nb'>$drvCount</span></a>")
        $null = $sb.AppendLine("<a href='#events'>&#x1F4CB;&nbsp; &Eacute;v&eacute;nements <span class='nb $badgeEvt'>$evtCount</span></a>")
        $null = $sb.AppendLine('<a href="#log">&#x1F5D2;&nbsp; Journal</a>')
        $null = $sb.AppendLine('</div>')
        $null = $sb.AppendLine("<div class='sb-foot'>Device Repair Tool PRO</div>")
        $null = $sb.AppendLine('</nav>')

        # ── Main ──────────────────────────────────────────────
        $null = $sb.AppendLine('<main class="main">')

        # Page header
        $null = $sb.AppendLine('<div class="ph">')
        $null = $sb.AppendLine('<div>')
        $null = $sb.AppendLine("<h1>Device Repair Tool PRO v$($Script:APP_VERSION)</h1>")
        $null = $sb.AppendLine("<div class='sub'>Rapport de diagnostic &mdash; G&eacute;n&eacute;r&eacute; le $now</div>")
        $null = $sb.AppendLine('</div>')
        $null = $sb.AppendLine('<div class="ph-badge">')
        if ($devCount  -gt 0) { $null = $sb.AppendLine("<span class='tag err'>$devCount erreur(s)</span>") }
        if ($critEvt   -gt 0) { $null = $sb.AppendLine("<span class='tag warn'>$critEvt evt critique(s)</span>") }
        if ($devCount  -eq 0) { $null = $sb.AppendLine("<span class='tag ok'>&#x2713; Aucun probl&egrave;me</span>") }
        $null = $sb.AppendLine('</div></div>')

        # ── Stat cards ────────────────────────────────────────
        $null = $sb.AppendLine('<div class="stat-row">')

        $null = $sb.AppendLine("<div class='sc $cardDev'>")
        $null = $sb.AppendLine('<div class="sc-lbl">P&eacute;riph&eacute;riques en erreur</div>')
        $null = $sb.AppendLine("<div class='sc-val'>$devCount</div>")
        $null = $sb.AppendLine('<div class="sc-sub">d&eacute;tect&eacute;s lors du scan</div></div>')

        $null = $sb.AppendLine('<div class="sc info">')
        $null = $sb.AppendLine('<div class="sc-lbl">Driver Store</div>')
        $null = $sb.AppendLine("<div class='sc-val'>$drvCount</div>")
        $null = $sb.AppendLine('<div class="sc-sub">pilotes index&eacute;s</div></div>')

        $null = $sb.AppendLine("<div class='sc $cardEvt'>")
        $null = $sb.AppendLine('<div class="sc-lbl">&Eacute;v&eacute;nements</div>')
        $null = $sb.AppendLine("<div class='sc-val'>$evtCount</div>")
        $null = $sb.AppendLine("<div class='sc-sub'>$critEvt critique(s) / erreur(s)</div></div>")

        $null = $sb.AppendLine('<div class="sc ok">')
        $null = $sb.AppendLine('<div class="sc-lbl">Architecture</div>')
        $null = $sb.AppendLine("<div class='sc-val' style='font-size:20px;padding-top:6px'>$(ConvertTo-HtmlEscaped $sysInfo.Architecture)</div>")
        $null = $sb.AppendLine("<div class='sc-sub'>$(ConvertTo-HtmlEscaped $sysInfo.RAM_GB) GB RAM &mdash; $(ConvertTo-HtmlEscaped $sysInfo.Uptime) up</div></div>")

        $null = $sb.AppendLine('</div>') # /stat-row

        # ── System Info ───────────────────────────────────────
        $null = $sb.AppendLine('<div class="section" id="sys">')
        $null = $sb.AppendLine('<div class="sec-hdr"><h2>Informations syst&egrave;me</h2></div>')
        $null = $sb.AppendLine('<div class="sysbox">')
        foreach ($kv in @(
            @('OS',           $sysInfo.OS),
            @('Architecture', $sysInfo.Architecture),
            @('CPU',          $sysInfo.CPU),
            @('RAM',          "$($sysInfo.RAM_GB) GB"),
            @('Uptime',       $sysInfo.Uptime),
            @('Dernier boot', $sysInfo.LastBoot)
        )) {
            $null = $sb.AppendLine("<div><div class='sl'>$($kv[0])</div><div class='sv'>$(ConvertTo-HtmlEscaped $kv[1])</div></div>")
        }
        $null = $sb.AppendLine('</div></div>')

        # ── Problem Devices ───────────────────────────────────
        $null = $sb.AppendLine('<div class="section" id="devices">')
        $bdCls = if ($devCount -gt 0) {'err'} else {'ok'}
        $null = $sb.AppendLine("<div class='sec-hdr'><h2>P&eacute;riph&eacute;riques en erreur</h2><span class='badge $bdCls'>$devCount</span></div>")
        if ($devCount -gt 0) {
            $null = $sb.AppendLine('<div class="tbl-wrap"><div class="tbl-scroll"><table><thead><tr>')
            $null = $sb.AppendLine('<th>Nom</th><th>Classe</th><th>Statut</th><th>Code</th><th>ID Mat&eacute;riel</th><th>Version pilote</th><th>Date</th><th>Fournisseur</th>')
            $null = $sb.AppendLine('</tr></thead><tbody>')
            foreach ($d in $Script:ProblemDevices) {
                $rc   = switch ($d.StatusCode) { 'Error' {'re'} 'Unknown' {'rk'} 'Degraded' {'rd'} default {''} }
                $chip = switch ($d.StatusCode) {
                    'Error'    { "<span class='chip c-err'>Error</span>" }
                    'Unknown'  { "<span class='chip c-warn'>Unknown</span>" }
                    'Degraded' { "<span class='chip c-crit'>Degraded</span>" }
                    default    { "<span class='chip c-ok'>$(ConvertTo-HtmlEscaped $d.StatusCode)</span>" }
                }
                $null = $sb.AppendLine("<tr class='$rc'><td>$(ConvertTo-HtmlEscaped $d.Name)</td><td>$(ConvertTo-HtmlEscaped $d.Class)</td><td>$chip</td><td class='mono'>$($d.ErrorCode)</td><td class='mono'>$(ConvertTo-HtmlEscaped $d.InstanceId)</td><td class='mono'>$(ConvertTo-HtmlEscaped $d.DriverVersion)</td><td style='white-space:nowrap'>$($d.DriverDate)</td><td>$(ConvertTo-HtmlEscaped $d.Provider)</td></tr>")
            }
            $null = $sb.AppendLine('</tbody></table></div></div>')
        } else {
            $null = $sb.AppendLine('<div class="empty-ok">&#x2713; Aucun p&eacute;riph&eacute;rique en erreur d&eacute;tect&eacute; (scan non effectu&eacute; ou syst&egrave;me sain)</div>')
        }
        $null = $sb.AppendLine('</div>')

        # ── Driver Store ──────────────────────────────────────
        $null = $sb.AppendLine('<div class="section" id="drivers">')
        $null = $sb.AppendLine("<div class='sec-hdr'><h2>Driver Store</h2><span class='badge'>$drvCount</span></div>")
        if ($drvCount -gt 0) {
            $null = $sb.AppendLine('<div class="tbl-wrap"><div class="tbl-scroll"><table><thead><tr>')
            $null = $sb.AppendLine('<th>Nom publi&eacute;</th><th>Fichier INF</th><th>Fournisseur</th><th>Classe</th><th>Version</th><th>Date</th>')
            $null = $sb.AppendLine('</tr></thead><tbody>')
            foreach ($drv in $Script:DriverStoreDrivers) {
                $null = $sb.AppendLine("<tr><td>$(ConvertTo-HtmlEscaped $drv.PublishedName)</td><td class='mono'>$(ConvertTo-HtmlEscaped $drv.OriginalName)</td><td>$(ConvertTo-HtmlEscaped $drv.Provider)</td><td>$(ConvertTo-HtmlEscaped $drv.Class)</td><td class='mono'>$(ConvertTo-HtmlEscaped $drv.Version)</td><td style='white-space:nowrap'>$(ConvertTo-HtmlEscaped $drv.Date)</td></tr>")
            }
            $null = $sb.AppendLine('</tbody></table></div></div>')
        } else {
            $null = $sb.AppendLine('<div class="empty">Analyse Driver Store non effectu&eacute;e.</div>')
        }
        $null = $sb.AppendLine('</div>')

        # ── Events ────────────────────────────────────────────
        $null = $sb.AppendLine('<div class="section" id="events">')
        $beCls = if ($critEvt -gt 0) {'warn'} else {''}
        $null = $sb.AppendLine("<div class='sec-hdr'><h2>&Eacute;v&eacute;nements pilotes</h2><span class='badge $beCls'>$evtCount</span></div>")
        if ($evtCount -gt 0) {
            $null = $sb.AppendLine('<div class="tbl-wrap"><div class="tbl-scroll"><table><thead><tr>')
            $null = $sb.AppendLine('<th>Niveau</th><th>Date / Heure</th><th>Event ID</th><th>Source</th><th>P&eacute;riph&eacute;rique</th><th>Description</th><th>Conseil</th>')
            $null = $sb.AppendLine('</tr></thead><tbody>')
            foreach ($evt in $Script:EventEntries) {
                $rc   = switch ($evt.Level) { 'Critique' {'rc'} 'Erreur' {'rr'} 'Avertissement' {'ra'} default {''} }
                $chip = switch ($evt.Level) {
                    'Critique'      { "<span class='chip c-crit'>Critique</span>" }
                    'Erreur'        { "<span class='chip c-err'>Erreur</span>"    }
                    'Avertissement' { "<span class='chip c-warn'>Avert.</span>"   }
                    default         { "<span class='chip c-ok'>$(ConvertTo-HtmlEscaped $evt.Level)</span>" }
                }
                $null = $sb.AppendLine("<tr class='$rc'><td>$chip</td><td style='white-space:nowrap'>$($evt.TimeCreated)</td><td class='mono'>$($evt.EventId)</td><td>$(ConvertTo-HtmlEscaped $evt.ProviderName)</td><td>$(ConvertTo-HtmlEscaped $evt.DeviceHint)</td><td>$(ConvertTo-HtmlEscaped $evt.Message)</td><td>$(ConvertTo-HtmlEscaped $evt.Advice)</td></tr>")
            }
            $null = $sb.AppendLine('</tbody></table></div></div>')
        } else {
            $null = $sb.AppendLine('<div class="empty">Analyse &eacute;v&eacute;nements non effectu&eacute;e.</div>')
        }
        $null = $sb.AppendLine('</div>')

        # ── Full log ──────────────────────────────────────────
        $null = $sb.AppendLine('<div class="section" id="log">')
        $null = $sb.AppendLine('<div class="sec-hdr"><h2>Journal complet</h2></div>')
        $null = $sb.AppendLine('<div class="log-wrap">')
        $null = $sb.AppendLine("<pre>$(ConvertTo-HtmlEscaped (Get-FormattedLogs))</pre>")
        $null = $sb.AppendLine('</div></div>')

        $null = $sb.AppendLine('</main></div>')
        $null = $sb.AppendLine("<div class='foot'>Device Repair Tool PRO v$($Script:APP_VERSION) &mdash; G&eacute;n&eacute;r&eacute; le $now</div>")
        $null = $sb.AppendLine('</body></html>')

        # Ecriture UTF-8 sans BOM
        [System.IO.File]::WriteAllText($reportPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

        Write-Log "Rapport HTML exporte: $reportPath" "SUCCESS"
        Set-Status $StatusLabel "Rapport HTML exporte sur le Bureau"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }

        Start-Process $reportPath
    }
    catch {
        Write-Log "Erreur lors de l'export du rapport: $($_.Exception.Message)" "ERROR"
        Set-Status $StatusLabel "Erreur d'export"
    }
}

# ============================================================
# 15. Interface WPF (XAML)
# ============================================================
$xamlString = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Device Repair Tool PRO v4.1"
        Height="750" Width="1150" WindowStartupLocation="CenterScreen"
        Background="#1E1E1E" Foreground="White" FontFamily="Segoe UI">
    <Window.Resources>

        <Style TargetType="TabItem">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#3E3E42"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1E1E1E"/>
                    <Setter Property="Foreground" Value="#4EC9B0"/>
                    <Setter Property="FontWeight" Value="Bold"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="3"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Opacity" Value="0.85"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.4"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#1E1E1E"/>
            <Setter Property="Foreground" Value="#D4D4D4"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="RowBackground" Value="#252526"/>
            <Setter Property="AlternatingRowBackground" Value="#2A2A2D"/>
            <Setter Property="BorderBrush" Value="#3E3E42"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#3E3E42"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="Foreground" Value="#4EC9B0"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="BorderBrush" Value="#3E3E42"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#D4D4D4"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="4,2"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0D0D0D"/>
            <Setter Property="Foreground" Value="#9CDCFE"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="BorderBrush" Value="#3E3E42"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="4"/>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Height" Value="6"/>
            <Setter Property="Foreground" Value="#4EC9B0"/>
            <Setter Property="Background" Value="#2D2D30"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>

    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#252526" CornerRadius="4" Padding="12,8" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock x:Name="txtAppTitle" FontSize="18" FontWeight="Bold">
                    <Run Text="⚙ " Foreground="#FFB700"/>
                    <Run Text="Device Repair Tool PRO" Foreground="#FFB700"/>
                    <Run Text=" v4.1" Foreground="#888888" FontSize="13" FontWeight="Normal"/>
                </TextBlock>
                <TextBlock Text="Gestionnaire de pilotes - Diagnostic et reparation" FontSize="11" Foreground="#666666" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <TabControl x:Name="MainTabControl" Grid.Row="1" Background="#1E1E1E" BorderBrush="#3E3E42" BorderThickness="1">

            <!-- ═══════════════════════════════════════════════════ -->
            <!-- Onglet 1 : Scan peripheriques                      -->
            <!-- ═══════════════════════════════════════════════════ -->
            <TabItem Header="Scan peripheriques">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="120" MinHeight="40"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnScan"      Content="Scanner maintenant" Background="#0E639C" Width="160"/>
                        <Button x:Name="btnRepairSel" Content="Reparer selection"  Background="#CE9178" Width="130" IsEnabled="False"/>
                        <Button x:Name="btnRepairAll" Content="Reparer tout"       Background="#D16969" Width="110" IsEnabled="False"/>
                        <Button x:Name="btnSelectAll" Content="Tout selectionner"  Background="#3E3E42" Width="130"/>
                        <TextBox x:Name="txtCatalogSearch"
                                 Width="220" Margin="5,0,0,0" Padding="4,3"
                                 Background="#0D0D0D" Foreground="#4EC9B0"
                                 BorderBrush="#3E3E42" FontSize="11"
                                 ToolTip="Terme pre-rempli a la selection. Modifiable avant recherche."
                                 VerticalContentAlignment="Center"/>
                        <Button x:Name="btnSearchCatalog" Content="Catalogue" Background="#3e80ac" Width="80" Margin="4,0,0,0"/>
                    </StackPanel>

                    <ProgressBar x:Name="progressBar" Grid.Row="1" Margin="0,0,0,8"/>

                    <!-- DataGrid peripheriques -->
                    <DataGrid x:Name="dgDevices" Grid.Row="2" AutoGenerateColumns="False"
                              SelectionMode="Extended" SelectionUnit="FullRow"
                              CanUserAddRows="False" CanUserDeleteRows="False">

                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Error"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#2D1A0E"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Error"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#4A2F2F"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Unknown"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#25230A"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Unknown"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#3D3A15"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Degraded"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#3A1515"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="Degraded"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#5C2F2F"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="OK"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#162416"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding StatusCode}" Value="OK"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#223522"/>
                                    </MultiDataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>

                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Header="Copier le nom"/>
                                <MenuItem Header="Copier la version du pilote"/>
                                <MenuItem Header="Copier l'ID materiel"/>
                                <Separator/>
                                <MenuItem Header="Rechercher dans le catalogue"/>
                            </ContextMenu>
                        </DataGrid.ContextMenu>

                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                                    Header="Sel" Width="40"/>
                            <DataGridTemplateColumn Header="Peripherique" Width="*" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding Name, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Classe" Width="110" SortMemberPath="Class">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding Class, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Statut" Width="90" SortMemberPath="StatusCode">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding StatusCode, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Code" Width="55" SortMemberPath="ErrorCode">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding ErrorCode, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Version" Width="120" SortMemberPath="DriverVersion">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding DriverVersion, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Date" Width="90" SortMemberPath="DriverDate">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding DriverDate, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Fournisseur" Width="140" SortMemberPath="Provider">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding Provider, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="ID Materiel" Width="*" IsReadOnly="True" SortMemberPath="InstanceId">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <TextBox Text="{Binding InstanceId, Mode=OneWay}" IsReadOnly="True"
                                                 BorderThickness="0" Background="Transparent" Padding="2,0" VerticalAlignment="Center"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                        </DataGrid.Columns>
                    </DataGrid>

                    <GridSplitter Grid.Row="3" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <TextBox x:Name="txtLogScan" Grid.Row="4" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <TextBlock x:Name="txtStatusScan" Grid.Row="5" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

            <!-- ═══════════════════════════════════════════════════ -->
            <!-- Onglet 2 : Driver Store                            -->
            <!-- ═══════════════════════════════════════════════════ -->
            <TabItem Header="Driver Store">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="120" MinHeight="40"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnScanStore"   Content="Analyser le Driver Store" Background="#0E639C" Width="180"/>
                        <Button x:Name="btnRemoveStore" Content="Supprimer selection"       Background="#D16969" Width="150" IsEnabled="False"/>
                    </StackPanel>

                    <DataGrid x:Name="dgDriverStore" Grid.Row="1" AutoGenerateColumns="False">
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Binding="{Binding Selected, Mode=TwoWay}" Header="Sel" Width="40"/>
                            <DataGridTextColumn Binding="{Binding PublishedName}" Header="Nom publie"    Width="90"  IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding OriginalName}"  Header="Nom d'origine" Width="150" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Provider}"      Header="Fournisseur"   Width="*"   IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Class}"         Header="Classe"        Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Date}"          Header="Date"          Width="90"  IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Version}"       Header="Version"       Width="150" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <TextBox x:Name="txtLogStore" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <TextBlock x:Name="txtStatusStore" Grid.Row="4" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

            <!-- ═══════════════════════════════════════════════════ -->
            <!-- Onglet 3 : Evenements Pilotes (NOUVEAU v4.1)       -->
            <!-- ═══════════════════════════════════════════════════ -->
            <TabItem Header="Evenements">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="120" MinHeight="40"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Boutons -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnEventScan7"      Content="Analyser (7 jours)"       Background="#0E639C" Width="150"/>
                        <Button x:Name="btnEventScan30"     Content="Analyser (30 jours)"      Background="#0E639C" Width="150"/>
                        <Button x:Name="btnEventClear"      Content="Vider liste"              Background="#3E3E42" Width="100"/>
                        <Button x:Name="btnEventCopy"       Content="Copier evenement"         Background="#3E3E42" Width="140"/>
                        <Button x:Name="btnEventOpenViewer" Content="Observateur evenements"   Background="#2D5986" Width="190" Margin="5,0,0,0"/>
                    </StackPanel>

                    <!-- Barre de synthese compteurs -->
                    <Border Grid.Row="1" Background="#252526" CornerRadius="4" Padding="10,5" Margin="0,0,0,8">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock x:Name="txtEvtCritical" Text="Critiques : 0"      Foreground="#F44747" FontWeight="Bold" Margin="0,0,20,0"/>
                            <TextBlock x:Name="txtEvtError"    Text="Erreurs : 0"         Foreground="#CE9178" FontWeight="Bold" Margin="0,0,20,0"/>
                            <TextBlock x:Name="txtEvtWarning"  Text="Avertissements : 0"  Foreground="#DCDCAA" FontWeight="Bold" Margin="0,0,20,0"/>
                            <TextBlock x:Name="txtEvtTotal"    Text="Total : 0"            Foreground="#9CDCFE" FontWeight="Bold" Margin="0,0,30,0"/>
                            <TextBlock Text="Double-clic = message complet  |  Correle = peripherique identifie par le scan onglet 1"
                                       Foreground="#555555" FontSize="10" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Border>

                    <!-- DataGrid evenements -->
                    <DataGrid x:Name="dgEvents" Grid.Row="2" AutoGenerateColumns="False" SelectionMode="Single">
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Critique"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#3A1515"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Critique"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#5C2F2F"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Erreur"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#2D1A0E"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Erreur"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#4A2F2F"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Avertissement"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="False"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#25230A"/>
                                    </MultiDataTrigger>
                                    <MultiDataTrigger>
                                        <MultiDataTrigger.Conditions>
                                            <Condition Binding="{Binding Level}" Value="Avertissement"/>
                                            <Condition Binding="{Binding RelativeSource={RelativeSource Self}, Path=IsMouseOver}" Value="True"/>
                                        </MultiDataTrigger.Conditions>
                                        <Setter Property="Background" Value="#3D3A15"/>
                                    </MultiDataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Binding="{Binding Level}"        Header="Niveau"           Width="110" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding TimeCreated}"  Header="Date / Heure"     Width="140" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding EventId}"      Header="Event ID"         Width="70"  IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding ProviderName}" Header="Source"           Width="180" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding DeviceHint}"   Header="Peripherique lie" Width="160" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Message}"      Header="Description"      Width="*"   IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Advice}"       Header="Conseil"          Width="230" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <GridSplitter Grid.Row="3" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <TextBox x:Name="txtLogEvent" Grid.Row="4" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <TextBlock x:Name="txtStatusEvent" Grid.Row="5" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

            <!-- ═══════════════════════════════════════════════════ -->
            <!-- Onglet 4 : Windows Update (pilotes)                -->
            <!-- ═══════════════════════════════════════════════════ -->
            <TabItem Header="Update">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="120" MinHeight="40"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnWUSearch"      Content="Rechercher les mises a jour" Background="#0E639C" Width="210" Margin="0,0,5,0"/>
                        <Button x:Name="btnWUOpenCatalog" Content="Catalogue Microsoft"         Background="#0E639C" Width="160"/>
                        <Button x:Name="btnAdGuard"       Content="store.rg-adguard.net"        Background="#2D5986" Width="170" Margin="5,0,0,0"/>
                    </StackPanel>

                    <DataGrid x:Name="dgWUUpdates" Grid.Row="1" AutoGenerateColumns="False" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Binding="{Binding Title}"       Header="Titre"       Width="*"  IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Size_MB}"     Header="Taille MB"   Width="80" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding KBArticle}"   Header="KB"          Width="80" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Description}" Header="Description" Width="*"  IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>

                    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <TextBox x:Name="txtLogWU" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <TextBlock x:Name="txtStatusWU" Grid.Row="4" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

            <!-- ═══════════════════════════════════════════════════ -->
            <!-- Onglet 5 : Outils Systeme                          -->
            <!-- ═══════════════════════════════════════════════════ -->
            <TabItem Header="Outils Systeme">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="180" MinHeight="60"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnSFC"          Content="SFC /scannow"        Background="#CE9178" Width="120"/>
                        <Button x:Name="btnDISM"         Content="DISM /RestoreHealth" Background="#CE9178" Width="160"/>
                        <Button x:Name="btnResetUSB"     Content="Reset USB"           Background="#CE9178" Width="100"/>
                        <Button x:Name="btnExportReport" Content="Exporter rapport"    Background="#0E639C" Width="130"/>
                        <Button x:Name="btnRefreshSys"   Content="Actualiser"          Background="#3E3E42" Width="90"  Margin="5,0,0,0"/>
                    </StackPanel>

                    <!-- Zone infos systeme + mini tuto -->
                    <TextBox x:Name="txtSysInfo" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" FontSize="11"
                             Background="#0D0D0D" Foreground="#D4D4D4"/>

                    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <!-- Log redimensionnable -->
                    <TextBox x:Name="txtLogSys" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <TextBlock x:Name="txtStatusSys" Grid.Row="4" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

        </TabControl>

        <TextBlock x:Name="txtFooter" Grid.Row="2" FontSize="9" Foreground="#444444"
                   Margin="0,6,0,0" HorizontalAlignment="Center"
                   Text="Device Repair Tool PRO v4.1 | Gestionnaire de pilotes"/>
    </Grid>
</Window>
"@

# ============================================================
# 16. Chargement de l'interface WPF
# ============================================================
try {
    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($xamlString)
    $stream = [System.IO.MemoryStream]::new($bytes)
    $window = [System.Windows.Markup.XamlReader]::Load($stream)
    $stream.Dispose()
    $window.Title = "$Script:APP_NAME v$Script:APP_VERSION"
    Write-Log "Interface chargee avec succes" "SUCCESS"
}
catch {
    Write-Log "Erreur de chargement de l'interface: $($_.Exception.Message)" "ERROR"
    [System.Console]::WriteLine("ERREUR XAML : $($_.Exception.Message)")
    Read-Host "Appuyez sur Entree pour quitter"
    exit 1
}

# ============================================================
# 17. Recuperation des controles
# ============================================================

# --- Onglet Scan peripheriques ---
$btnScan          = $window.FindName("btnScan")
$btnRepairSel     = $window.FindName("btnRepairSel")
$btnRepairAll     = $window.FindName("btnRepairAll")
$btnSelectAll     = $window.FindName("btnSelectAll")
$btnSearchCatalog = $window.FindName("btnSearchCatalog")
$txtCatalogSearch = $window.FindName("txtCatalogSearch")
$dgDevices        = $window.FindName("dgDevices")
$progressBar      = $window.FindName("progressBar")
$txtStatusScan    = $window.FindName("txtStatusScan")
$txtLogScan       = $window.FindName("txtLogScan")

# --- Onglet Driver Store ---
$btnScanStore     = $window.FindName("btnScanStore")
$btnRemoveStore   = $window.FindName("btnRemoveStore")
$dgDriverStore    = $window.FindName("dgDriverStore")
$txtStatusStore   = $window.FindName("txtStatusStore")
$txtLogStore      = $window.FindName("txtLogStore")

# --- Onglet Evenements ---
$btnEventScan7      = $window.FindName("btnEventScan7")
$btnEventScan30     = $window.FindName("btnEventScan30")
$btnEventClear      = $window.FindName("btnEventClear")
$btnEventCopy       = $window.FindName("btnEventCopy")
$btnEventOpenViewer = $window.FindName("btnEventOpenViewer")
$dgEvents           = $window.FindName("dgEvents")
$txtEvtCritical     = $window.FindName("txtEvtCritical")
$txtEvtError        = $window.FindName("txtEvtError")
$txtEvtWarning      = $window.FindName("txtEvtWarning")
$txtEvtTotal        = $window.FindName("txtEvtTotal")
$txtLogEvent        = $window.FindName("txtLogEvent")
$txtStatusEvent     = $window.FindName("txtStatusEvent")

# --- Onglet Windows Update ---
$btnWUSearch      = $window.FindName("btnWUSearch")
$btnWUOpenCatalog = $window.FindName("btnWUOpenCatalog")
$btnAdGuard       = $window.FindName("btnAdGuard")
$dgWUUpdates      = $window.FindName("dgWUUpdates")
$txtStatusWU      = $window.FindName("txtStatusWU")
$txtLogWU         = $window.FindName("txtLogWU")

# --- Onglet Outils Systeme ---
$btnSFC           = $window.FindName("btnSFC")
$btnDISM          = $window.FindName("btnDISM")
$btnResetUSB      = $window.FindName("btnResetUSB")
$btnExportReport  = $window.FindName("btnExportReport")
$btnRefreshSys    = $window.FindName("btnRefreshSys")
$txtSysInfo       = $window.FindName("txtSysInfo")
$txtStatusSys     = $window.FindName("txtStatusSys")
$txtLogSys        = $window.FindName("txtLogSys")

# ============================================================
# 18. Fonction + initialisation infos systeme
# ============================================================
function Update-SysInfoBox {
    param([System.Windows.Controls.TextBox]$txtBox)

    $sysInfo = Get-SystemInfo
    $sb = [System.Text.StringBuilder]::new()

    # --- RESUME DIAGNOSTIC (si scan effectue) ---
    if ($Script:ScanDone -and $Script:ProblemDevices.Count -gt 0) {
        $errCount = @($Script:ProblemDevices | Where-Object { $_.StatusCode -in @('Error','Degraded') }).Count
        $unkCount = @($Script:ProblemDevices | Where-Object { $_.StatusCode -eq 'Unknown' }).Count
        $okCount  = @($Script:ProblemDevices | Where-Object { $_.StatusCode -eq 'OK' }).Count
        $null = $sb.AppendLine("RESUME DIAGNOSTIC")
        $null = $sb.AppendLine("-----------------")
        $null = $sb.AppendLine("Peripheriques totaux : $($Script:ProblemDevices.Count)  |  En erreur : $errCount  |  Inconnu : $unkCount  |  OK : $okCount")
        if ($Script:EventEntries.Count -gt 0) {
            $evtCrit = @($Script:EventEntries | Where-Object { $_.Level -eq 'Critique' }).Count
            $evtErr  = @($Script:EventEntries | Where-Object { $_.Level -eq 'Erreur' }).Count
            $evtWarn = @($Script:EventEntries | Where-Object { $_.Level -eq 'Avertissement' }).Count
            $null = $sb.AppendLine("Evenements :          $($Script:EventEntries.Count) total  |  $evtCrit critique(s)  $evtErr erreur(s)  $evtWarn avert.")
        } else {
            $null = $sb.AppendLine("Evenements :          non analyses (cliquez sur onglet Evenements)")
        }
        $null = $sb.AppendLine("")
    }

    # --- SYSTEME ---
    $null = $sb.AppendLine("SYSTEME")
    $null = $sb.AppendLine("-------")
    $null = $sb.AppendLine("OS:             $($sysInfo.OS)")
    $null = $sb.AppendLine("Version:        $($Script:OSInfo.Version)  (Build $($Script:OSInfo.BuildNumber))")
    $null = $sb.AppendLine("Edition:        $($Script:OSInfo.Caption)")
    $null = $sb.AppendLine("Architecture:   $($sysInfo.Architecture)")
    $null = $sb.AppendLine("Installe le:    $(try { $Script:OSInfo.InstallDate.ToString('yyyy-MM-dd') } catch { 'N/A' })")
    $null = $sb.AppendLine("CPU:            $($sysInfo.CPU)")
    $null = $sb.AppendLine("RAM:            $($sysInfo.RAM_GB) GB")
    $null = $sb.AppendLine("Uptime:         $($sysInfo.Uptime)")
    $null = $sb.AppendLine("Dernier boot:   $($sysInfo.LastBoot)")
    $null = $sb.AppendLine("PowerShell:     $($sysInfo.PSVersion)")

    # --- BIOS ---
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("BIOS / FIRMWARE")
    $null = $sb.AppendLine("---------------")
    try {
        $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
        $null = $sb.AppendLine("Fabricant:      $($bios.Manufacturer)")
        $null = $sb.AppendLine("Version:        $($bios.SMBIOSBIOSVersion)")
        $null = $sb.AppendLine("Date release:   $(try { $bios.ReleaseDate.ToString('yyyy-MM-dd') } catch { 'N/A' })")
        $null = $sb.AppendLine("Numero serie:   $($bios.SerialNumber)")
    } catch { $null = $sb.AppendLine("Impossible de lire les infos BIOS") }

    # --- GPU ---
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("GPU")
    $null = $sb.AppendLine("---")
    try {
        $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
        foreach ($gpu in $gpus) {
            $vramMB = if ($gpu.AdapterRAM -and $gpu.AdapterRAM -gt 0) { "$([math]::Round($gpu.AdapterRAM / 1MB)) MB" } else { "N/A" }
            $null = $sb.AppendLine("$($gpu.Name) | VRAM : $vramMB | Statut : $($gpu.Status)")
        }
    } catch { $null = $sb.AppendLine("Impossible de lire les infos GPU") }

    # --- DISQUES PHYSIQUES ---
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("DISQUES PHYSIQUES")
    $null = $sb.AppendLine("-----------------")
    try {
        $disks = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
        foreach ($disk in $disks) {
            $sizeGB = [math]::Round($disk.Size / 1GB, 1)
            $null = $sb.AppendLine("$($disk.Model) | $sizeGB GB | $($disk.InterfaceType) | $($disk.Status)")
        }
    } catch { $null = $sb.AppendLine("Impossible de lire les disques") }

    # --- PARTITIONS ---
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("PARTITIONS (espace libre)")
    $null = $sb.AppendLine("-------------------------")
    try {
        $vols = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
        foreach ($vol in $vols) {
            $totalGB = [math]::Round($vol.Size / 1GB, 1)
            $freeGB  = [math]::Round($vol.FreeSpace / 1GB, 1)
            $usedPct = if ($vol.Size -gt 0) { [math]::Round((($vol.Size - $vol.FreeSpace) / $vol.Size) * 100) } else { 0 }
            $bar     = "[" + ("=" * [math]::Round($usedPct / 5)) + (" " * (20 - [math]::Round($usedPct / 5))) + "]"
            $null = $sb.AppendLine("$($vol.DeviceID)  $bar  $freeGB GB libres / $totalGB GB  ($usedPct% utilise)")
        }
    } catch { $null = $sb.AppendLine("Impossible de lire les partitions") }

    # --- GUIDE ---
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("================================================================")
    $null = $sb.AppendLine("GUIDE DES OUTILS")
    $null = $sb.AppendLine("================================================================")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("[ SFC /scannow ]  (Verificateur de fichiers systeme)")
    $null = $sb.AppendLine("  Analyse et repare les fichiers Windows corrompus ou manquants.")
    $null = $sb.AppendLine("  Duree : 5 a 15 minutes. Ne pas eteindre le PC pendant l'analyse.")
    $null = $sb.AppendLine("  Code 0 = sain  |  Code 1 = repare  |  Code 2 = non repare (lancer DISM d'abord)")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("[ DISM /RestoreHealth ]  (Reparation de l'image systeme)")
    $null = $sb.AppendLine("  Repare l'image Windows. Necessite Internet. Duree : 10 a 30 min.")
    $null = $sb.AppendLine("  Ordre recommande : DISM --> SFC --> Redemarrer")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("[ Reset USB ]")
    $null = $sb.AppendLine("  Reinitialise les peripheriques USB non-HID (sans deconnecter clavier/souris).")
    $null = $sb.AppendLine("")
    $null = $sb.AppendLine("[ Exporter rapport ]")
    $null = $sb.AppendLine("  Genere un rapport HTML complet sur le Bureau.")

    $txtBox.Text = $sb.ToString()
}

Update-SysInfoBox -txtBox $txtSysInfo


# ============================================================
# 19. Gestionnaires d'evenements
# ============================================================

# ── Onglet Scan peripheriques ──────────────────────────────────

$btnScan.Add_Click({
    $btnScan.IsEnabled      = $false
    $btnRepairSel.IsEnabled = $false
    $btnRepairAll.IsEnabled = $false
    Invoke-SystemScan -DataGrid $dgDevices -StatusLabel $txtStatusScan -ProgressBar $progressBar `
        -BtnStart $btnScan -BtnRepairSel $btnRepairSel -BtnRepairAll $btnRepairAll -TxtLog $txtLogScan
})

$btnSelectAll.Add_Click({
    if ($Script:ProblemDevices.Count -eq 0) { return }
    $allChecked = (@($Script:ProblemDevices | Where-Object { -not $_.Selected }).Count -eq 0)
    foreach ($device in $Script:ProblemDevices) {
        $device.Selected = -not $allChecked
    }
    $dgDevices.Items.Refresh()
})

$btnRepairSel.Add_Click({
    $selectedDevices = @($Script:ProblemDevices | Where-Object { $_.Selected })
    if ($selectedDevices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun peripherique selectionne", "Information", "OK", "Information")
        return
    }
    Invoke-RepairDevices -Devices $selectedDevices -LogBox $txtLogScan -StatusLabel $txtStatusScan
    $txtLogScan.Text = Get-FormattedLogs
})

$btnRepairAll.Add_Click({
    $errorDevices = @($Script:ProblemDevices | Where-Object { $_.StatusCode -in @('Error','Degraded') })
    if ($errorDevices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun peripherique en erreur detecte", "Information", "OK", "Information")
        return
    }
    $confirm = [System.Windows.MessageBox]::Show(
        "Reparer TOUS les $($errorDevices.Count) peripheriques en erreur ?",
        "Confirmation", "YesNo", "Warning"
    )
    if ($confirm -eq "Yes") {
        Invoke-RepairDevices -Devices $errorDevices -LogBox $txtLogScan -StatusLabel $txtStatusScan
        $txtLogScan.Text = Get-FormattedLogs
    }
})

# Menu contextuel dgDevices : copie des valeurs de cellules
$ctxItems = @($dgDevices.ContextMenu.Items | Where-Object { $_ -is [System.Windows.Controls.MenuItem] })
$ctxItems[0].Add_Click({   # Copier le nom
    $item = $dgDevices.SelectedItem
    if ($item) { [System.Windows.Clipboard]::SetText($item.Name) }
})
$ctxItems[1].Add_Click({   # Copier la version
    $item = $dgDevices.SelectedItem
    if ($item) { [System.Windows.Clipboard]::SetText($item.DriverVersion) }
})
$ctxItems[2].Add_Click({   # Copier l'ID materiel
    $item = $dgDevices.SelectedItem
    if ($item) { [System.Windows.Clipboard]::SetText($item.InstanceId) }
})
$ctxItems[3].Add_Click({   # Rechercher dans le catalogue
    $item = $dgDevices.SelectedItem
    if (-not $item) { return }
    $officialUrl = Get-OfficialDriverUrl $item
    if ($officialUrl) {
        Start-Process $officialUrl
        return
    }
    $term = Get-CatalogSearchTerm $item
    if ($txtCatalogSearch) { $txtCatalogSearch.Text = $term }
    Start-Process "https://catalog.update.microsoft.com/Search.aspx?q=$([System.Uri]::EscapeDataString($term))"
})

# Double-clic sur la colonne ID Materiel (index 8) -> pre-remplit le champ + ouvre le catalogue
$dgDevices.Add_PreviewMouseDoubleClick({
    # OriginalSource hit-test : fiable en tunneling, contrairement a CurrentCell
    $node = $_.OriginalSource -as [System.Windows.DependencyObject]
    while ($node -and $node -isnot [System.Windows.Controls.DataGridCell]) {
        $node = [System.Windows.Media.VisualTreeHelper]::GetParent($node)
    }
    if (-not $node) { return }
    if ($dgDevices.Columns.IndexOf($node.Column) -ne 8) { return }
    $item = $dgDevices.SelectedItem
    if (-not $item) { return }
    if ($txtCatalogSearch) {
        $txtCatalogSearch.Text = Get-HardwareIdTerm $item
        $txtCatalogSearch.Focus()
        $txtCatalogSearch.SelectAll()
    }
    $_.Handled = $true
})

# Construit le terme de recherche catalogue
# Le NOM du peripherique donne de bien meilleurs resultats que VEN/DEV sur catalog.update.microsoft.com
# Retourne l'URL officielle du fabricant selon VEN/VID - fallback $null si inconnu
function Get-OfficialDriverUrl {
    param($device)
    $id = $device.InstanceId

    # Peripheriques ACPI (processeurs, EC, firmware) - pas de VEN_ dans l'InstanceId
    if ($id -match '^ACPI\\AuthenticAMD')   { return 'https://www.amd.com/en/support' }
    if ($id -match '^ACPI\\GenuineIntel')   { return 'https://www.intel.com/content/www/us/en/download-center/home.html' }
    if ($id -match '^ACPI\\INTC')           { return 'https://www.intel.com/content/www/us/en/download-center/home.html' }
    if ($id -match '^ACPI\\AMDI|^ACPI\\AMD') { return 'https://www.amd.com/en/support' }

    $pciUrls = @{
        # GPU
        '10DE' = 'https://www.nvidia.com/Download/Find.aspx'
        '1002' = 'https://www.amd.com/en/support'
        # CPU / Chipset
        '1022' = 'https://www.amd.com/en/support'
        '8086' = 'https://www.intel.com/content/www/us/en/download-center/home.html'
        # Audio
        '10EC' = 'https://www.realtek.com/en/downloads'
        '1102' = 'https://support.creative.com/'
        '13F6' = 'https://www.cmedia.com.tw/driver'
        '1412' = 'https://www.viatech.com/en/support/drivers/'
        # Reseau / WiFi
        '14E4' = 'https://www.broadcom.com/support'
        '168C' = 'https://www.qualcomm.com/support'
        '1969' = 'https://www.qualcomm.com/support'
        '1814' = 'https://www.mediatek.com/consumer'
        '14C3' = 'https://www.mediatek.com/consumer'
        '11AB' = 'https://www.marvell.com/support/'
        '1D6A' = 'https://www.marvell.com/support/'
        # Chipset / Controleurs
        '1106' = 'https://www.viatech.com/en/support/drivers/'
        '1039' = 'https://www.sis.com/support/'
        '1131' = 'https://www.nxp.com/support/'
        '1B21' = 'https://www.asmedia.com.tw/product/asmedia-usb-controller'
        '1912' = 'https://www.renesas.com/en/support'
        '1095' = 'https://www.latticesemi.com/Support'
        # Stockage / RAID
        '1000' = 'https://docs.broadcom.com/'
        '1B4B' = 'https://www.marvell.com/support/'
        '144D' = 'https://semiconductor.samsung.com/us/consumer-storage/support/tools/'
        '15B7' = 'https://support-en.wd.com/app/answers/detailweb/a_id/25929'
        '1987' = 'https://www.phison.com/en/support/support-list'
        '1179' = 'https://www.kioxia.com/en-us/support.html'
        # OEM systeme
        '1043' = 'https://www.asus.com/support/'
        '1028' = 'https://www.dell.com/support/home/en-us'
        '103C' = 'https://support.hp.com/us-en/drivers'
        '17AA' = 'https://support.lenovo.com/'
        '1025' = 'https://www.acer.com/us-en/support'
        '1462' = 'https://www.msi.com/support/'
        '1458' = 'https://www.gigabyte.com/Support'
        '196D' = 'https://www.clevo.com.tw/Support.asp'
        '1071' = 'https://www.toshiba.com/support'
    }

    $usbUrls = @{
        # Peripheriques gaming / HID
        '046D' = 'https://support.logi.com/hc/en-us'
        '1532' = 'https://support.razer.com/'
        '1B1C' = 'https://www.corsair.com/us/en/s/downloads'
        '1038' = 'https://steelseries.com/engine'
        '0951' = 'https://www.kingston.com/us/support/technical/downloads'
        '045E' = 'https://www.microsoft.com/accessories/en-us'
        '054C' = 'https://www.playstation.com/en-us/support/'
        '057E' = 'https://en-americas-support.nintendo.com/'
        '1E71' = 'https://nzxt.com/downloads'
        '044F' = 'https://support.thrustmaster.com/'
        '0F0D' = 'https://hori.jp/us/support/'
        '0079' = 'https://www.speedlink-gaming.com/support/'
        # Audio
        '041E' = 'https://support.creative.com/'
        '0582' = 'https://www.roland.com/us/support/'
        '0944' = 'https://www.korg.com/us/support/'
        '0D8C' = 'https://www.cmedia.com.tw/driver'
        # Chipsets USB / Serie
        '0BDA' = 'https://www.realtek.com/en/downloads'
        '0483' = 'https://www.st.com/en/development-tools.html'
        '1A86' = 'https://www.wch-ic.com/downloads/category/30.html'
        '0403' = 'https://ftdichip.com/drivers/'
        '04D8' = 'https://www.microchip.com/support'
        '0CF3' = 'https://www.qualcomm.com/support'
        # Biometrie
        '06CB' = 'https://www.synaptics.com/products/biometric-solutions'
        '138A' = 'https://www.synaptics.com/products/biometric-solutions'
        '27C6' = 'https://www.goodixtech.com/support/'
        # Stockage USB
        '04E8' = 'https://www.samsung.com/us/support/'
        '0781' = 'https://support-en.wd.com/'
        '152D' = 'https://www.jmicron.com/supportDownload'
        '090C' = 'https://www.siliconmotion.com/tw/customer.php'
        '1BCF' = 'https://www.sunplusmm.com/support'
        # Imprimantes / Scan
        '03F0' = 'https://support.hp.com/us-en/drivers/printers'
        '04A9' = 'https://www.usa.canon.com/support'
        '04B8' = 'https://epson.com/Support/sl/s'
        '04F9' = 'https://support.brother.com/'
        '04DA' = 'https://panasonic.net/cns/sav/support/'
        '04DD' = 'https://sharp-nec-displays.com/us/support/'
        # Mobile / Android
        '18D1' = 'https://developer.android.com/tools/adb'
        '2717' = 'https://www.mi.com/global/service/supportdetail'
        '22B8' = 'https://motorola-global-portal.custhelp.com/'
        '04BB' = 'https://www.iodata.jp/support/'
        # Cameras / Webcams
        '0C45' = 'https://www.microdia.com/'
        '058F' = 'https://www.alcormicro.com/support/'
        # Gaming plateforme
        '28DE' = 'https://help.steampowered.com/'
    }

    if ($id -match 'VEN_([0-9A-Fa-f]{4})') {
        $key = $matches[1].ToUpper()
        if ($pciUrls.ContainsKey($key)) { return $pciUrls[$key] }
    }
    if ($id -match 'VID_([0-9A-Fa-f]{4})') {
        $key = $matches[1].ToUpper()
        if ($usbUrls.ContainsKey($key)) { return $usbUrls[$key] }
    }
    return $null
}

# Retourne l'ID hardware brut (VEN/DEV) - utilise pour le double-clic col 8
function Get-HardwareIdTerm {
    param($device)
    $id = $device.InstanceId
    if ($id -match 'VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})') {
        return "VEN_$($matches[1].ToUpper()) DEV_$($matches[2].ToUpper())"
    }
    if ($id -match 'VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})') {
        return "VID_$($matches[1].ToUpper()) PID_$($matches[2].ToUpper())"
    }
    return ($id -split '\\')[1]
}

function Get-CatalogSearchTerm {
    param($device)
    $id   = $device.InstanceId
    $name = $device.Name

    $generics = @('Peripherique inconnu','PCI Device','USB Device','Unknown Device',
                  'Base System Device','SM Bus Controller','Coprocesseur')
    if ($name -and $name -notin $generics -and $name.Length -gt 4) {
        return $name
    }
    if ($id -match 'VEN_([0-9A-Fa-f]{4})&DEV_([0-9A-Fa-f]{4})') {
        return "VEN_$($matches[1].ToUpper()) DEV_$($matches[2].ToUpper())"
    }
    if ($id -match 'VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})') {
        return "VID_$($matches[1].ToUpper()) PID_$($matches[2].ToUpper())"
    }
    return ($id -split '\\')[1]
}

# Bouton Catalogue : fabricant officiel si reconnu, sinon Microsoft Catalog
if ($btnSearchCatalog) {
    $btnSearchCatalog.Add_Click({
        $item = $dgDevices.SelectedItem
        $officialUrl = if ($item) { Get-OfficialDriverUrl $item } else { $null }
        if ($officialUrl) {
            Start-Process $officialUrl
            return
        }
        $term = if ($txtCatalogSearch -and $txtCatalogSearch.Text.Trim()) {
            $txtCatalogSearch.Text.Trim()
        } elseif ($item) {
            Get-CatalogSearchTerm $item
        } else { '' }
        if (-not $term) {
            [System.Windows.MessageBox]::Show("Selectionnez un peripherique d'abord.", "Info", "OK", "Information") | Out-Null
            return
        }
        if ($txtCatalogSearch) { $txtCatalogSearch.Text = $term }
        Start-Process "https://catalog.update.microsoft.com/Search.aspx?q=$([System.Uri]::EscapeDataString($term))"
    })
}

# ── Onglet Driver Store ────────────────────────────────────────

$btnScanStore.Add_Click({
    Invoke-DriverStoreScan -DataGrid $dgDriverStore -StatusLabel $txtStatusStore
    $btnRemoveStore.IsEnabled = ($Script:DriverStoreDrivers.Count -gt 0)
    $txtLogStore.Text = Get-FormattedLogs
})

$btnRemoveStore.Add_Click({
    $selectedDrivers = @($Script:DriverStoreDrivers | Where-Object { $_.Selected })
    if ($selectedDrivers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun pilote selectionne", "Information", "OK", "Information")
        return
    }
    Invoke-RemoveDrivers -Drivers $selectedDrivers -LogBox $txtLogStore -StatusLabel $txtStatusStore
    Set-Status $txtStatusStore "Rafraichissement du Driver Store..."
    Invoke-DriverStoreScan -DataGrid $dgDriverStore -StatusLabel $txtStatusStore
})

# ── Onglet Evenements ──────────────────────────────────────────

$btnEventScan7.Add_Click({
    Invoke-EventScan -DataGrid $dgEvents -StatusLabel $txtStatusEvent -LogBox $txtLogEvent `
        -LblCritical $txtEvtCritical -LblError $txtEvtError -LblWarning $txtEvtWarning `
        -LblTotal $txtEvtTotal -DaysBack 7
})

$btnEventScan30.Add_Click({
    Invoke-EventScan -DataGrid $dgEvents -StatusLabel $txtStatusEvent -LogBox $txtLogEvent `
        -LblCritical $txtEvtCritical -LblError $txtEvtError -LblWarning $txtEvtWarning `
        -LblTotal $txtEvtTotal -DaysBack 30
})

$btnEventClear.Add_Click({
    $Script:EventEntries.Clear()
    $dgEvents.ItemsSource = $null
    $txtEvtCritical.Text = "Critiques : 0"
    $txtEvtError.Text    = "Erreurs : 0"
    $txtEvtWarning.Text  = "Avertissements : 0"
    $txtEvtTotal.Text    = "Total : 0"
    Set-Status $txtStatusEvent "Liste videe"
    Write-Log "Liste evenements videe" "INFO"
    $txtLogEvent.Text = Get-FormattedLogs
})

$btnEventCopy.Add_Click({
    $item = $dgEvents.SelectedItem
    if (-not $item) {
        [System.Windows.MessageBox]::Show("Selectionnez un evenement d'abord.", "Info", "OK", "Information") | Out-Null
        return
    }
    $txt = "Niveau:           $($item.Level)`r`n" +
           "Date:             $($item.TimeCreated)`r`n" +
           "Event ID:         $($item.EventId)`r`n" +
           "Source:           $($item.ProviderName)`r`n" +
           "Peripherique lie: $($item.DeviceHint)`r`n" +
           "Description:      $($item.Message)`r`n" +
           "Conseil:          $($item.Advice)`r`n`r`n" +
           "--- Message complet ---`r`n$($item.FullMessage)"
    [System.Windows.Clipboard]::SetText($txt)
    Set-Status $txtStatusEvent "Evenement copie dans le presse-papiers"
})

$btnEventOpenViewer.Add_Click({
    try {
        Start-Process "eventvwr.msc"
        Write-Log "Observateur d'evenements ouvert" "INFO"
    }
    catch {
        Write-Log "Impossible d'ouvrir l'Observateur d'evenements: $_" "ERROR"
    }
})

# Double-clic : popup message complet de l'evenement
$dgEvents.Add_MouseDoubleClick({
    $item = $dgEvents.SelectedItem
    if (-not $item) { return }
    $detail = "Event ID         : $($item.EventId)`n" +
              "Source           : $($item.ProviderName)`n" +
              "Date             : $($item.TimeCreated)`n" +
              "Niveau           : $($item.Level)`n" +
              "Peripherique lie : $($item.DeviceHint)`n" +
              "Conseil          : $($item.Advice)`n`n" +
              "--- Message complet ---`n" +
              "$($item.FullMessage)"
    [System.Windows.MessageBox]::Show($detail, "Detail evenement ID $($item.EventId)", "OK", "Information") | Out-Null
})

# ── Onglet Windows Update ──────────────────────────────────────

$btnWUSearch.Add_Click({
    Invoke-WindowsUpdateSearch -DataGrid $dgWUUpdates -StatusLabel $txtStatusWU -LogBox $txtLogWU
})

$btnWUOpenCatalog.Add_Click({
    try {
        Start-Process "https://catalog.update.microsoft.com/home.aspx"
        Write-Log "Catalogue Microsoft ouvert dans le navigateur." "INFO"
    }
    catch {
        Write-Log "Impossible d'ouvrir le catalogue : $($_.Exception.Message)" "ERROR"
    }
})

$btnAdGuard.Add_Click({
    try {
        Start-Process "https://store.rg-adguard.net/"
        Write-Log "store.rg-adguard.net ouvert dans le navigateur." "INFO"
    }
    catch {
        Write-Log "Impossible d'ouvrir store.rg-adguard.net : $($_.Exception.Message)" "ERROR"
    }
})

# ── Onglet Outils Systeme ──────────────────────────────────────

$btnSFC.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        "Lancer SFC /scannow ?`n(Cela peut prendre plusieurs minutes)",
        "SFC", "YesNo", "Question"
    )
    if ($confirm -eq "Yes") { Invoke-SFCScan -LogBox $txtLogSys -StatusLabel $txtStatusSys }
})

$btnDISM.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        "Lancer DISM /RestoreHealth ?`n(Necessite une connexion Internet, peut prendre du temps)",
        "DISM", "YesNo", "Question"
    )
    if ($confirm -eq "Yes") { Invoke-DISMRepair -LogBox $txtLogSys -StatusLabel $txtStatusSys }
})

$btnResetUSB.Add_Click({
    $confirm = [System.Windows.MessageBox]::Show(
        "Reinitialiser tous les peripheriques USB ?`n(Peut deconnecter temporairement vos appareils)",
        "Reset USB", "YesNo", "Warning"
    )
    if ($confirm -eq "Yes") { Invoke-ResetUSBDevices -LogBox $txtLogSys -StatusLabel $txtStatusSys }
})

$btnExportReport.Add_Click({
    Export-Report -LogBox $txtLogSys -StatusLabel $txtStatusSys
})

$btnRefreshSys.Add_Click({
    Update-SysInfoBox -txtBox $txtSysInfo
    Set-Status $txtStatusSys "Informations actualisees"
})

# ============================================================
# 20. Lancement
# ============================================================
# ============================================================
# 20. Lancement (with improvement #6: nested ShowDialog exception guard)
# ============================================================
try {
    if (-not $window) {
        Write-Host "CRITICAL ERROR: Window object is null!" -ForegroundColor Red
        Write-Host "Script initialization failed." -ForegroundColor Red
        Read-Host "Press Enter to close"
        exit 1
    }
    
    Hide-ConsoleWindow
    
    try {
        $window.ShowDialog() | Out-Null
    }
    catch {
        $hwnd = [Win32]::GetConsoleWindow()
        if ($hwnd -ne [IntPtr]::Zero) { 
            [Win32]::ShowWindow($hwnd, 5) | Out-Null
        }
        Write-Host "CRITICAL ERROR: Window display failed" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Erreur lors de l'affichage de la fenetre: $_" "ERROR"
        Read-Host "Press Enter to close"
        exit 1
    }
}
catch {
    $hwnd = [Win32]::GetConsoleWindow()
    if ($hwnd -ne [IntPtr]::Zero) { 
        [Win32]::ShowWindow($hwnd, 5) | Out-Null
    }
    Write-Host "CRITICAL ERROR: Window launch failed" -ForegroundColor Red
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Erreur lors du lancement de la fenetre: $_" "ERROR"
    Read-Host "Press Enter to close"
    exit 1
}