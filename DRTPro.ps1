#Requires -Version 5.1
<#
.SYNOPSIS
    Device Repair Tool PRO v3.1 - Outil de gestion des pilotes.
.NOTES
	Auto-Signer avec certificat ou
	powershell -ExecutionPolicy Bypass -File "Chemindufichier\V2.ps1"
    Corrections appliquees :
    - C1 : Guard PSCommandPath null (boucle infinie)
    - C2 : DoEvents() WinForms remplace par Dispatcher WPF
    - C3 : Division par zero dans le scan
    - C4 : DriverDate null-safe
    - C5 : SFC/DISM non-bloquants via Runspace + timer
    - M1/M2 : Parsing pnputil locale-independant
    - M3 : Validation PublishedName avant suppression
    - M4/M5 : Where-Object Count null-safe avec @()
    - M6/M7 : Dead code TEMP_DIR et HISTORY_PATH supprimes
    - m1 : Calcul RAM corrige (KB -> GB)
    - m3 : UTF8 sans BOM
#>

# ============================================================
# 1. Verification admin + relance auto
# ============================================================
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $PSCommandPath) {
        [System.Windows.MessageBox]::Show(
            "Lancez ce script depuis un fichier .ps1, pas depuis la console.",
            "Erreur de lancement", "OK", "Error"
        )
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
# WinForms n'est plus charge (DoEvents supprime)

# ============================================================
# 4. Variables globales
# ============================================================
$Script:APP_VERSION  = "3.1"
$Script:APP_NAME     = "Device Repair Tool PRO"
$Script:DESKTOP_DIR  = [Environment]::GetFolderPath("Desktop")
$Script:Logs         = [System.Collections.Generic.List[object]]::new()
$Script:ProblemDevices    = [System.Collections.Generic.List[object]]::new()
$Script:DriverStoreDrivers= [System.Collections.Generic.List[object]]::new()
$Script:WindowsUpdates    = [System.Collections.Generic.List[object]]::new()
$Script:ScanDone     = $false

# ============================================================
# 5. Detection systeme
# ============================================================
try {
    $Script:OSInfo      = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    $Script:OSVersion   = [version]$Script:OSInfo.Version
    $Script:PSVersion   = $PSVersionTable.PSVersion
    $Script:CPU         = (Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1).Name
    # CORR m1 : TotalVisibleMemorySize est en KB -> diviser par 1024*1024 pour GB
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

# Dispatcher yield non-bloquant (remplace DoEvents WinForms - CORR C2)
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

function Invoke-SystemScan {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [System.Windows.Controls.ProgressBar]$ProgressBar
    )

    try {
        Set-Status $StatusLabel "Scan en cours..."
        if ($ProgressBar) { $ProgressBar.Value = 0 }

        Write-Log "Debut du scan des peripheriques" "INFO"

        # ✅ FIX 1 : init safe
        if (-not $Script:ProblemDevices) {
            $Script:ProblemDevices = New-Object System.Collections.ArrayList
        } else {
            $Script:ProblemDevices.Clear()
        }

        # ✅ FIX 2 : charger UNE fois (perf enorme)
        $allDrivers = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue

        $devices = @(Get-PnpDevice -ErrorAction SilentlyContinue)
        $totalDevices = $devices.Count
        $problemCount = 0
        $i = 0

        foreach ($device in $devices) {
            $i++

            if ($ProgressBar -and $totalDevices -gt 0) {
                $ProgressBar.Value = [math]::Round(($i / $totalDevices) * 100)
            }

            if ($device.Status -ne "OK") {

                # ✅ FIX 3 : utiliser cache drivers
                $driverInfo = $allDrivers | Where-Object {
                    $_.DeviceID -eq $device.InstanceId
                } | Select-Object -First 1

                $driverDateStr = if ($driverInfo -and $driverInfo.DriverDate) {
                    $driverInfo.DriverDate.ToString("yyyy-MM-dd")
                } else { "-" }

                # ✅ FIX 4 : ErrorCode safe
                $errorCode = if ($device.PSObject.Properties["ConfigManagerErrorCode"]) {
                    $device.ConfigManagerErrorCode
                } else { "-" }

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

                [void]$Script:ProblemDevices.Add($deviceInfo)
                $problemCount++

                Write-Log "Peripherique en erreur: $($device.FriendlyName) (Code: $errorCode)" "WARNING"
            }
        }

        Write-Log "Scan termine: $problemCount peripherique(s) en erreur sur $totalDevices" "SUCCESS"
        Set-Status $StatusLabel "$problemCount peripherique(s) en erreur"

        if ($DataGrid) {
            $DataGrid.ItemsSource = $null
            $DataGrid.ItemsSource = $Script:ProblemDevices
        }

        if ($ProgressBar) { $ProgressBar.Value = 100 }

        $Script:ScanDone = $true
    }
    catch {
        Write-Log "Erreur lors du scan: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
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

        # CORR M1/M2 : parsing locale-independant
        # On detecte les blocs par la presence de "oem" ou "*.inf" en debut de valeur
        # apres le premier ":" de chaque ligne, sans valider le label en langue
        # $fieldIndex = 0  # compteur de champs dans le bloc courant
        # Ordre attendu des champs pnputil : PublishedName, OriginalName, Provider, Class, Date, Version, SignerName, InfFile
        # On utilise une heuristique : si la valeur ressemble a oemXX.inf -> debut de bloc

        foreach ($line in $lines) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            # Separator de champ : "Libelle : valeur" ou "Libelle    : valeur"
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
            # FR: "nom du fournisseur" / EN: "provider name"
            if ($label -match "fournisseur|provider") {
                $currentDriver.Provider = $value
            }
            # FR: "version du pilote" contient date ET version sur la meme ligne
            # ex: "03/30/2023 31.0.12027.9001"
            elseif ($label -match "version") {
                $parts = $value -split '\s+', 2
                if ($parts.Count -eq 2) {
                    $currentDriver.Date    = $parts[0]
                    $currentDriver.Version = $parts[1]
                } else {
                    $currentDriver.Version = $value
                }
            }
            # FR: "nom de la classe" / EN: "class name" — on cherche "classe" ou "class" n'importe ou dans le label
            elseif ($label -match "classe|class") {
                $currentDriver.Class = $value
            }
            # FR: "nom d'origine" / EN: "original name"
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
        # CORR M3 : validation du nom avant suppression
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
            Invoke-DispatcherYield  # CORR C2
        }
    }

    Write-Log "Suppression terminee: $success succes, $failed echecs" "INFO"
    Set-Status $StatusLabel "Suppression terminee: $success/$($Drivers.Count) pilotes"
}

# ============================================================
# 10. Fonction Reparation peripheriques — asynchrone WPF complete
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

    # Confirmation avant toute action
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

    # Timer WPF pour traiter un peripherique toutes les 500ms
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
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
            # 1️⃣ Reinitialisation du peripherique
            Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Start-Sleep -Milliseconds 500
            Enable-PnpDevice  -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
            Write-Log "  -> Peripherique reinitialise" "SUCCESS"

            # 2️⃣ Reinstallation du driver si InfFile disponible
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

        # Mise a jour log et UI
        if ($LogBox) {
            $LogBox.Text = Get-FormattedLogs
            $LogBox.ScrollToEnd()
            Invoke-DispatcherYield  # UI responsive
        }

        Write-Log "=== Fin reparation de $($device.Name) ===" "INFO"
    })

    $timer.Start()
}

# ============================================================
# 11. Fonction Windows Update (pilotes)
# ============================================================

function Invoke-WindowsUpdateSearch {
    param(
        [System.Windows.Controls.DataGrid]$DataGrid,
        [System.Windows.Controls.TextBlock]$StatusLabel,
        [System.Windows.Controls.TextBox]$LogBox
    )

    try {
        Set-Status $StatusLabel "Recherche des mises a jour..."
        Write-Log "Recherche Windows Update..." "INFO"

        # ✅ init safe
        if (-not $Script:WindowsUpdates) {
            $Script:WindowsUpdates = New-Object System.Collections.ArrayList
        } else {
            $Script:WindowsUpdates.Clear()
        }

        $session  = New-Object -ComObject "Microsoft.Update.Session"
        $searcher = $session.CreateUpdateSearcher()

        # ✅ FIX IMPORTANT
        $result = $searcher.Search("IsInstalled=0")

        foreach ($update in $result.Updates) {

            # ✅ filtre drivers ici
            if ($update.Type -ne 2) { continue }

            $Script:WindowsUpdates.Add([PSCustomObject]@{
                Title       = $update.Title
                Description = $update.Description
                Size_MB     = [math]::Round($update.MaxDownloadSize / 1MB, 1)
                KBArticle   = if ($update.KBArticleIDs.Count -gt 0) { $update.KBArticleIDs[0] } else { "-" }
            }) | Out-Null
        }

        if ($DataGrid) {
            $DataGrid.ItemsSource = $null
            $DataGrid.ItemsSource = $Script:WindowsUpdates
        }

        Set-Status $StatusLabel "$($Script:WindowsUpdates.Count) mise(s) a jour disponible(s)"

        if ($LogBox) {
            $LogBox.Text = Get-FormattedLogs
            $LogBox.ScrollToEnd()
        }
    }
    catch {
        Write-Log "Erreur Windows Update: $_" "ERROR"
        Set-Status $StatusLabel "Erreur: $_"
    }
}

# ========================================
# Start-BackgroundProcess pour SFC/DISM
# Pattern : Runspace + ConcurrentQueue + DispatcherTimer
# Evite les problemes de closure des events .NET en PowerShell
# ========================================
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

        # Runspace isole qui execute le process et pousse les lignes dans la queue
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()

        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs

        # On passe les variables au Runspace via InitialSessionState n'est pas dispo
        # -> on utilise AddScript avec parametres
        $ps.AddScript({
            param($exe, $exeArgs, $q)
            try {
                $isSFC = ($exe -match 'sfc' -or $exeArgs -match 'scannow')
                $proc = $null

                if ($isSFC) {
                    # SFC needs a console environment, but we don't want to see it.
                    # The most reliable method is to launch it via a new, hidden PowerShell process.
                    # This new process will run SFC and then pass its exit code back to us.
                    $psi = [System.Diagnostics.ProcessStartInfo]::new()
                    $psi.FileName = "powershell.exe"
                    # The command to execute in the hidden window: run sfc, wait for it, and then exit with sfc's exit code.
                    $psi.Arguments = "-NoProfile -WindowStyle Hidden -Command `"sfc.exe /scannow; exit `$LASTEXITCODE`""
                    $psi.UseShellExecute = $false
                    $psi.CreateNoWindow = $true

                    $proc = [System.Diagnostics.Process]::Start($psi)
                    $q.Enqueue("[INFO] Lancement de SFC dans un processus PowerShell cache (PID: $($proc.Id))...")
                    $proc.WaitForExit() # This waits for the hidden powershell.exe to finish

                    # Now, $proc.ExitCode holds the REAL exit code from sfc.exe.
                    # We can proceed with reading the log file.
                    $cbsLog = "$env:windir\Logs\CBS\CBS.log"
                    if (Test-Path $cbsLog) {
                        $lines = Get-Content $cbsLog -Encoding Unicode -ErrorAction SilentlyContinue
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

                    # Interpretation of SFC's exit code
                    switch ($proc.ExitCode) {
                        0 { $q.Enqueue("[SFC_OK] Aucune violation d'integrite detectee.") }
                        1 { $q.Enqueue("[SFC_FIXED] Fichiers corrompus detectes et repares avec succes.") }
                        2 { $q.Enqueue("[SFC_FAIL] Fichiers corrompus detectes mais NON repares. Lancez DISM puis relancez SFC.") }
                        default { $q.Enqueue("[SFC_UNK] Code de sortie inconnu : $($proc.ExitCode)") }
                    }
                } else {
                    # Original, working DISM logic
                    $psi = [System.Diagnostics.ProcessStartInfo]::new()
                    $psi.FileName               = $exe
                    $psi.Arguments              = $exeArgs
                    $psi.UseShellExecute        = $false
                    $psi.CreateNoWindow         = $true
                    $psi.RedirectStandardOutput = $true
                    $psi.RedirectStandardError  = $true

                    $proc = [System.Diagnostics.Process]::Start($psi)
                    $q.Enqueue("[INFO] PID $($proc.Id) lance")

                    # Read DISM output
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

                # Finally, signal completion with the exit code
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
        $timer = [System.Windows.Threading.DispatcherTimer]::new()
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

                    # Popup de fin
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

# ========================================
# Invoke-SFCScan
# ========================================
function Invoke-SFCScan {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    Start-BackgroundProcess -Exe "sfc.exe" -ExeArgs "/scannow" -LogBox $LogBox -StatusLabel $StatusLabel -Label "SFC"
}

# ========================================
# Invoke-DISMRepair
# ========================================
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

        $usbDevices = @(Get-PnpDevice -Class USB -ErrorAction SilentlyContinue)
        $count = 0

        foreach ($device in $usbDevices) {
            try {
                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                Start-Sleep -Milliseconds 200
                Enable-PnpDevice  -InstanceId $device.InstanceId -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
                Write-Log "  -> Reset: $($device.FriendlyName)" "SUCCESS"
                $count++
            }
            catch {
                Write-Log "  -> Echec: $($device.FriendlyName)" "WARNING"
            }
            # Mise a jour log en temps reel
            if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
            Set-Status $StatusLabel "Reset USB... ($count/$($usbDevices.Count))"
            Invoke-DispatcherYield
        }

        Write-Log "Reset USB termine: $count peripherique(s) traite(s)" "SUCCESS"
        Set-Status $StatusLabel "Reset USB termine ($count peripheriques)"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
    }
    catch {
        Write-Log "Erreur lors du reset USB: $($_.Exception.Message)" "ERROR"
        Set-Status $StatusLabel "Erreur reset USB"
    }
}

# ============================================================
# 14. Fonction Export Rapport
# ============================================================
function Export-Report {
    param(
        [System.Windows.Controls.TextBox]$LogBox,
        [System.Windows.Controls.TextBlock]$StatusLabel
    )
    try {
        $reportPath = "$Script:DESKTOP_DIR\DeviceRepairTool_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $sysInfo    = Get-SystemInfo

        $sb = [System.Text.StringBuilder]::new()
        $null = $sb.AppendLine("========================================")
        $null = $sb.AppendLine("RAPPORT DE DIAGNOSTIC - $Script:APP_NAME v$Script:APP_VERSION")
        $null = $sb.AppendLine("========================================")
        $null = $sb.AppendLine("Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("INFORMATIONS SYSTEME")
        $null = $sb.AppendLine("---------------------")
        $null = $sb.AppendLine("OS:           $($sysInfo.OS)")
        $null = $sb.AppendLine("Architecture: $($sysInfo.Architecture)")
        $null = $sb.AppendLine("CPU:          $($sysInfo.CPU)")
        $null = $sb.AppendLine("RAM:          $($sysInfo.RAM_GB) GB")
        $null = $sb.AppendLine("Uptime:       $($sysInfo.Uptime)")
        $null = $sb.AppendLine("Dernier boot: $($sysInfo.LastBoot)")
        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("PERIPHERIQUES EN ERREUR ($($Script:ProblemDevices.Count))")
        $null = $sb.AppendLine("-----------------------------------------------------")

        foreach ($device in $Script:ProblemDevices) {
            $null = $sb.AppendLine($device.Name)
            $null = $sb.AppendLine("  Classe:      $($device.Class)")
            $null = $sb.AppendLine("  Statut:      $($device.StatusCode)")
            $null = $sb.AppendLine("  Code:        $($device.ErrorCode)")
            $null = $sb.AppendLine("  Version:     $($device.DriverVersion)")
            $null = $sb.AppendLine("  Fournisseur: $($device.Provider)")
            $null = $sb.AppendLine("  --------------------------------")
        }

        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("PILOTES DANS LE DRIVER STORE ($($Script:DriverStoreDrivers.Count))")
        $null = $sb.AppendLine("---------------------------------------------------------------")

        foreach ($driver in $Script:DriverStoreDrivers) {
            $null = $sb.AppendLine($driver.PublishedName)
            $null = $sb.AppendLine("  Fournisseur: $($driver.Provider)")
            $null = $sb.AppendLine("  Version:     $($driver.Version)")
            $null = $sb.AppendLine("  Date:        $($driver.Date)")
            $null = $sb.AppendLine("  Classe:      $($driver.Class)")
            $null = $sb.AppendLine("  Fichier INF: $($driver.InfFile)")
            $null = $sb.AppendLine("  --------------------------------")
        }

        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("JOURNAL COMPLET")
        $null = $sb.AppendLine("---------------")
        $null = $sb.AppendLine($(Get-FormattedLogs))
        $null = $sb.AppendLine("")
        $null = $sb.AppendLine("========================================")
        $null = $sb.AppendLine("Fin du rapport - Genere le $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        $null = $sb.AppendLine("========================================")

        # CORR m3 : UTF8 sans BOM
        [System.IO.File]::WriteAllText($reportPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

        Write-Log "Rapport exporte: $reportPath" "SUCCESS"
        Set-Status $StatusLabel "Rapport exporte sur le Bureau"
        if ($LogBox) { $LogBox.Text = Get-FormattedLogs; $LogBox.ScrollToEnd() }
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
        Title="Device Repair Tool PRO v3.1"
        Height="750" Width="1100" WindowStartupLocation="CenterScreen"
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
                <TextBlock x:Name="txtAppTitle" Text="Device Repair Tool PRO v3.1" FontSize="18" FontWeight="Bold" Foreground="#4EC9B0"/>
                <TextBlock Text="Gestionnaire de pilotes - Diagnostic et reparation" FontSize="11" Foreground="#666666" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>

        <TabControl x:Name="MainTabControl" Grid.Row="1" Background="#1E1E1E" BorderBrush="#3E3E42" BorderThickness="1">

            <!-- Onglet 1 : Scan peripheriques -->
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
                        <Button x:Name="btnScan" Content="Scanner maintenant" Background="#0E639C" Width="150"/>
                        <Button x:Name="btnRepairSel" Content="Reparer selection" Background="#CE9178" Width="130" IsEnabled="False"/>
                        <Button x:Name="btnRepairAll" Content="Reparer tout" Background="#D16969" Width="110" IsEnabled="False"/>
                        <Button x:Name="btnSelectAll" Content="Tout selectionner" Background="#3E3E42" Width="130"/>
                        <Button x:Name="btnSearchCatalog" Content="Rechercher dans le catalogue" Background="#3e80ac" Width="210" Margin="5,0,0,0"/>
                    </StackPanel>

                    <ProgressBar x:Name="progressBar" Grid.Row="1" Margin="0,0,0,8"/>

                    <DataGrid x:Name="dgDevices" Grid.Row="2" AutoGenerateColumns="False"
                              SelectionMode="Extended"
                              SelectionUnit="FullRow" 
                              CanUserAddRows="False"
                              CanUserDeleteRows="False">
                
                    <DataGrid.Columns>
                        <DataGridCheckBoxColumn Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Header=" " Width="36"/>
                        <DataGridTemplateColumn Header="Peripherique" Width="*">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding Name}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Classe" Width="100">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding Class}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Statut" Width="90">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding StatusCode}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Code" Width="55">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding ErrorCode}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="ID Materiel" Width="*">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding InstanceId}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Version" Width="110">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding DriverVersion}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Date" Width="90">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding DriverDate}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                        <DataGridTemplateColumn Header="Fournisseur" Width="120">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <TextBox Text="{Binding Provider}" IsReadOnly="True" BorderThickness="0" Background="Transparent" Foreground="White" Padding="2"/>
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

            <!-- Onglet 2 : Driver Store -->
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
                        <Button x:Name="btnScanStore" Content="Analyser le Driver Store" Background="#0E639C" Width="180"/>
                        <Button x:Name="btnRemoveStore" Content="Supprimer selection" Background="#D16969" Width="150" IsEnabled="False"/>
                    </StackPanel>

                    <DataGrid x:Name="dgDriverStore" Grid.Row="1" AutoGenerateColumns="False">
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Binding="{Binding Selected, Mode=TwoWay}" Header=" " Width="36"/>
                            <DataGridTextColumn Binding="{Binding PublishedName}" Header="Nom publié" Width="90" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding OriginalName}" Header="Nom d'origine" Width="150" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Provider}" Header="Fournisseur" Width="*" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Class}" Header="Classe" Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Date}" Header="Date" Width="90" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Version}" Header="Version" Width="150" IsReadOnly="True"/>
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

            <!-- Onglet 3 : Windows Update -->
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
                        <Button x:Name="btnWUSearch" Content="Rechercher les mises a jour" Background="#0E639C" Width="210" Margin="0,0,5,0"/>
                        <Button x:Name="btnWUOpenCatalog" Content="Ouvrir Catalogue Windows" Background="#0E639C" Width="180"/>
                    </StackPanel>

                    <DataGrid x:Name="dgWUUpdates" Grid.Row="1" AutoGenerateColumns="False" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Binding="{Binding Title}" Header="Titre" Width="*" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Size_MB}" Header="Taille MB" Width="80" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding KBArticle}" Header="KB" Width="80" IsReadOnly="True"/>
                            <DataGridTextColumn Binding="{Binding Description}" Header="Description" Width="*" IsReadOnly="True"/>
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

            <!-- Onglet 4 : Outils Systeme -->
            <TabItem Header="Outils Systeme">
                <Grid Margin="10" Background="#1E1E1E">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="5"/>
                        <RowDefinition Height="180" MinHeight="60"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Boutons -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="btnSFC" Content="SFC /scannow" Background="#CE9178" Width="120"/>
                        <Button x:Name="btnDISM" Content="DISM /RestoreHealth" Background="#CE9178" Width="160"/>
                        <Button x:Name="btnResetUSB" Content="Reset USB" Background="#CE9178" Width="100"/>
                        <Button x:Name="btnExportReport" Content="Exporter rapport" Background="#0E639C" Width="130"/>
                    </StackPanel>

                    <!-- Zone infos systeme + mini tuto -->
                    <TextBox x:Name="txtSysInfo" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" FontSize="11"
                             Background="#0D0D0D" Foreground="#D4D4D4"/>

                    <!-- GridSplitter : glisser pour redimensionner le log -->
                    <GridSplitter Grid.Row="2" Height="5" HorizontalAlignment="Stretch"
                                  Background="#3E3E42" Cursor="SizeNS" ResizeDirection="Rows"
                                  ResizeBehavior="PreviousAndNext" Margin="0,1"/>

                    <!-- Log redimensionnable -->
                    <TextBox x:Name="txtLogSys" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,2,0,0"
                             Background="#0D0D0D" Foreground="#9CDCFE"/>

                    <!-- Barre de statut -->
                    <TextBlock x:Name="txtStatusSys" Grid.Row="4" FontSize="11" Foreground="#888888" Margin="0,4,0,0"/>
                </Grid>
            </TabItem>

        </TabControl>

        <TextBlock x:Name="txtFooter" Grid.Row="2" FontSize="9" Foreground="#444444" Margin="0,6,0,0" HorizontalAlignment="Center"
                   Text="Device Repair Tool PRO v3.1 | Gestionnaire de pilotes"/>
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
$btnScan        = $window.FindName("btnScan")
$btnRepairSel   = $window.FindName("btnRepairSel")
$btnRepairAll   = $window.FindName("btnRepairAll")
$btnSelectAll   = $window.FindName("btnSelectAll")
$dgDevices      = $window.FindName("dgDevices")
$progressBar    = $window.FindName("progressBar")
$txtStatusScan  = $window.FindName("txtStatusScan")
$txtLogScan     = $window.FindName("txtLogScan")

$btnScanStore   = $window.FindName("btnScanStore")
$btnRemoveStore = $window.FindName("btnRemoveStore")
$dgDriverStore  = $window.FindName("dgDriverStore")
$txtStatusStore = $window.FindName("txtStatusStore")
$txtLogStore    = $window.FindName("txtLogStore")

$btnWUSearch    = $window.FindName("btnWUSearch")
$dgWUUpdates    = $window.FindName("dgWUUpdates")
$txtStatusWU    = $window.FindName("txtStatusWU")
$txtLogWU       = $window.FindName("txtLogWU")

$btnSFC         = $window.FindName("btnSFC")
$btnDISM        = $window.FindName("btnDISM")
$btnResetUSB    = $window.FindName("btnResetUSB")
$btnExportReport= $window.FindName("btnExportReport")
$txtSysInfo     = $window.FindName("txtSysInfo")
$txtStatusSys   = $window.FindName("txtStatusSys")
$txtLogSys      = $window.FindName("txtLogSys")

$btnWUOpenCatalog = $window.FindName("btnWUOpenCatalog")
$btnSearchCatalog = $window.FindName("btnSearchCatalog")



# ============================================================
# 18. Initialisation infos systeme + mini tuto
# ============================================================
$sysInfo = Get-SystemInfo
$sb = [System.Text.StringBuilder]::new()

# --- SYSTEME ---
$null = $sb.AppendLine("SYSTEME")
$null = $sb.AppendLine("-------")
$null = $sb.AppendLine("OS:           $($sysInfo.OS)")
$null = $sb.AppendLine("Architecture: $($sysInfo.Architecture)")
$null = $sb.AppendLine("CPU:          $($sysInfo.CPU)")
$null = $sb.AppendLine("RAM:          $($sysInfo.RAM_GB) GB")
$null = $sb.AppendLine("Uptime:       $($sysInfo.Uptime)")
$null = $sb.AppendLine("Dernier boot: $($sysInfo.LastBoot)")

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

# --- DISQUES ---
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

# --- MINI TUTO ---
$null = $sb.AppendLine("")
$null = $sb.AppendLine("================================================================")
$null = $sb.AppendLine("GUIDE DES OUTILS - A QUOI CA SERT ?")
$null = $sb.AppendLine("================================================================")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("[ SFC /scannow ]  (Verificateur de fichiers systeme)")
$null = $sb.AppendLine("  Analyse et repare les fichiers Windows corrompus ou manquants.")
$null = $sb.AppendLine("  Duree : 5 a 15 minutes. Ne pas eteindre le PC pendant l'analyse.")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("  Resultats possibles :")
$null = $sb.AppendLine("  [OK] Code 0 -> Aucun probleme detecte. Votre systeme est sain.")
$null = $sb.AppendLine("  [OK] Code 1 -> Des fichiers corrompus ont ete trouves ET repares automatiquement.")
$null = $sb.AppendLine("  [!!] Code 2 -> Fichiers corrompus IMPOSSIBLES a reparer par SFC seul.")
$null = $sb.AppendLine("                 Solution : lancer DISM d'abord, puis relancer SFC.")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("[ DISM /RestoreHealth ]  (Reparation de l'image systeme)")
$null = $sb.AppendLine("  Repare l'image Windows elle-meme (la 'source' dont SFC a besoin).")
$null = $sb.AppendLine("  Necessite une connexion Internet (telecharge les fichiers manquants).")
$null = $sb.AppendLine("  Duree : 10 a 30 minutes selon votre connexion.")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("  Ordre recommande en cas de probleme grave :")
$null = $sb.AppendLine("  1. Lancer DISM  -->  2. Lancer SFC  -->  3. Redemarrer")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("  Resultats possibles :")
$null = $sb.AppendLine("  [OK] Code 0 -> Image reparee avec succes.")
$null = $sb.AppendLine("  [!!] Autre  -> Erreur. Verifiez votre connexion Internet et relancez.")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("[ Reset USB ]")
$null = $sb.AppendLine("  Reinitialise tous les peripheriques USB (clavier, souris, cles USB...).")
$null = $sb.AppendLine("  Utile si un peripherique USB n'est plus reconnu sans raison apparente.")
$null = $sb.AppendLine("  Vos appareils se deconnectent 1 seconde puis se reconnectent.")
$null = $sb.AppendLine("  Conseil : ne pas faire pendant un transfert de fichiers USB.")
$null = $sb.AppendLine("")
$null = $sb.AppendLine("[ Exporter rapport ]")
$null = $sb.AppendLine("  Genere un fichier .txt sur le Bureau avec toutes les infos du diagnostic.")
$null = $sb.AppendLine("  Pratique pour envoyer a un technicien ou garder une trace.")

$txtSysInfo.Text = $sb.ToString()


# ============================================================
# 19. Gestionnaires d'evenements
# ============================================================

$btnScan.Add_Click({
    $btnScan.IsEnabled = $false
    try {
        Invoke-SystemScan -DataGrid $dgDevices -StatusLabel $txtStatusScan -ProgressBar $progressBar
        $btnRepairSel.IsEnabled = ($Script:ProblemDevices.Count -gt 0)
        $btnRepairAll.IsEnabled = ($Script:ProblemDevices.Count -gt 0)
        $txtLogScan.Text = Get-FormattedLogs
    }
    finally {
        $btnScan.IsEnabled = $true
    }
})

$btnSelectAll.Add_Click({
    if ($Script:ProblemDevices.Count -eq 0) { return }
    # CORR M5 : @() pour eviter $null.Count
    $allChecked = (@($Script:ProblemDevices | Where-Object { -not $_.Selected }).Count -eq 0)
    foreach ($device in $Script:ProblemDevices) {
        $device.Selected = -not $allChecked
    }
    $dgDevices.Items.Refresh()
})

$btnRepairSel.Add_Click({
    # CORR M4 : @() pour eviter $null.Count
    $selectedDevices = @($Script:ProblemDevices | Where-Object { $_.Selected })
    if ($selectedDevices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun peripherique selectionne", "Information", "OK", "Information")
        return
    }
    Invoke-RepairDevices -Devices $selectedDevices -LogBox $txtLogScan -StatusLabel $txtStatusScan
    $txtLogScan.Text = Get-FormattedLogs
})

$btnRepairAll.Add_Click({
    if ($Script:ProblemDevices.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun peripherique en erreur detecte", "Information", "OK", "Information")
        return
    }
    $confirm = [System.Windows.MessageBox]::Show(
        "Reparer TOUS les $($Script:ProblemDevices.Count) peripheriques en erreur ?",
        "Confirmation", "YesNo", "Warning"
    )
    if ($confirm -eq "Yes") {
        Invoke-RepairDevices -Devices $Script:ProblemDevices -LogBox $txtLogScan -StatusLabel $txtStatusScan
        $txtLogScan.Text = Get-FormattedLogs
    }
})

$btnScanStore.Add_Click({
    Invoke-DriverStoreScan -DataGrid $dgDriverStore -StatusLabel $txtStatusStore
    $btnRemoveStore.IsEnabled = ($Script:DriverStoreDrivers.Count -gt 0)
    $txtLogStore.Text = Get-FormattedLogs
})

$btnRemoveStore.Add_Click({
    # CORR M4
    $selectedDrivers = @($Script:DriverStoreDrivers | Where-Object { $_.Selected })
    if ($selectedDrivers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Aucun pilote selectionne", "Information", "OK", "Information")
        return
    }
    Invoke-RemoveDrivers -Drivers $selectedDrivers -LogBox $txtLogStore -StatusLabel $txtStatusStore
    Invoke-DriverStoreScan -DataGrid $dgDriverStore -StatusLabel $txtStatusStore
})

$btnWUSearch.Add_Click({
    Invoke-WindowsUpdateSearch -DataGrid $dgWUUpdates -StatusLabel $txtStatusWU -LogBox $txtLogWU
})

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

# ============================================================
# Bouton "Ouvrir Catalogue Windows"
# ============================================================

$btnWUOpenCatalog.Add_Click({
    try {
        Start-Process "https://catalog.update.microsoft.com/home.aspx"
        Write-Log "Catalogue Windows ouvert dans le navigateur par defaut." "INFO"
    }
    catch {
        Write-Log "Impossible d'ouvrir le catalogue Windows : $($_.Exception.Message)" "ERROR"
    }
})

# ============================================================
# Double-clic pour copier l'InstanceId
# ============================================================
# $dgDevices.Add_MouseDoubleClick({
#     $cell = $dgDevices.CurrentCell
#     if ($cell.Item -and $cell.Column.Header -eq "ID Materiel") {
#         [System.Windows.Clipboard]::SetText($cell.Item.InstanceId)
#         [System.Windows.MessageBox]::Show("InstanceId copie : $($cell.Item.InstanceId)")
#     }
# })

# ============================================================
# Checkbox indépendante du highlight bleu — pas de SelectionChanged
# Les checkboxes fonctionnent via le binding TwoWay directement.
# Le bouton "Réparer sélection" est activé dès qu'un scan trouve des erreurs.
# ============================================================

# ============================================================
# Double-clic : Copie UNIQUE (sans interférer avec la coche)
# ============================================================
$dgDevices.Add_MouseDoubleClick({
    # On récupère l'item sous la souris
    $selectedItem = $dgDevices.SelectedItem
    if ($selectedItem -and $selectedItem.InstanceId) {
        # Copie SANS message box (pour éviter de perdre le focus)
        [System.Windows.Clipboard]::SetText($selectedItem.InstanceId)
        
        # On informe l'utilisateur discrètement dans la barre de statut
        if ($txtStatusScan) { 
            $txtStatusScan.Text = "ID Copié : $($selectedItem.InstanceId)" 
        }
    }
})

# ============================================================
# Bouton Recherche (avec sécurité anti-erreur Null)
# ============================================================
if ($btnSearchCatalog) {
    $btnSearchCatalog.Add_Click({
        $selectedDevice = $dgDevices.SelectedItem
        if (-not $selectedDevice) {
            [System.Windows.MessageBox]::Show("Sélectionnez un périphérique d'abord.", "Info")
            return
        }

        $instanceId = $selectedDevice.InstanceId
        $searchTerm = if ($instanceId -match "VEN_([^&]+)&PROD_([^\\]+)") {
            "$($matches[1]) $($matches[2])"
        } else {
            $instanceId.Split('\')[-1]
        }

        if ($searchTerm) {
            $url = "https://catalog.update.microsoft.com/Search.aspx?q=$($searchTerm -replace ' ', '+')"
            Start-Process $url
        }
    })
}


# ============================================================
# 20. Lancement
# ============================================================
$window.ShowDialog() | Out-Nu
