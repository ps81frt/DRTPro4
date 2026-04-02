# Device Repair Tool PRO — v4.1

A single-file PowerShell WPF application for diagnosing and repairing Windows driver and device issues. No installation, no dependencies beyond what ships with Windows.

**Repository:** https://github.com/ps81frt/driverRep  
**Latest release:** [v4.1](https://github.com/ps81frt/driverRep/releases/tag/4.1)

---

## Requirements

| Requirement | Minimum |
|---|---|
| PowerShell | 5.1 |
| OS | Windows 10 / Windows 11 |
| Privileges | Administrator (auto-escalation built in) |
| .NET | WPF stack (.NET Framework 4.x — included with Windows) |

No external modules, no NuGet packages, no internet access required (except for the Windows Update scan tab, which uses the local Windows Update Agent COM object).

---

## How to run

```powershell
powershell.exe -ExecutionPolicy Bypass -File DRTPro4.ps1
```

If the script is not launched as administrator it silently re-launches itself elevated via `Start-Process -Verb RunAs`. If it is launched directly from a console (not from a `.ps1` file path), it shows an error dialog and exits — this is intentional because WPF requires a proper script context.

The console window is hidden immediately after elevation using `kernel32!GetConsoleWindow` + `user32!ShowWindow(0)` so the user only ever sees the WPF window.

---

## Architecture

The script is structured into numbered sections:

1. Admin check and auto-elevation
2. Console window suppression via P/Invoke
3. WPF assembly loading (`PresentationFramework`, `PresentationCore`, `WindowsBase`)
4. Global script-scoped state (`$Script:` variables, typed generic lists)
5. CIM-based system detection at startup (OS version, CPU, RAM)
6. Utility functions (logging, HTML escaping, WPF dispatcher yield)
7–14. Feature functions (one per tab)
15–19. XAML UI definition and control binding
20. Window launch with nested exception guards

### Threading model

All long-running operations use the same pattern:

```
DispatcherTimer (UI thread, 200 ms tick)
    ↕ ConcurrentQueue<string>
PowerShell Runspace (background thread)
```

The runspace does the heavy work, enqueues tagged string messages (`[INIT]`, `[DEVICE]`, `[PROGRESS]`, `[DONE]`, `[ERR]`). The timer drains the queue on the UI thread, updates controls, and disposes the runspace when `[DONE]` is received. This avoids `Invoke`/`BeginInvoke` marshalling and `DoEvents` hacks.

- **Device scan** and **SFC**: STA runspace (PnP COM apartment requirement)
- **Windows Update search**: MTA runspace (Windows Update Agent COM requirement)
- **DISM**: real-time stdout line streaming via `RedirectStandardOutput`
- **SFC**: launched in a hidden child PowerShell process, then CBS.log is parsed after exit

Closure variable capture inside `.NET` event handlers is explicitly handled — `$Script:` scope is unreliable inside `DispatcherTimer.Tick` closures, so list references are captured into local variables before the closure is created.

---

## Tabs

### Tab 1 — Device Scan

Queries all PnP devices via `Get-PnpDevice` and cross-references against `Win32_PnPSignedDriver` using a pre-built hashtable keyed on `DeviceID` (avoids repeated `Where-Object` filtering, one pass).

Displayed per device:
- Friendly name, device class
- PnP status (`OK`, `Error`, `Degraded`, `Unknown`)
- ConfigManagerErrorCode
- InstanceId
- Driver version, driver date, provider name

Status bar shows total device count with a breakdown: `N erreur(s) / N inconnu(s) / N OK`.

After scan, "Repair selected" and "Repair all errors" buttons are enabled only if error-status devices exist.

**Repair logic** (per device, via DispatcherTimer at 500 ms intervals):
1. `Disable-PnpDevice` → 500 ms wait → `Enable-PnpDevice`
2. If the device has an associated INF file: `pnputil /delete-driver <inf> /force` then `pnputil /add-driver <inf> /install`

**Catalog search:** populates a search term from the selected device (friendly name or class), opens `catalog.update.microsoft.com/Search.aspx?q=<term>` in the default browser.

![drt 1](./images/drt1.png)

---

### Tab 2 — Driver Store

Calls `pnputil /enum-drivers` and parses the output in a locale-independent way: block boundaries are detected by the `oem\d+.inf` value pattern rather than by label text. Field mapping uses regex against the label substring to work across English and French Windows installs (e.g. `fournisseur|provider`, `version`, `classe|class`).

Displayed per driver package:
- Published name (`oem0.inf` … `oem999.inf`)
- Original INF name
- Provider
- Version
- Date
- Class

Driver removal calls `pnputil /delete-driver <name> /force` only after validating that the name matches `^oem\d+\.inf$`. The store is re-scanned automatically after removal.

![drt 2](./images/drt2.png)
---

### Tab 3 — Event Log

Collects driver-related Windows events from three sources:

| Source | Filter |
|---|---|
| `System` | Level 1/2/3, provider name matches a 30-entry driver source regex |
| `Microsoft-Windows-Kernel-PnP/Configuration` | Level 1/2/3 |
| `Application` | Level 1/2, Event IDs 1000/1001/1002, provider matches `wer\|fault\|crash` |

The provider pattern covers: kernel-pnp, disk, volmgr, nvlddmkm, amdkmpfd, atikmdag, usbhub, USBXHCI, Netwtw, HidUsb, HDAudBus, storahci, stornvme, Wacom, Synaptics, and ~20 others.

A set of Event IDs is force-included regardless of provider: `9, 11, 15, 20, 43, 51, 129, 153, 157, 219, 411, 1001, 6008, 7026, 7031, 7034, 7045`.

Events are deduplicated by `(EventId | TimeCreated.Ticks | ProviderName)` using a `HashSet<string>`, sorted descending, capped at 500 entries.

**Built-in knowledge base** (17 Event IDs):

| ID | Description | Default advice |
|---|---|---|
| 9 | Device timeout | Check physical connection |
| 11 | Disk controller error | Test S.M.A.R.T., check SATA/NVMe cables |
| 15 | Device not ready | USB reset or reinstall driver |
| 20 | Driver installation failed | Check OS/architecture compatibility |
| 43 | USB problem | USB reset or disable selective suspend |
| 51 | Disk paging error | Run `chkdsk /r` |
| 129 | Controller reset | Update firmware/storage controller driver |
| 153 | I/O delay | Drive may be failing — clone and replace |
| 157 | Disk abnormal ejection | Check power/cables |
| 219 | Driver not loaded | Reinstall or update, check signature |
| 411 | Device install failure | Check Driver Store, re-run scan |
| 1001 | WER crash report | Analyze dump in `%localappdata%\CrashDumps` |
| 6008 | Unexpected shutdown | Check temps, RAM (memtest), PSU |
| 7026 | Boot driver failed | Check failing services via msconfig |
| 7031 | Service auto-restarted | Analyze Application log |
| 7034 | Service stopped unexpectedly | Check crashes, update component |
| 7045 | New service installed | Verify origin (possible malware) |

**Cross-correlation with Tab 1:** if a scan has been run, each event's full message text is checked against the InstanceId fragments and friendly names of all devices that were in error state. Matching devices are shown in a `DeviceHint` column.

Scan range: 7 days or 30 days (two buttons).  
Double-click on any row opens a popup with the full event message.  
"Copy" button puts the full event detail to the clipboard.

![drt 3](./images/drt3.png)
---

### Tab 4 — Windows Update (Driver Updates)

Uses `Microsoft.Update.Session` COM object (requires MTA apartment) to query pending updates filtered to `Type = 2` (drivers only). Runs in a background MTA runspace.

Displayed per available update:
- Title
- Description
- Size (MB)
- KB article ID

The tab provides buttons to open the Microsoft Update Catalog and `store.rg-adguard.net` in the browser.

---

### Tab 5 — System Tools

**SFC /scannow**  
Launched in a hidden child `powershell.exe` process to get a real exit code (SFC requires a console environment). After the process exits, `%windir%\Logs\CBS\CBS.log` is read (Unicode encoding), the last 100 lines matching `[SR]` are extracted, and the exit code is interpreted:

| Exit code | Meaning |
|---|---|
| 0 | No integrity violations found |
| 1 | Violations found and repaired |
| 2 | Violations found but could not be repaired — run DISM first |

**DISM /RestoreHealth**  
Launched with `RedirectStandardOutput = true`. Output is streamed line by line into the log box in real time via the queue/timer pattern.

**USB Reset**  
Calls `Get-PnpDevice` filtered by `InstanceId -match '^USB\\'`, excluding classes `HIDClass`, `Keyboard`, `Mouse`, `System`, `SecurityDevices`, `Biometric` to avoid losing input during the reset. For each remaining device: `Disable-PnpDevice` → 300 ms → `Enable-PnpDevice`.

**System Info**  
Displays: OS caption, architecture, CPU name, total RAM (GB), PowerShell version, uptime (`Xd Yh Zm`), last boot timestamp. Refreshable.

**HTML Report Export**  
Generates a dark-theme HTML file to the desktop (`DeviceRepairTool_Report_<timestamp>.html`). The report includes:
- Sidebar navigation with anchor links and color-coded badges
- Stat cards (device count, driver store count, event count, critical event count)
- System info grid
- Full table of problem devices (from Tab 1 scan)
- Full table of Driver Store entries (from Tab 2 scan)
- Full events table (from Tab 3 scan)
- Full activity log

All values are HTML-escaped before insertion. The report is built with `StringBuilder` to avoid repeated string concatenation. Layout uses CSS grid and flexbox, sticky table headers, responsive columns.

![drt 4](./images/drt4.png)
---

## Logging

Every operation appends to an in-memory `List<PSCustomObject>` with fields `Timestamp`, `Level` (`INFO` / `SUCCESS` / `WARNING` / `ERROR`), and `Message`. Each tab has a read-only text box that shows the formatted log and auto-scrolls to the latest entry. Logs can be cleared from the UI.

---

## Known constraints

- Requires a `.ps1` file path to be present (`$PSCommandPath` must not be null) — the script cannot be pasted into an interactive console.
- The Windows Update tab uses the Windows Update Agent COM object; it will return no results on systems where WUA is disabled or behind WSUS with different policies.
- SFC output is read from CBS.log after the fact, not streamed in real time.
- Driver repair via INF reinstall only applies if the device has an associated INF file recorded in `Win32_PnPSignedDriver`; otherwise only the disable/enable cycle is performed.
- The Driver Store removal is irreversible. The script validates the `oem\d+.inf` name pattern before calling `pnputil /delete-driver` but performs no dependency check.

---

## File structure

```
driverRep/
├── DRTPro.ps1          # Main script (also released as DRTPro4.ps1 in v4.1)
├── cd-rom-driver_25179.ico
└── .gitattributes
```

The entire application is a single `.ps1` file. The icon is a companion resource; the script does not reference it at runtime (it is used for the release packaging).
