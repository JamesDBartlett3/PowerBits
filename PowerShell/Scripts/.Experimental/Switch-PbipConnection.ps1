
<#
.SYNOPSIS
  Switch a PBIP thin report between:
   - Local Power BI Desktop XMLA → schema-compliant byPath to ../{ReportName}.SemanticModel (modelReference.json)
   - Published Power BI Service semantic model → byConnection (service only)

.DESCRIPTION
  - Targets a specific PBIP via -PbipFilePath (handles folders containing many PBIPs).
  - LOCAL:
      * Finds Desktop XMLA port via Get-NetTCPConnection (msmdsrv) with deterministic selection.
      * Creates/updates adjacent folder: {ReportName}.SemanticModel.
      * Writes modelReference.json (Data Source=localhost:<port>; Initial Catalog=<Catalog>; connectionType=analysisServicesDatabaseLive).
      * Rewrites {ReportName}.Report\definition.pbir to byPath("../{ReportName}.SemanticModel").
  - PUBLISHED:
      * Restores previously-saved byConnection (exact JSON), OR
      * Builds a valid byConnection from -WorkspaceConnectionString + -DatasetId (semantic model GUID).
  - Stores presets outside PBIP in {ReportName}.Connections\:
      * definition.pbir.workspace  (exact byConnection JSON template to restore)
      * local.meta.json            (last detected port/catalog/folder)

.REFERENCES
  - Report definition (definition.pbir) schema & rules (byPath vs byConnection; service vs non-service):
    https://github.com/microsoft/powerbi-desktop-samples/blob/main/item-schemas/report/definition.pbir.md
    https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json
  - PBIP Semantic Model folder & required files (.platform, definition.pbism, modelReference.json):
    https://learn.microsoft.com/power-bi/developer/projects/projects-dataset
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $PbipFilePath,     # Full path to a .pbip file

    [Parameter(Mandatory = $true)]
    [ValidateSet("Local","Published")]
    [string] $Mode,

    # --- Published inputs (service) ---
    [string] $WorkspaceConnectionString,  # e.g., powerbi://...; Initial Catalog=...; Integrated Security=ClaimsToken
    [string] $DatasetId,                  # semantic model GUID required to build a compliant byConnection

    # --- Local discovery overrides ---
    [int]    $Port,                # force a specific local XMLA port
    [int]    $MsmdsrvPid,          # pin to a specific Desktop engine PID
    [string] $Catalog = "Model"    # initial catalog for local engine (usually 'Model')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------- Paths ----------
if (-not (Test-Path $PbipFilePath)) { throw "PBIP not found: $PbipFilePath" }
if (-not ($PbipFilePath.ToLower().EndsWith(".pbip"))) { throw "Not a .pbip file: $PbipFilePath" }

$pbipFullPath = (Resolve-Path $PbipFilePath).Path
$pbipFolder   = Split-Path $pbipFullPath -Parent
$pbipFilename = Split-Path $pbipFullPath -Leaf
$reportName   = [IO.Path]::GetFileNameWithoutExtension($pbipFilename)

$reportFolder   = Join-Path $pbipFolder "$reportName.Report"
$definitionFile = Join-Path $reportFolder "definition.pbir"
if (-not (Test-Path $reportFolder))   { throw "Report folder not found: $reportFolder" }
if (-not (Test-Path $definitionFile)) { throw "definition.pbir not found at: $definitionFile" }

# Per your naming rule: do NOT include "local" or "workspace" in the artifact name
$semModelFolderName = "$reportName.SemanticModel"
$semModelFolder     = Join-Path $pbipFolder $semModelFolderName

# External presets (outside PBIP to avoid Desktop cleanup)
$externalFolder       = Join-Path $pbipFolder "$reportName.Connections"
if (-not (Test-Path $externalFolder)) { New-Item -ItemType Directory -Path $externalFolder | Out-Null }
$workspacePresetPath  = Join-Path $externalFolder "definition.pbir.workspace" # FULL JSON (byConnection)
$localMetaPath        = Join-Path $externalFolder "local.meta.json"           # last resolved port/catalog/folder
$backupFolder         = Join-Path $externalFolder "backups"                   # backup of report metadata files

# Files to preserve (Power BI Desktop may delete these when switching connections)
$filesToPreserve = @(
  "semanticModelDiagramLayout.json"
)

# Cache expensive checks at script level (DRY principle)
$script:hasGetNetTCPConnection = [bool](Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue)
$script:FabricSchemaBase = "https://developer.microsoft.com/json-schemas/fabric"

# ---------- Helpers ----------
function Read-JsonFile([string] $filePath) {
  (Get-Content -Path $filePath -Raw -Encoding UTF8) | ConvertFrom-Json
}
function Write-JsonFile([string] $filePath, $psObject) {
  $json = $psObject | ConvertTo-Json -Depth 30
  Set-Content -Path $filePath -Value $json -Encoding UTF8
}

# Use Windows API to enumerate windows - more reliable than Get-Process.MainWindowTitle
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public class WindowEnumerator {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    public static List<string> GetWindowTitlesForProcess(string processName) {
        var titles = new List<string>();
        var pids = new HashSet<uint>();

        foreach (var proc in System.Diagnostics.Process.GetProcessesByName(processName)) {
            pids.Add((uint)proc.Id);
        }

        EnumWindows((hWnd, lParam) => {
            if (IsWindowVisible(hWnd)) {
                uint pid;
                GetWindowThreadProcessId(hWnd, out pid);
                if (pids.Contains(pid)) {
                    int len = GetWindowTextLength(hWnd);
                    if (len > 0) {
                        var sb = new StringBuilder(len + 1);
                        GetWindowText(hWnd, sb, sb.Capacity);
                        string title = sb.ToString();
                        if (!string.IsNullOrEmpty(title)) {
                            titles.Add(title);
                        }
                    }
                }
            }
            return true;
        }, IntPtr.Zero);

        return titles;
    }
}
"@ -ErrorAction SilentlyContinue

# Check if any Power BI Desktop instance has a file with the given basename open (from window title)
function Test-BaseNameOpenInDesktop([string] $baseName) {
  $windowTitles = [WindowEnumerator]::GetWindowTitlesForProcess("PBIDesktop")
  foreach ($title in $windowTitles) {
    Write-Verbose "Found PBIDesktop window: '$title'"
    if ($title -match '^(.+?)\*?\s*-\s*Power BI Desktop') {
      $openBaseName = $Matches[1].Trim()
      Write-Verbose "  Extracted basename: '$openBaseName' (looking for: '$baseName')"
      if ($openBaseName -eq $baseName) {
        return $true
      }
    }
  }
  return $false
}

# Backup report metadata files that Power BI Desktop may delete when switching connections
function Backup-ReportMetadata() {
  if (-not (Test-Path $backupFolder)) { New-Item -ItemType Directory -Path $backupFolder | Out-Null }
  foreach ($fileName in $filesToPreserve) {
    $srcFile = Join-Path $reportFolder $fileName
    if (Test-Path $srcFile) {
      $dstFile = Join-Path $backupFolder $fileName
      Copy-Item -Path $srcFile -Destination $dstFile -Force
      Write-Verbose "Backed up $fileName"
    }
  }
}

# Restore report metadata files if they were deleted by Power BI Desktop
function Restore-ReportMetadata() {
  if (-not (Test-Path $backupFolder)) { return }
  $restored = @()
  foreach ($fileName in $filesToPreserve) {
    $srcFile = Join-Path $reportFolder $fileName
    $bakFile = Join-Path $backupFolder $fileName
    if ((-not (Test-Path $srcFile)) -and (Test-Path $bakFile)) {
      Copy-Item -Path $bakFile -Destination $srcFile -Force
      $restored += $fileName
    }
  }
  if ($restored.Count -gt 0) {
    Write-Host "Restored from backup: $($restored -join ', ')" -ForegroundColor Yellow
  }
}
function Get-CurrentDefinition() { Read-JsonFile $definitionFile }
function Get-CurrentConnectionMode() {
  try {
    $def = Get-CurrentDefinition
    if ($def.datasetReference.byPath)       { return "ByPath" }
    if ($def.datasetReference.byConnection) { return "ByConnection" }
    return "Unknown"
  } catch { return "Unknown" }
}
function New-ReportByPathDefinition([string] $relativePath) {
  # Preserve schema and version from existing definition.pbir
  $current = Get-CurrentDefinition
  return @{
    '$schema' = $current.'$schema'
    version   = $current.version
    datasetReference = @{
      byPath       = @{ path = $relativePath }
      byConnection = $null
    }
  }
}
function New-ReportByConnectionDefinition([string] $connectionString, [string] $datasetGuid) {
  # Preserve schema and version from existing definition.pbir
  # Build connection string with semantic model ID embedded
  $current = Get-CurrentDefinition
  $fullConnString = "$connectionString;semanticmodelid=$datasetGuid"
  return @{
    '$schema' = $current.'$schema'
    version   = $current.version
    datasetReference = @{
      byConnection = @{ connectionString = $fullConnString }
    }
  }
}
function Test-TcpPort([int] $testPort) {
  try {
    $client = New-Object System.Net.Sockets.TcpClient
    $async  = $client.BeginConnect("127.0.0.1", $testPort, $null, $null)
    $wait   = $async.AsyncWaitHandle.WaitOne(300)
    if ($wait -and $client.Connected) { $client.Close(); return $true }
    $client.Close(); return $false
  } catch { return $false }
}
function Get-ListeningPortsForProcess([int] $processId) {
  $ports = @()
  if ($script:hasGetNetTCPConnection) {
    try {
      $conns = Get-NetTCPConnection -State Listen -OwningProcess $processId -ErrorAction Stop
      foreach ($c in $conns) { if ($c.LocalPort -gt 0) { $ports += [int]$c.LocalPort } }
    } catch { }
  }
  if (-not $ports -or $ports.Count -eq 0) {
    try {
      $lines = netstat -ano -p tcp 2>$null | Select-String -Pattern "LISTENING"
      foreach ($line in $lines) {
        $parts = $line.ToString() -split '\s+'
        if ($parts.Length -ge 5) {
          $pidCandidate = [int]$parts[-1]
          if ($pidCandidate -eq $processId) {
            $local = $parts[2]
            if ($local -match ':(\d+)$') { $ports += [int]$Matches[1] }
          }
        }
      }
    } catch { }
  }
  return ($ports | Sort-Object -Unique)
}

function Get-MsmdsrvCandidates() {
  $procs = @(Get-Process -Name "msmdsrv" -ErrorAction SilentlyContinue)
  foreach ($p in $procs) {
    $processId = $p.Id
    $ports     = Get-ListeningPortsForProcess -processId $processId

    # Build per-port metadata (does this PID listen on both IPv4 and IPv6 loopback?)
    $portMeta = @()
    if ($script:hasGetNetTCPConnection) {
      $ipv4 = (Get-NetTCPConnection -State Listen -OwningProcess $processId -ErrorAction SilentlyContinue |
               Where-Object { $_.LocalAddress -eq "127.0.0.1" } | Select-Object -ExpandProperty LocalPort -Unique)
      $ipv6 = (Get-NetTCPConnection -State Listen -OwningProcess $processId -ErrorAction SilentlyContinue |
               Where-Object { $_.LocalAddress -eq "::1" }       | Select-Object -ExpandProperty LocalPort -Unique)
      foreach ($prt in $ports) {
        $hasV4 = $ipv4 -contains $prt
        $hasV6 = $ipv6 -contains $prt
        $portMeta += [pscustomobject]@{
          Port = $prt; HasV4 = $hasV4; HasV6 = $hasV6; Both = ($hasV4 -and $hasV6)
        }
      }
    } else {
      foreach ($prt in $ports) {
        $portMeta += [pscustomobject]@{ Port = $prt; HasV4 = $true; HasV6 = $false; Both = $false }
      }
    }

    [pscustomobject]@{
      Pid       = $processId
      StartTime = $p.StartTime
      Ports     = $portMeta
    }
  }
}

function Select-DesktopPortStable([int] $preferredPort, [int] $preferredPid) {
  # 1) Prefer last-known port if it's still listening
  if ($preferredPort -gt 0 -and (Test-TcpPort -testPort $preferredPort)) {
    Write-Verbose "Using last-known port $preferredPort"
    return $preferredPort
  }

  $candidates = Get-MsmdsrvCandidates | Sort-Object StartTime -Descending

  # Helper to iterate a candidate's ports in a stable priority order
  function Get-OrderedPorts($cand) {
    $ordered = @()
    # Prefer ports that listen on BOTH ::1 and 127.0.0.1
    $ordered += ($cand.Ports | Where-Object { $_.Both } | Sort-Object Port -Descending)
    # Then the rest (highest port first)
    $ordered += ($cand.Ports | Sort-Object Port -Descending)
    # Remove duplicates that may appear in both lists
    return ($ordered | Select-Object -Unique)
  }

  # 2) If a PID is pinned, try it first
  if ($preferredPid -gt 0) {
    $cand = $candidates | Where-Object { $_.Pid -eq $preferredPid } | Select-Object -First 1
    if ($cand) {
      foreach ($entry in (Get-OrderedPorts $cand)) {
        if (Test-TcpPort -testPort $entry.Port) { return $entry.Port }
      }
    }
  }

  # 3) Otherwise: newest engine → (both-loopback first) → highest port
  foreach ($cand in $candidates) {
    foreach ($entry in (Get-OrderedPorts $cand)) {
      if (Test-TcpPort -testPort $entry.Port) { return $entry.Port }
    }
  }

  return $null
}

function Get-LocalCatalogName([int] $port) {
  # Query the local AS instance to discover available databases/catalogs
  # When no Initial Catalog is specified, we can query $SYSTEM.DBSCHEMA_CATALOGS
  try {
    $connStr = "Data Source=localhost:$port;Provider=MSOLAP"
    $conn = New-Object System.Data.OleDb.OleDbConnection($connStr)
    $conn.Open()

    # Query for available catalogs
    $schemaTable = $conn.GetOleDbSchemaTable([System.Data.OleDb.OleDbSchemaGuid]::Catalogs, $null)
    $catalogs = @()
    foreach ($row in $schemaTable.Rows) {
      $catalogs += $row["CATALOG_NAME"]
    }
    $conn.Close()

    if ($catalogs.Count -eq 0) {
      Write-Verbose "No catalogs found on localhost:$port"
      return $null
    }

    # Return the first non-system catalog (typically there's only one for Desktop)
    Write-Verbose "Found catalogs: $($catalogs -join ', ')"
    return $catalogs[0]
  } catch {
    Write-Verbose "Failed to query catalogs: $_"
    return $null
  }
}

function Get-DesktopFilePath([int] $port) {
  # Find the full path of the .pbix/.pbip file open in Power BI Desktop
  # Returns a hashtable with: FullPath (if available), BaseName, IsFullPath (bool)
  try {
    # Find which msmdsrv process owns this port
    $msmdsrvPid = $null
    if ($script:hasGetNetTCPConnection) {
      $conn = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue | Select-Object -First 1
      if ($conn) { $msmdsrvPid = $conn.OwningProcess }
    }
    if (-not $msmdsrvPid) {
      Write-Verbose "Could not find msmdsrv process for port $port"
      return $null
    }

    # Walk up the process tree to find PBIDesktop.exe
    $currentPid = $msmdsrvPid
    $maxDepth = 10
    $pbiDesktopPid = $null
    $pbiDesktopProc = $null
    for ($i = 0; $i -lt $maxDepth; $i++) {
      $proc = Get-Process -Id $currentPid -ErrorAction SilentlyContinue
      if (-not $proc) { break }
      if ($proc.ProcessName -eq "PBIDesktop") {
        $pbiDesktopPid = $proc.Id
        $pbiDesktopProc = $proc
        break
      }
      # Get parent process ID via CIM
      $wmiProc = Get-CimInstance Win32_Process -Filter "ProcessId = $currentPid" -ErrorAction SilentlyContinue
      if (-not $wmiProc -or -not $wmiProc.ParentProcessId) { break }
      $currentPid = $wmiProc.ParentProcessId
    }

    if (-not $pbiDesktopPid) {
      Write-Verbose "Could not find parent PBIDesktop.exe process"
      return $null
    }

    # Try to get the full file path using native Windows API
    $fullPath = Get-ProcessOpenFiles -processId $pbiDesktopPid | Where-Object { $_ -match '\.(?:pbix|pbip)$' } | Select-Object -First 1

    # Get basename from window title as fallback/verification
    $baseName = $null
    $title = $pbiDesktopProc.MainWindowTitle
    if ($title -match '^(.+?)\*?\s*-\s*Power BI Desktop') {
      $baseName = $Matches[1].Trim()
    }

    if ($fullPath) {
      Write-Verbose "Found file via Windows API: $fullPath"
      return @{ FullPath = $fullPath; BaseName = [IO.Path]::GetFileNameWithoutExtension($fullPath); IsFullPath = $true }
    } elseif ($baseName) {
      return @{ FullPath = $null; BaseName = $baseName; IsFullPath = $false }
    }
    return $null
  } catch {
    Write-Verbose "Failed to get Desktop file path: $_"
    return $null
  }
}

function Get-ProcessOpenFiles([int] $processId) {
  # Enumerate file handles for a process using native Windows APIs (P/Invoke)
  # Returns an array of file paths
  # Uses a background thread with timeout for NtQueryObject to avoid hanging on pipes
  Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public class NtHandleQuery {
    [DllImport("ntdll.dll")]
    private static extern int NtQuerySystemInformation(int infoClass, IntPtr info, int size, out int length);

    [DllImport("ntdll.dll")]
    private static extern int NtQueryObject(IntPtr handle, int infoClass, IntPtr info, int size, out int length);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(int access, bool inherit, int processId);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DuplicateHandle(IntPtr sourceProcess, IntPtr sourceHandle,
        IntPtr targetProcess, out IntPtr targetHandle, int access, bool inherit, int options);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentProcess();

    private const int SystemHandleInformation = 16;
    private const int ObjectNameInformation = 1;
    private const int ObjectTypeInformation = 2;
    private const int PROCESS_DUP_HANDLE = 0x0040;
    private const int DUPLICATE_SAME_ACCESS = 0x0002;

    [StructLayout(LayoutKind.Sequential)]
    private struct SYSTEM_HANDLE_ENTRY {
        public int OwnerPid;
        public byte ObjectType;
        public byte Flags;
        public short Handle;
        public IntPtr Object;
        public int Access;
    }

    public static List<string> GetOpenFiles(int processId) {
        var files = new List<string>();
        int length = 0x10000;
        IntPtr ptr = Marshal.AllocHGlobal(length);

        try {
            int returnLength;
            while (NtQuerySystemInformation(SystemHandleInformation, ptr, length, out returnLength) == unchecked((int)0xC0000004)) {
                length *= 2;
                Marshal.FreeHGlobal(ptr);
                ptr = Marshal.AllocHGlobal(length);
            }

            int handleCount = Marshal.ReadInt32(ptr);
            IntPtr handlePtr = ptr + IntPtr.Size;
            int entrySize = Marshal.SizeOf(typeof(SYSTEM_HANDLE_ENTRY));

            IntPtr processHandle = OpenProcess(PROCESS_DUP_HANDLE, false, processId);
            if (processHandle == IntPtr.Zero) return files;

            try {
                for (int i = 0; i < handleCount; i++) {
                    var entry = (SYSTEM_HANDLE_ENTRY)Marshal.PtrToStructure(handlePtr + (i * entrySize), typeof(SYSTEM_HANDLE_ENTRY));
                    if (entry.OwnerPid != processId) continue;

                    IntPtr dupHandle;
                    if (!DuplicateHandle(processHandle, (IntPtr)entry.Handle, GetCurrentProcess(), out dupHandle, 0, false, DUPLICATE_SAME_ACCESS))
                        continue;

                    try {
                        string name = GetObjectNameWithTimeout(dupHandle, 100);
                        if (!string.IsNullOrEmpty(name) && name.StartsWith("\\Device\\")) {
                            string dosPath = ConvertDevicePath(name);
                            if (!string.IsNullOrEmpty(dosPath)) files.Add(dosPath);
                        }
                    } finally {
                        CloseHandle(dupHandle);
                    }
                }
            } finally {
                CloseHandle(processHandle);
            }
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
        return files;
    }

    private static string GetObjectNameWithTimeout(IntPtr handle, int timeoutMs) {
        string result = null;
        var thread = new Thread(() => { result = GetObjectNameDirect(handle); });
        thread.IsBackground = true;
        thread.Start();
        if (!thread.Join(timeoutMs)) {
            // Thread is stuck (likely on a pipe handle) - abandon it
            return null;
        }
        return result;
    }

    private static string GetObjectNameDirect(IntPtr handle) {
        int length = 0x1000;
        IntPtr ptr = Marshal.AllocHGlobal(length);
        try {
            int returnLength;
            if (NtQueryObject(handle, ObjectNameInformation, ptr, length, out returnLength) != 0)
                return null;

            int nameLength = Marshal.ReadInt16(ptr);
            if (nameLength == 0) return null;

            IntPtr namePtr = Marshal.ReadIntPtr(ptr + IntPtr.Size);
            return Marshal.PtrToStringUni(namePtr, nameLength / 2);
        } catch {
            return null;
        } finally {
            Marshal.FreeHGlobal(ptr);
        }
    }

    private static string ConvertDevicePath(string devicePath) {
        foreach (var drive in System.IO.DriveInfo.GetDrives()) {
            if (drive.DriveType != System.IO.DriveType.Fixed && drive.DriveType != System.IO.DriveType.Removable)
                continue;

            var sb = new StringBuilder(256);
            if (QueryDosDevice(drive.Name.Substring(0, 2), sb, 256) > 0) {
                string deviceName = sb.ToString();
                if (devicePath.StartsWith(deviceName, StringComparison.OrdinalIgnoreCase)) {
                    return drive.Name.Substring(0, 2) + devicePath.Substring(deviceName.Length);
                }
            }
        }
        return null;
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int QueryDosDevice(string deviceName, StringBuilder targetPath, int maxSize);
}
"@ -ErrorAction SilentlyContinue

  try {
    return [NtHandleQuery]::GetOpenFiles($processId)
  } catch {
    Write-Verbose "Failed to query process handles: $_"
    return @()
  }
}

function Initialize-SemanticModelReference([string] $server, [string] $db, [string] $folderPath, [string] $displayName) {
  if (-not (Test-Path $folderPath)) {
    New-Item -ItemType Directory -Path $folderPath | Out-Null
  } else {
    # If this looks like a *real* authored semantic model (model.bim or definition/ folder), abort rather than mutate it
    if ((Test-Path (Join-Path $folderPath "model.bim")) -or (Test-Path (Join-Path $folderPath "definition"))) {
      throw "Existing '$folderPath' looks like a full semantic model (model.bim/definition). Choose a different report or remove the folder first."
    }
  }

  # .platform (required in PBIP semantic model folder)
  $platformFile = Join-Path $folderPath ".platform"
  if (-not (Test-Path $platformFile)) {
    $platform = @{
      '$schema' = "$script:FabricSchemaBase/gitIntegration/platformProperties/2.0.0/schema.json"
      metadata  = @{ type = "SemanticModel"; displayName = $displayName }
      config    = @{ version = "2.0"; logicalId = ([guid]::NewGuid().ToString()) }
    }
    Write-JsonFile -filePath $platformFile -psObject $platform
  }

  # definition.pbism (minimal)
  $defPbismFile = Join-Path $folderPath "definition.pbism"
  if (-not (Test-Path $defPbismFile)) {
    $defPbism = @{
      '$schema' = "$script:FabricSchemaBase/item/semanticModel/definitionProperties/1.0.0/schema.json"
      version   = "4.2"
      settings  = @{}
    }
    Write-JsonFile -filePath $defPbismFile -psObject $defPbism
  }

  # modelReference.json (the XMLA link to Desktop)
  $modelRefFile = Join-Path $folderPath "modelReference.json"
  $connString   = "Data Source=$server;Initial Catalog=$db"
  $modelRef = @{
    '$schema'          = "$script:FabricSchemaBase/item/semanticModel/modelReference/2.0.0/schema.json"
    connectionString   = $connString
    connectionType     = "analysisServicesDatabaseLive"
    isMultiDimentional = $false
  }
  Write-JsonFile -filePath $modelRefFile -psObject $modelRef
}

function Save-WorkspacePresetIfByConnection() {
  try {
    $def = Get-CurrentDefinition
    if ($def.datasetReference.byConnection) {
      Write-JsonFile -filePath $workspacePresetPath -psObject $def
      Write-Host "Saved current published binding to preset: $workspacePresetPath" -ForegroundColor DarkGray
    }
  } catch { }
}

# ---------- SAFETY CHECK: Ensure target PBIP is not open in Desktop ----------
if (Test-BaseNameOpenInDesktop -baseName $reportName) {
  throw "Cannot switch connection: A file named '$reportName' is currently open in Power BI Desktop. Close the file first and retry."
}

# ---------- RESTORE METADATA (if deleted by Desktop on previous run) ----------
Restore-ReportMetadata

# ---------- LOCAL MODE ----------
if ($Mode -eq "Local") {
  Write-Host "Switching to LOCAL XMLA (schema-compliant byPath pointing to modelReference.json)..." -ForegroundColor Cyan

  # Backup metadata files before switching (Desktop may delete them when report is opened)
  Backup-ReportMetadata

  # Preserve the exact service binding for easy restore later
  if (Get-CurrentConnectionMode -eq "ByConnection") { Save-WorkspacePresetIfByConnection }

  # Load last-known port if available
  $preferredPort = $null
  if (Test-Path $localMetaPath) {
    try { $prev = Read-JsonFile $localMetaPath; $preferredPort = [int]$prev.lastPort } catch { $preferredPort = $null }
  }

  # Resolve port
  $resolvedPort = $null
  if ($PSBoundParameters.ContainsKey('Port') -and $Port -gt 0) {
    if (Test-TcpPort -testPort $Port) { $resolvedPort = $Port }
    else { Write-Verbose "User-specified -Port $Port is not listening." }
  }
  if (-not $resolvedPort) {
    $resolvedPort = Select-DesktopPortStable -preferredPort $preferredPort -preferredPid $MsmdsrvPid
  }
  if (-not $resolvedPort) { throw "No local Power BI Desktop XMLA endpoint detected. Open a semantic model in Desktop and retry." }

  Write-Host "Using XMLA port: $resolvedPort" -ForegroundColor Green

  # Discover the file open in Power BI Desktop
  $desktopFile = Get-DesktopFilePath -port $resolvedPort

  # Discover the actual catalog/database name from the running instance
  $resolvedCatalog = $Catalog
  if (-not $PSBoundParameters.ContainsKey('Catalog') -or $Catalog -eq "Model") {
    $discoveredCatalog = Get-LocalCatalogName -port $resolvedPort
    if ($discoveredCatalog) {
      $resolvedCatalog = $discoveredCatalog
      Write-Host "Discovered catalog: $resolvedCatalog" -ForegroundColor Green
    } else {
      Write-Warning "Could not discover catalog name; using default '$Catalog'. You may need to specify -Catalog explicitly."
    }
  }

  # Create/update {ReportName}.SemanticModel reference (NOT a full semantic model)
  Initialize-SemanticModelReference -server "localhost:$resolvedPort" -db $resolvedCatalog -folderPath $semModelFolder -displayName $reportName

  # Update definition.pbir → byPath ("../{ReportName}.SemanticModel")
  $relative = "../$semModelFolderName"
  $byPathDef = New-ReportByPathDefinition -relativePath $relative
  Write-JsonFile -filePath $definitionFile -psObject $byPathDef

  # Persist meta so we prefer the same engine on the next run
  $meta = @{
    lastPort     = $resolvedPort
    catalog      = $resolvedCatalog
    semModelPath = $semModelFolderName
    timestamp    = (Get-Date).ToString("o")
  }
  if ($desktopFile) {
    $meta.desktopFile     = $desktopFile.FullPath
    $meta.desktopBaseName = $desktopFile.BaseName
  }
  Write-JsonFile -filePath $localMetaPath -psObject $meta

  # Display success message with Desktop file info
  if ($desktopFile) {
    $displayName = if ($desktopFile.IsFullPath) { $desktopFile.FullPath } else { $desktopFile.BaseName }
    Write-Host "Bound to semantic model in '$displayName', running locally in Power BI Desktop (queries over XMLA: localhost:$resolvedPort)" -ForegroundColor Green
  } else {
    Write-Host "Bound to local Power BI Desktop (queries over XMLA: localhost:$resolvedPort)" -ForegroundColor Green
  }
  exit 0
}

# ---------- PUBLISHED MODE ----------
if ($Mode -eq "Published") {
  Write-Host "Switching to PUBLISHED semantic model (byConnection)..." -ForegroundColor Cyan

  # Backup metadata files before switching (Desktop may delete them when report is opened)
  Backup-ReportMetadata

  $publishDef = $null

  # Preferred: restore the exact JSON we saved before
  if (Test-Path $workspacePresetPath) {
    Write-Host "Restoring saved .workspace preset..." -ForegroundColor Yellow
    $publishDef = Read-JsonFile $workspacePresetPath
  }

  # Or construct a compliant byConnection from inputs
  if (-not $publishDef) {
    if (-not $WorkspaceConnectionString -or -not $DatasetId) {
      throw "Need a saved preset OR both -WorkspaceConnectionString and -DatasetId."
    }
    $publishDef = New-ReportByConnectionDefinition -connectionString $WorkspaceConnectionString -datasetGuid $DatasetId
  }

  Write-JsonFile -filePath $definitionFile      -psObject $publishDef
  Write-JsonFile -filePath $workspacePresetPath -psObject $publishDef

  Write-Host "Bound to PUBLISHED semantic model (powerbi:// …)." -ForegroundColor Green
  exit 0
}
