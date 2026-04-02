#region HEADER & GLOBALS

<#
    Version 7.4.0
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Resolve script directory and core paths
$ScriptDir  = Split-Path -Parent $PSCommandPath
$ConfigPath = Join-Path $ScriptDir 'config.json'
$GamesPath  = Join-Path $ScriptDir 'games.json'
$LogFile    = Join-Path $ScriptDir 'sim_launcher.log'

# State Memory: Track running processes before terminating them
$Global:PreSessionRunningApps = @()
$Global:SteamWasRunningBeforeNonSteamLaunch = $false
$Global:SteamClosedForNonSteamLaunch = $false
$Global:DisplaySessionState = $null

#endregion HEADER & GLOBALS

#region ADMIN ELEVATION

function Ensure-Admin {
    $currentIdentity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal        = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    $isAdmin          = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "Requesting administrative privileges..." -ForegroundColor Yellow
        $psi = @{
            FilePath     = 'powershell.exe'
            ArgumentList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
            Verb         = 'RunAs'
        }
        try {
            Start-Process @psi
        } catch {
            Write-Host "Elevation cancelled or failed. Exiting." -ForegroundColor Red
        }
        exit
    }
}

Ensure-Admin

#endregion ADMIN ELEVATION

#region BASIC UTILITIES

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    Add-Content -Path $LogFile -Value $line
}

#endregion BASIC UTILITIES
#region UI FRAMEWORK
<#
    UI Framework
    - Light box-drawing borders
    - Color helpers
    - Input helpers
    - Menu rendering utilities
#>

# ASCII box characters (encoding-safe for Windows PowerShell 5.1)
$UI = @{
    TopLeft     = "+"
    TopRight    = "+"
    BottomLeft  = "+"
    BottomRight = "+"
    Horizontal  = "-"
    Vertical    = "|"
}

function New-BoxLine {
    param(
        [Parameter(Mandatory)]
        [string]$Text,
        [ValidateSet('Header','Footer','Line')]
        [string]$Type = 'Line',
        [int]$Width = 60
    )

    switch ($Type) {
        'Header' {
            $h = $UI.Horizontal * $Width
            return "$($UI.TopLeft)$h$($UI.TopRight)"
        }
        'Footer' {
            $h = $UI.Horizontal * $Width
            return "$($UI.BottomLeft)$h$($UI.BottomRight)"
        }
        'Line' {
            # Center text inside the box
            $padding = $Width - $Text.Length
            if ($padding -lt 0) { $padding = 0 }
            $leftPad  = [math]::Floor($padding / 2)
            $rightPad = $padding - $leftPad
            return "$($UI.Vertical)$(' ' * $leftPad)$Text$(' ' * $rightPad)$($UI.Vertical)"
        }
    }
}

function Show-Box {
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [int]$Width = 60
    )

    Write-Host (New-BoxLine -Text $Title -Type Header -Width $Width) -ForegroundColor Cyan
    Write-Host (New-BoxLine -Text $Title -Type Line   -Width $Width) -ForegroundColor Cyan
    Write-Host (New-BoxLine -Text $Title -Type Footer -Width $Width) -ForegroundColor Cyan
    Write-Host ""
}

# Color helpers
function Write-Info    { param($t) Write-Host $t -ForegroundColor Cyan }
function Write-Warn    { param($t) Write-Host $t -ForegroundColor Yellow }
function Write-ErrorUI { param($t) Write-Host $t -ForegroundColor Red }
function Write-Success { param($t) Write-Host $t -ForegroundColor Green }
function Write-White   { param($t) Write-Host $t -ForegroundColor White }

function Get-RemovalIndex {
    param(
        [Parameter(Mandatory)]
        [string]$Text,
        [Parameter(Mandatory)]
        [int]$MaxCount
    )

    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }

    $trimmed = $Text.Trim()
    if ($trimmed -notmatch '^[0-9]+$') { return $null }

    $index = [int]$trimmed
    if ($index -lt 1 -or $index -gt $MaxCount) { return $null }

    return ($index - 1)
}

function Remove-ArrayItemAtIndex {
    param(
        [Parameter(Mandatory)]
        [object[]]$Items,
        [Parameter(Mandatory)]
        [int]$IndexToRemove
    )

    $result = @()
    for ($i = 0; $i -lt $Items.Count; $i++) {
        if ($i -ne $IndexToRemove) {
            $result += ,$Items[$i]
        }
    }

    return $result
}

function ConvertTo-HashtableDeep {
    param(
        [Parameter(ValueFromPipeline = $true)]
        $InputObject
    )

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $hash = @{}
        foreach ($key in $InputObject.Keys) {
            $hash[$key] = ConvertTo-HashtableDeep -InputObject $InputObject[$key]
        }
        return $hash
    }

    if ($InputObject -is [System.Array] -or ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string]))) {
        $items = @()
        foreach ($item in $InputObject) {
            $items += ,(ConvertTo-HashtableDeep -InputObject $item)
        }
        return ,$items
    }

    if ($InputObject -is [pscustomobject]) {
        $hash = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $hash[$property.Name] = ConvertTo-HashtableDeep -InputObject $property.Value
        }
        return $hash
    }

    return $InputObject
}

# Input helper
function Read-Choice {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )
    Write-Host ""
    Write-Host -NoNewline "$Prompt "
    return Read-Host
}

function ConvertTo-NormalizedProcessName {
    param(
        [Parameter(Mandatory)]
        [string]$InputValue
    )

    if ([string]::IsNullOrWhiteSpace($InputValue)) { return $null }

    $leaf = Split-Path -Leaf $InputValue
    if ([string]::IsNullOrWhiteSpace($leaf)) {
        $leaf = $InputValue
    }

    $normalized = ($leaf -replace '\.exe$','').Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }

    return $normalized
}

function Test-ExecutablePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }

    return ([string]::Equals([System.IO.Path]::GetExtension($Path), '.exe', [System.StringComparison]::OrdinalIgnoreCase))
}

function Select-ExecutableFile {
    param(
        [string]$Title = 'Select executable file',
        [string]$InitialDirectory = $null
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms
    }
    catch {
        Write-ErrorUI 'File picker is not available on this system.'
        Write-Log "Failed to load System.Windows.Forms for file picker: $_" -Level ERROR
        return $null
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $Title
    $dialog.Filter = 'Executable files (*.exe)|*.exe|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false

    if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory -PathType Container)) {
        $dialog.InitialDirectory = $InitialDirectory
    }
    elseif (Test-Path -LiteralPath $env:ProgramFiles -PathType Container) {
        $dialog.InitialDirectory = $env:ProgramFiles
    }

    $result = $dialog.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    if (-not (Test-ExecutablePath -Path $dialog.FileName)) {
        Write-ErrorUI 'Please select a valid .exe file.'
        return $null
    }

    return $dialog.FileName
}

#endregion UI FRAMEWORK
#region CONFIG SYSTEM
<#
    JSON Configuration System
    - Loads config.json from script directory
    - Creates default config if missing
    - Provides Get/Set helpers
    - Ensures strong structure and validation
#>

# Default configuration structure
$DefaultConfig = @{
    Kill = @{
        OneDrive        = $true
        Edge            = $true
        CCleaner        = $true
        iCloudServices  = $true
        iCloudDrive     = $true
        Discord         = $true 
        Custom          = @()   # array of process names
    }
    Restart = @{
        Edge            = $true
        Discord         = $true
        OneDrive        = $true
        CCleaner        = $true
        iCloud          = $true
        Custom          = @()   # array of @{ Command=""; Args="" }
    }
    Exception = @{
        VSCode          = $false
        Spotify         = $false
        Steam           = $false
        Discord         = $false
        TeamSpeak       = $false
        Custom          = @()   # array of process names
    }
    Display = @{
        DisableSecondMonitor = $false
        LowestQuality        = $false
    }
    DefaultSim              = $null
    AutoRunOnStart          = $false
    RestoreOnlyActiveApps   = $true
    CustomAppsToStartAfter  = ""
}

function Normalize-CustomConfigLists {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    # Kill.Custom must always be an array of process-name strings.
    $killItems = @($Config.Kill.Custom)
    if ($null -eq $Config.Kill.Custom) { $killItems = @() }

    $normalizedKill = @()
    foreach ($item in $killItems) {
        if ([string]::IsNullOrWhiteSpace([string]$item)) { continue }
        $name = ConvertTo-NormalizedProcessName -InputValue ([string]$item)
        if (-not $name) { continue }

        $nameKey = $name.ToLowerInvariant()
        $existing = @($normalizedKill | ForEach-Object { $_.ToLowerInvariant() })
        if ($nameKey -in $existing) { continue }

        $normalizedKill += ,$name
    }
    $Config.Kill.Custom = $normalizedKill

    # Exception.Custom must always be an array of process-name strings.
    $exceptionItems = @($Config.Exception.Custom)
    if ($null -eq $Config.Exception.Custom) { $exceptionItems = @() }

    $normalizedException = @()
    foreach ($item in $exceptionItems) {
        if ([string]::IsNullOrWhiteSpace([string]$item)) { continue }
        $name = ConvertTo-NormalizedProcessName -InputValue ([string]$item)
        if (-not $name) { continue }

        $nameKey = $name.ToLowerInvariant()
        $existing = @($normalizedException | ForEach-Object { $_.ToLowerInvariant() })
        if ($nameKey -in $existing) { continue }

        $normalizedException += ,$name
    }
    $Config.Exception.Custom = $normalizedException

    # Restart.Custom must always be an array of @{ Command=''; Args='' } objects.
    $restartItems = @($Config.Restart.Custom)
    if ($null -eq $Config.Restart.Custom) { $restartItems = @() }

    $normalizedRestart = @()
    foreach ($item in $restartItems) {
        if ($null -eq $item) { continue }

        $command = $null
        $args = ''

        if ($item -is [System.Collections.IDictionary]) {
            if ($item.Contains('Command')) { $command = [string]$item['Command'] }
            if ($item.Contains('Args')) { $args = [string]$item['Args'] }
        }
        elseif ($item -is [pscustomobject]) {
            if ($null -ne $item.PSObject.Properties['Command']) { $command = [string]$item.Command }
            if ($null -ne $item.PSObject.Properties['Args']) { $args = [string]$item.Args }
        }
        else {
            $command = [string]$item
        }

        if ([string]::IsNullOrWhiteSpace($command)) { continue }
        $normalizedRestart += ,([ordered]@{
            Command = $command
            Args    = $args
        })
    }
    $Config.Restart.Custom = $normalizedRestart
}

function Save-Config {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    try {
        Normalize-CustomConfigLists -Config $Config
        $json = $Config | ConvertTo-Json -Depth 10 -Compress
        Set-Content -Path $ConfigPath -Value $json -Encoding UTF8
        Write-Log "Configuration saved to $ConfigPath"
    }
    catch {
        Write-ErrorUI "Failed to save configuration."
        Write-Log "Failed to save configuration: $_" -Level ERROR
    }
}

function Load-Config {
    if (-not (Test-Path $ConfigPath)) {
        Write-Warn "Config file not found. Creating default config.json..."
        Save-Config -Config $DefaultConfig
        return $DefaultConfig
    }

    try {
        $json = Get-Content -Path $ConfigPath -Raw
        $config = ConvertTo-HashtableDeep -InputObject ($json | ConvertFrom-Json)

        # Validate structure and fill missing keys
        foreach ($key in $DefaultConfig.Keys) {
            if (-not $config.ContainsKey($key)) {
                $config[$key] = $DefaultConfig[$key]
            }
        }

        # Validate nested keys
        foreach ($section in @('Kill','Restart','Exception','Display')) {
            foreach ($key in $DefaultConfig[$section].Keys) {
                if (-not $config[$section].ContainsKey($key)) {
                    $config[$section][$key] = $DefaultConfig[$section][$key]
                }
            }
        }

        Normalize-CustomConfigLists -Config $config

        Write-Log "Configuration loaded from $ConfigPath"
        return $config
    }
    catch {
        Write-ErrorUI "Config file is corrupted. Creating a new one."
        Write-Log "Config corrupted. Resetting: $_" -Level ERROR
        return $DefaultConfig
    }
}

# Load config into global variable
$Config = Load-Config

# Helper: Get config value
function Get-ConfigValue {
    param(
        [Parameter(Mandatory)][string]$Path
    )
    return $Config.$Path
}

# Helper: Set config value
function Set-ConfigValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Value
    )
    $Config.$Path = $Value
    Save-Config -Config $Config
}

#endregion CONFIG SYSTEM
#region LOGGING SYSTEM
<#
    Logging System
    - Timestamped log entries
    - Automatic log rotation
    - Session markers
    - Integrated with Write-Log helper
#>

# Maximum log size before rotation (2 MB)
$MaxLogSizeBytes = 2MB

function Initialize-Log {
    if (Test-Path $LogFile) {
        $size = (Get-Item $LogFile).Length
        if ($size -ge $MaxLogSizeBytes) {
            $backup = "$LogFile.old"
            try {
                Move-Item -Path $LogFile -Destination $backup -Force
                Write-Host "Log rotated (size exceeded 2MB)" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Failed to rotate log file." -ForegroundColor Red
            }
        }
    }

    # Start a new session entry
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogFile -Value "============================================================"
    Add-Content -Path $LogFile -Value "[$timestamp] [SESSION START]"
}

function Close-LogSession {
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $LogFile -Value "[$timestamp] [SESSION END]"
    Add-Content -Path $LogFile -Value ""
}

# Override Write-Log to include rotation checks
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )

    # Rotate if needed
    if (Test-Path $LogFile) {
        $size = (Get-Item $LogFile).Length
        if ($size -ge $MaxLogSizeBytes) {
            $backup = "$LogFile.old"
            try {
                Move-Item -Path $LogFile -Destination $backup -Force
                Write-Host "Log rotated (size exceeded 2MB)" -ForegroundColor Yellow
            }
            catch {
                Write-Host "Failed to rotate log file." -ForegroundColor Red
            }
        }
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    Add-Content -Path $LogFile -Value $line
}

# Initialize log at script start
Initialize-Log

#endregion LOGGING SYSTEM
#region PROCESS TOOLS
<#
    Process Tools
    - Kill built-in apps
    - Kill custom apps
    - Restart built-in apps
    - Restart custom apps
    - Safe wrappers around process control
#>

function Stop-ProcessSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (Test-IsExceptionProcess -Name $Name) {
        Write-Log "Skip kill for exception process: $Name"
        return
    }

    try {
        $proc = Get-Process -Name $Name -ErrorAction SilentlyContinue
        if ($proc) {
            # Capture the process name to state memory before stopping it
            if ($Name -notin $Global:PreSessionRunningApps) {
                $Global:PreSessionRunningApps += $Name
                Write-Log "Pre-session app captured: $Name" -Level DEBUG
            }
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
            Write-Log "Killed process: $Name"
            Write-Success "Killed: $Name"
        }
    }
    catch {
        Write-Log "Failed to kill process ${Name}: $_" -Level ERROR
        Write-Warn "Could not kill: $Name"
    }
}

function Test-IsExceptionProcess {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    if (-not $Config.ContainsKey('Exception')) { return $false }

    $normalized = ($Name -replace '\.exe$','').ToLowerInvariant()
    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add($normalized)

    switch ($normalized) {
        'edge' { $candidates.Add('msedge'); $candidates.Add('microsoftedge') }
        'msedge' { $candidates.Add('edge'); $candidates.Add('microsoftedge') }
        'microsoftedge' { $candidates.Add('edge'); $candidates.Add('msedge') }
        'ccleaner' { $candidates.Add('ccleaner64') }
        'ccleaner64' { $candidates.Add('ccleaner') }
        'icloud' { $candidates.Add('icloudservices'); $candidates.Add('iclouddrive') }
        'icloudservices' { $candidates.Add('icloud'); $candidates.Add('iclouddrive') }
        'iclouddrive' { $candidates.Add('icloud'); $candidates.Add('icloudservices') }
    }

    foreach ($custom in $Config.Exception.Custom) {
        if ([string]::IsNullOrWhiteSpace($custom)) { continue }
        $customNormalized = ($custom -replace '\.exe$','').ToLowerInvariant()
        if ($candidates -contains $customNormalized) {
            return $true
        }
    }

    switch ($normalized) {
        'code' { return [bool]$Config.Exception.VSCode }
        'spotify' { return [bool]$Config.Exception.Spotify }
        'steam' { return [bool]$Config.Exception.Steam }
        'steamwebhelper' { return [bool]$Config.Exception.Steam }
        'discord' { return [bool]$Config.Exception.Discord }
        'ts3client_win64' { return [bool]$Config.Exception.TeamSpeak }
        'ts3client_win32' { return [bool]$Config.Exception.TeamSpeak }
        'teamspeak' { return [bool]$Config.Exception.TeamSpeak }
        'teamspeak_client' { return [bool]$Config.Exception.TeamSpeak }
        'teamspeak5' { return [bool]$Config.Exception.TeamSpeak }
        'teamspeak6' { return [bool]$Config.Exception.TeamSpeak }
    }

    return $false
}

function Test-ProcessRunning {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    return ($null -ne (Get-Process -Name $Name -ErrorAction SilentlyContinue))
}

function Get-ProcessNameFromCommand {
    param(
        [Parameter(Mandatory)]
        [string]$Command
    )

    if ([string]::IsNullOrWhiteSpace($Command)) { return $null }

    $leaf = Split-Path -Leaf $Command
    if ([string]::IsNullOrWhiteSpace($leaf)) {
        $leaf = $Command
    }

    return ($leaf -replace '\.exe$','')
}

function Start-ProcessSafe {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [string]$Args = ""
    )

    try {
        if ($Command -ieq 'explorer.exe') {
            if ($Args) {
                Start-Process -FilePath 'explorer.exe' -ArgumentList $Args | Out-Null
            }
            else {
                Start-Process -FilePath 'explorer.exe' | Out-Null
            }
        }
        elseif (Test-Path $Command) {
            $argLine = "`"$Command`""
            if ($Args) {
                $argLine = "$argLine $Args"
            }
            # Launch via Explorer to avoid inheriting elevated token for restored user apps.
            Start-Process -FilePath 'explorer.exe' -ArgumentList $argLine | Out-Null
        }
        else {
            if ($Args) {
                Start-Process -FilePath $Command -ArgumentList $Args | Out-Null
            }
            else {
                Start-Process -FilePath $Command | Out-Null
            }
        }
        Write-Log "Started process (user context): $Command $Args"
        Write-Success "Started: $Command"
    }
    catch {
        Write-Log "Failed to start ${Command}: $_" -Level ERROR
        Write-Warn "Could not start: $Command"
    }
}

# ----------------------------
# Display session controls - WinAPI + DisplaySwitch approach
# ----------------------------

function Initialize-DisplayInterop {
    if ('DisplayNative' -as [type]) { return }

    $displayNativeCode = @"
using System;
using System.Runtime.InteropServices;

public static class DisplayNative
{
    public const int ENUM_CURRENT_SETTINGS = -1;
    public const int ENUM_REGISTRY_SETTINGS = -2;
    public const int EDS_RAWMODE = 0x00000002;

    public const int CDS_UPDATEREGISTRY = 0x00000001;
    public const int CDS_TEST = 0x00000002;

    public const int DISP_CHANGE_SUCCESSFUL = 0;

    public const int DM_PELSWIDTH = 0x00080000;
    public const int DM_PELSHEIGHT = 0x00100000;
    public const int DM_DISPLAYFREQUENCY = 0x00400000;
    public const int DM_BITSPERPEL = 0x00040000;

    public const int DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = 0x00000001;
    public const int DISPLAY_DEVICE_PRIMARY_DEVICE = 0x00000004;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DEVMODE
    {
        private const int CCHDEVICENAME = 32;
        private const int CCHFORMNAME = 32;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;

        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;

        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DISPLAY_DEVICE
    {
        public int cb;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceString;
        public int StateFlags;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceID;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceKey;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplaySettingsEx(string lpszDeviceName, int iModeNum, ref DEVMODE lpDevMode, int dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int ChangeDisplaySettingsEx(string lpszDeviceName, ref DEVMODE lpDevMode, IntPtr hwnd, int dwflags, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);
}
"@

    Add-Type -TypeDefinition $displayNativeCode -Language CSharp
}

function New-DevMode {
    $dm = New-Object DisplayNative+DEVMODE
    $dm.dmSize = [System.Runtime.InteropServices.Marshal]::SizeOf([type]([DisplayNative+DEVMODE]))
    return $dm
}

function Get-DisplayDevices {
    $devices = @()

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $screens = [System.Windows.Forms.Screen]::AllScreens
        foreach ($screen in $screens) {
            $devices += ,([ordered]@{
                DeviceName   = [string]$screen.DeviceName
                DeviceString = [string]$screen.DeviceName
                IsPrimary    = [bool]$screen.Primary
                StateFlags   = 0
            })
        }
    }
    catch {
        Write-Log "Get-DisplayDevices: Screen API failed, falling back to EnumDisplayDevices: $_" -Level WARN
    }

    if (@($devices).Count -gt 0) {
        return @($devices)
    }

    for ($index = 0; $index -lt 16; $index++) {
        $dd = New-Object DisplayNative+DISPLAY_DEVICE
        $dd.cb = [System.Runtime.InteropServices.Marshal]::SizeOf([type]([DisplayNative+DISPLAY_DEVICE]))

        $ok = [DisplayNative]::EnumDisplayDevices($null, [uint32]$index, [ref]$dd, 0)
        if (-not $ok) { break }

        $attached = (($dd.StateFlags -band [DisplayNative]::DISPLAY_DEVICE_ATTACHED_TO_DESKTOP) -ne 0)
        if (-not $attached) { continue }

        $isPrimary = (($dd.StateFlags -band [DisplayNative]::DISPLAY_DEVICE_PRIMARY_DEVICE) -ne 0)
        $devices += ,([ordered]@{
            DeviceName   = [string]$dd.DeviceName
            DeviceString = [string]$dd.DeviceString
            IsPrimary    = $isPrimary
            StateFlags   = $dd.StateFlags
        })
    }

    return @($devices)
}

function Get-CurrentDisplayMode {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceName
    )

    $dm = New-DevMode
    $ok = [DisplayNative]::EnumDisplaySettingsEx($DeviceName, [DisplayNative]::ENUM_CURRENT_SETTINGS, [ref]$dm, 0)
    if (-not $ok) { return $null }

    return [ordered]@{
        Width       = [int]$dm.dmPelsWidth
        Height      = [int]$dm.dmPelsHeight
        Frequency   = [int]$dm.dmDisplayFrequency
        BitsPerPel  = [int]$dm.dmBitsPerPel
    }
}

function Get-LowestDisplayMode {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceName,
        [int]$PreferredWidth = 800,
        [int]$PreferredHeight = 600
    )

    $modeIndex = 0
    $allModes = @()

    while ($true) {
        $dm = New-DevMode
        $ok = [DisplayNative]::EnumDisplaySettingsEx($DeviceName, $modeIndex, [ref]$dm, [DisplayNative]::EDS_RAWMODE)
        if (-not $ok) { break }

        if ($dm.dmPelsWidth -ge 640 -and $dm.dmPelsHeight -ge 480) {
            $allModes += ,([ordered]@{
                Width      = [int]$dm.dmPelsWidth
                Height     = [int]$dm.dmPelsHeight
                Frequency  = [int]$dm.dmDisplayFrequency
                BitsPerPel = [int]$dm.dmBitsPerPel
            })
        }

        $modeIndex++
    }

    if ($allModes.Count -eq 0) { return $null }

    $dedup = @{}
    foreach ($m in $allModes) {
        $key = "{0}x{1}@{2}x{3}" -f $m.Width, $m.Height, $m.Frequency, $m.BitsPerPel
        if (-not $dedup.ContainsKey($key)) {
            $dedup[$key] = $m
        }
    }

    $modes = @($dedup.Values)
    $preferred = $modes | Where-Object { $_.Width -eq $PreferredWidth -and $_.Height -eq $PreferredHeight } | Select-Object -First 1
    if ($preferred) {
        return $preferred
    }

    return $modes |
        Sort-Object `
            @{ Expression = { $_.Width * $_.Height }; Ascending = $true },
            @{ Expression = { $_.Width }; Ascending = $true },
            @{ Expression = { $_.Height }; Ascending = $true },
            @{ Expression = { if ($_.Frequency -gt 0) { $_.Frequency } else { 9999 } }; Ascending = $true } |
        Select-Object -First 1
}

function Set-DisplayMode {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceName,
        [Parameter(Mandatory)]
        [int]$Width,
        [Parameter(Mandatory)]
        [int]$Height,
        [int]$Frequency = 0,
        [int]$BitsPerPel = 0
    )

    $dm = New-DevMode

    $currentOk = [DisplayNative]::EnumDisplaySettingsEx($DeviceName, [DisplayNative]::ENUM_CURRENT_SETTINGS, [ref]$dm, 0)
    if (-not $currentOk) {
        Write-Log "Set-DisplayMode: Could not read current mode for $DeviceName" -Level WARN
        return $false
    }

    $dm.dmPelsWidth = $Width
    $dm.dmPelsHeight = $Height
    $dm.dmFields = [DisplayNative]::DM_PELSWIDTH -bor [DisplayNative]::DM_PELSHEIGHT

    if ($Frequency -gt 0) {
        $dm.dmDisplayFrequency = $Frequency
        $dm.dmFields = $dm.dmFields -bor [DisplayNative]::DM_DISPLAYFREQUENCY
    }

    if ($BitsPerPel -gt 0) {
        $dm.dmBitsPerPel = $BitsPerPel
        $dm.dmFields = $dm.dmFields -bor [DisplayNative]::DM_BITSPERPEL
    }

    $testResult = [DisplayNative]::ChangeDisplaySettingsEx($DeviceName, [ref]$dm, [IntPtr]::Zero, [DisplayNative]::CDS_TEST, [IntPtr]::Zero)
    if ($testResult -ne [DisplayNative]::DISP_CHANGE_SUCCESSFUL) {
        Write-Log "Set-DisplayMode: Test failed for $DeviceName to ${Width}x${Height}. Result=$testResult" -Level WARN
        return $false
    }

    $applyResult = [DisplayNative]::ChangeDisplaySettingsEx($DeviceName, [ref]$dm, [IntPtr]::Zero, [DisplayNative]::CDS_UPDATEREGISTRY, [IntPtr]::Zero)
    if ($applyResult -ne [DisplayNative]::DISP_CHANGE_SUCCESSFUL) {
        Write-Log "Set-DisplayMode: Apply failed for $DeviceName to ${Width}x${Height}. Result=$applyResult" -Level WARN
        return $false
    }

    Write-Log "Set-DisplayMode: Applied ${Width}x${Height} to $DeviceName (Hz=$Frequency, Bpp=$BitsPerPel)."
    return $true
}

function Invoke-DisplaySwitch {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('/internal','/extend','/external')]
        [string]$Mode
    )

    $exe = Join-Path $env:WINDIR 'System32\DisplaySwitch.exe'
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) {
        $exe = 'DisplaySwitch.exe'
    }

    try {
        Start-Process -FilePath $exe -ArgumentList $Mode -WindowStyle Hidden -Wait
        Start-Sleep -Seconds 3
        Write-Log "DisplaySwitch applied: $Mode"
        return $true
    }
    catch {
        Write-Log "DisplaySwitch failed for mode ${Mode}: $_" -Level WARN
        return $false
    }
}

function Save-DisplaySessionState {
    $state = [ordered]@{
        OriginalModes = @{}
        SecondMonitorApplied = $false
        LowQualityApplied = $false
    }

    try {
        Initialize-DisplayInterop
        $devices = @(Get-DisplayDevices)
        foreach ($device in $devices) {
            $mode = Get-CurrentDisplayMode -DeviceName $device.DeviceName
            if ($null -ne $mode) {
                $state.OriginalModes[$device.DeviceName] = $mode
                Write-Log "Captured display mode $($device.DeviceName): $($mode.Width)x$($mode.Height) @ $($mode.Frequency)Hz"
            }
        }
    }
    catch {
        Write-Log "Failed to capture original display state: $_" -Level WARN
    }
    
    $Global:DisplaySessionState = $state
}

function Invoke-SecondMonitorSessionChange {
    if (-not $Config.Display.DisableSecondMonitor) {
        return
    }

    try {
        Initialize-DisplayInterop
        $activeDisplays = @(Get-DisplayDevices)
        $activeCountBefore = @($activeDisplays).Count

        if ($activeCountBefore -lt 2) {
            Write-Log 'Second monitor disable requested, but fewer than two active displays are attached.' -Level WARN
            return
        }

        $changed = Invoke-DisplaySwitch -Mode '/internal'

        $activeCountAfter = @((Get-DisplayDevices)).Count
        if (($activeCountAfter -ge 2) -and $changed) {
            Write-Log 'DisplaySwitch /internal did not reduce active displays; trying /external fallback.' -Level WARN
            $changed = Invoke-DisplaySwitch -Mode '/external'
            $activeCountAfter = @((Get-DisplayDevices)).Count
        }

        if ($changed) {
            Write-Log "Display topology reduced from $activeCountBefore to $activeCountAfter active display(s)."
            if ($null -ne $Global:DisplaySessionState) {
                $Global:DisplaySessionState.SecondMonitorApplied = $true
            }
        }
        else {
            Write-Log 'Failed to disable secondary displays via DisplaySwitch.' -Level WARN
        }
    }
    catch {
        Write-Log "Failed to apply second-monitor change: $_" -Level ERROR
    }
}

function Invoke-LowestQualitySessionChange {
    if (-not $Config.Display.LowestQuality) {
        return
    }

    try {
        Initialize-DisplayInterop
        $devices = @(Get-DisplayDevices)
        if (@($devices).Count -eq 0) {
            Write-Log 'No active display devices found for low-quality mode.' -Level WARN
            return
        }

        Write-Log "Low-quality mode: found $(@($devices).Count) active display(s)."

        $changedCount = 0

        foreach ($device in $devices) {
            $targetMode = Get-LowestDisplayMode -DeviceName $device.DeviceName
            if ($null -eq $targetMode) {
                Write-Log "No suitable low mode found for $($device.DeviceName)." -Level WARN
                continue
            }

            Write-Log "Low-quality target for $($device.DeviceName): $($targetMode.Width)x$($targetMode.Height) @ $($targetMode.Frequency)Hz"
            $applied = Set-DisplayMode -DeviceName $device.DeviceName -Width $targetMode.Width -Height $targetMode.Height -Frequency $targetMode.Frequency -BitsPerPel $targetMode.BitsPerPel
            if ($applied) {
                $changedCount++
            }
        }

        if ($changedCount -gt 0) {
            if ($null -ne $Global:DisplaySessionState) {
                $Global:DisplaySessionState.LowQualityApplied = $true
            }
            Write-Log "Applied low-quality display mode on $changedCount display(s)."
        }
        else {
            Write-Log 'Low-quality mode requested, but no display mode changes were applied.' -Level WARN
        }
    }
    catch {
        Write-Log "Failed to apply low-quality display change: $_" -Level ERROR
    }
}

function Invoke-DisplaySessionPrep {
    Write-Info "Preparing display session for VR..."
    Save-DisplaySessionState
    Invoke-SecondMonitorSessionChange
    Invoke-LowestQualitySessionChange
}

function Restore-DisplaySessionState {
    if ($null -eq $Global:DisplaySessionState) {
        return
    }

    $state = $Global:DisplaySessionState
    
    try {
            Initialize-DisplayInterop

            if ($state.SecondMonitorApplied) {
                $restoredTopology = Invoke-DisplaySwitch -Mode '/extend'
                if ($restoredTopology) {
                    Write-Log 'Restored secondary displays via DisplaySwitch /extend.'
                }
                else {
                    Write-Log 'Failed to restore secondary displays via DisplaySwitch /extend.' -Level WARN
                }
            }

            if ($state.OriginalModes.Count -gt 0) {
                foreach ($deviceName in $state.OriginalModes.Keys) {
                    $mode = $state.OriginalModes[$deviceName]
                    $ok = Set-DisplayMode -DeviceName $deviceName -Width $mode.Width -Height $mode.Height -Frequency $mode.Frequency -BitsPerPel $mode.BitsPerPel
                    if (-not $ok) {
                        Write-Log "Failed to restore display mode for $deviceName" -Level WARN
                }
            }
        }
    }
    catch {
        Write-Log "Display restoration encountered an error: $_" -Level WARN
    }
    finally {
        $Global:DisplaySessionState = $null
    }
}

# ------------------------------------------------------------
# Built-in KILL actions
# ------------------------------------------------------------
function Invoke-KillBuiltIn {
    Write-Info "Applying built-in kill rules..."

    if ($Config.Kill.OneDrive)       { Stop-ProcessSafe -Name "OneDrive" }
    if ($Config.Kill.Edge)           { Stop-ProcessSafe -Name "msedge" }
    if ($Config.Kill.CCleaner)       { Stop-ProcessSafe -Name "CCleaner64" }
    if ($Config.Kill.iCloudServices) { Stop-ProcessSafe -Name "iCloudServices" }
    if ($Config.Kill.iCloudDrive)    { Stop-ProcessSafe -Name "iCloudDrive" }
    if ($Config.Kill.Discord) 		 { Stop-ProcessSafe -Name "Discord" }
}

# ------------------------------------------------------------
# Custom KILL actions
# ------------------------------------------------------------
function Invoke-KillCustom {
    if ($Config.Kill.Custom.Count -eq 0) { return }

    Write-Info "Applying custom kill rules..."

    foreach ($procName in $Config.Kill.Custom) {
        if ([string]::IsNullOrWhiteSpace($procName)) { continue }
        Stop-ProcessSafe -Name $procName
    }
}

# ------------------------------------------------------------
# Built-in RESTART actions
# ------------------------------------------------------------
function Invoke-RestartBuiltIn {
    Write-Info "Applying built-in restart rules..."

    if ($Config.Restart.Edge) {
        # Only restart if enabled AND either RestoreOnlyActiveApps is false OR Edge was previously running
        if ($Config.RestoreOnlyActiveApps -eq $false -or 'msedge' -in $Global:PreSessionRunningApps) {
            $edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
            if (Test-IsExceptionProcess -Name 'msedge') {
                Write-Log "Skip restart for exception app: Edge"
            }
            elseif (Test-Path $edgePath) {
                if (Test-ProcessRunning -Name 'msedge') {
                    Write-Log "Skip restart for Edge (already running)."
                }
                else {
                    Start-ProcessSafe -Command $edgePath
                }
            }
        }
    }

    if ($Config.Restart.Discord) {
        # Only restart if enabled AND either RestoreOnlyActiveApps is false OR Discord was previously running
        if ($Config.RestoreOnlyActiveApps -eq $false -or 'Discord' -in $Global:PreSessionRunningApps) {
            $discordUpdater = Join-Path $env:LOCALAPPDATA "Discord\Update.exe"
            if (Test-IsExceptionProcess -Name 'Discord') {
                Write-Log "Skip restart for exception app: Discord"
            }
            elseif (Test-Path $discordUpdater) {
                if (Test-ProcessRunning -Name 'Discord') {
                    Write-Log "Skip restart for Discord (already running)."
                }
                else {
                    Start-ProcessSafe -Command $discordUpdater -Args "--processStart Discord.exe"
                }
            }
        }
    }

    if ($Config.Restart.OneDrive) {
        # Only restart if enabled AND either RestoreOnlyActiveApps is false OR OneDrive was previously running
        if ($Config.RestoreOnlyActiveApps -eq $false -or 'OneDrive' -in $Global:PreSessionRunningApps) {
            $oneDrivePath = "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"
            if (-not (Test-Path $oneDrivePath)) { 
                $oneDrivePath = "C:\Program Files\Microsoft OneDrive\OneDrive.exe" 
            }

            if (Test-IsExceptionProcess -Name 'OneDrive') {
                Write-Log "Skip restart for exception app: OneDrive"
            }
            elseif (Test-Path $oneDrivePath) {
                if (Test-ProcessRunning -Name 'OneDrive') {
                    Write-Log "Skip restart for OneDrive (already running)."
                }
                else {
                    # automatic start via Explorer
                    # /background silent mode
                    Start-Process "explorer.exe" -ArgumentList "`"$oneDrivePath`" /background"

                    Write-Log "OneDrive started via Explorer (non-elevated)."
                    Write-Success "OneDrive succesful started."
                }
            } else {
                Write-Log "OneDrive Pfad not found." -Level WARN
                Write-Warn "OneDrive executable not found."
            }
        }
    }

    if ($Config.Restart.CCleaner) {
        # Only restart if enabled AND either RestoreOnlyActiveApps is false OR CCleaner was previously running
        if ($Config.RestoreOnlyActiveApps -eq $false -or 'CCleaner64' -in $Global:PreSessionRunningApps) {
            $ccleaner = "C:\Program Files\CCleaner\CCleaner64.exe"
            if (Test-IsExceptionProcess -Name 'CCleaner64') {
                Write-Log "Skip restart for exception app: CCleaner"
            }
            elseif (Test-Path $ccleaner) {
                if (Test-ProcessRunning -Name 'CCleaner64') {
                    Write-Log "Skip restart for CCleaner (already running)."
                }
                else {
                    Start-ProcessSafe -Command $ccleaner -Args "/MONITOR"
                }
            }
        }
    }

    if ($Config.Restart.iCloud) {
        # Only restart if enabled AND either RestoreOnlyActiveApps is false OR iCloud was previously running
        if ($Config.RestoreOnlyActiveApps -eq $false -or 'iCloud' -in $Global:PreSessionRunningApps) {
            $storePath = "C:\Program Files\WindowsApps\AppleInc.iCloud_*"
            $desktopPath = "C:\Program Files (x86)\Common Files\Apple\Internet Services\iCloud.exe"

            if (Test-IsExceptionProcess -Name 'iCloud') {
                Write-Log "Skip restart for exception app: iCloud"
            }
            elseif (Test-ProcessRunning -Name 'iCloud') {
                Write-Log "Skip restart for iCloud (already running)."
            }
            else {
                if (Get-ChildItem $storePath -ErrorAction SilentlyContinue) {
                    Start-ProcessSafe -Command "explorer.exe" -Args "shell:AppsFolder\AppleInc.iCloud_skh98v6769f6t!iCloud"
                }
                elseif (Test-Path $desktopPath) {
                    Start-ProcessSafe -Command $desktopPath
                }
            }
        }
    }
}

# ------------------------------------------------------------
# Custom RESTART actions
# ------------------------------------------------------------
function Invoke-RestartCustom {
    if ($Config.Restart.Custom.Count -eq 0) { return }

    Write-Info "Applying custom restart rules..."

    foreach ($entry in $Config.Restart.Custom) {
        if (-not $entry.Command) { continue }
        $procName = Get-ProcessNameFromCommand -Command $entry.Command
        if ($procName -and (Test-IsExceptionProcess -Name $procName)) {
            Write-Log "Skip restart for exception custom app '$($entry.Command)' (process '$procName')."
            continue
        }
        if ($procName -and (Test-ProcessRunning -Name $procName)) {
            Write-Log "Skip restart for custom app '$($entry.Command)' (already running as process '$procName')."
            continue
        }
        Start-ProcessSafe -Command $entry.Command -Args $entry.Args
    }
}

#endregion PROCESS TOOLS

#region CPU AFFINITY SYSTEM
<#
    CPU Affinity System
    - Universal Intel 12th gen+ and AMD Ryzen support
    - Detects P-core threads automatically
    - Applies affinity mask to simulator process
#>

function Get-PCoreAffinityMask {
    # Query CPU topology
    $cpu = Get-CimInstance Win32_Processor

    $logical  = $cpu.NumberOfLogicalProcessors
    $physical = $cpu.NumberOfCores

    # P-core threads = physical cores * 2 (SMT)
    $pThreads = $physical * 2

    # Safety clamp
    if ($pThreads -gt $logical) {
        $pThreads = $logical
    }

    # Build mask: first $pThreads bits = 1
    $mask = 0
    for ($i = 0; $i -lt $pThreads; $i++) {
        $mask = $mask -bor (1 -shl $i)
    }

    return $mask
}

function Set-PCoreAffinity {
    param(
        [Parameter(Mandatory)][string]$ProcessName
    )

    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $proc) {
        Write-Log "Set-PCoreAffinity: Process $ProcessName not found" -Level WARN
        return
    }

    $mask = Get-PCoreAffinityMask

    try {
        $proc.ProcessorAffinity = $mask
        $binary = [Convert]::ToString($mask, 2)
        Write-Log "Applied P-core affinity mask ($mask) binary=[$binary] to $ProcessName"
        Write-Info "Applied P-core affinity to $ProcessName"
    }
    catch {
        Write-Log "Failed to apply affinity to ${ProcessName}: $_" -Level ERROR
        Write-Warn "Failed to apply CPU affinity."
    }
}

function Set-HighPriority {
    param(
        [Parameter(Mandatory)][string]$ProcessName
    )

    $proc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
    if (-not $proc) {
        Write-Log "{Set-HighPriority}: Process ${ProcessName} not found" -Level WARN
        return
    }

    try {
        $proc.PriorityClass = "High"
        Write-Log "Set process priority to HIGH for $ProcessName"
        Write-Info "Set HIGH priority for $ProcessName"
    }
    catch {
        Write-Log "Failed to set HIGH priority for ${ProcessName}: $_" -Level ERROR
        Write-Warn "Failed to set HIGH priority."
    }
}

#endregion CPU AFFINITY SYSTEM

#region POWER PLAN TOOLS
<#
    Power Plan Tools
    - Detect active plan
    - Switch to Ultimate Performance
    - Restore previous plan
    - Logging integration
#>

# GUIDs for known power plans
$PowerPlans = @{
    HighPerformance     = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
    Balanced            = "381b4222-f694-41f0-9685-ff5bb260df2e"
    UltimatePerformance = "e9a42b02-d5df-448d-aa00-03f14749eb61"
}

# ------------------------------------------------------------
# Get the currently active power plan GUID
# ------------------------------------------------------------
function Get-ActivePowerPlan {
    try {
        $output = powercfg /getactivescheme
        if ($output -match 'GUID:\s+([a-fA-F0-9\-]+)') {
            return $Matches[1]
        }
    }
    catch {
        Write-Log "Failed to read active power plan: $_" -Level ERROR
    }
    return $null
}

# ------------------------------------------------------------
# Switch to a specific power plan
# ------------------------------------------------------------
function Set-PowerPlan {
    param(
        [Parameter(Mandatory)]
        [string]$Guid
    )

    try {
        powercfg /setactive $Guid | Out-Null
        Write-Log "Power plan switched to $Guid"
        Write-Info "Power plan set to: $Guid"
    }
    catch {
        Write-Log "Failed to set power plan ${Guid}: $_" -Level ERROR
        Write-Warn "Could not switch power plan."
    }
}

# ------------------------------------------------------------
# Ensure Ultimate Performance is active
# ------------------------------------------------------------
function Ensure-UltimatePerformance {
    Write-Info "Switching to Ultimate Performance power plan..."

    $ultimate = $PowerPlans.UltimatePerformance

    # Check if Ultimate Performance exists
    $plans = powercfg /list
    if ($plans -notmatch $ultimate) {
        Write-Warn "Ultimate Performance plan not found. Attempting to enable it..."
        try {
            powercfg -duplicatescheme $ultimate | Out-Null
            Write-Log "Ultimate Performance plan duplicated/created."
        }
        catch {
            Write-Warn "Failed to create Ultimate Performance plan. Falling back to High Performance."
            Set-PowerPlan -Guid $PowerPlans.HighPerformance
            return
        }
    }

    # Activate Ultimate Performance
    Set-PowerPlan -Guid $ultimate
}

# ------------------------------------------------------------
# Restore previous power plan
# ------------------------------------------------------------
function Restore-PowerPlan {
    param(
        [Parameter(Mandatory)]
        [string]$PreviousGuid
    )

    Write-Info "Restoring previous power plan..."
    Set-PowerPlan -Guid $PreviousGuid
}

#endregion POWER PLAN TOOLS
#region SYSTEM PREP
<#
    System Prep
    - Kill built-in apps
    - Kill custom apps
    - Stop services
    - Enable NVIDIA persistence mode
    - Flush DNS
    - Launch Virtual Desktop Streamer
    - Logging integration
#>

function Stop-ServicesForVR {
    Write-Info "Stopping unnecessary services..."

    $services = @(
        "SysMain",
        "Spooler"
    )

    foreach ($svc in $services) {
        try {
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            Write-Log "Stopped service: $svc"
        }
        catch {
            Write-Log "Failed to stop service ${svc}: $_" -Level WARN
        }
    }
}

function Start-ServicesAfterVR {
    Write-Info "Restoring system services..."

    $services = @(
        "SysMain",
        "Spooler"
    )

    foreach ($svc in $services) {
        try {
            Start-Service -Name $svc -ErrorAction SilentlyContinue
            Write-Log "Started service: $svc"
        }
        catch {
            Write-Log "Failed to start service ${svc}: $_" -Level WARN
        }
    }
}

function Enable-NvidiaPersistence {
    Write-Info "Enabling NVIDIA persistence mode..."
    try {
        nvidia-smi -pm 1 | Out-Null
        Write-Log "NVIDIA persistence mode enabled."
    }
    catch {
        Write-Log "Failed to enable NVIDIA persistence mode: $_" -Level WARN
    }
}

function Disable-NvidiaPersistence {
    Write-Info "Disabling NVIDIA persistence mode..."
    try {
        nvidia-smi -pm 0 | Out-Null
        Write-Log "NVIDIA persistence mode disabled."
    }
    catch {
        Write-Log "Failed to disable NVIDIA persistence mode: $_" -Level WARN
    }
}

function Flush-DNS {
    Write-Info "Flushing DNS..."
    try {
        ipconfig /flushdns | Out-Null
        Write-Log "DNS flushed."
    }
    catch {
        Write-Log "Failed to flush DNS: $_" -Level WARN
    }
}

function Launch-VirtualDesktopStreamer {
    $vdPath = "C:\Program Files\Virtual Desktop Streamer\VirtualDesktop.Streamer.exe"

    if (Test-Path $vdPath) {
        Write-Info "Launching Virtual Desktop Streamer..."
        try {
            Start-Process -FilePath $vdPath | Out-Null
            Write-Log "Virtual Desktop Streamer launched."
            Start-Sleep -Seconds 8
        }
        catch {
            Write-Log "Failed to launch Virtual Desktop Streamer: $_" -Level WARN
        }
    }
}

# ------------------------------------------------------------
# MAIN PREP FUNCTION
# ------------------------------------------------------------
function Invoke-SystemPrep {
    Write-Info "Running system preparation steps..."
    Write-Log "System prep started."

    # Kill apps
    Invoke-KillBuiltIn
    Invoke-KillCustom

    # Stop services
    Stop-ServicesForVR

    # NVIDIA persistence
    Enable-NvidiaPersistence

    # DNS flush
    Flush-DNS

    # Launch VR streamer
    Launch-VirtualDesktopStreamer

    Write-Log "System prep completed."
}

#endregion SYSTEM PREP
#region SIMULATOR LAUNCHER
<#
    Simulator Launcher
    - Steam launch
    - Microsoft Store / GamePass launch
    - Standalone DCS
    - Standalone X-Plane
    - Process detection
    - CPU priority + affinity
    - Logging integration
#>

# ------------------------------------------------------------
# GAME LIBRARY (loaded from games.json)
# ------------------------------------------------------------
function Load-Games {
    if (-not (Test-Path $GamesPath)) {
        Write-ErrorUI "games.json not found at $GamesPath"
        Write-Log "games.json not found at $GamesPath" -Level ERROR
        return @()
    }

    try {
        $json = Get-Content -Path $GamesPath -Raw
        $games = $json | ConvertFrom-Json
        Write-Log "Game library loaded from $GamesPath ($($games.Count) entries)"
        return $games
    }
    catch {
        Write-ErrorUI "Failed to parse games.json. Check the file for syntax errors."
        Write-Log "Failed to parse games.json: $_" -Level ERROR
        return @()
    }
}

function Get-GameById {
    param(
        [Parameter(Mandatory)]
        [string]$Id
    )
    return $Global:GamesList | Where-Object { $_.Id -eq $Id } | Select-Object -First 1
}

# Load game library at script start
$Global:GamesList = Load-Games

# ------------------------------------------------------------
# Resolve Store URI
# ------------------------------------------------------------
function Get-StoreURI {
    param(
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    try {
        $pkg = Get-AppxPackage |
            Where-Object { $_.Name -match $Pattern } |
            Select-Object -First 1

        if ($pkg) {
            return "shell:AppsFolder\$($pkg.PackageFamilyName)!App"
        }
    }
    catch {
        Write-Log "Failed to resolve Store URI: $_" -Level WARN
    }

    return $null
}

# ------------------------------------------------------------
# Launch Steam sim
# ------------------------------------------------------------
function Launch-SteamSim {
    param(
        [Parameter(Mandatory)]
        [string]$AppId
    )

    Write-Info "Launching Steam simulator..."
    Write-Log "Launching Steam appid $AppId"

    try {
        # Use Explorer URI dispatch so Steam launches in normal user context.
        Start-ProcessSafe -Command 'explorer.exe' -Args "steam://run/$AppId"
    }
    catch {
        Write-Log "Failed to launch Steam app ${AppId}: $_" -Level ERROR
        Write-ErrorUI "Failed to launch Steam simulator."
    }
}

# ------------------------------------------------------------
# Launch Store sim
# ------------------------------------------------------------
function Launch-StoreSim {
    param(
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    Write-Info "Launching Microsoft Store simulator..."
    $uri = Get-StoreURI -Pattern $Pattern

    if (-not $uri) {
        Write-ErrorUI "Could not resolve Store app. Is it installed?"
        Write-Log "Store app not found for pattern $Pattern" -Level ERROR
        return
    }

    Write-Log "Launching Store URI: $uri"

    try {
        Start-Process "explorer.exe" $uri
    }
    catch {
        Write-Log "Failed to launch Store sim: $_" -Level ERROR
        Write-ErrorUI "Failed to launch Store simulator."
    }
}

# ------------------------------------------------------------
# Launch DCS Standalone
# ------------------------------------------------------------
function Launch-DCSStandalone {
    Write-Info "Launching DCS Standalone..."

    # First check if a custom path is provided in games.json
    $gameEntry = $Global:GamesList | Where-Object { $_.Id -eq "7" } | Select-Object -First 1
    $possiblePaths = @()
    
    if ($gameEntry -and $gameEntry.Path) {
        Write-Log "Custom DCS path found in games.json: $($gameEntry.Path)" -Level DEBUG
        if (Test-Path $gameEntry.Path) {
            $customItem = Get-Item $gameEntry.Path -ErrorAction SilentlyContinue
            if ($customItem -and $customItem.PSIsContainer) {
                $possiblePaths += (Join-Path $gameEntry.Path 'DCS.exe')
            }
            else {
                $possiblePaths += $gameEntry.Path
            }
        }
    }
    
    # Always include standard paths as fallback
    $possiblePaths += @(
        "C:\Program Files\Eagle Dynamics\DCS World\bin\DCS.exe",
        "C:\Eagle Dynamics\DCS World\bin\DCS.exe",
        "C:\DCS World\bin\DCS.exe"
    )

    # Search all available filesystem drives
    foreach ($drive in (Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name)) {
        $root = "$drive`:"

        $path1 = Join-Path $root 'Eagle Dynamics\DCS World\bin\DCS.exe'
        if (Test-Path $path1) { $possiblePaths += $path1 }

        $path2 = Join-Path $root 'DCS World\bin\DCS.exe'
        if (Test-Path $path2) { $possiblePaths += $path2 }

        $path3 = Join-Path $root 'Program Files\Eagle Dynamics\DCS World\bin\DCS.exe'
        if (Test-Path $path3) { $possiblePaths += $path3 }
    }

    $exe = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $exe) {
        Write-ErrorUI "DCS Standalone not found."
        Write-Log "DCS Standalone not found. Searched paths: $($possiblePaths -join '; ')" -Level ERROR
        return
    }

    Write-Log "Found DCS Standalone at: $exe" -Level DEBUG
    Write-Log "Launching DCS Standalone: $exe"

    try {
        $workDir = Split-Path -Parent $exe
        Write-Log "Working directory: $workDir" -Level DEBUG
        Start-Process -FilePath $exe -WorkingDirectory $workDir
        Write-Success "DCS Standalone launched"
        Write-Log "DCS Standalone process started successfully"
        Start-Sleep -Seconds 3
    }
    catch {
        Write-Log "Failed to launch DCS Standalone: $_" -Level ERROR
        Write-ErrorUI "Failed to launch DCS Standalone: $_"
    }
}

# ------------------------------------------------------------
# Launch X-Plane Standalone
# ------------------------------------------------------------
function Launch-XPlaneStandalone {
    Write-Info "Launching X-Plane 12 Standalone..."

    $paths = @()

    foreach ($drive in 'C'..'J') {
        $candidate = "$drive`:\X-Plane 12\X-Plane.exe"
        if (Test-Path $candidate) { $paths += $candidate }
    }

    $exe = $paths | Select-Object -First 1

    if (-not $exe) {
        Write-ErrorUI "X-Plane 12 Standalone not found."
        Write-Log "X-Plane Standalone not found. Searched paths: $($paths -join '; ')" -Level ERROR
        return
    }

    Write-Log "Found X-Plane Standalone at: $exe" -Level DEBUG
    Write-Log "Launching X-Plane Standalone: $exe"

    try {
        $workDir = Split-Path -Parent $exe
        Write-Log "Working directory: $workDir" -Level DEBUG
        Start-Process -FilePath $exe -WorkingDirectory $workDir
        Write-Success "X-Plane Standalone launched"
        Write-Log "X-Plane Standalone process started successfully"
        Start-Sleep -Seconds 3
    }
    catch {
        Write-Log "Failed to launch X-Plane Standalone: $_" -Level ERROR
        Write-ErrorUI "Failed to launch X-Plane Standalone: $_"
    }
}

function Close-SteamForNonSteamLaunch {
    if (Test-IsExceptionProcess -Name 'steam') {
        Write-Log "Skip closing Steam for non-Steam launch because Steam is marked as an exception."
        return
    }

    $Global:SteamWasRunningBeforeNonSteamLaunch = Test-ProcessRunning -Name 'steam'
    $Global:SteamClosedForNonSteamLaunch = $false

    if (-not $Global:SteamWasRunningBeforeNonSteamLaunch) {
        Write-Log "Steam is not running; no close needed for non-Steam launch."
        return
    }

    Write-Info "Non-Steam title selected. Closing Steam processes..."
    Write-Log "Closing Steam processes for non-Steam launch."

    foreach ($name in @('steam', 'steamwebhelper')) {
        Stop-ProcessSafe -Name $name
    }

    $Global:SteamClosedForNonSteamLaunch = -not (Test-ProcessRunning -Name 'steam')
    if ($Global:SteamClosedForNonSteamLaunch) {
        Write-Log "Steam closed for non-Steam launch; it will be restored at session end."
    }
}

function Restore-SteamAfterNonSteamLaunch {
    if (-not $Global:SteamWasRunningBeforeNonSteamLaunch) {
        return
    }

    if (-not $Global:SteamClosedForNonSteamLaunch) {
        return
    }

    if (Test-ProcessRunning -Name 'steam') {
        Write-Log "Skip Steam restore (already running)."
        return
    }

    Write-Info "Restoring Steam (normal user context)..."
    Write-Log "Restoring Steam after non-Steam launch."
    Start-ProcessSafe -Command 'explorer.exe' -Args 'steam://open/main'
}

# ------------------------------------------------------------
# Detect running sim process
# ------------------------------------------------------------
function Wait-ForSimProcess {
    param(
        [Parameter(Mandatory)]
        [string]$ExeName
    )

    Write-Info "Waiting for simulator process to start..."
    Write-Log "Waiting for process: $ExeName"

    for ($i = 1; $i -le 7; $i++) {
        $proc = Get-Process -Name ($ExeName -replace ".exe","") -ErrorAction SilentlyContinue
        if ($proc) {
            Write-Success "Simulator detected: $ExeName"
            Write-Log "Simulator detected: $ExeName"
            return $proc
        }

        Write-Host "  Attempt $i/7..."
        Start-Sleep -Seconds 5
    }

    Write-ErrorUI "Simulator did not start."
    Write-Log "Simulator failed to start." -Level ERROR
    return $null
}

function Wait-ForDcsSession {
    param(
        [Parameter(Mandatory)]
        [string]$ExeName
    )

    $procName = ($ExeName -replace '.exe','')
    $sessionObserved = $false
    $goneSince = $null
    $goneGraceSeconds = 20
    $pollSeconds = 3
    $startTime = Get-Date
    $maxSessionHours = 12

    Write-Info "Monitoring DCS launcher/game lifecycle..."
    Write-Log "DCS monitor started. ProcessName=$procName GoneGraceSeconds=$goneGraceSeconds"

    while ($true) {
        $all = @()
        try {
            $all = @(Get-Process -Name $procName -ErrorAction SilentlyContinue)
        }
        catch {
            Write-Log "DCS monitor process query failed: $_" -Level WARN
        }

        if ($all.Count -gt 0) {
            if (-not $sessionObserved) {
                Write-Log "DCS process detected. Session considered active."
            }
            $sessionObserved = $true
            $goneSince = $null
        }

        if ($all.Count -eq 0) {
            if (-not $goneSince) {
                $goneSince = Get-Date
                if ($sessionObserved) {
                    Write-Log "DCS process temporarily gone. Starting exit grace timer (${goneGraceSeconds}s)."
                }
                else {
                    Write-Log "No DCS process detected yet. Waiting with grace timer (${goneGraceSeconds}s)." -Level WARN
                }
            }

            $elapsed = ((Get-Date) - $goneSince).TotalSeconds
            if ($elapsed -ge $goneGraceSeconds) {
                if ($sessionObserved) {
                    Write-Log "DCS session ended (process stayed absent past grace timer)."
                    return $true
                }

                Write-Warn "DCS closed before stable runtime (launcher exit edge case)."
                Write-Log "DCS session ended early: no process observed through grace timer." -Level WARN
                return $false
            }
        }

        if (((Get-Date) - $startTime).TotalHours -ge $maxSessionHours) {
            Write-Log "DCS monitor safety timeout reached (${maxSessionHours}h). Forcing restore." -Level WARN
            return $true
        }

        Start-Sleep -Seconds $pollSeconds
    }
}

# ------------------------------------------------------------
# MAIN LAUNCH FUNCTION
# ------------------------------------------------------------
function Launch-Simulator {
    param(
        [Parameter(Mandatory)]
        [string]$SimId
    )

    $sim = Get-GameById -Id $SimId
    if ($null -eq $sim) {
        Write-ErrorUI "Invalid simulator selection."
        return $null
    }

    Write-Info "Launching: $($sim.Name)"
    Write-Log "Launching simulator: $($sim.Name)"

    if ($sim.Method -ne 'Steam') {
        Close-SteamForNonSteamLaunch
    }

    switch ($sim.Method) {
        "Steam"            { Launch-SteamSim -AppId $sim.AppId }
        "Store"            { Launch-StoreSim -Pattern $sim.StorePattern }
        "DCSStandalone"    { Launch-DCSStandalone }
        "XPlaneStandalone" { Launch-XPlaneStandalone }
    }

    # Wait for process
    $proc = Wait-ForSimProcess -ExeName $sim.ExeName
    if (-not $proc) { return $null }
    # Apply HIGH priority 
    Set-HighPriority -ProcessName ($sim.ExeName -replace ".exe","")
    # Apply P-core affinity 
    Set-PCoreAffinity -ProcessName ($sim.ExeName -replace ".exe","")

    return $proc
}

#endregion SIMULATOR LAUNCHER
#region RESTORE LOGIC
<#
    Restore Logic
    - Restore services
    - Restore power plan
    - Disable NVIDIA persistence mode
    - Restart built-in apps
    - Restart custom apps
    - Logging integration
#>

function Invoke-SystemRestore {
    param(
        [Parameter(Mandatory)]
        [string]$PreviousPowerPlan
    )

    Write-Info "Restoring system state..."
    Write-Log "System restore started."

    # Log pre-session app state
    if ($Global:PreSessionRunningApps.Count -gt 0) {
        $appList = $Global:PreSessionRunningApps -join ', '
        Write-Info "Restoring pre-session apps: $appList"
        Write-Log "Pre-session apps to restore: $appList"
    } else {
        Write-Info "No pre-session apps to restore (RestoreOnlyActiveApps: $($Config.RestoreOnlyActiveApps))"
        Write-Log "No pre-session apps captured."
    }

    # Restore services
    Start-ServicesAfterVR

    # Disable NVIDIA persistence mode
    Disable-NvidiaPersistence

    # Restore display state changed for the session
    Restore-DisplaySessionState

    # Restore previous power plan
    if ($PreviousPowerPlan) {
        Restore-PowerPlan -PreviousGuid $PreviousPowerPlan
    }
    else {
        Write-Warn "Previous power plan unknown - skipping restore."
        Write-Log "Previous power plan missing; restore skipped." -Level WARN
    }

    # Restart built-in apps
    Invoke-RestartBuiltIn

    # Restore Steam only if this session closed it for a non-Steam launch
    Restore-SteamAfterNonSteamLaunch

    # Restart custom apps
    Invoke-RestartCustom

    # Reset per-run Steam state flags
    $Global:SteamWasRunningBeforeNonSteamLaunch = $false
    $Global:SteamClosedForNonSteamLaunch = $false

    Write-Log "System restore completed."
    Write-Success "System restored."
}

#endregion RESTORE LOGIC
#region MENUS & MAIN FLOW
<#
    Menus & Main Flow
    - Main menu
    - Sim selection
    - Config menu
    - Custom app management
    - Launch + prep + restore orchestration
#>

function Show-MainMenu {
    Clear-Host
    Show-Box -Title "VR AUTO-OPTIMIZER - MAIN MENU"

    Write-White "  1) Launch Simulator (manual selection)"
    Write-White "  2) Configure App Controls"
    Write-White ""
    Write-White "  X) Exit"
}

function Show-SimMenu {
    Clear-Host
    Show-Box -Title "SELECT YOUR SIMULATOR"

    foreach ($game in $Global:GamesList) {
        Write-White "  $($game.Id)) $($game.Name)"
    }
    Write-White ""
    Write-White "  B) Back"
    Write-White "  X) Exit"
}

function Toggle-Flag {
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $parts = $Path.Split('.')
    
    if ($parts.Count -eq 2) {
        # Nested property like "Kill.OneDrive"
        $section = $parts[0]
        $key     = $parts[1]
        $current = $Config[$section][$key]
        
        if ($current -is [bool]) {
            $Config[$section][$key] = -not $current
            Save-Config -Config $Config
        }
    }
    elseif ($parts.Count -eq 1) {
        # Top-level property like "RestoreOnlyActiveApps"
        $key     = $parts[0]
        $current = $Config[$key]
        
        if ($current -is [bool]) {
            $Config[$key] = -not $current
            Save-Config -Config $Config
        }
    }
}


function Show-ConfigMenu {
    while ($true) {
        Clear-Host
        Show-Box -Title "CONFIGURATION - APP CONTROLS"

        Write-White "  Kill Flags:"
        Write-White "    [1] OneDrive        = $($Config.Kill.OneDrive)"
        Write-White "    [2] Edge            = $($Config.Kill.Edge)"
        Write-White "    [3] CCleaner        = $($Config.Kill.CCleaner)"
        Write-White "    [4] iCloudServices  = $($Config.Kill.iCloudServices)"
        Write-White "    [5] iCloudDrive     = $($Config.Kill.iCloudDrive)"
        Write-White "    [6] Discord         = $($Config.Kill.Discord)"
        Write-Host ""

        Write-White "  Restart Flags:"
        Write-White "    [7] Restart Edge     = $($Config.Restart.Edge)"
        Write-White "    [8] Restart Discord  = $($Config.Restart.Discord)"
        Write-White "    [9] Restart OneDrive = $($Config.Restart.OneDrive)"
        Write-White "    [10] Restart CCleaner = $($Config.Restart.CCleaner)"
        Write-White "    [11] Restart iCloud   = $($Config.Restart.iCloud)"
        Write-Host ""

        Write-White "  Exception Flags (high priority - never kill/restart):"
        Write-White "    [12] VS Code   = $($Config.Exception.VSCode)"
        Write-White "    [13] Spotify   = $($Config.Exception.Spotify)"
        Write-White "    [14] Steam     = $($Config.Exception.Steam)"
        Write-White "    [15] Discord   = $($Config.Exception.Discord)"
        Write-White "    [16] TeamSpeak = $($Config.Exception.TeamSpeak)"
        Write-Host ""

        Write-White "  D) Set default sim (current: $($Config.DefaultSim))"
        Write-White "  A) Toggle auto-run on start (AutoRunOnStart = $($Config.AutoRunOnStart))"
        Write-White "  R) Toggle restore only active apps (RestoreOnlyActiveApps = $($Config.RestoreOnlyActiveApps))"
        Write-White "  M) Disable second monitor (Display.DisableSecondMonitor = $($Config.Display.DisableSecondMonitor))"
        Write-White "  L) Lowest quality display mode (Display.LowestQuality = $($Config.Display.LowestQuality))"
        Write-White ""
        Write-White "  C) Manage custom apps"
        Write-White "  E) Manage exception apps"
        Write-White "  S) Save and return"
        Write-White "  B) Back without saving"

        $choice = Read-Choice -Prompt "Selection:"
        switch -Regex ($choice) {
            '^1$' { Toggle-Flag 'Kill.OneDrive' }
            '^2$' { Toggle-Flag 'Kill.Edge' }
            '^3$' { Toggle-Flag 'Kill.CCleaner' }
            '^4$' { Toggle-Flag 'Kill.iCloudServices' }
            '^5$' { Toggle-Flag 'Kill.iCloudDrive' }
            '^6$' { Toggle-Flag 'Kill.Discord' }
            '^7$' { Toggle-Flag 'Restart.Edge' }
            '^8$' { Toggle-Flag 'Restart.Discord' }
            '^9$' { Toggle-Flag 'Restart.OneDrive' }
            '^10$' { Toggle-Flag 'Restart.CCleaner' }
            '^11$' { Toggle-Flag 'Restart.iCloud' }
            '^12$' { Toggle-Flag 'Exception.VSCode' }
            '^13$' { Toggle-Flag 'Exception.Spotify' }
            '^14$' { Toggle-Flag 'Exception.Steam' }
            '^15$' { Toggle-Flag 'Exception.Discord' }
            '^16$' { Toggle-Flag 'Exception.TeamSpeak' }
            '^[mM]$' { Toggle-Flag 'Display.DisableSecondMonitor' }
            '^[lL]$' { Toggle-Flag 'Display.LowestQuality' }
            '^[dD]$' { Set-DefaultSim }
            '^[aA]$' {
                $Config.AutoRunOnStart = -not $Config.AutoRunOnStart
                Save-Config -Config $Config
            }
            '^[rR]$' { Toggle-Flag 'RestoreOnlyActiveApps' }
            '^[cC]$' { Manage-CustomApps }
            '^[eE]$' { Manage-ExceptionApps }
            '^[sS]$' { Save-Config -Config $Config; return }
            '^[bB]$' { return }
        }
    }
}

function Set-DefaultSim {
    Show-SimMenu
    $sel = Read-Choice -Prompt "Enter default sim ID (or blank to cancel):"
    if ([string]::IsNullOrWhiteSpace($sel)) { return }
    if ($null -eq (Get-GameById -Id $sel)) {
        Write-ErrorUI "Invalid sim ID."
        Start-Sleep -Seconds 2
        return
    }
    $Config.DefaultSim = $sel
    Save-Config -Config $Config
}

function Manage-CustomApps {
    while ($true) {
        Clear-Host
        Show-Box -Title "CUSTOM APPS - KILL / RESTART"

        $killCustom = @($Config.Kill.Custom)
        $restartCustom = @($Config.Restart.Custom)

        Write-White "  Custom Kill List (process names):"
        if ($killCustom.Count -eq 0) {
            Write-White "    (none)"
        } else {
            $i = 1
            foreach ($p in $killCustom) {
                Write-White ("    [{0}] {1}" -f $i, $p)
                $i++
            }
        }
        Write-Host ""

        Write-White "  Custom Restart List (Command + Args):"
        if ($restartCustom.Count -eq 0) {
            Write-White "    (none)"
        } else {
            $i = 1
            foreach ($entry in $restartCustom) {
                Write-White ("    [{0}] {1} {2}" -f $i, $entry.Command, $entry.Args)
                $i++
            }
        }
        Write-Host ""

        Write-White "  1) Add custom kill process"
        Write-White "  2) Remove custom kill process"
        Write-White "  3) Add custom restart entry"
        Write-White "  4) Remove custom restart entry"
        Write-White ""
        Write-White "  B) Back"

        $choice = Read-Choice -Prompt "Selection:"
        switch ($choice) {
            '1' { Add-CustomKill }
            '2' { Remove-CustomKill }
            '3' { Add-CustomRestart }
            '4' { Remove-CustomRestart }
            'B' { return }
            'b' { return }
        }
    }
}

function Add-CustomKill {
    Write-White "  Add kill app via:"
    Write-White "    [1] Manual process name"
    Write-White "    [2] Browse for executable"
    $mode = Read-Choice -Prompt "Select option (blank to cancel):"

    $name = $null
    switch ($mode.Trim()) {
        '1' { $name = Read-Choice -Prompt "Enter process name to kill (without .exe):" }
        '2' {
            $selectedPath = Select-ExecutableFile -Title 'Select executable to add to custom kill list'
            if ([string]::IsNullOrWhiteSpace($selectedPath)) { return }
            $name = ConvertTo-NormalizedProcessName -InputValue $selectedPath
        }
        default { return }
    }

    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $normalized = ($name -replace '\.exe$','').ToLowerInvariant()
    $existing = @($Config.Kill.Custom | ForEach-Object { ($_ -replace '\.exe$','').ToLowerInvariant() })
    if ($normalized -in $existing) { return }

    $Config.Kill.Custom += ($name -replace '\.exe$','')
    Save-Config -Config $Config
}

function Remove-CustomKill {
    $killCustom = @($Config.Kill.Custom)
    if ($killCustom.Count -eq 0) { return }
    $idx = Read-Choice -Prompt "Enter index to remove:"
    $index = Get-RemovalIndex -Text $idx -MaxCount $killCustom.Count
    if ($null -eq $index) { return }
    $Config.Kill.Custom = Remove-ArrayItemAtIndex -Items $killCustom -IndexToRemove $index
    Save-Config -Config $Config
}

function Add-CustomRestart {
    Write-White "  Add restart app via:"
    Write-White "    [1] Manual command path"
    Write-White "    [2] Browse for executable"
    $mode = Read-Choice -Prompt "Select option (blank to cancel):"

    $cmd = $null
    switch ($mode.Trim()) {
        '1' { $cmd = Read-Choice -Prompt "Enter full command path:" }
        '2' {
            $cmd = Select-ExecutableFile -Title 'Select executable to add to custom restart list'
            if ([string]::IsNullOrWhiteSpace($cmd)) { return }
        }
        default { return }
    }

    if ([string]::IsNullOrWhiteSpace($cmd)) { return }
    $args = Read-Choice -Prompt "Enter arguments (optional):"

    $entry = [ordered]@{
        Command = $cmd
        Args    = $args
    }
    $Config.Restart.Custom += $entry
    Save-Config -Config $Config
}

function Remove-CustomRestart {
    $restartCustom = @($Config.Restart.Custom)
    if ($restartCustom.Count -eq 0) { return }
    $idx = Read-Choice -Prompt "Enter index to remove:"
    $index = Get-RemovalIndex -Text $idx -MaxCount $restartCustom.Count
    if ($null -eq $index) { return }
    $Config.Restart.Custom = Remove-ArrayItemAtIndex -Items $restartCustom -IndexToRemove $index
    Save-Config -Config $Config
}

function Manage-ExceptionApps {
    while ($true) {
        Clear-Host
        Show-Box -Title "EXCEPTION APPS - HIGH PRIORITY"

        $exceptionCustom = @($Config.Exception.Custom)

        Write-White "  These apps are never killed or restarted by the optimizer."
        Write-Host ""

        Write-White "  Custom Exception List (process names):"
        if ($exceptionCustom.Count -eq 0) {
            Write-White "    (none)"
        }
        else {
            $i = 1
            foreach ($p in $exceptionCustom) {
                Write-White ("    [{0}] {1}" -f $i, $p)
                $i++
            }
        }
        Write-Host ""

        Write-White "  1) Add custom exception process"
        Write-White "  2) Remove custom exception process"
        Write-White ""
        Write-White "  B) Back"

        $choice = Read-Choice -Prompt "Selection:"
        switch ($choice) {
            '1' { Add-ExceptionCustom }
            '2' { Remove-ExceptionCustom }
            'B' { return }
            'b' { return }
        }
    }
}

function Add-ExceptionCustom {
    Write-White "  Add exception app via:"
    Write-White "    [1] Manual process name"
    Write-White "    [2] Browse for executable"
    $mode = Read-Choice -Prompt "Select option (blank to cancel):"

    $name = $null
    switch ($mode.Trim()) {
        '1' { $name = Read-Choice -Prompt "Enter process name to EXCLUDE (without .exe):" }
        '2' {
            $selectedPath = Select-ExecutableFile -Title 'Select executable to add to exception list'
            if ([string]::IsNullOrWhiteSpace($selectedPath)) { return }
            $name = ConvertTo-NormalizedProcessName -InputValue $selectedPath
        }
        default { return }
    }

    if ([string]::IsNullOrWhiteSpace($name)) { return }

    $normalized = ($name -replace '\.exe$','').ToLowerInvariant()
    $existing = @($Config.Exception.Custom | ForEach-Object { ($_ -replace '\.exe$','').ToLowerInvariant() })
    if ($normalized -in $existing) { return }

    $Config.Exception.Custom += $name
    Save-Config -Config $Config
}

function Remove-ExceptionCustom {
    $exceptionCustom = @($Config.Exception.Custom)
    if ($exceptionCustom.Count -eq 0) { return }
    $idx = Read-Choice -Prompt "Enter index to remove:"
    $index = Get-RemovalIndex -Text $idx -MaxCount $exceptionCustom.Count
    if ($null -eq $index) { return }
    $Config.Exception.Custom = Remove-ArrayItemAtIndex -Items $exceptionCustom -IndexToRemove $index
    Save-Config -Config $Config
}

function Run-SimFlow {
    param(
        [Parameter(Mandatory)]
        [string]$SimId
    )

    # Capture current power plan
    $prevPlan = Get-ActivePowerPlan

    # Switch to Ultimate Performance
    Ensure-UltimatePerformance

    # System prep
    Invoke-SystemPrep

    # Apply display session changes
    Invoke-DisplaySessionPrep

    # Launch sim
    $proc = Launch-Simulator -SimId $SimId
    if (-not $proc) {
        Write-ErrorUI "Launch failed; restoring system..."
        Invoke-SystemRestore -PreviousPowerPlan $prevPlan
        return
    }

    Show-Box -Title "$((Get-GameById -Id $SimId).Name) RUNNING"
    Write-White "  Do not close this window while the simulator is running."
    Write-Host ""

    $sim = Get-GameById -Id $SimId

    try {
        if ($sim.Method -eq 'DCSStandalone') {
            $completed = Wait-ForDcsSession -ExeName $sim.ExeName
            if (-not $completed) {
                Write-Log "DCS session ended before stable game runtime." -Level WARN
            }
        }
        else {
            # Wait for sim exit
            while (-not $proc.HasExited) {
                Start-Sleep -Seconds 15
                try {
                    $proc.Refresh()
                } catch {
                    break
                }
            }
        }
    }
    catch {
        Write-Warn "Runtime monitor failed unexpectedly. Proceeding to restore."
        Write-Log "Runtime monitor failure for sim '$($sim.Name)': $_" -Level ERROR
    }
    finally {
        Write-Info "Simulator exited. Restoring system..."
        Invoke-SystemRestore -PreviousPowerPlan $prevPlan
    }
}

function Main-Loop {
    # Optional auto-run
    if ($Config.AutoRunOnStart -and $Config.DefaultSim) {
        Write-Info "Auto-run enabled. Launching default sim: $($Config.DefaultSim)"
        Run-SimFlow -SimId $Config.DefaultSim
    }

    while ($true) {
        Show-MainMenu
        $choice = Read-Choice -Prompt "Selection:"

        switch -Regex ($choice) {
            '^1$' {
                while ($true) {
                    Show-SimMenu
                    $sel = Read-Choice -Prompt "Selection (1-$($Global:GamesList.Count)/B/X):"
                    if ($sel -match '^[xX]$') { Close-LogSession; exit }
                    if ($sel -match '^[bB]$') { break }
                    if ($null -eq (Get-GameById -Id $sel)) {
                        Write-ErrorUI "Invalid selection."
                        Start-Sleep -Seconds 2
                        continue
                    }
                    Run-SimFlow -SimId $sel
                }
            }
            '^2$' { Show-ConfigMenu }
            '^[xX]$' { Close-LogSession; exit }
        }
    }
}

#endregion MENUS & MAIN FLOW

# Entry point
Main-Loop
