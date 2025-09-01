# Copyright 2023 Google Inc. All Rights Reserved.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Password Sync diagnostics tool
# VBScript version by Liron Newman <lironn@google.com>
# PowerShell conversion by Jules

#==============================================================================
# Script Parameters
#==============================================================================
param(
    # If provided, the script runs diagnostics on the specified Domain Controller.
    # If not provided, the script runs in orchestrator mode.
    [string]$DCName
)

#==============================================================================
# Initial Variables and Configuration
#==============================================================================

# Const Ver = "2.0.3.0"
$Ver = "2.0.3.0"
# Const ToolName = "PasswordSyncSupportTool"
$ToolName = "PasswordSyncSupportTool"

# Set objShell = WScript.CreateObject("Wscript.Shell")
$objShell = New-Object -ComObject Wscript.Shell

# On Error Resume Next  ' Errors will be handled by the code
$ErrorActionPreference = "Continue"

#==============================================================================
# Helper Functions
#==============================================================================

# Function GetCurrentTimeString()
function Get-CurrentTimeString {
    # GetCurrentTimeString = Year(Now) & Right("0" & Month(Now), 2) ...
    return Get-Date -Format "yyyyMMdd_HHmmss"
}

# Sub LogStr(str)
function Log-Str {
    param(
        [string]$Message,
        [switch]$SkipEcho # Used when logging output from jobs, which is already shown once.
    )
    # LogFile.WriteLine Now & " " & str
    $timestampedMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    # Use a global variable for the log file path, defined in the main script body.
    $timestampedMessage | Add-Content -Path $Global:LogFileName

    # WScript.Echo Now & " " & str
    if (-not $SkipEcho) {
        Write-Host $timestampedMessage
    }
}

# Function GetWritableDCs()
function Get-WritableDCs {
    Log-Str "A:Getting list of writable DCs..."
    $dcs = @()
    try {
        # Set conn = CreateObject("ADODB.Connection")
        $rootDSE = [adsi]"LDAP://RootDSE"
        $searchRoot = [adsi]"LDAP://$($rootDSE.defaultNamingContext)"

        # Query = ... "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192)(!(msDS-IsRodc=true)));dNSHostName;subtree"
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
        $searcher.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192)(!(msDS-IsRodc=true)))"
        $searcher.PropertiesToLoad.Add("dNSHostName") | Out-Null
        $searcher.PageSize = 1000

        $results = $searcher.FindAll()

        if ($results.Count -eq 0) {
            # LogStr "W:No DCs found - maybe msDS-IsRodc is missing from the schema (Windows 2003)? Trying without it."
            Log-Str "W:No writable DCs found with msDS-IsRodc attribute. Trying legacy query for Win2003."
            $searcher.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))"
            $results = $searcher.FindAll()
        }

        foreach ($result in $results) {
            $dcName = $result.Properties["dnshostname"][0]
            Log-Str "A:Found DC: $dcName"
            $dcs += $dcName
        }
    }
    catch {
        Log-Str "E:Error querying for Domain Controllers: $($_.Exception.Message)"
    }
    return $dcs
}

# Function CheckComputerAndUserDetails()
function Test-ComputerAndUserDetails {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem

        Log-Str "A:Computer Name: $($cs.Name)"
        Log-Str "A:Computer's Domain: $($cs.Domain)"
        Log-Str "A:Part Of Domain: $($cs.PartOfDomain)"

        if (-not $cs.PartOfDomain) {
            $msg = "This machine isn't part of a domain. Make sure you are logged in as a domain admin, and run this tool again."
            Log-Str "E:$msg"
            [System.Windows.MessageBox]::Show($msg, $ToolName, 'Ok', 'Error')
            return $false
        }

        $userName = $env:USERNAME
        $userDnsDomain = $env:USERDNSDOMAIN

        if ([string]::IsNullOrEmpty($userDnsDomain)) {
             $msg = "The logged in user ($userName) isn't a domain user. Make sure you are logged in as a domain admin, and run this tool again."
             Log-Str "E:$msg"
             [System.Windows.MessageBox]::Show($msg, $ToolName, 'Ok', 'Error')
             return $false
        }

        Log-Str "A:Current user's AD DNS domain: $($userDnsDomain.ToLower())"

        if ($cs.Domain.ToLower() -ne $userDnsDomain.ToLower()) {
            $msg = "The current user's DNS domain ($userDnsDomain) doesn't match the machine's DNS domain ($($cs.Domain)). This will cause issues."
            Log-Str "E:$msg"
            [System.Windows.MessageBox]::Show($msg, $ToolName, 'Ok', 'Error')
            return $false
        }
    }
    catch {
        Log-Str "E:Failed to check computer and user details: $($_.Exception.Message)"
        return $false
    }
    return $true
}

# Function CheckIfRunningAsDomainAdmin()
function Test-IsRunningAsDomainAdmin {
    $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)

    try {
        $domainSid = $identity.User.AccountDomainSid
        # Well-known RID for Domain Admins is 512
        $domainAdminsSid = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-21" + $domainSid.Value.Substring(7) + "-512")

        if ($principal.IsInRole($domainAdminsSid)) {
            Log-Str "A:The current user is a member of Domain Admins."
            return $true
        } else {
            $msg = "The current user isn't a member of the Domain Admins group. To successfully use this tool, you must be a Domain Admin."
            Log-Str "E:$msg"
            [System.Windows.MessageBox]::Show($msg, $ToolName, 'Ok', 'Error')
            return $false
        }
    }
    catch {
        $msg = "Could not determine if user is a Domain Admin. Error: $($_.Exception.Message)"
        Log-Str "E:$msg"
        [System.Windows.MessageBox]::Show($msg, $ToolName, 'Ok', 'Error')
        return $false
    }
}

# Sub PrintLine(Text)
function Print-Line {
    param([string]$Text)
    # WScript.StdOut.WriteLine Text
    Write-Output $Text
}

# Sub PrintErrorIfNeeded(Text)
function Print-ErrorIfNeeded {
    param([string]$Text)
    if ($LASTEXITCODE -ne 0 -or $Error) {
        $lastError = $Error[0] | Out-String
        Print-Line "E:$Text : $lastError"
        $Error.Clear()
    }
}

# Sub RunCommand(Command, OutputFileNameBase)
function Invoke-NativeCommand {
    param([string]$Command, [string]$OutputFileNameBase)

    Print-Line "Running command: $Command"

    $stdOutFile = "$($OutputFileNameBase).txt"
    $stdErrFile = "$($OutputFileNameBase).err"

    $process = Start-Process cmd.exe -ArgumentList "/c $Command" -NoNewWindow -Wait -RedirectStandardOutput $stdOutFile -RedirectStandardError $stdErrFile -PassThru

    if ($process.ExitCode -ne 0) {
        $errorContent = if (Test-Path $stdErrFile) { Get-Content $stdErrFile | Out-String } else { "Unknown error" }
        Print-Line "E:Running command '$Command' failed with exit code $($process.ExitCode). Error: $errorContent"
    }
}

# Sub RunCopyCommand(Source, Target)
function Invoke-CopyCommand {
    param([string]$Source, [string]$Target)

    Print-Line "Copying from '$Source' to '$Target'"
    try {
        # The VBScript's xcopy parameters /C /E /F /H /Y /I /G translate roughly to this:
        Copy-Item -Path $Source -Destination $Target -Recurse -Force -ErrorAction Stop
    }
    catch {
        Print-Line "E:Failed to copy from '$Source' to '$Target'. Error: $($_.Exception.Message)"
    }
}

# Sub DecodeWinHTTPSettings(CompName, OutputFileName)
function Decode-WinHTTPSettings {
    param([string]$CompName, [string]$OutputFileName)

    $LogLinePrefix = "Current WinHTTP proxy settings:`r`n`r`n"
    $OutputContent = ""

    try {
        $regKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
        $remoteReg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $CompName)
        $connectionsKey = $remoteReg.OpenSubKey($regKeyPath)
        $winHttpSettings = $connectionsKey.GetValue("WinHttpSettings")

        if ($winHttpSettings) {
            $proxyLength = $winHttpSettings[12]
            if ($proxyLength -gt 0) {
                $proxyServer = [System.Text.Encoding]::Unicode.GetString($winHttpSettings, 16, $proxyLength)
                $bypassListLength = $winHttpSettings[16 + $proxyLength]
                $bypassList = "(none)"
                if ($bypassListLength -gt 0) {
                    $bypassList = [System.Text.Encoding]::Unicode.GetString($winHttpSettings, 20 + $proxyLength, $bypassListLength)
                }
                $OutputContent = $LogLinePrefix + "    Proxy Server(s) :  $proxyServer`r`n" + "    Bypass List     :  $bypassList"
            }
            else {
                $OutputContent = $LogLinePrefix + "    Direct access (no proxy server)."
            }
        } else {
            $OutputContent = $LogLinePrefix + "    WinHttpSettings registry value not found."
        }
    }
    catch {
        $OutputContent = $LogLinePrefix + "    Error decoding WinHttpSettings: $($_.Exception.Message)"
    }

    $OutputContent | Out-File -FilePath $OutputFileName -Encoding "utf8"
}

# Sub RunDiagnostics(CompName)
function Run-Diagnostics {
    param(
        [string]$CompName
    )

    $ErrorActionPreference = 'Continue'
    Print-Line "Starting diagnostics on $CompName"

    $workDir = Join-Path $Global:TempDir $CompName
    Set-Location $workDir
    Print-ErrorIfNeeded "Error changing to work folder for this DC file"

    Invoke-NativeCommand -Command "reg query \\$CompName\HKLM\SYSTEM\CurrentControlSet\Control\Lsa /v ""Notification Packages""" -OutputFileNameBase "dll-reg"
    Invoke-NativeCommand -Command "tasklist /S $CompName /m password_sync_dll.dll" -OutputFileNameBase "dll-loaded"

    foreach ($serviceName in @("Google Apps Password Sync", "G Suite Password Sync", "Password Sync")) {
        $shortName = $serviceName.Split(' ')[0].ToLower()
        Print-Line "Getting status for service '$serviceName'"
        try {
            $scOutput = Get-Service -ComputerName $CompName -Name $serviceName -ErrorAction Stop | Select-Object * | Format-List | Out-String
            $scOutput += "`n"
            $scOutput += Get-CimInstance -ComputerName $CompName -Class Win32_Service -Filter "Name='$serviceName'" | Select-Object * | Format-List | Out-String
            $scOutput | Out-File -FilePath "service_$shortName.txt"
        }
        catch {
            Print-Line "I:Service '$serviceName' not found or error querying: $($_.Exception.Message)"
        }
    }

    Print-Line "Copying logs and XML..."
    $userHomePaths = @(
        "\\$CompName\c$\Users\%username%\AppData\Local",
        "\\$CompName\c$\Documents and Settings\%username%\Local Settings\Application Data"
    )
    $networkServicePaths = @(
        "\\$CompName\c$\Windows\ServiceProfiles\NetworkService\AppData\Local",
        "\\$CompName\c$\Documents and Settings\NetworkService\Local Settings\Application Data"
    )
    $systemProfilePaths = @(
        "\\$CompName\c$\WINDOWS\system32\config\systemprofile\AppData\Local",
        "\\$CompName\c$\WINDOWS\system32\config\systemprofile\Local Settings\Application Data"
    )
    $programDataPaths = @(
        "\\$CompName\c$\ProgramData",
        "\\$CompName\c$\Documents and Settings\All Users\Application Data"
    )

    Invoke-CopyCommand -Source "$($userHomePaths[0])\Google\Google Apps Password Sync\Tracing" -Target "UI"
    Invoke-CopyCommand -Source "$($userHomePaths[1])\Google\Google Apps Password Sync\Tracing" -Target "UI"
    Invoke-CopyCommand -Source "$($userHomePaths[0])\Google\Identity" -Target "Identity"
    Invoke-CopyCommand -Source "$($userHomePaths[1])\Google\Identity" -Target "Identity"
    Invoke-CopyCommand -Source "$($networkServicePaths[0])\Google\Google Apps Password Sync\Tracing\password_sync_service" -Target "Service"
    Invoke-CopyCommand -Source "$($networkServicePaths[1])\Google\Google Apps Password Sync\Tracing\password_sync_service" -Target "Service"
    Invoke-CopyCommand -Source "$($networkServicePaths[0])\Google\Identity" -Target "ServiceAuth"
    Invoke-CopyCommand -Source "$($networkServicePaths[1])\Google\Identity" -Target "ServiceAuth"
    Invoke-CopyCommand -Source "$($systemProfilePaths[0])\Google\Google Apps Password Sync\Tracing\lsass" -Target "DLL"
    Invoke-CopyCommand -Source "$($systemProfilePaths[1])\Google\Google Apps Password Sync\Tracing\lsass" -Target "DLL"
    Invoke-CopyCommand -Source "$($programDataPaths[0])\Google\Google Apps Password Sync\config.xml" -Target "."
    Invoke-CopyCommand -Source "$($programDataPaths[1])\Google\Google Apps Password Sync\config.xml" -Target "."

    Print-Line "Getting list of installed files..."
    $installPaths = @(
        "\\$CompName\c$\Program Files\Google\Google Apps Password Sync",
        "\\$CompName\c$\Program Files\Google\G Suite Password Sync",
        "\\$CompName\c$\Program Files\Google\Password Sync"
    )
    $installPathsX86 = @(
        "\\$CompName\c$\Program Files (x86)\Google\Google Apps Password Sync",
        "\\$CompName\c$\Program Files (x86)\Google\G Suite Password Sync",
        "\\$CompName\c$\Program Files (x86)\Google\Password Sync"
    )
    $installPaths | ForEach-Object { Get-ChildItem -Path $_ -Recurse -ErrorAction SilentlyContinue } | Select-Object -ExpandProperty FullName | Out-File -FilePath "install.txt" -Append
    $installPathsX86 | ForEach-Object { Get-ChildItem -Path $_ -Recurse -ErrorAction SilentlyContinue } | Select-Object -ExpandProperty FullName | Out-File -FilePath "instx86.txt" -Append

    Invoke-NativeCommand -Command "reg query ""\\$CompName\HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"" /v ProxySettingsPerUser" -OutputFileNameBase "proxy"
    Invoke-NativeCommand -Command "reg query ""\\$CompName\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"" /v WinHttpSettings" -OutputFileNameBase "winhttp"
    Invoke-NativeCommand -Command "reg query ""\\$CompName\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v Email" -OutputFileNameBase "admin-and-serviceaccount-emails"
    Invoke-NativeCommand -Command "reg query ""\\$CompName\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v ServiceAccountEmail" -OutputFileNameBase "admin-and-serviceaccount-emails"
    Invoke-NativeCommand -Command "reg query ""\\$CompName\HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"" /s" -OutputFileNameBase "tls-registry-settings"

    Decode-WinHTTPSettings -CompName $CompName -OutputFileName "winhttp_decoded.txt"

    Print-Line "Getting local time and last boot time from remote machine..."
    try {
        $osInfo = Get-CimInstance -ComputerName $CompName -ClassName Win32_OperatingSystem -ErrorAction Stop
        $localTime = $osInfo.ConvertToDateTime($osInfo.LocalDateTime)
        $lastBootTime = $osInfo.ConvertToDateTime($osInfo.LastBootUpTime)
        $timeZone = $osInfo.CurrentTimeZone / 60
        $timeZoneStr = if ($timeZone -ge 0) { "+$timeZone" } else { "$timeZone" }
        "Local Time: $localTime, Time Zone: $timeZoneStr" | Out-File "time.txt"
        "Last boot time: $lastBootTime" | Out-File "time.txt" -Append
    }
    catch {
        Print-Line "E:Error getting time from WMI: $($_.Exception.Message)"
    }

    Print-Line "Getting versions of important executables..."
    $filePaths = @(
        'c:\Windows\System32\lsass.exe',
        'c:\Windows\System32\password_sync_dll.dll',
        'c:\Program Files\Google\Google Apps Password Sync\GoogleAppsPasswordSync.exe',
        'c:\Program Files\Google\Google Apps Password Sync\password_sync_service.exe',
        'c:\Program Files\Google\Google Apps Password Sync\unifiedlogin.dll',
        'c:\Program Files\Google\Password Sync\PasswordSync.exe',
        'c:\Program Files\Google\Password Sync\password_sync_service.exe',
        'c:\Program Files\Google\Password Sync\unifiedlogin.dll',
        'c:\Program Files\Google\G Suite Password Sync\PasswordSync.exe',
        'c:\Program Files\Google\G Suite Password Sync\password_sync_service.exe',
        'c:\Program Files\Google\G Suite Password Sync\unifiedlogin.dll'
    )
    $whereClause = $filePaths.ForEach({ "Name='$($_ -replace '\\', '\\\\')'" }) -join ' OR '
    try {
        Get-CimInstance -ComputerName $CompName -ClassName CIM_Datafile -Filter $whereClause |
            Select-Object Name, Version |
            Format-Table |
            Out-File "file_versions.txt"
    }
    catch {
        Print-Line "E:Error getting file versions via WMI: $($_.Exception.Message)"
    }

    Print-Line "Finished diagnostics on $CompName"
}


#==============================================================================
# Main Script Execution
#==============================================================================

# Check for administrative privileges and re-launch if necessary.
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID)

if (-not ($myWindowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))) {
    Write-Warning "Administrator privileges are required. Attempting to re-launch with elevation..."

    $ArgumentList = @("-NoProfile", "-File", "`"$($MyInvocation.MyCommand.Path)`"")
    $ArgumentList += $MyInvocation.UnboundArguments
    $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object {
        $ArgumentList += "-$($_.Key)"
        if ($_.Value) { $ArgumentList += "`"$($_.Value)`"" }
    }

    Start-Process powershell.exe -Verb RunAs -ArgumentList $ArgumentList
    exit
}

# If -DCName is specified, this script is a child process for a single DC.
if ($PSBoundParameters.ContainsKey('DCName')) {
    $Global:TempDir = Join-Path $env:TEMP $ToolName
    Run-Diagnostics -CompName $DCName
    exit
}

# --- Main Orchestrator Logic ---

$Global:TempDir = Join-Path $env:TEMP $ToolName
$Global:LogFileName = Join-Path $Global:TempDir "$($ToolName).log"
$CurrentTimeString = Get-CurrentTimeString

if (Test-Path $Global:TempDir) {
    Remove-Item -Path $Global:TempDir -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -Path $Global:TempDir -ItemType Directory | Out-Null

Log-Str "A:Starting $ToolName version $Ver from $($MyInvocation.MyCommand.Path)"

if (-not (Test-ComputerAndUserDetails)) { exit }
if (-not (Test-IsRunningAsDomainAdmin)) { exit }

$arrWritableDCs = Get-WritableDCs
Log-Str "A:Got $($arrWritableDCs.Count) writable DCs"

$jobs = @()
foreach ($dc in $arrWritableDCs) {
    if (-not [string]::IsNullOrWhiteSpace($dc)) {
        $dcWorkDir = Join-Path $Global:TempDir $dc
        Log-Str "A:Creating $dcWorkDir"
        New-Item -Path $dcWorkDir -ItemType Directory | Out-Null

        Log-Str "A:Starting job for $dc"
        $job = Start-Job -FilePath $MyInvocation.MyCommand.Path -ArgumentList @{ DCName = $dc } -Name $dc
        $jobs += $job
    }
    else {
        Log-Str "W:Skipping empty DC name"
    }
}

Log-Str "A:Waiting for all diagnostic jobs to complete..."
Wait-Job -Job $jobs | Out-Null

Log-Str "A:All jobs completed. Collecting results."
foreach ($job in $jobs) {
    Log-Str "A:--- Results for job $($job.Name) (State: $($job.State)) ---"

    $output = Receive-Job -Job $job
    if ($output) {
        $output | ForEach-Object { Log-Str $_ -SkipEcho }
    }

    if ($job.State -eq 'Failed') {
        Log-Str "E:Job $($job.Name) failed."
        if ($job.Error) {
            $job.Error | ForEach-Object { Log-Str "E: $($_.Exception.Message)" -SkipEcho }
        }
    }
    else {
        Log-Str "A:Job $($job.Name) finished successfully."
    }

    Log-Str "A:--- End of results for job $($job.Name) ---"
    Remove-Job -Job $job
}

Log-Str "A:Finished collecting information, creating ZIP"

$NewFolderName = "${Global:TempDir}_${CurrentTimeString}"
Rename-Item -Path $Global:TempDir -NewName $NewFolderName

$ZipName = "$($ToolName)-report_${CurrentTimeString}.zip"
$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$zipPath = Join-Path $desktopPath $ZipName
Compress-Archive -Path "$NewFolderName\*" -DestinationPath $zipPath -Force

$Message = "Please send the file `"$ZipName`" from your Desktop to Google Support for investigation."
Write-Host "`n$Message"
[System.Windows.MessageBox]::Show($Message, $ToolName, 'Ok', 'Information')

Read-Host "Press Enter to close this window"
