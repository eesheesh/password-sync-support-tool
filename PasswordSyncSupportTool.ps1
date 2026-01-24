<#
' Copyright 2012 Google Inc. All Rights Reserved.
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

' Password Sync diagnostics tool
' Liron Newman lironn@google.com
#>

# Const Ver = "2.0.3.0"
$Ver = "2.0.3.0"

# Const ToolName = "PasswordSyncSupportTool"
$ToolName = "PasswordSyncSupportTool"

# Dim fso, objShell, objShellApplication, CurrentComputerName
# Set fso = WScript.CreateObject("Scripting.FileSystemObject")
# Set objShell = WScript.CreateObject("Wscript.Shell")
# Set objShellApplication = CreateObject("Shell.Application")
# Const ForReading = 1, ForWriting = 2, ForAppending = 8
# Const WshRunning = 0, WshFinished = 1, WshFailed = 2
$ForReading = 1
$ForWriting = 2
$ForAppending = 8
$WshRunning = 0
$WshFinished = 1
$WshFailed = 2

# Global variables to mimic VBS scope
$script:LogFileName = ""
$script:TempDir = ""
$script:CurrentTimeString = ""
$script:CurrentComputerName = ""
$script:CurrentScriptPath = $PSCommandPath


function Get-CurrentTimeString {
    # GetCurrentTimeString = _
    #     Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
    #     "_" & _
    #     Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & _
    #     Right("0" & Second(Now), 2)
    return (Get-Date -Format "yyyyMMdd_HHmmss")
}

function LogStr($str) {
    # Dim LogFile  ' As Stream
    # Set LogFile = fso.OpenTextFile(LogFileName, ForAppending, True)
    # ' TODO: Prettier date/time with DatePart(), and add timezone from http://social.technet.microsoft.com/Forums/en-US/ITCG/thread/daf4b666-fcb6-46ad-becc-689e6daf49ed
    # LogFile.WriteLine Now & " " & str
    $line = "$(Get-Date) $str"
    try {
        $line | Out-File -FilePath $script:LogFileName -Append -Encoding utf8
    } catch {
        # Fallback or ignore if log file inaccessible
    }
    # LogFile.Close
    # WScript.Echo Now & " " & str
    Write-Host $line
}

function LogErr {
    # LogErr = "Error #" & Err & " (hex 0x" & Right("00000000" & Hex(Err), 8) & _
    #          "), Source: " & Err.Source & ", Description: " & Err.Description
    if ($Error.Count -gt 0) {
        $e = $Error[0]
        return "Error #$($e.Exception.HResult) (hex 0x$($e.Exception.HResult.ToString('X8'))), Source: $($e.InvocationInfo.ScriptName), Description: $($e.Exception.Message)"
    }
    # Err.Clear
    $Error.Clear()
    return ""
}

function PrintLine($Text) {
    # WScript.StdOut.WriteLine Text
    Write-Output $Text
}

function PrintErrorIfNeeded($Text) {
    # If Err <> 0 Then PrintLine " E:" & Text & " " & LogErr
    if ($Error.Count -gt 0) {
        PrintLine " E:$Text $(LogErr)"
    }
}

function LogErrorIfNeeded($Text) {
    # If Err <> 0 Then LogStr "E:" & Text & ": " & LogErr
    if ($Error.Count -gt 0) {
        LogStr "E:${Text}: $(LogErr)"
    }
}

function ErrorMsgBox($Text) {
    # MsgBox "Error: " & Text & vbCrLf & vbCrLf & _
    #            "Please share this with Google Support.", _
    #        vbOKOnly Or vbExclamation, _
    #        ToolName
    if ($IsWindows) {
        # MsgBox is not available in core without loading assembly
        # [System.Windows.Forms.MessageBox]::Show("Error: $Text `n`nPlease share this with Google Support.", $ToolName, "OK", "Exclamation")
    } else {
        Write-Host "Error: $Text `n`nPlease share this with Google Support."
    }
}

function RunCommand($Command, $OutputFileNameBase) {
    # On Error Resume Next
    $ErrorActionPreference = "SilentlyContinue"

    # ' Always use bWaitOnReturn=True to make sure the subproccess returns after
    # ' all data was collected.
    # PrintLine "Running command: " & Command
    PrintLine "Running command: $Command"
    # objShell.Run "cmd /c " & Command & " 1>>" & OutputFileNameBase & ".txt " & _
    #                  "2>>" & OutputFileNameBase & ".err", _
    #              0, _
    #              True

    $outFile = "$OutputFileNameBase.txt"
    $errFile = "$OutputFileNameBase.err"

    # We use cmd /c for compatibility
    $processArgs = @{
        FilePath = "cmd"
        ArgumentList = "/c $Command 1>>""$outFile"" 2>>""$errFile"""
        Wait = $true
        PassThru = $true
    }
    if ($IsWindows) { $processArgs.WindowStyle = 'Hidden' }
    $p = Start-Process @processArgs

    # PrintErrorIfNeeded "Running command '" & Command & "' failed. "
    if ($p.ExitCode -ne 0) {
         if ($Error.Count -gt 0) {
             PrintErrorIfNeeded "Running command '$Command' failed. "
         }
    }
}

function RunCopyCommand($Source, $Target) {
    $localSource = $Source
    # ' Checking if we are copying from the local machine and current user.
    # ' If we are, use %userprofile% which is more reliable.
    # CurrentMachineAndUserPrefix2008 = _
    #     "\\" & CurrentComputerName & "\C$\USERS\%USERNAME%\"
    $CurrentMachineAndUserPrefix2008 = "\\$script:CurrentComputerName\C$\USERS\$env:USERNAME\"
    # CurrentMachineAndUserPrefix2003 = _
    #     "\\" & CurrentComputerName & "\C$\DOCUMENTS AND SETTINGS\%USERNAME%\"
    $CurrentMachineAndUserPrefix2003 = "\\$script:CurrentComputerName\C$\DOCUMENTS AND SETTINGS\$env:USERNAME\"

    # If UCase(Left(Source, Len(CurrentMachineAndUserPrefix2008))) = _
    #     CurrentMachineAndUserPrefix2008 Then
    if ($localSource.ToUpper().StartsWith($CurrentMachineAndUserPrefix2008)) {
    #   Source = "%userprofile%" & _
    #            Mid(Source, Len(CurrentMachineAndUserPrefix2008))
         $localSource = "%userprofile%" + $localSource.Substring($CurrentMachineAndUserPrefix2008.Length)
    # ElseIf UCase(Left(Source, Len(CurrentMachineAndUserPrefix2003))) = _
    #     CurrentMachineAndUserPrefix2003 Then
    } elseif ($localSource.ToUpper().StartsWith($CurrentMachineAndUserPrefix2003)) {
    #   Source = "%userprofile%" & _
    #            Mid(Source, Len(CurrentMachineAndUserPrefix2003))
         $localSource = "%userprofile%" + $localSource.Substring($CurrentMachineAndUserPrefix2003.Length)
    # End If
    }

    # RunCommand _
    #     "xcopy """ & Source & """ """ & Target & """ " & "/C /E /F /H /Y /I /G", _
    #     "copying"
    RunCommand "xcopy ""$localSource"" ""$Target"" /C /E /F /H /Y /I /G" "copying"
}

function DecodeWinHTTPSettings($CompName, $OutputFileName) {
    # On Error Resume Next
    $ErrorActionPreference = "SilentlyContinue"

    # LogLinePrefix = "Current WinHTTP proxy settings:" & vbCRLF & vbCRLF
    $LogLinePrefix = "Current WinHTTP proxy settings:`r`n`r`n"

    # ' Create a WMI StdRegProv.
    # Dim objStdRegProv
    # Set objStdRegProv = GetObject( _
    #     "winmgmts:{impersonationLevel=impersonate}!\\" & CompName & _
    #     "\root\default:StdRegProv")

    # PrintErrorIfNeeded "Error opening WMI StdRegProv on " & CompName & ": "

    # ' Retrieve the value of WinHTTPSettings from the registry.
    # ' Note that GetBinaryValue returns an array, where each element in the array
    # ' is a DECIMAL value of the octets.
    # Dim WinHTTPSettingsArray
    # Const HKEY_LOCAL_MACHINE = &H80000002  ' From https://msdn.microsoft.com/en-us/library/aa394600(v=vs.85).aspx?cs-lang=vb
    # objStdRegProv.GetBinaryValue HKEY_LOCAL_MACHINE, _
    #                              "SOFTWARE\Microsoft\Windows\CurrentVersion\" & _
    #                                  "Internet Settings\Connections", _
    #                              "WinHttpSettings", _
    #                              WinHTTPSettingsArray

    $WinHTTPSettingsArray = $null
    try {
        $WinHTTPSettingsArray = Invoke-CimMethod -ComputerName $CompName -Namespace "root/default" -ClassName "StdRegProv" -MethodName "GetBinaryValue" -Arguments @{
            hDefKey = [uint32]2147483650
            sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"
            sValueName = "WinHttpSettings"
        } | Select-Object -ExpandProperty uValue
    } catch {
         # PrintErrorIfNeeded _
         #     "Error retrieving HKLM\SOFTWARE\Microsoft\Windows\" & _
         #     "CurrentVersion\Internet Settings\Connections\WinHttpSettings on " & _
         #     CompName & ": "
         PrintErrorIfNeeded "Error retrieving HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections\WinHttpSettings on ${CompName}: "
         return
    }

    # ' The WinHttpSettings registry value appears to be formatted as follows:
    # ...
    # ' Start by getting the proxy string length.
    # Dim WinHTTPProxyLength
    # WinHTTPProxyLength = WinHTTPSettingsArray(12)
    $WinHTTPProxyLength = 0
    if ($WinHTTPSettingsArray) {
        $WinHTTPProxyLength = $WinHTTPSettingsArray[12]
    }

    # ' Prepare the output file.
    # Dim WinHTTPParsedFile
    # Set WinHTTPParsedFile = fso.OpenTextFile(OutputFileName, ForAppending, True)
    # PrintErrorIfNeeded "Error opening " & OutputFileName

    # ' If the proxy string length is greater than 0, a proxy is set. If not, the
    # ' connection is direct.
    $output = ""
    # If WinHTTPProxyLength > 0 Then
    if ($WinHTTPProxyLength -gt 0) {
        # Dim WinHTTPProxy, WinHTTPBypassList, WinHTTPBypassListLength
        $WinHTTPProxy = ""
        # ' Concatenate the proxy, starting from 16, through the proxy length.
        # For Index = 16 To (16 + WinHTTPProxyLength - 1)
        #   WinHTTPProxy = WinHTTPProxy & ChrW(WinHTTPSettingsArray(Index))
        # Next
        for ($i = 16; $i -lt (16 + $WinHTTPProxyLength); $i++) {
            $WinHTTPProxy += [char]$WinHTTPSettingsArray[$i]
        }

        # ' Get the bypass list string length. We know its position is 12 + 1 + 3 +
        # ' the length of the proxy string.
        # WinHTTPBypassListLength = WinHTTPSettingsArray((16 + WinHTTPProxyLength))
        $WinHTTPBypassListLength = $WinHTTPSettingsArray[(16 + $WinHTTPProxyLength)]

        # ' If the length of the list is greater than 0, concatenate it.
        $WinHTTPBypassList = ""
        # If WinHTTPBypassListLength > 0 Then
        if ($WinHTTPBypassListLength -gt 0) {
            # ' Start from 12 + 1 + 3 + proxy string length + 1 + 3.
            # For Index = (20 + WinHTTPProxyLength) To _
            #     (20 + WinHTTPProxyLength + WinHTTPBypassListLength - 1)
            #   WinHTTPBypassList = WinHTTPBypassList & _
            #                       ChrW(WinHTTPSettingsArray(Index))
            # Next
            for ($i = (20 + $WinHTTPProxyLength); $i -lt (20 + $WinHTTPProxyLength + $WinHTTPBypassListLength); $i++) {
                 $WinHTTPBypassList += [char]$WinHTTPSettingsArray[$i]
            }
        # Else
        } else {
            # WinHTTPBypassList = "(none)"
            $WinHTTPBypassList = "(none)"
        # End If
        }

        # PrintErrorIfNeeded "Error decoding WinHttpSettings on " & CompName & ": "
        # WinHTTPParsedFile.WriteLine LogLinePrefix & _
        #     "    Proxy Server(s) :  " & WinHTTPProxy & vbCRLF & _
        #     "    Bypass List     :  " & WinHTTPBypassList
        $output = "$LogLinePrefix    Proxy Server(s) :  $WinHTTPProxy`r`n    Bypass List     :  $WinHTTPBypassList"

    # Else
    } else {
        # WinHTTPParsedFile.WriteLine LogLinePrefix & _
        #     "    Direct access (no proxy server)."
        $output = "$LogLinePrefix    Direct access (no proxy server)."
    # End If
    }

    # PrintErrorIfNeeded "Error writing to " & OutputFileName
    # WinHTTPParsedFile.Close
    $output | Out-File -FilePath $OutputFileName -Append -Encoding utf8
}

function WMIDateStringToTime($strDate) {
  # On Error Resume Next
  # If Len(strDate) < 15 Then
  if ($strDate.Length -lt 15) {
    # WMIDateStringToTime = "**Can't parse WMI time string " & strDate & "**"
    return "**Can't parse WMI time string $strDate**"
  # Else
  } else {
    # WMIDateStringToTime = Left(strDate, 4) & "-" & _
    #                       Mid(strDate, 5, 2) & "-" & _
    #                       Mid(strDate, 7, 2) & " " & _
    #                       Mid (strDate, 9, 2) & ":" & _
    #                       Mid(strDate, 11, 2) & ":" & _
    #                       Mid(strDate, 13, 2)
    return $strDate.Substring(0, 4) + "-" + $strDate.Substring(4, 2) + "-" + $strDate.Substring(6, 2) + " " + $strDate.Substring(8, 2) + ":" + $strDate.Substring(10, 2) + ":" + $strDate.Substring(12, 2)
  # End If
  }
  # PrintErrorIfNeeded "Error converting WMI time: "
}

function RunDiagnostics($CompName) {
  # On Error Resume Next
  $ErrorActionPreference = "SilentlyContinue"

  # PrintLine "Starting diagnostics on " & CompName
  PrintLine "Starting diagnostics on $CompName"
  # objShell.CurrentDirectory = TempDir & "\" & CompName  ' Change current dir
  Set-Location "$script:TempDir\$CompName"
  # PrintErrorIfNeeded "Error changing to work folder for this DC file: "
  PrintErrorIfNeeded "Error changing to work folder for this DC file: "

  # PrintLine "Getting Notification Package DLL reg entry - dll-reg.txt"
  PrintLine "Getting Notification Package DLL reg entry - dll-reg.txt"
  # RunCommand "reg query \\" & CompName & _
  #                "\HKLM\SYSTEM\CurrentControlSet\Control\Lsa " & _
  #                "/v ""Notification Packages""", _
  #            "dll-reg"
  RunCommand "reg query \\$CompName\HKLM\SYSTEM\CurrentControlSet\Control\Lsa /v ""Notification Packages""" "dll-reg"

  # PrintLine "Running tasklist.exe to see if the DLL is loaded - dll-loaded.txt"
  PrintLine "Running tasklist.exe to see if the DLL is loaded - dll-loaded.txt"
  # RunCommand "tasklist /S " & CompName & " /m password_sync_dll.dll", _
  #            "dll-loaded"
  RunCommand "tasklist /S $CompName /m password_sync_dll.dll" "dll-loaded"

  # PrintLine "Getting service status - service_*.txt"
  PrintLine "Getting service status - service_*.txt"
  # RunCommand "(sc \\" & CompName & " query ""Google Apps Password Sync"" && " & _
  #                "sc \\" & CompName & " qc ""Google Apps Password Sync"")", _
  #            "service_gaps"
  RunCommand "(sc \\$CompName query ""Google Apps Password Sync"" && sc \\$CompName qc ""Google Apps Password Sync"")" "service_gaps"
  # RunCommand "(sc \\" & CompName & " query ""G Suite Password Sync"" && " & _
  #                "sc \\" & CompName & " qc ""G Suite Password Sync"")", _
  #            "service_gsps"
  RunCommand "(sc \\$CompName query ""G Suite Password Sync"" && sc \\$CompName qc ""G Suite Password Sync"")" "service_gsps"
  # RunCommand "(sc \\" & CompName & " query ""Password Sync"" && " & _
  #                "sc \\" & CompName & " qc ""Password Sync"")", _
  #            "service_password_sync"
  RunCommand "(sc \\$CompName query ""Password Sync"" && sc \\$CompName qc ""Password Sync"")" "service_password_sync"

  # ' Get logs (from default locations - v1) using XCOPY to get the full tree
  # ' Assume the username is the same as the current username for the UI logs.
  # ' It doesn't matter for the other paths (they don't depend on the username).
  # PrintLine "Copying logs and XML - copying.txt"
  PrintLine "Copying logs and XML - copying.txt"

  # ' C:\Users\username\AppData\Local\Google\Google Apps Password Sync\Tracing
  # RunCopyCommand "\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Google Apps Password Sync\Tracing", _
  #                "UI"
  RunCopyCommand "\\$CompName\c$\Users\%username%\AppData\Local\Google\Google Apps Password Sync\Tracing" "UI"

  # ' C:\Documents and Settings\username\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing
  # RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\%username%\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing", _
  #                "UI"
  RunCopyCommand "\\$CompName\c$\Documents and Settings\%username%\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing" "UI"

  # ' C:\Users\username\AppData\Local\Google\Identity
  # RunCopyCommand "\\" & CompName & "\c$\Users\%username%\AppData\Local\Google\Identity", _
  #                "Identity"
  RunCopyCommand "\\$CompName\c$\Users\%username%\AppData\Local\Google\Identity" "Identity"

  # ' C:\Documents and Settings\username\Local Settings\Application Data\Google\Identity
  # RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\username\Local Settings\Application Data\Google\Identity", _
  #                "Identity"
  RunCopyCommand "\\$CompName\c$\Documents and Settings\username\Local Settings\Application Data\Google\Identity" "Identity"

  # ' C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service
  # RunCopyCommand "\\" & CompName & "\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service", _
  #                "Service"
  RunCopyCommand "\\$CompName\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Google Apps Password Sync\Tracing\password_sync_service" "Service"

  # 'C:\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service
  # RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service", _
  #                "Service"
  RunCopyCommand "\\$CompName\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\password_sync_service" "Service"

  # ' C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Identity
  # RunCopyCommand "\\" & CompName & "\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Identity", _
  #                "ServiceAuth"
  RunCopyCommand "\\$CompName\c$\Windows\ServiceProfiles\NetworkService\AppData\Local\Google\Identity" "ServiceAuth"

  # 'C:\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Identity
  # RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Identity", _
  #                "ServiceAuth"
  RunCopyCommand "\\$CompName\c$\Documents and Settings\NetworkService\Local Settings\Application Data\Google\Identity" "ServiceAuth"

  # ' C:\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass
  # RunCopyCommand "\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass", _
  #                "DLL"
  RunCopyCommand "\\$CompName\c$\WINDOWS\system32\config\systemprofile\AppData\Local\Google\Google Apps Password Sync\Tracing\lsass" "DLL"

  # 'C:\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass
  # RunCopyCommand "\\" & CompName & "\c$\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass", _
  #                "DLL"
  RunCopyCommand "\\$CompName\c$\WINDOWS\system32\config\systemprofile\Local Settings\Application Data\Google\Google Apps Password Sync\Tracing\lsass" "DLL"

  # ' C:\ProgramData\Google\Google Apps Password Sync\config.xml
  # RunCopyCommand "\\" & CompName & "\c$\ProgramData\Google\Google Apps Password Sync\config.xml", _
  #                "."
  RunCopyCommand "\\$CompName\c$\ProgramData\Google\Google Apps Password Sync\config.xml" "."

  # ' C:\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml
  # RunCopyCommand "\\" & CompName & "\c$\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml", _
  #                "."
  RunCopyCommand "\\$CompName\c$\Documents and Settings\All Users\Application Data\Google\Google Apps Password Sync\config.xml" "."

  # ' Get install path for Password Sync (x86 indicates that the x86 version was
  # ' installed on x64 - won't work). Just search for the files in both possible
  # ' paths.
  # PrintLine "Getting list of installed files - install.txt and instx86.txt"
  PrintLine "Getting list of installed files - install.txt and instx86.txt"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files\Google\Google Apps Password Sync"" /B /S", _
  #            "install"
  RunCommand "dir ""\\$CompName\c$\Program Files\Google\Google Apps Password Sync"" /B /S" "install"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files\Google\G Suite Password Sync"" /B /S", _
  #            "install"
  RunCommand "dir ""\\$CompName\c$\Program Files\Google\G Suite Password Sync"" /B /S" "install"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files\Google\Password Sync"" /B /S", _
  #            "install"
  RunCommand "dir ""\\$CompName\c$\Program Files\Google\Password Sync"" /B /S" "install"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files (x86)\Google\Google Apps Password Sync"" /B /S", _
  #            "instx86"
  RunCommand "dir ""\\$CompName\c$\Program Files (x86)\Google\Google Apps Password Sync"" /B /S" "instx86"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files (x86)\Google\G Suite Password Sync"" /B /S", _
  #            "instx86"
  RunCommand "dir ""\\$CompName\c$\Program Files (x86)\Google\G Suite Password Sync"" /B /S" "instx86"
  # RunCommand "dir ""\\" & CompName & "\c$\Program Files (x86)\Google\Password Sync"" /B /S", _
  #            "instx86"
  RunCommand "dir ""\\$CompName\c$\Program Files (x86)\Google\Password Sync"" /B /S" "instx86"

  # PrintLine "Getting system-wide proxy settings dump from registry - proxy.txt"
  PrintLine "Getting system-wide proxy settings dump from registry - proxy.txt"
  # RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"" /v ProxySettingsPerUser", _
  #            "proxy"
  RunCommand "reg query ""\\$CompName\HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings"" /v ProxySettingsPerUser" "proxy"

  # PrintLine "Getting system-wide WinHTTP settings dump from registry - winhttp.txt"
  PrintLine "Getting system-wide WinHTTP settings dump from registry - winhttp.txt"
  # RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"" /v WinHttpSettings", _
  #            "winhttp"
  RunCommand "reg query ""\\$CompName\HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections"" /v WinHttpSettings" "winhttp"

  # PrintLine "Getting admin email address and service account address (if applicable)"
  PrintLine "Getting admin email address and service account address (if applicable)"
  # RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v Email", _
  #            "admin-and-serviceaccount-emails"
  RunCommand "reg query ""\\$CompName\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v Email" "admin-and-serviceaccount-emails"
  # RunCommand "reg query ""\\" & CompName & "\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v ServiceAccountEmail", _
  #            "admin-and-serviceaccount-emails"
  RunCommand "reg query ""\\$CompName\HKLM\SOFTWARE\Google\Google Apps Password Sync"" /v ServiceAccountEmail" "admin-and-serviceaccount-emails"

  # ' https://docs.microsoft.com/en-us/windows-server/security/tls/tls-registry-settings
  # ' This is useful for understanding errors such as WINHTTP_CALLBACK_STATUS_FLAG_SECURITY_CHANNEL_ERROR
  # PrintLine "Getting TLS registry settings"
  PrintLine "Getting TLS registry settings"
  # RunCommand "reg query ""\\" & CompName & "\HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"" /s", _
  #            "tls-registry-settings"
  RunCommand "reg query ""\\$CompName\HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL"" /s" "tls-registry-settings"

  # PrintLine "Getting system-wide WinHTTP settings dump from registry, and decoding - winhttp_decoded.txt"
  PrintLine "Getting system-wide WinHTTP settings dump from registry, and decoding - winhttp_decoded.txt"
  # DecodeWinHTTPSettings CompName, "winhttp_decoded.txt"
  DecodeWinHTTPSettings $CompName "winhttp_decoded.txt"

  # ' Get remote system time using http://blogs.technet.com/b/heyscriptingguy/archive/2007/03/08/how-can-i-verify-the-system-time-on-a-remote-computer.aspx,
  # ' and last boot time using https://gallery.technet.microsoft.com/ScriptCenter/82588289-4e07-455e-8322-c635cc719f00/
  # PrintLine "Getting local time and last boot time from remote machine"
  PrintLine "Getting local time and last boot time from remote machine"
  # Set objWMIService = GetObject("winmgmts:\\" & CompName & "\root\cimv2")
  # PrintErrorIfNeeded "Error opening WMI on " & CompName & ": "

  try {
      $colItems = Get-CimInstance -ComputerName $CompName -Namespace "root/cimv2" -ClassName "Win32_OperatingSystem"
  } catch {
      PrintErrorIfNeeded "Error opening WMI on ${CompName}: "
      return
  }

  # Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
  # PrintErrorIfNeeded "Error querying Win32_OperatingSystem. "

  # For Each objItem in colItems
  foreach ($objItem in $colItems) {
    # strTimeZone = (objItem.CurrentTimeZone / 60)
    $strTimeZone = ($objItem.CurrentTimeZone / 60)
    # If objItem.CurrentTimeZone >= 0 Then
    if ($objItem.CurrentTimeZone -ge 0) {
      # strTimeZone = "+" & strTimeZone
      $strTimeZone = "+" + $strTimeZone
    # End If
    }
    # WScript.Echo "Local Time: " & WMIDateStringToTime(objItem.LocalDateTime) & _
    #              ", Time Zone: " & strTimeZone
    Write-Host "Local Time: $(WMIDateStringToTime $objItem.LocalDateTime), Time Zone: $strTimeZone"
    # WScript.Echo "Last boot time: " & _
    #              WMIDateStringToTime(objItem.LastBootUpTime)
    Write-Host "Last boot time: $(WMIDateStringToTime $objItem.LastBootUpTime)"
  # Next
  }
  # PrintErrorIfNeeded "Error printing time: "
  # Err.Clear
  $Error.Clear()

  # ' Get file versions using https://blogs.technet.microsoft.com/heyscriptingguy/2005/04/18/how-can-i-determine-the-version-number-of-a-file/
  # PrintLine "Getting versions of important executables"
  PrintLine "Getting versions of important executables"

  $fileNames = @(
      'c:\\Windows\\System32\\lsass.exe',
      'c:\\Windows\\System32\\password_sync_dll.dll',
      'c:\\Program Files\\Google\\Google Apps Password Sync\\GoogleAppsPasswordSync.exe',
      'c:\\Program Files\\Google\\Google Apps Password Sync\\password_sync_service.exe',
      'c:\\Program Files\\Google\\Google Apps Password Sync\\unifiedlogin.dll',
      'c:\\Program Files\\Google\\Password Sync\\PasswordSync.exe',
      'c:\\Program Files\\Google\\Password Sync\\password_sync_service.exe',
      'c:\\Program Files\\Google\\Password Sync\\unifiedlogin.dll',
      'c:\\Program Files\\Google\\G Suite Password Sync\\PasswordSync.exe',
      'c:\\Program Files\\Google\\G Suite Password Sync\\password_sync_service.exe',
      'c:\\Program Files\\Google\\G Suite Password Sync\\unifiedlogin.dll'
  )

  # Set colFiles = objWMIService.ExecQuery( ... )
  # We construct the query or just loop
  # CIM query:
  $query = "SELECT Name, Version FROM CIM_Datafile WHERE Name = '" + ($fileNames -join "' OR Name = '") + "'"

  try {
      $colFiles = Get-CimInstance -ComputerName $CompName -Query $query
  } catch {
       PrintErrorIfNeeded "Error querying CIM_Datafile. "
  }

  # PrintErrorIfNeeded "Error querying CIM_Datafile. "
  # For Each objFile in colFiles
  if ($colFiles) {
      foreach ($objFile in $colFiles) {
        # Wscript.Echo "[" & objFile.Name & "] " & objFile.Version
        Write-Host "[$($objFile.Name)] $($objFile.Version)"
      # Next
      }
  }
  # PrintErrorIfNeeded "Error printing files and versions: "

  # PrintLine "Finished diagnostics on " & CompName
  PrintLine "Finished diagnostics on $CompName"
}

function Get-WritableDCs {
  # On Error Resume Next
  $ErrorActionPreference = "SilentlyContinue"

  # ' Initialize ADSI ADO provider. This is used because we need to make a
  # ' subtree-scope query.
  # Set conn = CreateObject("ADODB.Connection")
  # conn.Provider = "ADSDSOObject"
  # conn.Open "ADs Provider"
  # LogErrorIfNeeded "Error opening ADSI ADO provider"

  # QueryBase = _
  #     "<LDAP://" & GetObject("LDAP://RootDSE").Get("defaultNamingContext") & ">;"

  $defaultNamingContext = ""
  try {
      $RootDSE = [adsi]"LDAP://RootDSE"
      $defaultNamingContext = $RootDSE.defaultNamingContext
  } catch {
      LogErrorIfNeeded "Error getting RootDSE"
      return @()
  }

  # ' Query for computer accounts where userAccountControl has
  # ' SERVER_TRUST_ACCOUNT bit set, meaning it's a DC, and not msDS-IsRodc=true,
  # ' meaning it isn't an RODC. See http://support.microsoft.com/kb/305144 for
  # ' reference.
  # Query = QueryBase & _
  #         "(&(objectCategory=computer)" & _
  #         "(userAccountControl:1.2.840.113556.1.4.803:=8192)" & _
  #         "(!(msDS-IsRodc=true)));" & _
  #         "dNSHostName;subtree"

  $filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192)(!(msDS-IsRodc=true)))"

  # LogStr "A:Getting list of writable DCs: " & Query
  LogStr "A:Getting list of writable DCs: $filter"
  # Set rs = conn.Execute(Query)
  # LogErrorIfNeeded "Error executing query"

  $searcher = New-Object System.DirectoryServices.DirectorySearcher
  $searcher.SearchRoot = [adsi]"LDAP://$defaultNamingContext"
  $searcher.Filter = $filter
  $searcher.PageSize = 1000
  $searcher.SearchScope = "Subtree"
  $searcher.PropertiesToLoad.Add("dNSHostName") | Out-Null

  $results = $null
  try {
      $results = $searcher.FindAll()
  } catch {
      LogErrorIfNeeded "Error executing query"
      return @()
  }

  # If rs.EOF Then
  if ($results.Count -eq 0) {
    # LogStr "W:No DCs found - maybe msDS-IsRodc is missing from the schema " & _
    #        "(Windows 2003)? Trying without it."
    LogStr "W:No DCs found - maybe msDS-IsRodc is missing from the schema (Windows 2003)? Trying without it."
    # Query = QueryBase & _
    #         "(&(objectCategory=computer)" & _
    #         "(userAccountControl:1.2.840.113556.1.4.803:=8192));" & _
    #         "dNSHostName;subtree"
    $filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=8192))"

    # LogStr "A:Getting list of all DCs: " & Query
    LogStr "A:Getting list of all DCs: $filter"
    # Set rs = conn.Execute(Query)
    # LogErrorIfNeeded "Error executing query"
    $searcher.Filter = $filter
    try {
        $results = $searcher.FindAll()
    } catch {
        LogErrorIfNeeded "Error executing query"
        return @()
    }
  # End If
  }

  # Dim DCs(), DCCount
  # DCCount = 0
  $DCs = @()
  # While Not rs.EOF
  foreach ($res in $results) {
    # LogStr "A:Found " & rs.Fields(0).Value
    $dcName = $res.Properties["dNSHostName"][0]
    LogStr "A:Found $dcName"
    # ReDim Preserve DCs(DCCount)
    # DCs(DCCount) = rs.Fields(0).Value
    $DCs += $dcName
    # LogErrorIfNeeded "Error getting DC name"
    # rs.MoveNext
    # DCCount = DCCount + 1
  # Wend
  }
  # GetWritableDCs = DCs
  return $DCs
}

function CheckComputerAndUserDetails {
  # On Error Resume Next
  $ErrorActionPreference = "SilentlyContinue"

  # Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
  # LogErrorIfNeeded "Getting WMI service object for computer details"
  # Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
  # LogErrorIfNeeded "Executing WMI query for computer details"

  $colItems = $null
  try {
      $colItems = Get-CimInstance -ClassName Win32_ComputerSystem
  } catch {
      LogErrorIfNeeded "Executing WMI query for computer details"
  }

  # For Each objItem In colItems
  foreach ($objItem in $colItems) {
    # LogStr "A:Computer Name: " & objItem.Name
    LogStr "A:Computer Name: $($objItem.Name)"
    # LogStr "A:Computer's Domain: " & objItem.Domain
    LogStr "A:Computer's Domain: $($objItem.Domain)"
    # LogStr "A:Part Of Domain: " & objItem.PartOfDomain
    LogStr "A:Part Of Domain: $($objItem.PartOfDomain)"
    # Select Case objItem.DomainRole
    $strDomainRole = "Unknown ($($objItem.DomainRole))"
    switch ($objItem.DomainRole) {
      # Case 0 strDomainRole = "Standalone Workstation"
      0 { $strDomainRole = "Standalone Workstation" }
      # Case 1 strDomainRole = "Member Workstation"
      1 { $strDomainRole = "Member Workstation" }
      # Case 2 strDomainRole = "Standalone Server"
      2 { $strDomainRole = "Standalone Server" }
      # Case 3 strDomainRole = "Member Server"
      3 { $strDomainRole = "Member Server" }
      # Case 4 strDomainRole = "Backup Domain Controller"
      4 { $strDomainRole = "Backup Domain Controller" }
      # Case 5 strDomainRole = "Primary Domain Controller"
      5 { $strDomainRole = "Primary Domain Controller" }
      # Case Else strDomainRole = "Unknown (" & objItem.DomainRole & ")"
    }
    # End Select
    # LogStr "A:Computer's Domain Role: " & strDomainRole
    LogStr "A:Computer's Domain Role: $strDomainRole"
    # LogStr "A:Computer's Roles: " & Join(objItem.Roles, ", ")
    LogStr "A:Computer's Roles: $($objItem.Roles -join ", ")"

    # If Not objItem.PartOfDomain Then
    if (-not $objItem.PartOfDomain) {
      # LogStr "E:This machine isn't part of a domain. Exiting."
      LogStr "E:This machine isn't part of a domain. Exiting."
      # ErrorMsgBox "This machine isn't part of a domain. Make sure you " & _
      #             "are logged in as a domain admin, and run this tool again."
      ErrorMsgBox "This machine isn't part of a domain. Make sure you are logged in as a domain admin, and run this tool again."
      # CheckComputerAndUserDetails = False
      # Exit Function
      return $false
    # End If
    }
    # UserName = objShell.ExpandEnvironmentStrings("%USERNAME%")
    $UserName = $env:USERNAME
    # LogStr "A:Current user's name: " & UserName
    LogStr "A:Current user's name: $UserName"
    # UserDNSDomain = LCase(objShell.ExpandEnvironmentStrings("%USERDNSDOMAIN%"))
    $UserDNSDomain = $env:USERDNSDOMAIN
    if ($UserDNSDomain) { $UserDNSDomain = $UserDNSDomain.ToLower() }
    # If UserDNSDomain = "%userdnsdomain%" Then
    if (-not $UserDNSDomain) {
      # LogStr "E:The logged in user isn't a domain user. Exiting."
      LogStr "E:The logged in user isn't a domain user. Exiting."
      # ErrorMsgBox "The logged in user (" & UserName & ") isn't a domain " & _
      #             "user. Make sure you are logged in as a domain admin, " & _
      #             "and run this tool again."
      ErrorMsgBox "The logged in user ($UserName) isn't a domain user. Make sure you are logged in as a domain admin, and run this tool again."
      # CheckComputerAndUserDetails = False
      # Exit Function
      return $false
    # End If
    }
    # LogStr "A:Current user's AD DNS domain: " & UserDNSDomain
    LogStr "A:Current user's AD DNS domain: $UserDNSDomain"
    # If LCase(objItem.Domain) <> UserDNSDomain Then
    if ($objItem.Domain.ToLower() -ne $UserDNSDomain) {
      # LogStr "E:The user's domain doesn't match the machine's domain. Exiting."
      LogStr "E:The user's domain doesn't match the machine's domain. Exiting."
      # ErrorMsgBox "The current user's DNS domain (" & UserDNSDomain & _
      #             ") doesn't match the machine's DNS domain (" & _
      #             objItem.Domain & "). This will cause Password Sync " & _
      #             "to fail. Make sure you are logged in as a " & _
      #             "domain admin from the same domain as the Domain " & _
      #             "Controller, and try the installation again."
      ErrorMsgBox "The current user's DNS domain ($UserDNSDomain) doesn't match the machine's DNS domain ($([string]$objItem.Domain)). This will cause Password Sync to fail. Make sure you are logged in as a domain admin from the same domain as the Domain Controller, and try the installation again."
      # CheckComputerAndUserDetails = False
      # Exit Function
      return $false
    # End If
    }
  # Next
  }
  # CheckComputerAndUserDetails = True
  return $true
}

function CompressFolder($strPath, $strFolder) {
  # On Error Resume Next
  $ErrorActionPreference = "SilentlyContinue"

  # PowerShell native equivalent
  Compress-Archive -Path "$strFolder\\*" -DestinationPath $strPath -Force

  # Logic below is VBS specific using Shell.Application and ADODB.Stream
  # Const adTypeBinary = 1
  # ...
}

function ByteArrToHexString($bytes) {
   # Dim i
   # ByteArrToHexString = ""
   # For i = 1 to LenB(bytes)
   #    ByteArrToHexString = _
   #        ByteArrToHexString & Right("0" & Hex(AscB(MidB(bytes, i, 1))), 2)
   #    LogErrorIfNeeded _
   #        "Error converting SID bytes to string at " & ByteArrToHexString
   # Next

   if ($bytes) {
       return ($bytes | ForEach-Object { "{0:X2}" -f $_ }) -join ""
   }
   return ""
}

function GetTokenGroups($dnObject) {
  # Dim adsObject

  # ' Setup query of tokenGroup SIDs from dnObject
  # Set adsObject = GetObject("LDAP://" & Replace(dnObject, "/", "\/"))
  # LogErrorIfNeeded "Error opening admin's DN using ADSI"
  # adsObject.GetInfoEx Array("tokenGroups"), 0
  # GetTokenGroups = adsObject.GetEx("tokenGroups")
  # LogErrorIfNeeded "Error getting current user's tokenGroups"

  try {
      $adsObject = [adsi]"LDAP://$($dnObject -replace '/', '\/')"
      $adsObject.GetInfoEx(@("tokenGroups"), 0)
      return $adsObject.GetEx("tokenGroups")
  } catch {
      LogErrorIfNeeded "Error getting current user's tokenGroups"
      return @()
  }
}

function CheckIfRunningAsDomainAdmin {
  # ' Written by Liron Newman based on Shawn Poulson's example
  # ' NOTE: This function doesn't take into account the actual token's groups,
  # ' meaning that if running unelevated on a system that uses UAC, the script
  # ' will not be able to actually use all the user's permissions.
  # On Error Resume Next
  $ErrorActionPreference = "SilentlyContinue"

  # Set oADSysInfo = CreateObject("ADSystemInfo")
  # LogErrorIfNeeded "Error creating ADSystemInfo object"
  # userDN = oADSysInfo.UserName  ' Get DN of user

  $userDN = ""
  try {
      $oADSysInfo = New-Object -ComObject ADSystemInfo
      $userDN = $oADSysInfo.UserName
  } catch {
      # Fallback for non-COM (Linux)
      LogErrorIfNeeded "Error creating ADSystemInfo object"
      # return $false
  }

  # LogStr "A:Current user DN: " & userDN
  LogStr "A:Current user DN: $userDN"
  # ' We shouldn't use the name "Domain Admins" to check membership because it
  # ' may be localized, we should use the Well-Known SID.

  # ' Define the Domain Admins group SID prefix and suffix in hex:
  # Const DomainAdminsSIDStart = "010500000000000515000000"
  # Const DomainAdminsSIDEnd = "00020000"
  $DomainAdminsSIDStart = "010500000000000515000000"
  $DomainAdminsSIDEnd = "00020000"

  # ' Enumerate all member group names
  # tkUser = GetTokenGroups(userDN)  ' Get tokens of member groups
  $tkUser = GetTokenGroups $userDN

  # ' See if the Domain Admins group SID is in the token groups
  # CheckIfRunningAsDomainAdmin = False
  $isAdmin = $false
  # For Each sid In tkUser
  foreach ($sid in $tkUser) {
    # Dim tmpstr
    # tmpstr = ByteArrToHexString(sid)
    $tmpstr = ByteArrToHexString $sid
    # If (Left(tmpstr, Len(DomainAdminsSIDStart)) = DomainAdminsSIDStart) _
    #     And (Right(tmpstr, Len(DomainAdminsSIDEnd)) = DomainAdminsSIDEnd) Then
    if ($tmpstr.StartsWith($DomainAdminsSIDStart) -and $tmpstr.EndsWith($DomainAdminsSIDEnd)) {
      # CheckIfRunningAsDomainAdmin = True
      $isAdmin = $true
      # Exit For
      break
    # End If
    }
    # LogErrorIfNeeded "Error checking SID " & tmpstr
    LogErrorIfNeeded "Error checking SID $tmpstr"
  # Next
  }

  # If CheckIfRunningAsDomainAdmin Then
  if ($isAdmin) {
    # LogStr "A:The current user is a member of Domain Admins"
    LogStr "A:The current user is a member of Domain Admins"
    return $true
  # Else
  } else {
    # LogStr "E:The current user is *not* a member of Domain Admins"
    LogStr "E:The current user is *not* a member of Domain Admins"
    # ErrorMsgBox "The current user isn't a member of the Domain Admins " & _
    #             "group. To successfully install and setup Password Sync, " & _
    #             "you must be a Domain Admin." & _
    #             vbNewLine & vbNewLine & _
    #             "Please contact a Domain Admin to continue. You can try " & _
    #             "running this command, it may add you to the Domain Admins " & _
    #             "group:" & vbNewLine & vbNewLine & _
    #             "net group ""Domain Admins"" " & _
    #             objShell.ExpandEnvironmentStrings("%username%") & " /add" & _
    #             vbNewLine & vbNewLine & _
    #             "After joining the Domain Admins group, log out and back " & _
    #             "in, and try again."
    ErrorMsgBox "The current user isn't a member of the Domain Admins group. To successfully install and setup Password Sync, you must be a Domain Admin.`n`nPlease contact a Domain Admin to continue."
    # ' TODO: Get the correct sAMAccountName for Domain Admins, as it may have
    # ' been localized... It can be obtained using:
    # ' GetObject("LDAP://<SID=" & ByteArrToHexString(objectSid) & ">").Get("sAMAccountName")
    return $false
  # End If
  }
}

function Main {
    param(
        [string]$Arg1,
        [string]$Arg2
    )

    # Check if the /ELEVATED parameter was provided, meaning that the script
    # re-invoked itself with elevation ("run as admin").
    # Dim ParameterProvided
    # ParameterProvided = False
    $ParameterProvided = $false

    # In VBScript, this check needs to be in a separate condition to prevent an
    # "index out of bounds" error.
    # If WScript.Arguments.Count > 0 Then
    #   If UCase(WScript.Arguments(0)) = "/ELEVATED" Or _
    #      UCase(WScript.Arguments(0)) = "/DC" Then
    #     ParameterProvided = True
    #   End If
    # End If
    if ($Arg1 -eq "/ELEVATED" -or $Arg1 -eq "/DC") {
        $ParameterProvided = $true
    }

    # If Not (UCase(Right(WScript.FullName, 12)) = "\CSCRIPT.EXE" And _
    #     ParameterProvided) Then
    #   objShellApplication.ShellExecute _
    #       "cmd.exe", _
    #       "/c title " & ToolName & " & cscript.exe //nologo """ & _
    #           WScript.ScriptFullName & """ /ELEVATED", _
    #       "", _
    #       "runas", _
    #       1
    #   WScript.Quit
    # End If

    if (-not $ParameterProvided) {
        if ($IsWindows -and -not (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
             $psExe = if ($PSVersionTable.PSEdition -eq 'Core') { "pwsh" } else { "powershell" }
             Start-Process $psExe -ArgumentList "-File `"$script:CurrentScriptPath`" /ELEVATED" -Verb RunAs
             exit
        }
    }

    # On Error Resume Next  ' Errors will be handled by the code
    $ErrorActionPreference = "SilentlyContinue"

    # Dim LogFileName, TempDir, CurrentTimeString
    # CurrentTimeString = GetCurrentTimeString
    $script:CurrentTimeString = Get-CurrentTimeString
    # TempDir = objShell.ExpandEnvironmentStrings("%temp%\" & ToolName)
    $script:TempDir = Join-Path $env:TEMP $ToolName
    # We can assume %userdnsdomain% is the computer's DNS domain too, because we
    # will enforce it in CheckComputerAndUserDetails().
    # CurrentComputerName = _
    #     UCase(objShell.ExpandEnvironmentStrings("%computername%.%userdnsdomain%"))
    $script:CurrentComputerName = ("$env:COMPUTERNAME.$env:USERDNSDOMAIN").ToUpper()
    # LogFileName = TempDir & "\" & ToolName & ".log"
    $script:LogFileName = Join-Path $script:TempDir "$ToolName.log"

    # Check if this instance was executed to diagnose a DC
    # If WScript.Arguments.Count > 0 Then
    #   ' Note that this will break commandline arguments if we plan to use them in
    #   ' the future
    #   If UCase(WScript.Arguments(0)) = "/DC" Then
    #     RunDiagnostics WScript.Arguments(1)
    #     WScript.Quit
    #   End If
    # End If

    if ($Arg1 -eq "/DC") {
        RunDiagnostics $Arg2
        return
    }

    # ' Delete old temporary folder
    # fso.DeleteFolder TempDir, True
    if (Test-Path $script:TempDir) {
        Remove-Item $script:TempDir -Recurse -Force
    }

    # ' Create new temporary folder
    # fso.CreateFolder TempDir
    New-Item -ItemType Directory -Path $script:TempDir | Out-Null

    # LogStr "A:Starting " & ToolName & " version " & Ver & " from " & _
    #        WScript.ScriptFullName
    LogStr "A:Starting $ToolName version $Ver from $PSCommandPath"

    # ' Check whether the current user is a Domain Admin and other machine/user
    # ' settings.
    # If Not CheckComputerAndUserDetails Then WScript.Quit
    if (-not (CheckComputerAndUserDetails)) { return }
    # If Not CheckIfRunningAsDomainAdmin Then WScript.Quit
    if (-not (CheckIfRunningAsDomainAdmin)) { return }


    # ' Get list of writable DCs
    # Dim arrWritableDCs
    # arrWritableDCs = GetWritableDCs
    $arrWritableDCs = Get-WritableDCs
    # LogStr "A:Got " & UBound(arrWritableDCs) + 1 & " writable DCs"
    LogStr ("A:Got " + $arrWritableDCs.Count + " writable DCs")

    # ' Instantiate additional arrays
    # Dim arrExec()  ' For Exec objects
    # ReDim arrExec(UBound(arrWritableDCs))
    # Dim arrBuffers()  ' For StdOut buffers
    # ReDim arrBuffers(UBound(arrWritableDCs))
    # Dim arrOutFiles()  ' For StdOut buffers
    # ReDim arrOutFiles(UBound(arrWritableDCs))
    $arrExec = New-Object "System.Object[]" $arrWritableDCs.Count
    $arrBuffers = New-Object "System.String[]" $arrWritableDCs.Count
    for ($k=0; $k -lt $arrBuffers.Length; $k++) { $arrBuffers[$k] = "" }
    $arrOutFiles = New-Object "System.Object[]" $arrWritableDCs.Count

    # For i = 0 To UBound(arrWritableDCs)
    for ($i = 0; $i -lt $arrWritableDCs.Count; $i++) {
    #   If arrWritableDCs(i) <> "" Then
        if ($arrWritableDCs[$i] -ne "") {
    #     ' Create folder for results
    #     LogStr "A:Creating " & TempDir & "\" & arrWritableDCs(i)
            LogStr "A:Creating $script:TempDir\$($arrWritableDCs[$i])"
    #     fso.CreateFolder TempDir & "\" & arrWritableDCs(i)
            New-Item -ItemType Directory -Path "$script:TempDir\$($arrWritableDCs[$i])" | Out-Null
    #     LogErrorIfNeeded "Error creating folder"
            LogErrorIfNeeded "Error creating folder"
    #     ' Call this script with DC name
    #     LogStr "A:Starting job for " & arrWritableDCs(i)
            LogStr "A:Starting job for $($arrWritableDCs[$i])"
    #     ' We need to redirect both stdout and stderr to a file instead of catching
    #     ' them directly with the StdOut/StdEr objects, because reading from these
    #     ' streams is blocking, and we want to do it concurrently.
    #     Set arrExec(i) = objShell.Exec( _
    #         "cmd /c cscript //NoLogo """ & WScript.ScriptFullName & """ /DC " & _
    #         arrWritableDCs(i) & " 1>" & TempDir & "\" & arrWritableDCs(i) & _
    #         ".txt 2>&1 ")
            $outFile = "$script:TempDir\$($arrWritableDCs[$i]).txt"
            $errFile = "$script:TempDir\$($arrWritableDCs[$i]).err"
            $psExe = if ($PSVersionTable.PSEdition -eq 'Core') { "pwsh" } else { "powershell" }
            $arrExec[$i] = Start-Process -FilePath $psExe -ArgumentList "-File", "`"$script:CurrentScriptPath`"", "/DC", $arrWritableDCs[$i] -RedirectStandardOutput $outFile -RedirectStandardError $errFile -PassThru
    #     LogErrorIfNeeded "Error starting job"
            LogErrorIfNeeded "Error starting job"
    #     WScript.Sleep 100
            Start-Sleep -Milliseconds 100
    #     ' Open the output file.
    #     Set arrOutFiles(i) = _
    #         fso.OpenTextFile(TempDir & "\" & arrWritableDCs(i) & ".txt", _
    #                          ForReading, _
    #                          0)
            try {
                $arrOutFiles[$i] = [System.IO.File]::Open("$outFile", [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            } catch {
                LogErrorIfNeeded "Error opening job output file"
            }
    #     LogErrorIfNeeded "Error opening job output file"

    #   Else
        } else {
    #     LogStr "W:Skipping empty DC name"
            LogStr "W:Skipping empty DC name"
    #   End If
        }
    # Next
    }

    # ' Process output from all instances until they're all gone
    # Dim NumCompleted
    # NumCompleted = 0
    $NumCompleted = 0
    # While NumCompleted <= UBound(arrWritableDCs)
    while ($NumCompleted -le ($arrWritableDCs.Count - 1)) {
    #   WScript.Sleep 10
        Start-Sleep -Milliseconds 100
    #   For i = 0 To UBound(arrWritableDCs)
        for ($i = 0; $i -lt $arrWritableDCs.Count; $i++) {
    #     ' We set completed Execs to Null, so we can skip them.
    #     If Not IsNull(arrExec(i)) Then
            if ($arrExec[$i] -ne $null) {

                $stream = $arrOutFiles[$i]
                if ($stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $content = $reader.ReadToEnd()
                    $arrBuffers[$i] = $arrBuffers[$i] + $content
                }

    #       Err.Clear  ' Ignore "Input past end of file" errors
                $Error.Clear()
    #       ' TODO: Improve logging here - some text files aren't being read on
    #       ' domains with many DCs.
    #       ' As long as we have full lines...
    #       While InStr(arrBuffers(i), vbNewLine) > 0
                while ($arrBuffers[$i].IndexOf([Environment]::NewLine) -ge 0) {
                    $newlineIdx = $arrBuffers[$i].IndexOf([Environment]::NewLine)
    #         If InStr(arrBuffers(i), vbNewLine) > 1 Then
                    if ($newlineIdx -ge 0) {
    #           LogStr "A:Job " & arrWritableDCs(i) & ": " & _
    #                  Left(arrBuffers(i), InStr(arrBuffers(i), vbNewLine) - 1)
                        $line = $arrBuffers[$i].Substring(0, $newlineIdx)
                        if ($line.Length -gt 0) {
                             LogStr "A:Job $($arrWritableDCs[$i]): $line"
                        }
    #         End If
                    }
    #         arrBuffers(i) = Mid(arrBuffers(i), _
    #                             InStr(arrBuffers(i), vbNewLine) + 2, _
    #                             Len(arrBuffers(i)))
                    if ($newlineIdx + [Environment]::NewLine.Length -lt $arrBuffers[$i].Length) {
                        $arrBuffers[$i] = $arrBuffers[$i].Substring($newlineIdx + [Environment]::NewLine.Length)
                    } else {
                        $arrBuffers[$i] = ""
                    }
    #       Wend
                }

    #       If arrExec(i).Status <> WshRunning Then
                if ($arrExec[$i].HasExited) {
    #         ' Write any leftover data
    #         arrBuffers(i) = arrBuffers(i) & arrOutFiles(i).ReadAll
                    if ($stream) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        $content = $reader.ReadToEnd()
                        $arrBuffers[$i] = $arrBuffers[$i] + $content
                    }

    #         Err.Clear  ' Ignore reading errors
                    $Error.Clear()
    #         If Len(arrBuffers(i)) > 0 Then
                    if ($arrBuffers[$i].Length -gt 0) {
    #           Dim arrTemp
    #           arrTemp = Split(arrBuffers(i), vbNewLine)
                        $arrTemp = $arrBuffers[$i] -split [Environment]::NewLine
    #           For j = 0 To UBound(arrTemp)
                        foreach ($line in $arrTemp) {
    #             If Len(arrTemp(j)) > 0 Then
                            if ($line.Length -gt 0) {
    #               LogStr "A:Job " & arrWritableDCs(i) & ": " & arrTemp(j)
                                LogStr "A:Job $($arrWritableDCs[$i]): $line"
    #             End If
                            }
    #           Next  ' j
                        }
    #         End If
                    }
    #         ' Close file we no longer need
    #         arrOutFiles(i).Close
                    if ($stream) { $stream.Close(); $stream.Dispose() }
    #         ' Log status
    #         If arrExec(i).Status = WshFailed Then
    #           LogStr "E:Job " & arrWritableDCs(i) & " failed with exit code " & _
    #                  arrExec(i).ExitCode
                    if ($arrExec[$i].ExitCode -ne 0) {
                         LogStr "E:Job $($arrWritableDCs[$i]) failed with exit code $($arrExec[$i].ExitCode)"
    #         ElseIf arrExec(i).Status = WshFinished Then
                    } else {
    #           LogStr "A:Job " & arrWritableDCs(i) & _
    #                  " finished successfully with exit code " & arrExec(i).ExitCode
                         LogStr "A:Job $($arrWritableDCs[$i]) finished successfully with exit code $($arrExec[$i].ExitCode)"
    #         End If
                    }
    #         NumCompleted = NumCompleted + 1
                    $NumCompleted++
    #         arrExec(i) = Null
                    $arrExec[$i] = $null
    #       End If
                }
    #     End If
            }
    #   Next  ' i
        }
    # Wend
    }

    # LogStr "A:Finished collecting information, creating ZIP"
    LogStr "A:Finished collecting information, creating ZIP"

    # ' Rename folder to include timestamp, for uniqueness
    # Dim NewFolderName
    # NewFolderName = TempDir & "_" & CurrentTimeString
    $NewFolderName = "${script:TempDir}_${script:CurrentTimeString}"
    # fso.MoveFolder TempDir, NewFolderName
    Move-Item -Path $script:TempDir -Destination $NewFolderName

    # ' Create ZIP with reports
    # Dim ZipName
    # ZipName = ToolName & "-report_" & CurrentTimeString & ".zip"
    $ZipName = "$ToolName-report_$script:CurrentTimeString.zip"
    # CompressFolder objShell.SpecialFolders("Desktop") & "\" & ZipName, NewFolderName
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    if (-not $DesktopPath -or -not (Test-Path $DesktopPath)) { $DesktopPath = $HOME + "/Desktop" }
    if (-not (Test-Path $DesktopPath)) { New-Item -ItemType Directory -Path $DesktopPath | Out-Null }
    CompressFolder (Join-Path $DesktopPath $ZipName) $NewFolderName

    # Message = "Please send the file """ & ZipName & _
    #           """ from your Desktop to Google Support for investigation."
    $Message = "Please send the file `"$ZipName`" from your Desktop to Google Support for investigation."
    # WScript.Echo VbNewLine & Message
    Write-Host ("`n" + $Message)
    # MsgBox Message, vbOKOnly, ToolName
    if ($IsWindows) {
        # [System.Windows.Forms.MessageBox]::Show($Message, $ToolName)
    }

    # WScript.Echo "Press Enter to close this window"
    Write-Host "Press Enter to close this window"
    # WScript.StdIn.Read(1)
    if ($Host.Name -eq "ConsoleHost") {
        Read-Host | Out-Null
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    Main $args[0] $args[1]
}
