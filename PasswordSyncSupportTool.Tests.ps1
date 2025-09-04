# Pester tests for PasswordSyncSupportTool.ps1
#
# Note: This test is designed to run in an environment where the script's
# dependencies (Active Directory, Windows Registry, etc.) are not available.
# It mocks key functions to verify the script's overall logic and control flow.

Describe "PasswordSyncSupportTool Script" {
    $script:scriptPath = "/app/PasswordSyncSupportTool.ps1"

    BeforeAll {
        # Dot-source the script to make its functions available for mocking and execution.
        . $script:scriptPath
    }

    It "should run the main orchestration logic, log errors for fake DCs, and produce a ZIP file" {
        # Mock functions to bypass environmental checks
        Mock Test-IsRunningAsDomainAdmin { return $true }
        Mock Test-ComputerAndUserDetails { return $true }

        # Mock the DC query to return a predictable list of fake DCs
        $fakeDcs = @('fake-dc-01.test.local', 'fake-dc-02.test.local')
        Mock Get-WritableDCs { return $fakeDcs }

        # Mock the message box to prevent it from blocking the test
        Mock [System.Windows.MessageBox]::Show { }

        # Mock Read-Host to prevent it from blocking the test at the end
        Mock Read-Host { }

        # --- Start of Main Orchestrator Logic ---
        $Global:TempDir = Join-Path $env:TEMP $ToolName
        $Global:LogFileName = Join-Path $Global:TempDir "$($ToolName).log"
        $CurrentTimeString = Get-CurrentTimeString

        if (Test-Path $Global:TempDir) {
            Remove-Item -Path $Global:TempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -Path $Global:TempDir -ItemType Directory | Out-Null

        Log-Str "A:Starting $ToolName version $Ver"

        if (-not (Test-ComputerAndUserDetails)) { throw "Test-ComputerAndUserDetails failed" }
        if (-not (Test-IsRunningAsDomainAdmin)) { throw "Test-IsRunningAsDomainAdmin failed" }

        $arrWritableDCs = Get-WritableDCs
        Log-Str "A:Got $($arrWritableDCs.Count) writable DCs"

        $jobs = @()
        foreach ($dc in $arrWritableDCs) {
            if (-not [string]::IsNullOrWhiteSpace($dc)) {
                $dcWorkDir = Join-Path $Global:TempDir $dc
                Log-Str "A:Creating $dcWorkDir"
                New-Item -Path $dcWorkDir -ItemType Directory | Out-Null

                Log-Str "A:Starting job for $dc"
                $job = Start-Job -FilePath $script:scriptPath -ArgumentList @{ DCName = $dc } -Name $dc
                $jobs += $job
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
        # --- End of Main Orchestrator Logic ---

        # Assertions
        $tempDir = Join-Path $env:TEMP $ToolName
        $logFile = Join-Path $tempDir "$($ToolName).log"

        $logFile | Should -Exist

        $logContent = Get-Content $logFile -Raw

        foreach ($dc in $fakeDcs) {
            $dcDir = Join-Path $tempDir $dc
            $dcDir | Should -Exist

            $logContent | Should -Match "Starting job for $dc"
            $logContent | Should -Match "Job $($dc) failed"
        }

        $logContent | Should -Match "Starting diagnostics on"
        $logContent | Should -Match "Running command: reg query"
        $logContent | Should -Match "Getting status for service 'Password Sync'"
        $logContent | Should -Match "Copying logs and XML"
        $logContent | Should -Match "Getting list of installed files"
        $logContent | Should -Match "Getting local time and last boot time"

        $zipFile = Get-ChildItem -Path ([System.Environment]::GetFolderPath('Desktop')) -Filter "*$($ToolName)-report_*.zip"
        $zipFile | Should -Not -BeNull
        $zipFile.Count | Should -Be 1
    }
}
