
$ScriptPath = Join-Path $PSScriptRoot "PasswordSyncSupportTool.ps1"

Describe "PasswordSyncSupportTool" {

    Context "Main Execution Flow" {
        BeforeAll {
            $ScriptPath = Join-Path $PSScriptRoot "PasswordSyncSupportTool.ps1"
            $TestTemp = Join-Path $PSScriptRoot "TestTemp"
            if (Test-Path $TestTemp) { Remove-Item $TestTemp -Recurse -Force }
            New-Item -ItemType Directory -Path $TestTemp | Out-Null
            $env:TEMP = $TestTemp

            # Source the script (without executing Main)
            . $ScriptPath

            # Ensure script:CurrentScriptPath is set correctly for Main to use in child processes
            $script:CurrentScriptPath = $ScriptPath
        }

        AfterAll {
            if (Test-Path $TestTemp) { Remove-Item $TestTemp -Recurse -Force }
        }

        It "Runs the main loop and triggers jobs for DCs" {
            # Mocks
            Mock CheckComputerAndUserDetails { return $true }
            Mock CheckIfRunningAsDomainAdmin { return $true }
            Mock Get-WritableDCs { return @("DC1", "DC2") }
            Mock Get-CurrentTimeString { return "MOCKED_TIME" }
            Mock CompressFolder { }
            Mock ErrorMsgBox { }
            Mock Write-Host { }
            # We verify Write-Host calls later or just ignore output

            # Run Main
            Main

            # Verify folders created
            $dc1Path = Join-Path $env:TEMP "PasswordSyncSupportTool_MOCKED_TIME"
            $dc1Path = Join-Path $dc1Path "DC1"
            Test-Path $dc1Path | Should -BeTrue

            $dc2Path = Join-Path $env:TEMP "PasswordSyncSupportTool_MOCKED_TIME"
            $dc2Path = Join-Path $dc2Path "DC2"
            Test-Path $dc2Path | Should -BeTrue
        }
    }

    Context "Worker Execution (/DC)" {
         BeforeAll {
            $ScriptPath = Join-Path $PSScriptRoot "PasswordSyncSupportTool.ps1"
            $TestTemp = Join-Path $PSScriptRoot "TestTempWorker"
            if (Test-Path $TestTemp) { Remove-Item $TestTemp -Recurse -Force }
            New-Item -ItemType Directory -Path $TestTemp | Out-Null
            $env:TEMP = $TestTemp

            . $ScriptPath
        }

        It "Runs diagnostics for a specific DC" {
            Mock DecodeWinHTTPSettings { }
            Mock Write-Host { }

            if ($IsWindows) {
                 # On Windows, we run without mocking system commands to verify real execution
                 # We assume the runner doesn't have the "C$" share accessible via network loopback by default,
                 # but we can try targeting localhost.

                 # We still mock RunCopyCommand because network shares require configuration
                 Mock RunCopyCommand { }

                 # Run Main with /DC pointing to localhost
                 Main "/DC" "localhost"

                 # Assertions
                 # We can check if log files were created and contain expected output
                 $logFile = "$env:TEMP\PasswordSyncSupportTool\localhost.txt"
                 Test-Path $logFile | Should -BeTrue

                 # Check for output from real commands
                 $content = Get-Content $logFile -Raw
                 $content | Should -Match "Image Name" # tasklist output
                 $content | Should -Match "SERVICE_NAME" # sc output (even if error, query might output something or error log)
                 # Wait, sc query might fail if service doesn't exist. "Enum: ... The specified service does not exist"
                 # But tasklist should run.

            } else {
                # Mock RunCommand to avoid actual execution failures and verify calls
                Mock RunCommand { }
                Mock RunCopyCommand { }
                function Get-CimInstance {}
                Mock Get-CimInstance { return @() }

                # Run Main with /DC
                Main "/DC" "DC1"

                # Assertions
                Should -Invoke RunCommand
                Should -Invoke RunCopyCommand
            }
        }
    }
}
