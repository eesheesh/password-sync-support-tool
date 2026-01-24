
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

            # Mock Start-Process
            # We need it to return an object with HasExited=$true
            # And we need to simulate the output file creation
            Mock Start-Process {
                $outFile = $null
                # Attempt to find RedirectStandardOutput in arguments

                if ($RedirectStandardOutput) {
                    $outFile = $RedirectStandardOutput
                    "Output from Mock" | Out-File -FilePath $outFile -Encoding utf8
                }

                return [PSCustomObject]@{
                    HasExited = $true
                    ExitCode = 0
                }
            }

            # Run Main
            Main

            # Assertions
            Should -Invoke Start-Process -Times 2

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

            Write-Host "PSScriptRoot: $PSScriptRoot"; Write-Host "ScriptPath: $ScriptPath"; . $ScriptPath
        }

        It "Runs diagnostics for a specific DC" {
            # Mock RunCommand to avoid actual execution failures and verify calls
            Mock RunCommand { }
            Mock RunCopyCommand { }
            Mock DecodeWinHTTPSettings { }
            function Get-CimInstance {}
            Mock Get-CimInstance { return @() }
            Mock Write-Host { }

            # Run Main with /DC
            Main "/DC" "DC1"

            # Assertions
            Should -Invoke RunCommand
            Should -Invoke RunCopyCommand
        }
    }
}
