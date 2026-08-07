param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'ALCLoader' {
    It 'Should load ExcelFast after System.IO.Packaging 9 is loaded in a PowerShell session' {
        if ($PSVersionTable.PSEdition -eq 'Desktop') {
            Set-ItResult -Skipped -Because 'ALC is unavailable in Windows PowerShell 5.1 and this test is invalid.'
            return
        }

        $packagePath = Join-Path $TestDrive 'System.IO.Packaging.9.0.2.nupkg.zip'
        $packageDirectory = Join-Path $TestDrive 'System.IO.Packaging.9.0.2'
        Invoke-WebRequest -Uri 'https://api.nuget.org/v3-flatcontainer/system.io.packaging/9.0.2/system.io.packaging.9.0.2.nupkg' -OutFile $packagePath
        Expand-Archive -Path $packagePath -DestinationPath $packageDirectory
        $packagingPath = Join-Path $packageDirectory 'lib/net8.0/System.IO.Packaging.dll'

        $script = {
            $ErrorActionPreference = 'Stop'
            $packagingContext = [System.Runtime.Loader.AssemblyLoadContext]::new('LegacyPackaging', $true)
            $packagingAssembly = $packagingContext.LoadFromAssemblyPath($env:__PESTER_EXCELFAST_PACKAGING_PATH)
            if ($packagingAssembly.GetName().Version.Major -ne 9) {
                throw "Expected System.IO.Packaging 9, got $($packagingAssembly.GetName().Version)."
            }
            Import-Module -Name $env:__PESTER_EXCELFAST_MODULE_PATH -Force
            $rows = Import-Workbook -Path $env:__PESTER_EXCELFAST_WORKBOOK_PATH
            if ($rows.Count -ne 10) {
                throw "Expected 10 rows, got $($rows.Count)."
            }
        }
        $env:__PESTER_EXCELFAST_MODULE_PATH = $ModulePath
        $env:__PESTER_EXCELFAST_WORKBOOK_PATH = $ValidExcelFile
        $env:__PESTER_EXCELFAST_PACKAGING_PATH = $packagingPath
        $output = & pwsh -NoProfile -NonInteractive -Command $script 2>&1

        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
    }
}
