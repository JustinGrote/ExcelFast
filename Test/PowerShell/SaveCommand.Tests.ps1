using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'Save-Workbook Command Tests' {
    BeforeEach {
        # Set up test paths in Pester's TestDrive
        $DestPath = Join-Path 'TestDrive:' 'test.xlsx'
        $TempDirPath = Join-Path 'TestDrive:' 'TestDir' 'test.xlsx'
    }

    Context 'When saving a workbook' {
        It 'Should successfully save to a new location' {
            $workbook = Open-Workbook -Path $ValidExcelFile
            Save-Workbook -Workbook $workbook -Destination $DestPath
            Test-Path $DestPath | Should -BeTrue
        }

        It 'Should save to its original location when no destination is specified' {
            # Copy the test file to a temp location first
            Copy-Item -Path $ValidExcelFile -Destination $DestPath
            $workbook = Open-Workbook -Path $DestPath
            Save-Workbook -Workbook $workbook
            Test-Path $DestPath | Should -BeTrue
        }

        It 'Should create directory path when Force is specified' {
            $workbook = Open-Workbook -Path $ValidExcelFile
            Save-Workbook -Workbook $workbook -Destination $TempDirPath -Force
            Test-Path $TempDirPath | Should -BeTrue
        }
    }

    Context 'Error handling' {
        It 'Should throw when directory does not exist and Force is not specified' {
            $workbook = Open-Workbook -Path $ValidExcelFile
            { Save-Workbook -Workbook $workbook -Destination $TempDirPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([DirectoryNotFoundException]) -ErrorId 'DirectoryNotFound,ExcelFast.PowerShell.Cmdlets.SaveCommand'
        }

        It 'Should throw when file exists and Force is not specified' {
            # Create the file first
            '' | Set-Content -Path $DestPath
            $workbook = Open-Workbook -Path $ValidExcelFile
            { Save-Workbook -Workbook $workbook -Destination $DestPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([IOException]) -ErrorId 'FileAlreadyExists,ExcelFast.PowerShell.Cmdlets.SaveCommand'
        }
    }

    Context 'Pipeline Input' {
        It 'Should accept pipeline input' {
            $workbook = Open-Workbook -Path $ValidExcelFile
            $workbook | Save-Workbook -Destination $DestPath
            Test-Path $DestPath | Should -BeTrue
        }

        It 'Should throw when using Destination with multiple workbooks' {
            $workbooks = @($ValidExcelFile, $SkippedRow) | ForEach-Object { Open-Workbook -Path $_ }
            { $workbooks | Save-Workbook -Destination $DestPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([System.Management.Automation.PSNotSupportedException]) -ErrorId 'MultipleWorkbooksWithDestinationParameter,ExcelFast.PowerShell.Cmdlets.SaveCommand'
        }
    }
}
