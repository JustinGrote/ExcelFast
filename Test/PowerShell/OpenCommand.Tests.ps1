using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'Open-Workbook Command Tests' {
    Context 'When opening a valid Excel file' {
        It 'Should successfully open Test10.xlsx' {
            $actual = Open-Workbook -Path $ValidExcelFile

            $actual | Should -Not -BeNullOrEmpty
            $actual.GetType().Name | Should -Be 'XLWorkbook'
            $actual.Worksheets.Count | Should -BeGreaterThan 0
        }
    }

    Context 'Path Parameter' {
        It 'Should throw FileNotFoundException for a non-existent file path' {
            { Open-Workbook -Path $InvalidPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([FileNotFoundException]) -ErrorId 'FileNotFound,ExcelFast.PowerShell.Cmdlets.OpenCommand'
        }

        It 'Should throw when opening a plaintext file with an xlsx extension' {
            { Open-Workbook -Path $NonExcelContent -ErrorAction Stop } |
                Should -Throw -ErrorId 'ImportExcelWorkbookError,ExcelFast.PowerShell.Cmdlets.OpenCommand'
        }
    }

    Context 'Pipeline Input' {
        It 'Should accept pipeline input' {
            $actual = $ValidExcelFile | Open-Workbook
            $actual | Should -Not -BeNullOrEmpty
            $actual.GetType().Name | Should -Be 'XLWorkbook'
        }

        It 'Should accept multiple file paths' {
            $actual = Open-Workbook -Path $ValidExcelFile, $SkippedRow
            $actual | Should -HaveCount 2
            $actual[0].GetType().Name | Should -Be 'XLWorkbook'
            $actual[1].GetType().Name | Should -Be 'XLWorkbook'
        }
    }
}
