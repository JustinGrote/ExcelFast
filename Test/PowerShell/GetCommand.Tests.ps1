using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'Get-Workbook Command Tests' {
    Context 'When opening a valid Excel file' {
        It 'Should successfully open Test10.xlsx' {
            $actual = Get-Workbook -Path $ValidExcelFile

            $actual | Should -Not -BeNullOrEmpty
            $actual.GetType().Name | Should -Be 'XLWorkbook'
            $actual.Worksheets.Count | Should -BeGreaterThan 0
        }
    }

    Context 'Path Parameter' {
        It 'Should throw FileNotFoundException for a non-existent file path' {
            { Get-Workbook -Path $InvalidPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([FileNotFoundException]) -ErrorId 'FileNotFound,ExcelFast.PowerShell.Cmdlets.GetCommand'
        }

        It 'Should throw when opening a plaintext file with an xlsx extension' {
            { Get-Workbook -Path $NonExcelContent -ErrorAction Stop } |
                Should -Throw -ErrorId 'ImportExcelWorkbookError,ExcelFast.PowerShell.Cmdlets.GetCommand'
        }
    }

    Context 'Pipeline Input' {
        It 'Should accept pipeline input by path' {
            $actual = $ValidExcelFile | Get-Workbook
            $actual | Should -Not -BeNullOrEmpty
            $actual.GetType().Name | Should -Be 'XLWorkbook'
        }

        It 'Should accept multiple file paths' {
            $actual = Get-Workbook -Path $ValidExcelFile, $SkippedRow
            $actual | Should -HaveCount 2
            $actual[0].GetType().Name | Should -Be 'XLWorkbook'
            $actual[1].GetType().Name | Should -Be 'XLWorkbook'
        }
    }

    Context 'Aliases' {
        It 'Should have alias gwb' {
            $actual = Get-Command -Name gwb
            $actual | Should -Not -BeNullOrEmpty
            $actual.ReferencedCommand | Should -Be 'Get-Workbook'
        }

        It 'Should have alias owb' {
            $actual = Get-Command -Name owb
            $actual | Should -Not -BeNullOrEmpty
            $actual.ReferencedCommand | Should -Be 'Get-Workbook'
        }
        It 'Should have alias Open-Workbook' {
            $actual = Get-Command -Name Open-Workbook
            $actual | Should -Not -BeNullOrEmpty
            $actual.ReferencedCommand | Should -Be 'Get-Workbook'
        }
    }
}