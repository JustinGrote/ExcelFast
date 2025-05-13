using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'Import-Excel Command Tests' {
    Context 'When importing a valid Excel file' {
        It 'Should successfully import data from Test10.xlsx' {
            $actual = Import-Workbook -Path $ValidExcelFile

            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 10
            $actual[0].Name | Should -Be 'Test1'
            $actual[0].Value | Should -Be 'Value1'
            $actual[-1].Name | Should -Be 'Test10'
            $actual[-1].Value | Should -Be 'Value10'
        }
    }

    Context 'Path Parameter' {
        It 'Should throw FileNotFoundException for a non-existent file path' {
            { Import-Workbook -Path $InvalidPath -ErrorAction Stop } |
                Should -Throw -ExceptionType ([FileNotFoundException]) -ErrorId 'FileNotFound,ExcelFast.PowerShell.Cmdlets.ImportCommand'
        }

        It 'Should throw ArgumentException for a non-Excel extension' {
            { Import-Workbook -Path $NonExcelExtension -ErrorAction Stop } |
                Should -Throw -ExceptionType ([ArgumentException]) -ErrorId 'UnsupportedFileType,ExcelFast.PowerShell.Cmdlets.ImportCommand'
        }

        It 'Should throw InvalidDataException for a plaintext file with an xlsx extension' {
            { Import-Workbook -Path $NonExcelContent -ErrorAction Stop } |
                Should -Throw -ExceptionType ([InvalidDataException]) -ErrorId 'UnknownFileContent,ExcelFast.PowerShell.Cmdlets.ImportCommand'
        }

        It 'Should throw InvalidDataException for a xlsx file that is a zip but is not a valid Excel file' {
            { Import-Workbook -Path $NonExcelZip -ErrorAction Stop } |
                Should -Throw -ExceptionType ([InvalidDataException]) -ErrorId 'UnknownFileContent,ExcelFast.PowerShell.Cmdlets.ImportCommand'
        }
    }

    Context 'When specifying sheet names' {
        It "Should throw ArgumentException when sheet doesn't exist" {
            { Import-Workbook -Path $ValidExcelFile -SheetName 'NonExistentSheet' -ErrorAction Stop } |
                Should -Throw -ExceptionType ([ArgumentException]) -ErrorId 'InvalidSheetName,ExcelFast.PowerShell.Cmdlets.ImportCommand'
        }

        It 'Should import data when specifying Sheet1' {
            $actual = Import-Workbook -Path $ValidExcelFile -SheetName 'Sheet1'

            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 10
            $actual[0].Name | Should -Be 'Test1'
            $actual[0].Value | Should -Be 'Value1'
            $actual[-1].Name | Should -Be 'Test10'
            $actual[-1].Value | Should -Be 'Value10'
        }
    }

    Context 'When importing multiple files' {
        It 'Should process an array of file paths' {
            # Use the same file twice
            $actual = Import-Workbook -Path @($ValidExcelFile, $ValidExcelFile)

            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 20
            $actual[0].Name | Should -Be 'Test1'
            $actual[0].Value | Should -Be 'Value1'
            $actual[-1].Name | Should -Be 'Test10'
            $actual[-1].Value | Should -Be 'Value10'
            $actual[10].Name | Should -Be 'Test1'
            $actual[10].Value | Should -Be 'Value1'
            $actual[19].Name | Should -Be 'Test10'
            $actual[19].Value | Should -Be 'Value10'
        }
    }

    Context 'Cell Range Parameters' {
        It 'Should import data from specified start cell' {
            $actual = Import-Workbook -Path $ValidExcelFile -StartCell 'B2'

            $actual | Should -Not -BeNullOrEmpty
            $actual.Count | Should -BeLessThan 10 # Since we're starting from B2, we should get fewer rows
        }

        It 'Should import data with specified range when NoHeaders is true' {
            $actual = Import-Workbook -Path $ValidExcelFile -StartCell 'A1' -EndCell 'B5' -NoHeaders

            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 5 # Get all rows in range
            $actual[0].A | Should -Not -BeNullOrEmpty
            $actual[0].B | Should -Not -BeNullOrEmpty
        }

        It 'Should import data with specified range when headers are used' {
            $actual = Import-Workbook -Path $ValidExcelFile -StartCell 'A1' -EndCell 'B5'

            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 4 # First row used as headers
            $actual[0].Name | Should -Not -BeNullOrEmpty
            $actual[0].Value | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Headers Parameter' {
        It 'Should use first row as headers by default' {
            $actual = Import-Workbook -Path $ValidExcelFile

            $actual | Should -Not -BeNullOrEmpty
            $actual[0].PSObject.Properties.Name | Should -Contain 'Name'
            $actual[0].PSObject.Properties.Name | Should -Contain 'Value'
        }

        It 'Should use excel column letters when NoHeaders is specified' {
            $actual = Import-Workbook -Path $ValidExcelFile -NoHeaders

            $actual | Should -Not -BeNullOrEmpty
            # Column names should be A, B, etc. when NoHeaders is used
            $actual[0].PSObject.Properties.Name | Should -Contain 'A'
            $actual[0].PSObject.Properties.Name | Should -Contain 'B'
        }
    }

    Context 'Empty Row Handling' {
	    # needs a Test file with empty rows
        It -Name 'Should skip empty rows by default' {
            $actual = Import-Workbook -Path $SkippedRow
            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 2
            $actual[1].Column1 | Should -Be 'ValueR4C1'
        }
        It -Name 'Should not skip empty rows when IncludeEmptyRows is specified' {
            $actual = Import-Workbook -Path $SkippedRow -IncludeEmptyRows
            $actual | Should -Not -BeNullOrEmpty
            $actual | Should -HaveCount 3
            $actual[0].Column1 | Should -Be 'ValueR2C1'
            $actual[1].Column1 | Should -BeNullOrEmpty
            $actual[1].PSObject.Properties | Should -HaveCount 3
            $actual[2].Column1 | Should -Be 'ValueR4C1'
        }
    }
}
