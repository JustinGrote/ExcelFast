using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'ClosedXML feature cmdlet tests' {
  BeforeEach {
    $WorkingFile = Join-Path 'TestDrive:' ("{0}.xlsx" -f [guid]::NewGuid())
    Copy-Item -Path $ValidExcelFile -Destination $WorkingFile
  }

  Context 'Organizing worksheets examples' {
    It 'Should add a worksheet with defaults' {
      $workbook = Get-Workbook -Path $WorkingFile

      $worksheet = Add-Worksheet -Workbook $workbook

      $worksheet | Should -Not -BeNullOrEmpty
      $workbook.Worksheets.Count | Should -Be 3
      ($workbook.Worksheets | Select-Object -Last 1).Name | Should -Be $worksheet.Name
    }

    It 'Should add a named worksheet at a specific position' {
      $workbook = Get-Workbook -Path $WorkingFile

      $worksheet = Add-Worksheet -Workbook $workbook -Name 'Import' -Position 2

      $worksheet.Position | Should -Be 2
      $workbook.Worksheet(2).Name | Should -Be 'Import'
    }

    It 'Should remove a worksheet by name' {
      $workbook = Get-Workbook -Path $WorkingFile

      Remove-Worksheet -Workbook $workbook -Name 'NamedSheet' | Out-Null

      $workbook.Worksheets.Name | Should -Not -Contain 'NamedSheet'
      $workbook.Worksheets.Count | Should -Be 1
    }

    It 'Should move worksheet to a new position' {
      $workbook = Get-Workbook -Path $WorkingFile

      Move-Worksheet -Workbook $workbook -Name 'Sheet1' -Position 2 | Out-Null

      $workbook.Worksheet('Sheet1').Position | Should -Be 2
    }

    It 'Should throw when removing a missing worksheet' {
      $workbook = Get-Workbook -Path $WorkingFile

      { Remove-Worksheet -Workbook $workbook -Name 'DoesNotExist' -ErrorAction Stop } |
        Should -Throw -ExceptionType ([ArgumentException]) -ErrorId 'WorksheetNotFound,ExcelFast.PowerShell.Cmdlets.RemoveWorksheetCommand'
    }
  }

  Context 'Table examples' {
    It 'Should create a table from worksheet range with theme' {
      $workbook = Get-Workbook -Path $WorkingFile
      $worksheet = $workbook.Worksheet('Sheet1')

      $table = Add-Table -Worksheet $worksheet -RangeAddress 'A1:B11' -Name 'PastrySales' -Theme 'TableStyleLight16'

      $table | Should -Not -BeNullOrEmpty
      $table.Name | Should -Be 'PastrySales'
      $table.Theme.ToString() | Should -Be 'TableStyleLight16'
    }

    It 'Should create table from range input and apply style options' {
      $workbook = Get-Workbook -Path $WorkingFile
      $range = $workbook.Worksheet('Sheet1').Range('A1:B11')

      $table = $range | Add-Table -Name 'StyledTable' -ShowTotalsRow -HideRowStripes -ShowColumnStripes

      $table.ShowTotalsRow | Should -BeTrue
      $table.ShowRowStripes | Should -BeFalse
      $table.ShowColumnStripes | Should -BeTrue
    }

    It 'Should throw for unsupported table theme' {
      $workbook = Get-Workbook -Path $WorkingFile
      $worksheet = $workbook.Worksheet('Sheet1')

      { Add-Table -Worksheet $worksheet -RangeAddress 'A1:B11' -Theme 'BadTheme' -ErrorAction Stop } |
        Should -Throw -ExceptionType ([ArgumentException]) -ErrorId 'InvalidTableTheme,ExcelFast.PowerShell.Cmdlets.AddTableCommand'
    }
  }
}
