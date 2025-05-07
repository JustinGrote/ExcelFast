BeforeAll {
	# Import the module - adjust path as needed for your module structure
	$ModulePath = (Resolve-Path "$PSScriptRoot\..\..\Release\ExcelFast.psd1").Path
	Import-Module $ModulePath -Force

	# Test file paths
	$TestDataPath = Join-Path $PSScriptRoot '..\Fixtures'
	$ValidExcelFile = Join-Path $TestDataPath 'Test10.xlsx'
	$NonExcelExtension = Join-Path $TestDataPath 'NotExcel.txt'
	$NonExcelContent = Join-Path $TestDataPath 'NotExcel.xlsx'
	$NonExcelZip = Join-Path $TestDataPath 'NotExcelZip.xlsx'
	$InvalidPath = Join-Path $TestDataPath 'DoesNotExist.xlsx'
}

Describe 'Import-Excel Command Tests' {
	Context 'When importing a valid Excel file' {
		It 'Should successfully import data from Test10.xlsx' {
			$actual = Import-Workbook -Path $ValidExcelFile

			$actual | Should -Not -BeNullOrEmpty
			$actual.Count | Should -Be 10
			$actual[0].Name | Should -Be 'Test1'
			$actual[0].Value | Should -Be 'Value1'
			$actual[-1].Name | Should -Be 'Test10'
			$actual[-1].Value | Should -Be 'Value10'
		}
	}

	Context 'Path Parameter' {
		It 'Should Error for a non-existent file path' {
			{ Import-Workbook -Path $InvalidPath -ErrorAction Stop } |
				Should -Throw -ExpectedMessage '*Excel file not found*'
		}

		It 'Should Error for a non-Excel extension' {
			{ Import-Workbook -Path $NonExcelExtension -ErrorAction Stop } |
				Should -Throw -ExpectedMessage 'Unsupported file type*'
		}

		It 'Should Error for a non-Excel content' {
			{ Import-Workbook -Path $NonExcelContent -ErrorAction Stop } |
				Should -Throw -ExpectedMessage '* is not a valid Excel file.'
		}

		It 'Should Error for a non-Excel zip file' {
			{ Import-Workbook -Path $NonExcelZip -ErrorAction Stop } |
				Should -Throw -ExpectedMessage 'Sequence contains no elements*'
		}
	}

	Context 'When specifying sheet names' {
		It "Should throw an error when sheet doesn't exist" {

			{ Import-Workbook -Path $ValidExcelFile -SheetName 'NonExistentSheet' -ErrorAction Stop } |
				Should -Throw -ExpectedMessage "*Sheet 'NonExistentSheet' does not exist*"
		}

		It 'Should import data when specifying Sheet1' {
			$actual = Import-Workbook -Path $ValidExcelFile -SheetName 'Sheet1'

			$actual | Should -Not -BeNullOrEmpty
			$actual.Count | Should -Be 10
			$actual[0].Name | Should -Be 'Test1'
			$actual[0].Value | Should -Be 'Value1'
			$actual[-1].Name | Should -Be 'Test10'
			$actual[-1].Value | Should -Be 'Value10'
		}
	}

	Context 'When importing multiple files' {
		It 'Should process an array of file paths' {
			# Use the same file twice for this test
			$actual = Import-Workbook -Path @($ValidExcelFile, $ValidExcelFile)

			# Should have twice as many results as we're reading the same file twice
			$actual | Should -Not -BeNullOrEmpty
			$actual.Count | Should -Be 20
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
}
