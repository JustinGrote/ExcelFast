# These are referenced elsewhere

BeforeAll {
	# Import the module - adjust path as needed for your module structure
	$ModulePath = (Resolve-Path "$PSScriptRoot\..\..\Artifacts\Module\ExcelFast.psd1").Path
	Import-Module $ModulePath -Force

	# Test file paths
	$TestDataPath = Join-Path $PSScriptRoot '..\Fixtures'
	$ValidExcelFile = Join-Path $TestDataPath 'Test10.xlsx'
	$NonExcelExtension = Join-Path $TestDataPath 'NotExcel.txt'
	$NonExcelContent = Join-Path $TestDataPath 'NotExcel.xlsx'
	$NonExcelZip = Join-Path $TestDataPath 'NotExcelZip.xlsx'
	$InvalidPath = Join-Path $TestDataPath 'DoesNotExist.xlsx'
	$SkippedRow = Join-Path $TestDataPath 'SkippedRow.xlsx'
}