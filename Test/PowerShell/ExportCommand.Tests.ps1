using namespace System.Collections.Generic
using namespace System.IO
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Scope='Script')]
param()

. $PSScriptRoot/../Fixtures/Fixtures.ps1

Describe 'Export-Workbook Command Tests' {
	BeforeEach {
		$DestPath = Join-Path 'TestDrive:' ("{0}.xlsx" -f [guid]::NewGuid())
	}

	Context 'Help examples' {
		It 'Example 1: Should export objects to Excel' {
			$data = @(
				@{ Name = 'Item1'; Value = 100 }
				@{ Name = 'Item2'; Value = 200 }
			)

			Export-Workbook -Destination $DestPath -InputObject $data
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Name | Should -Be 'Item1'
			$actual[0].Value | Should -Be 100
			$actual[1].Name | Should -Be 'Item2'
			$actual[1].Value | Should -Be 200
		}

		It 'Example 2: Should export pipeline data with Force' {
			# Prime the destination so -Force is required.
			[PSCustomObject]@{ Name = 'Seed'; CPU = 0; Memory = 0 } | Export-Workbook -Destination $DestPath

			Get-Process | Select-Object -First 5 Name, CPU, Memory | Export-Workbook -Destination $DestPath -Force
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -Not -BeNullOrEmpty
			$actual[0].PSObject.Properties.Name | Should -Contain 'Name'
			$actual[0].PSObject.Properties.Name | Should -Contain 'CPU'
		}

		It 'Example 3: Should export to a specific sheet' {
			$users = @(
				@{ Username = 'user1'; Email = 'user1@example.com' }
				@{ Username = 'user2'; Email = 'user2@example.com' }
			)

			Export-Workbook -Destination $DestPath -InputObject $users -SheetName 'Users'
			$workbook = Get-Workbook -Path $DestPath
			$actual = Import-Workbook -Path $DestPath -SheetName 'Users'

			$workbook.Worksheets.Name | Should -Contain 'Users'
			$actual | Should -HaveCount 2
			$actual[0].Username | Should -Be 'user1'
			$actual[1].Username | Should -Be 'user2'
		}

		It 'Example 4: Should export multiple data sets input without error' {
			$salesJsonPath = Join-Path 'TestDrive:' 'sales.json'
			$usersCsvPath = Join-Path 'TestDrive:' 'users.csv'

			@'
[
  { "OrderId": 1, "Amount": 10.5 },
  { "OrderId": 2, "Amount": 20.75 }
]
'@ | Set-Content -Path $salesJsonPath

			@'
Username,Email
user1,user1@example.com
user2,user2@example.com
'@ | Set-Content -Path $usersCsvPath

			$sales = Get-Content -Path $salesJsonPath | ConvertFrom-Json
			$users = Import-Csv -Path $usersCsvPath

			{ Export-Workbook -Destination $DestPath -InputObject @($sales, $users) -ErrorAction Stop } | Should -Not -Throw
			Test-Path $DestPath | Should -BeTrue
		}

		It 'Example 5: Should export cmdlet output' {
			$filesDir = Join-Path 'TestDrive:' 'Data'
			New-Item -ItemType Directory -Path $filesDir | Out-Null
			Set-Content -Path (Join-Path $filesDir 'a.txt') -Value 'a'
			Set-Content -Path (Join-Path $filesDir 'b.txt') -Value 'b'

			Get-ChildItem -Path $filesDir -File | Export-Workbook -Destination $DestPath -Force
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual.Name | Should -Contain 'a.txt'
			$actual.Name | Should -Contain 'b.txt'
		}
	}

	Context 'Object conversion' {
		It 'Should honor WhatIf and not create destination file' {
			$DestPath = Join-Path ([System.IO.Path]::GetTempPath()) ("{0}.xlsx" -f [guid]::NewGuid())
			$rows = @(
				[PSCustomObject]@{ Name = 'WhatIf1'; Value = 'Value1' }
			)

			$rows | Export-Workbook -Destination $DestPath -WhatIf *>&1 | Out-Null

			Test-Path $DestPath | Should -BeFalse
		}

		It 'Should export single hashtable' {
			$row = [ordered]@{ Name = 'Hash1'; Value = 'Value1' }

			$row | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 1
			$actual[0].Name | Should -Be 'Hash1'
			$actual[0].Value | Should -Be 'Value1'
		}
		It 'Should export simple psobject' {
			$row = [PSCustomObject]@{ Name = 'Obj1'; Value = 'Value1' }

			$row | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 1
			$actual[0].Name | Should -Be 'Obj1'
			$actual[0].Value | Should -Be 'Value1'
		}
		It 'Should export multiple psobjects' {
			$rows = @(
				[PSCustomObject]@{ Name = 'Obj1'; Value = 'Value1' }
				[PSCustomObject]@{ Name = 'Obj2'; Value = 'Value2' }
			)

			$rows | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Name | Should -Be 'Obj1'
			$actual[0].Value | Should -Be 'Value1'
			$actual[1].Name | Should -Be 'Obj2'
			$actual[1].Value | Should -Be 'Value2'
		}
		It 'Should export hashtable rows' {
			$rows = @(
				@{ Name = 'Hash1'; Value = 'Value1' }
				@{ Name = 'Hash2'; Value = 'Value2' }
			)

			$rows | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Name | Should -Be 'Hash1'
			$actual[0].Value | Should -Be 'Value1'
			$actual[1].Name | Should -Be 'Hash2'
			$actual[1].Value | Should -Be 'Value2'
		}

		It 'Should export multiple hashtables passed by InputObject parameter' {
			$rows = @(
				[ordered]@{ Name = 'ParamHash1'; Value = 10 }
				[ordered]@{ Name = 'ParamHash2'; Value = 20 }
			)

			Export-Workbook -Destination $DestPath -InputObject $rows
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Name | Should -Be 'ParamHash1'
			$actual[0].Value | Should -Be 10
			$actual[1].Name | Should -Be 'ParamHash2'
			$actual[1].Value | Should -Be 20
		}

		It 'Should export an array of arrays as indexed IDictionary-style columns' {
			$rows = @(
				,[object[]]@('R1', 1, 'X')
				,[object[]]@('R2', 2, 'Y')
			)

			Export-Workbook -Destination $DestPath -InputObject $rows
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Column1 | Should -Be 'R1'
			$actual[0].Column2 | Should -Be 1
			$actual[0].Column3 | Should -Be 'X'
			$actual[1].Column1 | Should -Be 'R2'
			$actual[1].Column2 | Should -Be 2
			$actual[1].Column3 | Should -Be 'Y'
		}

		It 'Should export generic dictionary rows' {
			$row1 = [Dictionary[string, object]]::new()
			$row1['Name'] = 'Dict1'
			$row1['Value'] = 'Value1'

			$row2 = [Dictionary[string, object]]::new()
			$row2['Name'] = 'Dict2'
			$row2['Value'] = 'Value2'

			@($row1, $row2) | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Name | Should -Be 'Dict1'
			$actual[0].Value | Should -Be 'Value1'
			$actual[1].Name | Should -Be 'Dict2'
			$actual[1].Value | Should -Be 'Value2'
		}

		It 'Should export scalar inputs into a Value column' {
			1, 2 | Export-Workbook -Destination $DestPath
			$actual = Import-Workbook -Path $DestPath

			$actual | Should -HaveCount 2
			$actual[0].Value | Should -Be 1
			$actual[1].Value | Should -Be 2
		}
	}
}