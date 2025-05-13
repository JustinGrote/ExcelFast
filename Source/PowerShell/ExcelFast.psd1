@{
	# Script module or binary module file associated with this manifest.
	RootModule = 'PowerShell.dll'

	# Version number of this module.
	ModuleVersion = '0.0.0'

	# Supported PSEditions
	CompatiblePSEditions = @('Core', 'Desktop')

	# ID used to uniquely identify this module
	GUID = '139f6aa7-d6df-450c-aebf-25b66f025812'

	# Author of this module
	Author = 'Justin Grote github:@JustinGrote bsky:@posh.guru'

	# Company or vendor of this module
	CompanyName = 'ExcelFast'

	# Copyright statement for this module
	Copyright = '(c) 2025 Justin Grote. All rights reserved.'

	# Description of the functionality provided by this module
	Description = 'High-performance Excel operations for PowerShell'

	# Minimum version of the PowerShell engine required by this module
	PowerShellVersion = '7.4'

	# Cmdlets to export from this module
	CmdletsToExport = @('*')

	# Variables to export from this module
	VariablesToExport = @()

	# Aliases to export from this module
	AliasesToExport      = @('*')

	# Private data to pass to the module specified in RootModule/ModuleToProcess
	PrivateData = @{
		PSData = @{
			# Tags applied to this module
			Tags = @('Excel', 'ImportExcel', 'Performance', 'Data')

			# License URI for this module
			LicenseUri = 'https://github.com/JustinGrote/ExcelFast/blob/main/LICENSE'

			# Project URI for this module
			ProjectUri = 'https://github.com/JustinGrote/ExcelFast'

			# A URL to an icon representing this module
			IconUri = 'https://raw.githubusercontent.com/JustinGrote/ExcelFast/main/Source/Media/icon.png'

			# ReleaseNotes of this module
			ReleaseNotes = 'https://github.com/JustinGrote/ExcelFast/blob/main/CHANGELOG.md'

			# Prerelease string of this module
			Prerelease   = 'Source'
		}
	}
}