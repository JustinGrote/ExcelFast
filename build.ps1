#This script is called by MSBuild
param(
	# The version of the module. Will be generated via gitversion if not specified.
	[Management.Automation.SemanticVersion]$Version,

	[string]$ModuleName = 'ExcelFast',

	[ValidateNotNullOrWhiteSpace()]
	[string]$ManifestPath = "Build/$ModuleName.psd1",

	#Specify this for a non-debug release
	[switch]$Production
)

$ErrorActionPreference = 'Stop'

# Build the module
try {
	Push-Location -Path $PSScriptRoot
	dotnet publish -c ($Production ? 'Release' : 'Debug') --version-suffix $($Production ? '' : 'dev')

	# Import the module to discover its commands and aliases
	$manifestPath = Resolve-Path $ManifestPath
	Import-Module -Name $manifestPath -Force

	# Get cmdlets to export
	$cmdletsToExport = (Get-Command -CommandType Cmdlet -Module $ModuleName).Name

	# Get aliases to export
	$aliasesToExport = (Get-Alias | Where-Object { $_.ResolvedCommand.Module.Name -eq $ModuleName }).Name

	if ($null -eq $Version) {
		# Get the module verison
		dotnet tool restore
		$versionInfo = dotnet gitversion | ConvertFrom-Json

		# Update the module version in the manifest
		$moduleVersion = $versionInfo.MajorMinorPatch
		$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $versionInfo.ShortSha
		$Version = $moduleVersion + '-' + $modulePrerelease
	}

	Write-Host -Fore Cyan "Module Version: $Version"

	# Update the module manifest
	Update-ModuleManifest -Path $manifestPath -CmdletsToExport $cmdletsToExport -AliasesToExport $aliasesToExport -ModuleVersion ([version]$Version) -Prerelease 'PRERELEASEPLACEHOLDER'

	#BUG: Update-ModuleManifest does not support build characters in the version string, hence this workaround.
	$manifestContent = Get-Content -Path $manifestPath -Raw
	$manifestContent = $manifestContent -replace 'PRERELEASEPLACEHOLDER', $Version.PreReleaseLabel
	Set-Content -Path $manifestPath -Value $manifestContent -NoNewline

	# Clean up by removing the imported module
	Remove-Module -Name $ModuleName -Force
	Write-Host "Module manifest '$manifestPath' updated with cmdlets and aliases."
} finally {
	# Return to the original location
	Pop-Location
}
