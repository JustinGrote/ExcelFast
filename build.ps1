#This script is called by MSBuild
param(
	# Path to the completed module. It must have a module manifest with the same name as the module.
	[ValidateNotNullOrWhiteSpace()]
	[string]$PublishDir = 'Release',

	[string]$ModuleName = 'ExcelFast',

	[ValidateNotNullOrWhiteSpace()]
	[string]$ManifestPath = (Join-Path $PublishDir "$ModuleName.psd1"),

	#Specify this for a non-debug release
	[switch]$Production
)

$ErrorActionPreference = "Stop"

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

	# Get the module verison
	dotnet tool restore
	$versionInfo = dotnet gitversion | ConvertFrom-Json

	# Update the module version in the manifest
	$moduleVersion = $versionInfo.MajorMinorPatch
	$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $versionInfo.ShortSha
	Write-Host -Fore Cyan "Module Version: $moduleVersion-$modulePrerelease"


	# Update the module manifest
	Update-ModuleManifest -Path $manifestPath -CmdletsToExport $cmdletsToExport -AliasesToExport $aliasesToExport -ModuleVersion $moduleVersion -Prerelease 'PRERELEASEPLACEHOLDER'

	#BUG: Update-ModuleManifest does not support build characters in the version string, hence this workaround.
	$manifestContent = Get-Content -Path $manifestPath -Raw
	$manifestContent = $manifestContent -replace 'PRERELEASEPLACEHOLDER', $modulePrerelease
	Set-Content -Path $manifestPath -Value $manifestContent -NoNewline

	# Clean up by removing the imported module
	Remove-Module -Name $ModuleName -Force
	Write-Host "Module manifest '$manifestPath' updated with cmdlets and aliases."
} finally {
	# Return to the original location
	Pop-Location
}
