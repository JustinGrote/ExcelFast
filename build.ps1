#requires -Modules @{ModuleName='Microsoft.PowerShell.Platyps'; ModuleVersion='1.0.0'}
using namespace System.Management.Automation

#This script is called by MSBuild
param(
	# The version of the module. Will be generated via gitversion if not specified.
	[Management.Automation.SemanticVersion]$Version = $ENV:MODULEVERSION,

	[string]$ModuleName = 'ExcelFast',

	[string]$PublishPath = "$PSScriptRoot/Build/Module",

	[ValidateNotNullOrWhiteSpace()]
	[string]$ManifestPath = "$PublishPath/$ModuleName.psd1",

	#Specify this for a non-debug release
	[switch]$Production,

	#Dont create a .nupkg
	[switch]$NoPackage,

	[string]$PackagePath = (Split-Path $PublishPath -Parent)
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
		# If this is a tagged build, use the version from the tag
		[SemanticVersion[]]$tag = git tag --points-at HEAD
		| ForEach-Object {
			try {
				[SemanticVersion]($_ -replace '^v')
			} catch {
				Write-Verbose -Fore Yellow "Tag '$_' is not a valid semantic version. Skipping."
			}
		}
		| Sort-Object -Descending

		if ($tag.Count -ge 1) {
			if ($tag.Count -gt 1) {
				Write-Warning "Multiple version tags ($($tag -join ', ')) found pointing to HEAD. Will build for the highest version found ($($tag[0]))."
			}
			$selectedTag = $tag[0]
			Write-Host -Fore Green "Using version from tag: $selectedTag"
			$Version = $selectedTag
		} else {
			Write-Host -Fore Yellow 'No tag found. Using GitVersion to determine the version.'
			# Get the module verison
			dotnet tool restore
			$versionInfo = dotnet gitversion | ConvertFrom-Json

			# Update the module version in the manifest
			$moduleVersion = $versionInfo.MajorMinorPatch

			# If this is running in Github Actions, use the run id and attempt ID as the prereleasenumber
			if ($env:GITHUB_RUN_NUMBER -and $env:GITHUB_RUN_ATTEMPT) {
				$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $env:GITHUB_RUN_NUMBER.ToString('D3') + '.' + $env:GITHUB_RUN_ATTEMPT.ToString('D3') + '.' + $versionInfo.ShortSha
			} else {
				# Otherwise, use the short sha as the prereleasenumber
				$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $versionInfo.ShortSha
			}

			$Version = $moduleVersion + '-' + $modulePrerelease
		}
	}

	Write-Host -Fore Cyan "Module Version: $Version"

	# Update the module manifest
	Update-ModuleManifest -Path $manifestPath -CmdletsToExport $cmdletsToExport -AliasesToExport $aliasesToExport -ModuleVersion ([version]$Version) -Prerelease 'PRERELEASEPLACEHOLDER'

	#BUG: Update-ModuleManifest does not support build characters in the version string, hence this workaround.
	$manifestContent = Get-Content -Path $manifestPath -Raw
	$manifestContent = $manifestContent -replace 'PRERELEASEPLACEHOLDER', $Version.PreReleaseLabel
	Set-Content -Path $manifestPath -Value $manifestContent -NoNewline
	Write-Host "Module manifest '$manifestPath' updated with cmdlets and aliases."

	# Generate PlatyPS Markdown files
	$newMarkdownCommandHelpSplat = @{
		ModuleInfo     = (Import-Module $manifestPath -Force -PassThru)
		OutputFolder   = "$PSScriptRoot/Docs/Commands"
		HelpVersion    = ([version]$Version)
		WithModulePage = $true
	}
	$helpResult = New-MarkdownCommandHelp @newMarkdownCommandHelpSplat

	# Clean up by removing the imported module
	Remove-Module -Name $ModuleName -Force

	# Create a nupkg with a dumb workaround
	if (-not $NoPackage) {
		try {
			try {
				Remove-Item $PackagePath/*.nupkg -ErrorAction Stop
			} catch {
				if ($_ -notmatch 'it does not exist') { throw }
			}
			Register-PSResourceRepository -Name 'PublishLocal' -Uri $PackagePath -Force
			Publish-PSResource -Path $PublishPath -Repository 'PublishLocal'
		} finally {
			Unregister-PSResourceRepository -Name 'PublishLocal'
		}
	}
	Write-Host "Module nupkg published to $PackagePath"


} finally {
	# Return to the original location
	Pop-Location
}
