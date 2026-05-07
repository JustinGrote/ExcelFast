#requires -Modules @{ModuleName='Microsoft.PowerShell.Platyps'; ModuleVersion='1.0.0'}
using namespace System.Management.Automation

#This script is called by MSBuild
param(
	# The version of the module. Will be generated via gitversion if not specified.
	[Management.Automation.SemanticVersion]$Version,

	[string]$ModuleName = 'ExcelFast',

	[string]$PublishPath = (Join-Path $PSScriptRoot 'Artifacts\Module'),

	[ValidateNotNullOrWhiteSpace()]
	[string]$ManifestPath = (Join-Path $PublishPath "$ModuleName.psd1"),

	#Specify this for a non-debug release
	[switch]$Production,

	#Dont create a .nupkg
	[switch]$NoPackage,

	[string]$PackagePath = (Split-Path $PublishPath -Parent)
)

# This gets awkward in the context of the cast parameter so we do it here.
if (-not $Version -and $ENV:MODULE_VERSION) {
	$Version = $ENV:MODULE_VERSION
}

$ErrorActionPreference = 'Stop'

#Clean the publish directory
Write-Host -Fore Cyan "Cleaning publish directory: $PublishPath"
git clean -fdx -- (Join-Path $PSScriptRoot 'Artifacts\Module') (Join-Path $PSScriptRoot 'Artifacts\*.nupkg')

Write-Host -Fore Cyan 'Building Module'

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
				$runId = ([int]$env:GITHUB_RUN_NUMBER).ToString('D3')
				$attemptId = ([int]$env:GITHUB_RUN_ATTEMPT).ToString('D3')

				$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $runId + '.' + $attemptId + '.' + $versionInfo.ShortSha
			} else {
				# Otherwise, use the short sha as the prereleasenumber
				$modulePrerelease = 'ci-' + $versionInfo.PreReleaseNumber.ToString('D3') + '+' + $versionInfo.ShortSha
			}

			$Version = $moduleVersion + '-' + $modulePrerelease
		}
	}

	Write-Host -Fore Cyan "Module Version: $Version"

	Write-Host -Fore Cyan "Updating module manifest '$manifestPath' with cmdlets aliases and types"
	$formatAndTypeSourcePath = Join-Path $PSScriptRoot 'Source\PowerShell\Formats'
	if (Test-Path $formatAndTypeSourcePath) {
		[string[]]$formatsToProcess = Get-ChildItem -Path $formatAndTypeSourcePath -Filter '*.format.ps1xml' -File | ForEach-Object { Join-Path 'Formats' $_.Name }
		[string[]]$typesToProcess = Get-ChildItem -Path $formatAndTypeSourcePath -Filter '*.types.ps1xml' -File | ForEach-Object { Join-Path 'Formats' $_.Name }
	}

	# Update the module manifest
	Update-ModuleManifest -Path $manifestPath -CmdletsToExport $cmdletsToExport -AliasesToExport $aliasesToExport -FormatsToProcess $formatsToProcess -TypesToProcess $typesToProcess -ModuleVersion ([version]$Version) -Prerelease 'PRERELEASEPLACEHOLDER'

	#BUG: Update-ModuleManifest does not support build characters in the version string, hence this workaround.
	$manifestContent = Get-Content -Path $manifestPath -Raw
	$manifestContent = $manifestContent -replace 'PRERELEASEPLACEHOLDER', $Version.PreReleaseLabel
	Set-Content -Path $manifestPath -Value $manifestContent -NoNewline

	Write-Host -Fore Cyan "Exporting MAML help to $PublishPath"
	# Generate PlatyPS Markdown files
	$newMarkdownCommandHelpSplat = @{
		ModuleInfo                  = (Import-Module $manifestPath -Force -PassThru)
		OutputFolder                = "$PSScriptRoot/Docs/Commands"
		HelpVersion                 = ([version]$Version)
		WithModulePage              = $true
		AbbreviateParameterTypeName = $true
	}
	#Generate for any net new modules or commands that dont have markdown files yet. This allows us to preserve any manual changes to existing markdown files.
	New-MarkdownCommandHelp @newMarkdownCommandHelpSplat | Out-Null

	Get-ChildItem -Recurse -Path $OutputFolder -Include '*.md'
	| Measure-PlatyPSMarkdown
	| Where-Object FileType -Match 'CommandHelp'
	| Import-MarkdownCommandHelp -Path { $_.FilePath }
	| Export-MamlCommandHelp -OutputFolder $PublishPath -Force

	#HACK: PlatyPS exports the help files to a subfolder named after the module, but to work properly it needs to be in a subfolder named after the culture (en-US). Hence this workaround.
	New-Item -ItemType Directory -Force (Join-Path $PublishPath 'en-US') | Out-Null
	Move-Item (Join-Path $PublishPath 'ExcelFast' '*.xml') (Join-Path $PublishPath 'en-US')
	Remove-Item (Join-Path $PublishPath 'ExcelFast') -Recurse | Out-Null

	# Clean up by removing the imported module
	Remove-Module -Name $ModuleName -Force

	#Package the nuget
	Compress-PSResource -Path $PublishPath -DestinationPath $PackagePath

	Write-Host "Module nupkg published to $PackagePath"

} finally {
	# Return to the original location
	Pop-Location
}

Write-Host ''
Write-Host -Fore Green '✅ Build completed successfully!'