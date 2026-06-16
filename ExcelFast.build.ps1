#requires -module InvokeBuild, Microsoft.PowerShell.PlatyPS
using namespace System.Management.Automation

#This script is called by MSBuild
param(
	# The version of the module. Will be generated via gitversion if not specified.
	[SemanticVersion]$Version,

	[string]$ModuleName = 'ExcelFast',

	[string]$PublishPath = (Join-Path $PSScriptRoot 'Artifacts\Module'),

	[ValidateNotNullOrWhiteSpace()]
	[string]$ManifestPath = (Join-Path $PublishPath "$ModuleName.psd1"),

	#Specify this for a non-debug release
	[switch]$Production,

	#Dont create a .nupkg
	[switch]$NoPackage,

	[string]$PackagePath = (Split-Path $PublishPath -Parent),

	#Don't build for Windows PowerShell 5.1. This is primarily intended as a workaround for CI builds on Linux where the .NET Framework assemblies cause issues.
	[switch]$SkipPS51
)

Set-BuildHeader {
	param($Path)
	"👷 $Path $(Get-BuildSynopsis $Task)"
}
Set-BuildFooter {
	param($Path)
	"✅ $Path $(Get-BuildSynopsis $Task)"
}

$ErrorActionPreference = 'Stop'

# This gets awkward in the context of the cast parameter so we do it here.
if (-not $Version -and $ENV:MODULE_VERSION) {
	$Version = $ENV:MODULE_VERSION
}

Task Clean {
	#Clean the publish directory
	Write-Host -Fore Cyan "Cleaning publish directory: $PublishPath"
	$env:GIT_ASK_YESNO = 'false'
	git clean -fdx --no-interactive -- (Join-Path $PSScriptRoot 'Artifacts\Module') (Join-Path $PSScriptRoot 'Artifacts\*.nupkg')
	if ($LASTEXITCODE -ne 0) {
    throw "Failed to clean publish directory. Git clean exited with code $LASTEXITCODE. A file is probably locked. Is the module loaded in a running PowerShell window?"
	}
}

Task CompilePS7 {
	$projectFile = Get-ChildItem -Path $PSScriptRoot -Filter '*.csproj' -Recurse | Select-Object -First 1
	if (-not $projectFile) {
		throw 'Could not find a .csproj file to determine target framework.'
	}

	[xml]$projectXml = Get-Content -Path $projectFile.FullName -Raw
	$targetFramework = @($projectXml.Project.PropertyGroup.TargetFramework | Where-Object { $_ })[0]
	if (-not $targetFramework) {
		$targetFrameworks = @($projectXml.Project.PropertyGroup.TargetFrameworks | Where-Object { $_ })[0]
		if ($targetFrameworks) {
			$targetFramework = ($targetFrameworks -split ';')[0]
		}
	}

	if (-not $targetFramework) {
		throw "Could not determine TargetFramework from $($projectFile.FullName)."
	}

	#TODO: Get framework from csproj
	$framework = 'net8.0'
	$publishArgs = @(
		'-c', ($Production ? 'Release' : 'Debug'),
		'--version-suffix', $($Production ? '' : 'dev'),
		'-f', $framework,
		'-o', (Join-Path $PSScriptRoot "Artifacts\Module\lib\$framework"),
		'-p:GenerateDocumentationFile=true',
		(Join-Path $PSScriptRoot 'Source\PowerShell\PowerShell.csproj')
	)
	dotnet publish @publishArgs
}

Task CompilePS51 {
	#TODO: Get framework from csproj
	$framework = 'net472'
	$publishArgs = @(
		'-c', ($Production ? 'Release' : 'Debug'),
		'--version-suffix', $($Production ? '' : 'dev'),
		'-f', $framework,
		'-o', (Join-Path $PSScriptRoot "Artifacts\Module\lib\$framework"),
		'-p:GenerateDocumentationFile=true',
		(Join-Path $PSScriptRoot 'Source\PowerShell\PowerShell.csproj')
	)
	dotnet publish @publishArgs
}

Task CopyModuleFiles {
	# Copy the files and preserve directores
	$sourcePath = Join-Path $PSScriptRoot 'Source\PowerShell\Module'
	Copy-Item -Path (Join-Path $sourcePath '*') -Destination $PublishPath -Force -Recurse
}

Task Build CompileAll, CopyModuleFiles, {
  [SemanticVersion]$BuildVersion = $Version
  try {
    Push-Location -Path $PSScriptRoot

    # Import the module to discover its commands and aliases
    $manifestPath = Resolve-Path $ManifestPath

    #HACK: Because importing the module loads the .NET assemblies and locks them to the session, we want it in a separate process.
    $job = Start-Job -ArgumentList $ManifestPath, $ModuleName -ScriptBlock {
      param(
        [string]$ManifestPath,
        [string]$ModuleName
      )

      Import-Module -Name $ManifestPath -Force

      return @{
        CmdletsToExport = (Get-Command -CommandType Cmdlet -Module $ModuleName).Name
        AliasesToExport = (Get-Alias | Where-Object { $_.ResolvedCommand.Module.Name -eq $ModuleName }).Name
      }
    }

    $jobOutput = Receive-Job -Job $job -Wait -AutoRemoveJob
    $cmdletsToExport = $jobOutput.CmdletsToExport
    $aliasesToExport = $jobOutput.AliasesToExport

    if ($null -eq $BuildVersion) {
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
        $BuildVersion = $selectedTag
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

        $BuildVersion = $moduleVersion + '-' + $modulePrerelease
      }
    }

    Write-Host -Fore Cyan "Module Version: $BuildVersion"

    Write-Host -Fore Cyan "Updating module manifest '$manifestPath' with cmdlets aliases and types"
    $formatAndTypeSourcePath = Join-Path $PSScriptRoot 'Source\PowerShell\Module\Formats'
    if (Test-Path $formatAndTypeSourcePath) {
      [string[]]$formatsToProcess = Get-ChildItem -Path $formatAndTypeSourcePath -Filter '*.format.ps1xml' -File
      | ForEach-Object { Join-Path 'Formats' $_.Name }

      [string[]]$typesToProcess = Get-ChildItem -Path $formatAndTypeSourcePath -Filter '*.types.ps1xml' -File
      | ForEach-Object { Join-Path 'Formats' $_.Name }
    }

    # Update the module manifest
    $updateModuleManifestSplat = @{
      Path             = $manifestPath
      CmdletsToExport  = $cmdletsToExport
      AliasesToExport  = $aliasesToExport
      FormatsToProcess = $formatsToProcess
      TypesToProcess   = $typesToProcess
      #Version becomes Version#WithLabel when cast from semantic version, hence the psbase work
      ModuleVersion    = [version]$BuildVersion
      Prerelease       = 'PRERELEASEPLACEHOLDER'
    }
    Update-ModuleManifest @updateModuleManifestSplat

    #BUG: Update-ModuleManifest does not support build characters in the version string, hence this workaround.
    $manifestContent = Get-Content -Path $manifestPath -Raw
    $manifestContent = $manifestContent -replace 'PRERELEASEPLACEHOLDER', $BuildVersion.PreReleaseLabel
    Set-Content -Path $manifestPath -Value $manifestContent -NoNewline

    #Package the nuget
    Compress-PSResource -Path $PublishPath -DestinationPath $PackagePath

    Remove-Item (Join-Path $PublishPath 'System*.dll') -Force

    Write-Host "Module nupkg published to $PackagePath"
    $SCRIPT:Version = $BuildVersion

  } finally {
    # Return to the original location
    Pop-Location
  }
}

Task Docs Build, {
  # Use only the numeric version for PlatyPS help generation; prerelease labels are not valid for [version].
  $helpVersion = [version](([string]$Version) -replace '[-+].*', '')

  #HACK: Because PlatyPS loads the .NET assemblies and locks them to the session, we want it in a separate process.
  Write-Host -Fore Cyan "Exporting MAML help for $ManifestPath to $PublishPath with version $helpVersion."
  Start-Job -ArgumentList $ManifestPath, $PublishPath, (Join-Path $PSScriptRoot 'Docs/Commands'), $helpVersion -ScriptBlock {
    #requires -Modules @{ModuleName='Microsoft.PowerShell.Platyps'; ModuleVersion='1.0.0'}

    param(
      [string]$ManifestPath,
      [string]$PublishPath,
      [string]$DocsPath,
      [version]$HelpVersion
    )
    Write-Host -Fore Cyan "Exporting MAML help to $PublishPath from job with manifest path: $ManifestPath and help version: $HelpVersion"
    $newMarkdownCommandHelpSplat = @{
      ModuleInfo                  = (Import-Module $ManifestPath -Force -PassThru)
      OutputFolder                = $DocsPath
      HelpVersion                 = [Version]$HelpVersion
      WithModulePage              = $true
      AbbreviateParameterTypeName = $true
      # Ignore warnings about existing markdown files
      WarningAction               = 'SilentlyContinue'
    }

    #Generate for any net new modules or commands that dont have markdown files yet. This allows us to preserve any manual changes to existing markdown files.
    Write-Host -Fore Cyan "Generating markdown command help for new or changed commands. Output folder: $($newMarkdownCommandHelpSplat.OutputFolder)"
    New-MarkdownCommandHelp @newMarkdownCommandHelpSplat | Out-Null

    Get-ChildItem -Recurse -Path $newMarkdownCommandHelpSplat.OutputFolder -Include '*.md'
		| Measure-PlatyPSMarkdown
		| Where-Object FileType -Match 'CommandHelp'
		| Import-MarkdownCommandHelp -Path { $_.FilePath }
		| Export-MamlCommandHelp -OutputFolder $PublishPath -Force
  }
  | Receive-Job -Wait -AutoRemoveJob

  #HACK: PlatyPS exports the help files to a subfolder named after the module, but to work properly it needs to be in a subfolder named after the culture (en-US). Hence this workaround.
  New-Item -ItemType Directory -Force (Join-Path $PublishPath 'en-US') | Out-Null
  Move-Item (Join-Path $PublishPath 'ExcelFast' '*.xml') (Join-Path $PublishPath 'en-US')
  Remove-Item (Join-Path $PublishPath 'ExcelFast') -Recurse | Out-Null
}

Task Pester {
  #Run in a separate job so as not to lock the assemblies
  Start-Job -ScriptBlock { Invoke-Pester } | Receive-Job -Wait -AutoRemoveJob
}

Task Pester-WinPS {
  if (-not $IsWindows) {
    Write-Host -ForegroundColor Yellow 'Skipping Pester-WinPS: non-Windows platform detected.'
    return
  }
  & powershell.exe -noprofile -c {
    $pester = Get-Module -FullyQualified @{ModuleName = 'Pester'; ModuleVersion = '5.0' } -ListAvailable -EA 0
    if (-not $pester) {
      Write-Host -ForegroundColor Cyan 'Pester not found. Installing Pester...'
      Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser -ErrorAction Stop
    }
    Invoke-Pester
  }
}

Task CompileAll CompilePS7, CompilePS51
Task Test Pester, Pester-WinPS
Task . Clean, Build, Docs