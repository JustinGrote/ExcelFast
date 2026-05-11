$importModule = Get-Command -Name Import-Module -Module Microsoft.PowerShell.Core
$moduleName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)

$libPath = if ($IsCoreCLR) {
	Join-Path $PSScriptRoot 'lib\net8.0'
} else {
	Join-Path $PSScriptRoot 'lib\net472'
}

if (-not $IsCoreClr) {
	# PowerShell 5.1 has no concept of an Assembly Load Context so it will
	# just load the module assembly directly.

	# Create a redirect handler for our newer versions of certain system Libraries
	$OnAssemblyResolve = [System.ResolveEventHandler] {
		param($null, $assemblyToResolve)
		# Check if this is the assembly causing the conflict
		if ($assemblyToResolve) {
			$assemblyName = $assemblyToResolve.Name.Split(',')[0]
			$assemblyPath = Join-Path $libPath "$assemblyName.dll"
			if (Test-Path $assemblyPath) {
				# Load and return the specific version you want to use instead
				Write-Debug "ExcelFast: Redirecting assembly load for $assemblyName to $assemblyPath"
				return [System.Reflection.Assembly]::LoadFrom($assemblyPath)
			}
		}
		return $null
	}

	# Register the handler to the current AppDomain
	[System.AppDomain]::CurrentDomain.add_AssemblyResolve($OnAssemblyResolve)
} else {
	# This is used to load the shared assembly in the Default ALC which then sets
	# an ALC for the module and any dependencies of that module to be loaded in
	# that ALC.

	$isReload = $true
	if (-not ('ExcelFast.ALCLoader.LoadContext' -as [type])) {
		$isReload = $false

		Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, 'lib', 'net8.0', "$moduleName.ALCLoader.dll"))
	}

	$mainModule = [ExcelFast.ALCLoader.LoadContext]::Initialize()
	$innerMod = &$importModule -Assembly $mainModule -PassThru:$isReload
}

if ($innerMod) {
	# Bug in pwsh, Import-Module in an assembly will pick up a cached instance
	# and not call the same path to set the nested module's cmdlets to the
	# current module scope. This is only technically needed if someone is
	# calling 'Import-Module -Name ALCLoader -Force' a second time. The first
	# import is still fine.
	# https://github.com/PowerShell/PowerShell/issues/20710
	$addExportedCmdlet = [System.Management.Automation.PSModuleInfo].GetMethod(
		'AddExportedCmdlet',
		[System.Reflection.BindingFlags]'Instance, NonPublic'
	)
	foreach ($cmd in $innerMod.ExportedCmdlets.Values) {
		$addExportedCmdlet.Invoke($ExecutionContext.SessionState.Module, @(, $cmd))
	}
}

#Load the binary module
Import-Module -Name (Join-Path $libPath 'ExcelFast.dll') -Force