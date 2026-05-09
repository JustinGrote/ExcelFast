# Proactively load dependency assemblies from the module directory.
# This avoids runtime probing differences across hosts/OS (notably CI on Linux).
$moduleRoot = Split-Path -Parent $PSCommandPath
$dependencyDlls = @(
	'Microsoft.Bcl.AsyncInterfaces.dll'
	'System.Threading.Tasks.Extensions.dll'
	'System.Runtime.CompilerServices.Unsafe.dll'
	'System.Memory.dll'
	'System.Buffers.dll'
	'System.Numerics.Vectors.dll'
	'System.IO.Packaging.dll'
	'DocumentFormat.OpenXml.Framework.dll'
	'DocumentFormat.OpenXml.dll'
	'ClosedXML.Parser.dll'
	'ClosedXML.dll'
	'ExcelNumberFormat.dll'
	'SixLabors.Fonts.dll'
	'RBush.dll'
	'MiniExcel.dll'
)

foreach ($dll in $dependencyDlls) {
	$dependencyPath = Join-Path $moduleRoot $dll
	if (-not (Test-Path $dependencyPath)) {
		continue
	}

	$assemblyName = [System.IO.Path]::GetFileNameWithoutExtension($dll)
	$alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
		Where-Object { $_.GetName().Name -eq $assemblyName } |
		Select-Object -First 1
	if ($null -ne $alreadyLoaded) {
		continue
	}

	try {
		[System.Reflection.Assembly]::LoadFrom($dependencyPath) | Out-Null
	} catch {
		# Best-effort load. Remaining dependencies can still resolve via default probing.
	}
}

# Register an assembly resolver for Windows PowerShell 5.1.
if ($PSEdition -eq 'Desktop' -and -not $global:ExcelFastAssemblyResolveRegistered) {
	Write-Verbose 'Registering assembly resolver for Windows PowerShell 5.1'
	$global:ExcelFastAssemblyResolveHandler = [System.ResolveEventHandler] {
		param($sender, $eventArgs)

		try {
			$requestedAssembly = New-Object System.Reflection.AssemblyName($eventArgs.Name)
			$requestedName = $requestedAssembly.Name

			$alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
				Where-Object { $_.GetName().Name -eq $requestedName } |
				Select-Object -First 1
			if ($null -ne $alreadyLoaded) {
				return $alreadyLoaded
			}

			$dependencyPath = Join-Path $moduleRoot ($requestedName + '.dll')
			if (-not (Test-Path $dependencyPath))
			{
				return $null
			}

			return [System.Reflection.Assembly]::LoadFrom($dependencyPath)
		}
		catch
		{
			return $null
		}
	}

	[AppDomain]::CurrentDomain.add_AssemblyResolve($global:ExcelFastAssemblyResolveHandler)
	$global:ExcelFastAssemblyResolveRegistered = $true
}
