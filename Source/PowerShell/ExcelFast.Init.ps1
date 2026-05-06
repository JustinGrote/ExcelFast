# Register an assembly resolver for Windows PowerShell 5.1.

if ($PSEdition -eq 'Desktop' -and -not $global:ExcelFastAssemblyResolveRegistered)
{
	Write-Verbose "Registering assembly resolver for Windows PowerShell 5.1"
	$global:ExcelFastAssemblyResolveHandler = [System.ResolveEventHandler]{
		param($sender, $eventArgs)

		try
		{
			$requestedAssembly = New-Object System.Reflection.AssemblyName($eventArgs.Name)
			$requestedName = $requestedAssembly.Name
			if (-not $requestedName.StartsWith('System.', [System.StringComparison]::OrdinalIgnoreCase))
			{
				return $null
			}

			$alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() |
				Where-Object { $_.GetName().Name -eq $requestedName } |
				Select-Object -First 1
			if ($null -ne $alreadyLoaded)
			{
				return $alreadyLoaded
			}

			$moduleRoot = Split-Path -Parent $PSCommandPath
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
