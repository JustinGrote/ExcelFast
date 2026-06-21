// AssemblyLoadContext won't work in net472 so we conditionally compile this
// for net5.0 or greater.
namespace ExcelFast.ALCLoader;

using System.Reflection;

#if NET
using System.Runtime.Loader;

public class LoadContext : AssemblyLoadContext
{
  private static LoadContext? _instance;
  private static readonly object _sync = new();

  private readonly Assembly _thisAssembly;
  private readonly AssemblyName _thisAssemblyName;
  private readonly string _assemblyDir;

  private LoadContext(string mainModulePathAssemblyPath)
      : base(name: "ALCLoader", isCollectible: false)
  {
    _assemblyDir = Path.GetDirectoryName(mainModulePathAssemblyPath) ?? "";
    _thisAssembly = typeof(LoadContext).Assembly;
    _thisAssemblyName = _thisAssembly.GetName();
  }

  protected override Assembly? Load(AssemblyName assemblyName)
  {
    // Checks to see if we are trying to access our current assembly
    // (ALCLoader.Shared). If so return the already loaded assembly object
    // as it provides a common interface between Pwsh and the ALC.
    if (AssemblyName.ReferenceMatchesDefinition(_thisAssemblyName, assemblyName))
    {
      return _thisAssembly;
    }

    // Checks to see if the assembly exists in our path, if so load it in
    // the ALC. Otherwise fallback to the default loading behaviour.
    string asmPath = Path.Join(_assemblyDir, $"{assemblyName.Name}.dll");
    if (File.Exists(asmPath))
    {
      return LoadFromAssemblyPath(asmPath);
    }
    else
    {
      return null;
    }
  }

  public static void Initialize()
  {
    lock (_sync)
    {
      string assemblyPath = typeof(LoadContext).Assembly.Location;
      string assemblyName = Path.GetFileNameWithoutExtension(assemblyPath);

      // Removes the '.ALCLoader' from the assembly name to refer to our main module.
      string moduleName = assemblyName.EndsWith(".ALCLoader") ? assemblyName.Substring(0, assemblyName.Length - ".ALCLoader".Length) : assemblyName;
      string modulePath = Path.Combine(
          Path.GetDirectoryName(assemblyPath)!,
          $"{moduleName}.dll"
      );

      // Creates the ALC which loads our module in the ALC and returns
      // the loaded Assembly object for the psm1 to load.
      _instance = new LoadContext(modulePath);
    }
  }
}
#else
/// <summary>
/// This is a fallback LoadContext for net472 where AssemblyLoadContext is not available. It works by hooking into the AppDomain.AssemblyResolve event and manually loading assemblies from the same directory as the ALCLoader assembly when they are requested. This allows us to load the same set of assemblies in both net472 and net5.0+ without needing to worry about the differences in assembly loading between the two runtimes.
/// </summary>
public static class LoadContext
{
  const Assembly? ASSEMBLY_NOT_FOUND = null;
  private static readonly object _sync = new();
  private static bool _initialized;
  public static void Initialize()
  {
    lock (_sync)
    {
      if (!_initialized)
      {
        AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
        _initialized = true;
      }
    }
  }

  private static Assembly? OnAssemblyResolve(object? sender, ResolveEventArgs? assemblyToResolve)
  {
    try
    {
      string assemblyNameToResolve = assemblyToResolve?.Name ?? string.Empty;
      if (string.IsNullOrWhiteSpace(assemblyNameToResolve)) return ASSEMBLY_NOT_FOUND;

      string libPath = Path.GetDirectoryName(typeof(LoadContext).Assembly.Location) ?? string.Empty;
      AssemblyName assemblyName = new(assemblyNameToResolve);
      string assemblyPath = Path.Combine(libPath, $"{assemblyName.Name}.dll");

      return File.Exists(assemblyPath)
        ? Assembly.LoadFrom(assemblyPath)
        : ASSEMBLY_NOT_FOUND;
    }
    catch (Exception ex)
    {
      Console.WriteLine($"ExcelFast: Failed to resolve assembly {assemblyToResolve?.Name}: {ex}");
      return ASSEMBLY_NOT_FOUND;
    }
  }
}
#endif
