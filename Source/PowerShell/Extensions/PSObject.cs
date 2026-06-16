namespace ExcelFast.Extensions;

using System.Collections;
using System.Management.Automation.Runspaces;

public static class PSObjectExtensions
{
	extension(PSObject psobject)
	{
		public Dictionary<string, object> ToFlatDictionary()
		{
			Dictionary<string, object> dictionary = new Dictionary<string, object>();

			object baseObject = psobject.BaseObject;

			// Handle IDictionary (Hashtable, OrderedDictionary, generic Dictionary, etc.)
			if (baseObject is IDictionary dict)
			{
				foreach (DictionaryEntry entry in dict)
				{
					string key = entry.Key?.ToString() ?? string.Empty;
					if (!string.IsNullOrWhiteSpace(key))
						dictionary[key] = entry.Value ?? string.Empty;
				}
				return dictionary;
			}

			// Handle arrays/lists (produce indexed Column1, Column2, ... columns)
			if (baseObject is not string && baseObject is IList list)
			{
				for (int i = 0; i < list.Count; i++)
					dictionary[$"Column{i + 1}"] = list[i]?.ToString() ?? string.Empty;
				return dictionary;
			}

      // Handle scalars (primitives, strings, DateTime, etc.) - map to "Value" column
      Type baseType = baseObject.GetType();
      if (baseType.IsPrimitive || baseObject is string || baseObject is decimal || baseObject is DateTime)
			{
				dictionary["Value"] = baseObject;
				return dictionary;
			}

			// Standard PSObject properties (PSCustomObject, .NET objects, etc.)
			foreach (PSPropertyInfo property in psobject.Properties)
			{
				if (string.IsNullOrWhiteSpace(property.Name))
				{
					continue;
				}

				string value;
        object? propertyValue = GetPropertyValue(psobject, property);
        value = propertyValue is not string && propertyValue is not IDictionary && propertyValue is IEnumerable<object> enumerable
          ? string.Join(", ", enumerable.Select(x => x?.ToString() ?? string.Empty))
          : propertyValue?.ToString() ?? string.Empty;

        dictionary[property.Name] = value;
      }

      return dictionary;
    }

    private static object? GetPropertyValue(PSObject targetObject, PSPropertyInfo property)
    {
      if (property is not (PSScriptProperty or PSCodeProperty))
      {
        return property.Value;
      }

      Runspace? originalRunspace = Runspace.DefaultRunspace;
      Runspace targetRunspace = originalRunspace ?? RunspaceFactory.CreateRunspace();
      if (targetRunspace.RunspaceStateInfo.State == RunspaceState.BeforeOpen)
      {
        targetRunspace.Open();
      }
      try
      {
        Runspace.DefaultRunspace = targetRunspace;
        return property.Value;
      }
      finally
      {
        Runspace.DefaultRunspace = originalRunspace;
        if (originalRunspace is null)
        {
          targetRunspace.Dispose();
        }
      }
    }
  }
}