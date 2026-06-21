namespace ExcelFast.Extensions;

using System.Collections;
using System.Management.Automation.Runspaces;
using System.Reflection;

public static class PSObjectExtensions
{
	extension(PSObject psobject)
	{
    public Dictionary<string, object> ToFlatDictionary(out Dictionary<string, string> conversionErrors, bool includeUnexportableProperties = false)
    {
      conversionErrors = new Dictionary<string, string>();
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
      PSMemberInfoCollection<PSPropertyInfo> properties = psobject.Properties;
      foreach (PSPropertyInfo property in properties)
      {
        if (string.IsNullOrWhiteSpace(property.Name))
        {
          continue;
        }

        try
        {
          object? propertyValue = property.Value;
          dictionary[property.Name] = LanguagePrimitives.ConvertTo<string>(propertyValue);
        }
        catch (Exception ex)
        {
          conversionErrors[property.Name] = $"could not be processed: {ex.Message}";
          if (includeUnexportableProperties)
          {
            dictionary[property.Name] = string.Empty;
          }
        }
      }

      if (dictionary.Count > 0)
      {
        return dictionary;
      }

      foreach (PropertyInfo property in baseObject.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
      {
        if (property.GetIndexParameters().Length > 0 || string.IsNullOrWhiteSpace(property.Name))
        {
          continue;
        }

        try
        {
          object? propertyValue = property.GetValue(baseObject);
          dictionary[property.Name] = LanguagePrimitives.ConvertTo<string>(propertyValue);
        }
        catch (Exception ex)
        {
          conversionErrors[property.Name] = $"could not be processed: {ex.Message}";
          if (includeUnexportableProperties)
          {
            dictionary[property.Name] = string.Empty;
          }
        }
      }

      return dictionary;
    }
  }
}