namespace ExcelFast.Extensions;

using System.Net;

public static class PSCmdletExtensions
{
  extension(PSCmdlet cmdlet)
  {
    public string ResolveWorkbookPath(string pathItem, out string? tempPath)
    {
      tempPath = null;

      if (Uri.TryCreate(pathItem, UriKind.Absolute, out Uri? uri) &&
        (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
      {
        string fileName = Path.GetFileName(uri.LocalPath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
          string extension = Path.GetExtension(uri.AbsolutePath);
          if (string.IsNullOrWhiteSpace(extension))
          {
            extension = ".tmp";
          }

          tempPath = Path.Combine(Path.GetTempPath(), $"ExcelFast-{Guid.NewGuid():N}{extension}");
        }
        else
        {
          tempPath = Path.Combine(Path.GetTempPath(), fileName);
          if (File.Exists(tempPath))
          {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            tempPath = Path.Combine(Path.GetTempPath(), $"{fileNameWithoutExtension}-{Guid.NewGuid():N}{extension}");
          }
        }

        cmdlet.WriteDebug($"Downloading workbook from '{uri}' to temporary file '{tempPath}'.");

#pragma warning disable SYSLIB0014 // WebRequest is used for Windows PowerShell 5.1 compatibility.
        HttpWebRequest request = WebRequest.CreateHttp(uri);
        request.Method = "GET";
#pragma warning restore SYSLIB0014

        using HttpWebResponse response = (HttpWebResponse)request.GetResponse();
        using Stream contentStream = response.GetResponseStream()
          ?? throw new CmdletInvocationException($"No response stream returned when downloading '{uri}'.");
        using FileStream outputStream = File.Create(tempPath);
        contentStream.CopyTo(outputStream);

        return tempPath;
      }

      return cmdlet.GetUnresolvedProviderPathFromPSPath(pathItem);
    }
  }
}
