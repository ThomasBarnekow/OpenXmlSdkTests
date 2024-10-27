using System.Reflection;

namespace Resources;

public static class ResourceHelper
{
    public static string GetString(Assembly assembly, string relativePath)
    {
        using Stream stream = GetManifestResourceStream(assembly, relativePath);
        return stream.ReadString();
    }

    public static MemoryStream GetMemoryStream(Assembly assembly, string relativePath)
    {
        using Stream resourceStream = GetManifestResourceStream(assembly, relativePath);

        // Create an editable MemoryStream from the read-only manifest resource stream.
        var memoryStream = new MemoryStream();
        resourceStream.CopyTo(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);

        return memoryStream;
    }

    private static Stream GetManifestResourceStream(Assembly assembly, string relativePath)
    {
        ArgumentNullException.ThrowIfNull(assembly);
        ArgumentNullException.ThrowIfNull(relativePath);

        string resourceName = GetResourceName(assembly, relativePath);
        return assembly.GetManifestResourceStream(resourceName) ?? Stream.Null;
    }

    private static string GetResourceName(Assembly assembly, string relativePath)
    {
        return assembly.GetName().Name + "." + relativePath.Replace("\\", ".").Replace("/", ".");
    }
}
