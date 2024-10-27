namespace Resources;

public static class StreamExtensions
{
    /// <summary>
    ///     Reads a <see cref="string" /> from the given <see cref="Stream" /> and does
    ///     not reset the <see cref="Stream.Position" /> property.
    /// </summary>
    /// <param name="source">The <see cref="Stream" /> from which to read the <see cref="string" />.</param>
    /// <returns>The <see cref="string" />.</returns>
    public static string ReadString(this Stream source)
    {
        ArgumentNullException.ThrowIfNull(source);

        var reader = new StreamReader(source);
        return reader.ReadToEnd();
    }
}
