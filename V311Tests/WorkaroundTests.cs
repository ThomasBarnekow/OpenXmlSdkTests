using System.Reflection;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using Resources;
using Xunit.Abstractions;

namespace V311Tests;

public class WorkaroundTests
{
    private static readonly Assembly ResourcesAssembly = typeof(ResourceHelper).Assembly;

    private const string HelloWorldXml = "HelloWorldFlatOpc.xml";

    private readonly ITestOutputHelper _output;

    public WorkaroundTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private void PrintPackageParts(IEnumerable<IPackagePart> packageParts, string title)
    {
        _output.WriteLine(title);

        foreach (IPackagePart part in packageParts)
        {
            _output.WriteLine($"- Uri: {part.Uri}, ContentType: {part.ContentType}");
        }

        _output.WriteLine("");
    }

    /// <summary>
    ///     This workaround creates a <see cref="WordprocessingDocument" /> that can be cloned.
    ///     It does that by creating a temporary <see cref="WordprocessingDocument" /> on a <see cref="Stream" />
    ///     and disposing it right away. It then creates a new <see cref="WordprocessingDocument" /> instance
    ///     from the <see cref="Stream" />.
    /// </summary>
    private static WordprocessingDocument FromFlatOpcString(string text, Stream stream, bool isEditable)
    {
        // Create and save a temporary instance on a stream.
        WordprocessingDocument tempDoc = WordprocessingDocument.FromFlatOpcString(text, stream, false);
        tempDoc.Dispose();

        // Create the final instance from the stream.
        return WordprocessingDocument.Open(stream, isEditable);
    }

    [Fact]
    public void FromFlatOpcStringOverloadsBehaveDifferently()
    {
        // Arrange
        string xml = ResourceHelper.GetString(ResourcesAssembly, HelloWorldXml);
        using var stream = new MemoryStream();

        // Act
        // Create a document WITH a stream.
        using WordprocessingDocument streamDoc = WordprocessingDocument.FromFlatOpcString(xml, stream, false);
        List<IPackagePart> streamDocParts = streamDoc.GetPackage().GetParts().OrderBy(p => p.Uri.OriginalString).ToList();
        PrintPackageParts(streamDocParts, "After FromFlatOpcString(xml, stream, false)");

        // Create a document WITHOUT a stream.
        using WordprocessingDocument doc = WordprocessingDocument.FromFlatOpcString(xml);
        List<IPackagePart> docParts = doc.GetPackage().GetParts().OrderBy(p => p.Uri.OriginalString).ToList();
        PrintPackageParts(docParts, "After FromFlatOpcString(xml)");

        // Assert
        // The two documents do not contain the same parts.
        Assert.NotEqual(streamDocParts, docParts, (p1, p2) => p1.Uri == p2.Uri);

        // The document created with FromFlatOpcString(xml, stream, false) contains the relationship parts.
        Assert.Contains(streamDocParts, part => part.Uri.OriginalString == "/_rels/.rels");
        Assert.Contains(streamDocParts, part => part.Uri.OriginalString == "/word/_rels/document.xml.rels");

        // The document created with FromFlatOpcString(xml) does not contain the relationship parts.
        // This shows that there is some bug in the Flat OPC feature.
        Assert.DoesNotContain(docParts, part => part.Uri.OriginalString == "/_rels/.rels");
        Assert.DoesNotContain(docParts, part => part.Uri.OriginalString == "/word/_rels/document.xml.rels");
    }

    [Fact]
    public void CanNotCloneFlatOpcDocumentCreatedWithStream()
    {
        // Arrange
        string xml = ResourceHelper.GetString(ResourcesAssembly, HelloWorldXml);
        using var stream = new MemoryStream();
        using WordprocessingDocument streamDoc = WordprocessingDocument.FromFlatOpcString(xml, stream, false);

        // Act and Assert
        // Even though the document created with a stream contains the expected relationship parts,
        // it still can't be cloned, demonstrating that, in this case, there is still some bug in
        // the Flat OPC and/or cloning feature.
        Assert.Throws<OpenXmlPackageException>(() => streamDoc.Clone().Dispose());
    }

    [Fact]
    public void CanUseHelperMethod()
    {
        // Arrange
        string xml = ResourceHelper.GetString(ResourcesAssembly, HelloWorldXml);
        using var stream = new MemoryStream();

        // Act
        using WordprocessingDocument doc = FromFlatOpcString(xml, stream, false);

        // Assert
        // Since the following does not throw, we can at least offer a workaround for the time being.
        doc.Clone().Dispose();
    }
}
