using System.Reflection;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Packaging;
using Resources;
using Xunit.Abstractions;

namespace V301Tests;

public class FlatOpcAndCloningTests
{
    private static readonly Assembly ResourcesAssembly = typeof(ResourceHelper).Assembly;

    private const string HelloWorldDocx = "HelloWorld.docx";
    private const string HelloWorldXml = "HelloWorldFlatOpc.xml";

    private readonly ITestOutputHelper _output;

    public FlatOpcAndCloningTests(ITestOutputHelper output)
    {
        _output = output;
    }

    private void PrintPackageParts(List<IPackagePart> packageParts, string title)
    {
        _output.WriteLine(title);

        foreach (IPackagePart part in packageParts)
        {
            _output.WriteLine($"- Uri: {part.Uri}, ContentType: {part.ContentType}");
        }

        _output.WriteLine("");
    }

    [Fact]
    public void DocumentsDoNotHaveIdenticalParts()
    {
        // Arrange
        using MemoryStream stream = ResourceHelper.GetMemoryStream(ResourcesAssembly, HelloWorldDocx);
        using WordprocessingDocument docxDocument = WordprocessingDocument.Open(stream, false);

        string xml = ResourceHelper.GetString(ResourcesAssembly, HelloWorldXml);
        using WordprocessingDocument flatOpcDocument = WordprocessingDocument.FromFlatOpcString(xml);

        // Act
        List<IPackagePart> docxPackageParts = docxDocument.GetPackage().GetParts().ToList();
        List<IPackagePart> flatOpcPackageParts = flatOpcDocument.GetPackage().GetParts().ToList();

        PrintPackageParts(docxPackageParts, HelloWorldDocx);
        PrintPackageParts(flatOpcPackageParts, HelloWorldXml);

        // Assert
        Assert.False(docxPackageParts.Select(p => p.Uri).SequenceEqual(flatOpcPackageParts.Select(p => p.Uri)));
    }

    [Fact]
    public void CanCloneDocxDocument()
    {
        // Arrange
        using MemoryStream stream = ResourceHelper.GetMemoryStream(ResourcesAssembly, HelloWorldDocx);
        using WordprocessingDocument docxDocument = WordprocessingDocument.Open(stream, false);

        // Act and Assert (no exception thrown)
        using OpenXmlPackage clone = docxDocument.Clone();
    }

    [Fact]
    public void CanNotCloneFlatOpcDocument()
    {
        // Arrange
        string xml = ResourceHelper.GetString(ResourcesAssembly, HelloWorldXml);
        using WordprocessingDocument flatOpcDocument = WordprocessingDocument.FromFlatOpcString(xml);

        // Act and Assert (no exception thrown)
        Assert.Throws<OpenXmlPackageException>(() =>
        {
            using OpenXmlPackage clone = flatOpcDocument.Clone();
        });
    }
}
