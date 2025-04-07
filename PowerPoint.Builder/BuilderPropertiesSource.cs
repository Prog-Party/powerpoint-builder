namespace PowerPoint.Builder;

internal record BuilderPropertiesSource
{
    internal string? Filepath { get; init; }
    internal Stream? Stream { get; init; }
    internal System.IO.Packaging.Package? Package { get; init; }

    internal BuilderPropertiesSource(string? filepath = null, Stream? stream = null, System.IO.Packaging.Package? package = null)
    {
        Filepath = filepath;
        Stream = stream;
        Package = package;

        if (filepath is null && stream is null && package is null)
            throw new ArgumentException("At least one of filePath, stream, or package must be provided.");
    }
}