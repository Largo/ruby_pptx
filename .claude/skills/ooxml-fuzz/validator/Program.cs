// Validates .pptx files with the Open XML SDK's OpenXmlValidator, the same
// schema + semantic checks the Open XML SDK Productivity Tool runs.
//
//   dotnet run -- [--version Office2007|...|Microsoft365] [--max N] file.pptx...
//   dotnet run -- --stdin        (one path per line in, one JSON line out)
//
// Prints one JSON object per file on stdout (JSON Lines):
//   {"file": "...", "ok": false, "open_error": null,
//    "errors": [{"id","type","description","part","path","node"}]}
// Exit status: 0 all valid, 1 at least one file invalid, 2 usage error.

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

var version = FileFormatVersions.Microsoft365;
var maxErrors = 1000;
var files = new List<string>();
var serve = false;

for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--version":
            if (!Enum.TryParse(args[++i], out version))
            {
                Console.Error.WriteLine($"unknown --version {args[i]}; one of {string.Join(", ", Enum.GetNames<FileFormatVersions>())}");
                return 2;
            }
            break;
        case "--max":
            maxErrors = int.Parse(args[++i]);
            break;
        case "--stdin":
            // Persistent mode for the minimizer: read one path per line,
            // answer with one JSON line each, so .NET starts only once.
            serve = true;
            break;
        case "-h" or "--help":
            Console.Error.WriteLine("usage: OoxmlValidate [--version Microsoft365] [--max 1000] file.pptx...");
            return 2;
        default:
            if (Directory.Exists(args[i]))
                files.AddRange(Directory.GetFiles(args[i], "*.pptx", SearchOption.AllDirectories).Order());
            else
                files.Add(args[i]);
            break;
    }
}

if (files.Count == 0 && !serve)
{
    Console.Error.WriteLine("no files given");
    return 2;
}

var validator = new OpenXmlValidator(version) { MaxNumberOfErrors = maxErrors };
var json = new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower };
var anyInvalid = false;

IEnumerable<string> Inputs()
{
    foreach (var f in files) yield return f;
    if (!serve) yield break;
    string? line;
    while ((line = Console.In.ReadLine()) is not null)
        if (line.Length > 0) yield return line;
}

foreach (var file in Inputs())
{
    var result = new FileResult { File = file };
    try
    {
        using var doc = PresentationDocument.Open(file, false);
        foreach (var e in validator.Validate(doc))
        {
            result.Errors.Add(new Error
            {
                Id = e.Id,
                Type = e.ErrorType.ToString(),
                Description = e.Description,
                Part = e.Part?.Uri.ToString(),
                Path = e.Path?.XPath,
                Node = e.Node?.LocalName,
            });
        }
    }
    catch (Exception ex)
    {
        // The package could not even be opened: broken zip, missing
        // relationship target, bad content type. Worse than a schema error.
        result.OpenError = $"{ex.GetType().Name}: {ex.Message}";
    }

    result.Ok = result.OpenError is null && result.Errors.Count == 0;
    anyInvalid |= !result.Ok;
    Console.WriteLine(JsonSerializer.Serialize(result, json));
    Console.Out.Flush();
}

return anyInvalid ? 1 : 0;

class FileResult
{
    public string File { get; set; } = "";
    public bool Ok { get; set; }
    public string? OpenError { get; set; }
    public List<Error> Errors { get; set; } = new();
}

class Error
{
    public string Id { get; set; } = "";
    public string Type { get; set; } = "";
    public string Description { get; set; } = "";
    public string? Part { get; set; }
    public string? Path { get; set; }
    public string? Node { get; set; }
}
