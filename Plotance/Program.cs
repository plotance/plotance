// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.CommandLine;
using System.CommandLine.Invocation;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Markdig;
using Plotance.Models;
using Plotance.Renderers.PowerPoint;
using Plotance.Renderers.Query;

var rootCommand = new RootCommand(
    "Convert markdown to PowerPoint presentation"
);

var inputArgument = new Argument<FileInfo>("input")
{
    Description = "Input markdown file"
};

var templateOption = new Option<FileInfo?>("--template")
{
    Description = "Template PowerPoint file"
};

var argumentOption = new Option<string[]>("--argument", "--arg")
{
    Description = "Additional arguments in NAME=VALUE format."
    + " Can be specified multiple times"
};

var outputOption = new Option<FileInfo?>("--output", "-o")
{
    Description = "Output PowerPoint file"
};

var verbosityOption = new Option<string>("--verbosity", "-v", "-verbose")
{
    Description = "Set verbosity. Quiet or Diagnostic (default)",
    DefaultValueFactory = _ => "Diagnostic"
};

verbosityOption.Arity = ArgumentArity.ZeroOrOne;

var quietOption = new Option<bool>("-q")
{
    Description = "Alias for --verbosity Quiet."
};

rootCommand.Arguments.Add(inputArgument);
rootCommand.Options.Add(templateOption);
rootCommand.Options.Add(outputOption);
rootCommand.Options.Add(argumentOption);
rootCommand.Options.Add(verbosityOption);
rootCommand.Options.Add(quietOption);

int Main(ParseResult parseResult)
{
    try
    {
        var input = parseResult.GetRequiredValue(inputArgument);
        var template = parseResult.GetValue(templateOption);
        var output = parseResult.GetValue(outputOption);
        var verbosity = parseResult.GetRequiredValue(verbosityOption);
        var quiet = parseResult.GetValue(quietOption);

        if (quiet)
        {
            verbosity = "Quiet";
        }

        quiet = verbosity.ToLowerInvariant() switch
        {
            "" or "diagnostic" or "diag" => false,

            "quiet" or "q"
                or "minimal" or "m"
                or "normal" or "n"
                or "detailed" or "d"
                => true,

            _ => throw new PlotanceException(
                null,
                null,
                "Invalid verbosity."
                    + " \"Quiet\", \"Q\","
                    + " \"Minimal\", \"M\","
                    + " \"Normal\", \"N\","
                    + " \"Detailed\", \"D\","
                    + " \"Diagnostic\", or \"Diag\""
                    + " is expected."
            )
        };

        PowerPointRenderer.Quiet = quiet;
        Spreadsheets.Quiet = quiet;

        var argumentsRaw = parseResult.GetValue(argumentOption) ?? [];
        var arguments = new Dictionary<string, string>();

        foreach (var pair in argumentsRaw)
        {
            var index = pair.IndexOf('=', StringComparison.InvariantCulture);

            if (index > 0 && index < pair.Length - 1)
            {
                var name = pair[..index].Trim();
                var value = pair[(index + 1)..];

                arguments[name] = value;
            }
            else
            {
                throw new ArgumentException(
                    $"Invalid argument format: '{pair}'."
                    + " Expected NAME=VALUE."
                );
            }
        }

        var currentDirectory = Directory.GetCurrentDirectory();
        var queryResult = QueryProcessor.Process(
            currentDirectory,
            Path.GetRelativePath(currentDirectory, input.FullName),
            arguments
        );
        var outputFile = (
            queryResult.Config.Output == null
                ? null
                : Path.GetFullPath(
                    queryResult.Config.Output.Path!,
                    queryResult.Config.Output.Value
                )
        )
            ?? output?.FullName
            ?? Path.ChangeExtension(input.FullName, ".pptx");
        var templateFile = (
            queryResult.Config.Template == null
                ? null
                : Path.GetFullPath(
                    queryResult.Config.Template.Path!,
                    queryResult.Config.Template.Value
                )
        ) ?? template?.FullName;
        using var presentationDocument = templateFile == null
            ? PresentationDocument.Create(
                new MemoryStream(),
                DocumentFormat.OpenXml.PresentationDocumentType.Presentation
            )
            : PresentationDocument.Open(templateFile, false);
        using var result = PowerPointRenderer.Render(
            presentationDocument,
            arguments,
            queryResult.Blocks
        );

        result.Clone(outputFile);

        if (!quiet)
        {
            Console.Error.WriteLine(
                "Successfully converted {0} to {1}.",
                input.Name,
                Path.GetRelativePath(currentDirectory, outputFile)
            );
        }

        return 0;
    }
    catch (PlotanceException e)
    {
        Console.Error.WriteLine(e.MessageWithLocation);
        // TODO show inner errors if debug level is high.
        return 1;
    }
    catch (Exception e)
    {
        Exception? ex = e;

        while (ex != null)
        {
            Console.Error.WriteLine(ex);
            ex = ex.InnerException;
        }

        return 1;
    }
}

rootCommand.SetAction(Main);

return rootCommand.Parse(args).Invoke();
