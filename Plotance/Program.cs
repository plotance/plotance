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
var inputArgument = new Argument<FileInfo>(
    name: "input",
    description: "Input markdown file"
);
var templateOption = new Option<FileInfo?>(
    name: "--template",
    description: "Template PowerPoint file"
);
var argumentOption = new Option<string[]>(
    aliases: ["--argument", "--arg"],
    description: "Additional arguments in NAME=VALUE format."
        + " Can be specified multiple times"
);

var outputOption = new Option<FileInfo?>(
    aliases: ["--output", "-o"],
    description: "Output PowerPoint file"
);

var verbosityOption = new Option<string>(
    aliases: ["--verbosity", "-v", "-verbose"],
    description: "Set verbosity. Quiet or Diagnostic (default)"
);

verbosityOption.SetDefaultValue("Diagnostic");
verbosityOption.Arity = ArgumentArity.ZeroOrOne;

var quietOption = new Option<bool>(
    aliases: ["-q"],
    description: "Alias for --verbosity Quiet."
);

rootCommand.AddArgument(inputArgument);
rootCommand.AddOption(templateOption);
rootCommand.AddOption(outputOption);
rootCommand.AddOption(argumentOption);
rootCommand.AddOption(verbosityOption);
rootCommand.AddOption(quietOption);

void Main(InvocationContext context)
{
    try
    {
        var input = context.ParseResult.GetValueForArgument(inputArgument);
        var template = context.ParseResult.GetValueForOption(templateOption);
        var output = context.ParseResult.GetValueForOption(outputOption);
        var verbosity = context.ParseResult.GetValueForOption(verbosityOption);
        var quiet = context.ParseResult.GetValueForOption(quietOption);

        if (quiet)
        {
            verbosity = "Quiet";
        }

        quiet = (verbosity ?? "Diagnostic").ToLowerInvariant() switch
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

        var argumentsRaw = context
            .ParseResult
            .GetValueForOption(argumentOption)
            ?? [];
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

        context.ExitCode = 0;
    }
    catch (PlotanceException e)
    {
        Console.Error.WriteLine(e.MessageWithLocation);
        // TODO show inner errors if debug level is high.
        context.ExitCode = 1;
    }
    catch (Exception e)
    {
        Exception? ex = e;

        while (ex != null)
        {
            Console.Error.WriteLine(ex);
            ex = ex.InnerException;
        }

        context.ExitCode = 1;
    }
}

rootCommand.SetHandler(Main);

return rootCommand.Invoke(args);
