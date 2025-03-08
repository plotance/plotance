// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Data.Common;
using System.Text;
using DuckDB.NET.Data;
using Markdig;
using Markdig.Syntax;
using Plotance.Models;

namespace Plotance.Renderers.Query;

/// <summary>
/// Processes Markdown documents containing SQL queries and plotter
/// configuration blocks. Parse configurations, load included files, executes
/// the queries, and attaches the configurations and results to the blocks.
/// <para>
/// This class attaches the following data to blocks using SetData:
/// <list type="table">
///   <listheader>
///     <term>Key</term>
///     <description>Description</description>
///   </listheader>
///   <item>
///     <term>path</term>
///     <description>
///     The path to the YAML/Markdown file the block is originated from
///     </description>
///   </item>
///   <item>
///     <term>plotter_config</term>
///     <description>
///     The unmarged plotter_config if the block contains a configuration
///     </description>
///   </item>
///   <item>
///     <term>included_configs</term>
///     <description>
///     The list of included configrations, from deepest to the root
///     </description>
///   </item>
///   <item>
///     <term>query_results</term>
///     <description>The list of query results</description>
///   </item>
/// </list>
/// </para>
/// </summary>
public sealed class QueryProcessor : IDisposable
{
    /// <summary>
    /// The original working directory when the program is invoked. File paths
    /// in error messages are converted to the relative paths from this
    /// directory.
    /// </summary>
    private string _baseDirectory;

    /// <summary>The DuckDB connection used to execute queries.</summary>
    private DuckDBConnection? _connection;

    /// <summary>The accumulated (merged) plotter configuration.</summary>
    private Configuration _accumulatedConfig;

    /// <summary>
    /// The list of Markdown blocks processed by the processor, including
    /// blocks from included files.
    /// </summary>
    private List<Block> _blocks;

    /// <summary>
    /// The dictionary of variables to expand in configurations and queries.
    /// </summary>
    private Dictionary<string, string> _variables;

    /// <summary>
    /// Initializes a new instance of the <see cref="QueryProcessor"/> class.
    /// </summary>
    /// <param name="baseDirectory">
    /// The original working directory when the program is invoked.
    /// </param>
    /// <param name="variables">
    /// Dictionary of variables to expand in configurations, bodies, and
    /// queries.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    private QueryProcessor(
        string baseDirectory,
        IReadOnlyDictionary<string, string> variables
    )
    {
        _baseDirectory = baseDirectory;
        _accumulatedConfig = new Configuration();
        _blocks = new List<Block>();
        _variables = new Dictionary<string, string>(variables);
    }

    /// <summary>
    /// Processes a Markdown file, executing any queries it contains.
    /// </summary>
    /// <param name="baseDirectory">
    /// The original working directory when the program is invoked.
    /// </param>
    /// <param name="path">The path to the Markdown file to process.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in configurations and queries.
    /// </param>
    /// <returns>The result of processing the Markdown file.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    public static QueryProcessorResult Process(
        string baseDirectory,
        string path,
        IReadOnlyDictionary<string, string> variables
    )
    {
        using var processor = new QueryProcessor(baseDirectory, variables);

        return processor.Process(new ValueWithLocation<string>(null, 0, path));
    }

    /// <summary>
    /// Processes an included Markdown file, executing any queries it contains.
    /// </summary>
    /// <param name="path">
    /// The path to the Markdown file to process, with source location where the
    /// file is included from.
    /// </param>
    /// <returns>The result of processing the Markdown file.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    public QueryProcessorResult Process(ValueWithLocation<string> path)
    {
        var source = ReadAllText(path);
        var document = ParseMarkdown(path, source);
        var relativePath = Path.GetRelativePath(_baseDirectory, path.Value);
        var directoryName = Path.GetDirectoryName(Path.GetFullPath(path.Value));
        var currentDirectory = Directory.GetCurrentDirectory();

        try
        {
            if (directoryName != null)
            {
                Directory.SetCurrentDirectory(directoryName);
            }

            foreach (var block in document)
            {
                ProcessBlock(relativePath, block);
            }
        }
        finally
        {
            Directory.SetCurrentDirectory(currentDirectory);
        }

        return new QueryProcessorResult(
            _blocks,
            _accumulatedConfig,
            _variables
        );
    }

    /// <summary>Reads the entire contents of a file.</summary>
    /// <param name="path">
    /// The path to the file to read, with source location where the file is
    /// included from.
    /// </param>
    /// <returns>The contents of the file as a string.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the file cannot be read.
    /// </exception>
    private string ReadAllText(ValueWithLocation<string> path)
    {
        try
        {
            return File.ReadAllText(path.Value, Encoding.UTF8);
        }
        catch (Exception e)
            when (e is IOException or UnauthorizedAccessException)
        {
            throw new PlotanceException(
                path.Path,
                path.Line,
                $"Cannot read file: {path.Value}",
                e
            );
        }
    }

    private MarkdownDocument ParseMarkdown(
        ValueWithLocation<string> path,
        string source
    )
    {
        try
        {
            return Markdown.Parse(
                source,
                new MarkdownPipelineBuilder().Build()
            );
        }
        catch (Exception e)
            when (e is InvalidOperationException or ArgumentException)
        {
            throw new PlotanceException(
                path.Path,
                path.Line,
                $"Cannot parse Markdown file: {path.Value}",
                e
            );
        }
    }

    /// <summary>
    /// Processes a Markdown block, if it is a plotter configuration block.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="block">The Markdown block to process.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    private void ProcessBlock(string path, Block block)
    {
        if (
            Configuration.TryCreate(path, block, _variables)
                is Configuration plotterConfig
        )
        {
            ProcessConfig(path, block, plotterConfig);
        }

        block.SetData("path", path);

        _blocks.Add(block);
    }

    /// <summary>Processes a plotter configuration block.</summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="block">
    /// The Markdown block containing the configuration.
    /// </param>
    /// <param name="config">The parsed plotter configuration.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    private void ProcessConfig(string path, Block block, Configuration config)
    {
        if (config.Include != null)
        {
            ProcessInclusion(path, block, config.Include);
        }

        if (config.Parameters != null)
        {
            ProcessParameters(config);
        }

        if (config.DataSource != null)
        {
            OpenConnection(config);
        }
        else if (config.DbConfig != null)
        {
            UpdateConnection(config);
        }

        if (config.Query != null || config.QueryFile != null)
        {
            ProcessQuery(block, config);
        }

        block.SetData("plotter_config", config);
    }

    /// <summary>
    /// Includes and processes a file, which may be another Markdown file or a
    /// YAML configuration file.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the inclusion directive, for error
    /// reporting.
    /// </param>
    /// <param name="block">
    /// The Markdown block containing the inclusion directive.
    /// </param>
    /// <param name="includingPath">
    /// The path to the file to include, with source location where the file is
    /// included from.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown when the file extension is not recognized, the file cannot be
    /// read, the connection is not open, or cannot execute the query.
    /// </exception>
    private void ProcessInclusion(
        string path,
        Block block,
        ValueWithLocation<string> includingPath
    )
    {
        var extension = Path.GetExtension(includingPath.Value);

        switch (extension)
        {
            case ".md" or ".markdown" or ".txt":
                Process(includingPath);
                break;

            case ".yaml" or ".yml":
                var source = ReadAllText(includingPath);
                var relativePath = Path.GetRelativePath(
                    _baseDirectory,
                    includingPath.Value
                );

                var config = Configuration.Create(
                    relativePath,
                    source,
                    0,
                    _variables
                );

                ProcessConfig(path, block, config);

                var existingIncludedConfigs = block
                    .GetData("included_configs")
                    as IReadOnlyList<Configuration>
                    ?? [];

                block.SetData(
                    "included_configs",
                    new List<Configuration>(
                        [.. existingIncludedConfigs, config]
                    )
                );

                break;

            default:
                throw new PlotanceException(
                    includingPath.Path,
                    includingPath.Line,
                    $"Unknown file extension: {extension}"
                );
        }
    }

    /// <summary>
    /// Processes parameter declarations. Updates the variables dictionary and
    /// database variables.
    /// </summary>
    /// <param name="config">
    /// The plotter configuration containing parameters.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown when a parameter name is missing, the connection is not open,
    /// or cannot execute the query.
    /// </exception>
    private void ProcessParameters(Configuration config)
    {
        var newParameters = new Dictionary<string, string>();

        foreach (var parameter in config.Parameters!.Value)
        {
            if (parameter.Name == null)
            {
                throw new PlotanceException(
                    config.Parameters.Path,
                    config.Parameters.Line,
                    "Parameter name is required."
                );
            }

            if (
                !_variables.ContainsKey(parameter.Name.Value)
                    && parameter.Default != null
            )
            {
                newParameters[parameter.Name.Value] = parameter.Default.Value;
                _variables[parameter.Name.Value] = parameter.Default.Value;
            }
        }

        if (_connection != null)
        {
            try
            {
                using var command = _connection.CreateCommand();

                foreach (var (name, value) in newParameters)
                {
                    var quotedName = QuoteDuckDbIdentifier(name);

                    command.CommandText
                        = $"SET VARIABLE {quotedName} = ?::TEXT";
                    command.Parameters.Clear();
                    command.Parameters.Add(new DuckDBParameter(value));
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception e)
                when (e is IOException or UnauthorizedAccessException)
            {
                throw new PlotanceException(
                    config.Parameters.Path,
                    config.Parameters.Line,
                    "Cannot set options.",
                    e
                );
            }
        }
    }

    /// <summary>
    /// Quotes an identifier for use in DuckDB SQL statements.
    /// </summary>
    /// <param name="name">The identifier to quote.</param>
    /// <returns>The quoted identifier.</returns>
    private string QuoteDuckDbIdentifier(string name)
        => "\"" + name.Replace("\"", "\"\"") + "\"";

    /// <summary>
    /// Quotes a string literal for use in DuckDB SQL statements.
    /// </summary>
    /// <remarks>
    /// We useally use prepared statements but DuckDB doesn't support prepared
    /// statements for SET statements.
    /// </remarks>
    /// <param name="value">The string value to quote.</param>
    /// <returns>The quoted string.</returns>
    private string QuoteDuckDbString(string value)
        => "'" + value.Replace("'", "''") + "'";

    /// <summary>
    /// Opens a DuckDB connection using the data source and configuration from
    /// the plotter config.  If the connection is already open, it will be
    /// closed first.
    /// </summary>
    /// <param name="config">
    /// The plotter configuration containing database settings.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown if the connection is not open or cannot execute the query.
    /// </exception>
    private void OpenConnection(Configuration config)
    {
        var builder = new DuckDBConnectionStringBuilder
        {
            DataSource = config.DataSource?.Value ?? ":memory:"
        };
        var dbConfig = config.DbConfig?.KeyValues
            ?? new Dictionary<string, ValueWithLocation<string>>();

        foreach (var (key, value) in dbConfig)
        {
            builder[key] = value;
        }

        if (_connection != null)
        {
            _connection.Dispose();
            _connection = null;
        }

        try
        {
            _connection = new DuckDBConnection(builder.ToString());
            _connection.Open();

            SetVariables(_connection);
        }
        catch (Exception e) when (e is InvalidOperationException or DbException)
        {
            throw new PlotanceException(
                config.DataSource?.Path ?? config.DbConfig?.Path,
                config.DataSource?.Line ?? config.DbConfig?.Line,
                "Cannot open and initialize connection.",
                e
            );
        }
    }

    /// <summary>Sets variables in a DuckDB connection.</summary>
    /// <param name="connection">
    /// The DuckDB connection to set variables in.
    /// </param>
    /// <exception cref="DbException">
    /// Thrown if the connection is not open or cannot execute the query.
    /// </exception>
    private void SetVariables(DuckDBConnection connection)
    {
        using var command = connection.CreateCommand();

        foreach (var (name, value) in _variables)
        {
            var quotedName = QuoteDuckDbIdentifier(name);

            command.CommandText = $"SET VARIABLE {quotedName} = ?::TEXT";
            command.Parameters.Clear();
            command.Parameters.Add(new DuckDBParameter(value));
            command.ExecuteNonQuery();
        }
    }

    /// <summary>
    /// Updates an existing DuckDB connection with settings from a plotter
    /// configuration.
    /// </summary>
    /// <param name="config">
    /// The plotter configuration containing database settings.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown if the connection is not open or cannot execute the query.
    /// </exception>
    private void UpdateConnection(Configuration config)
    {
        var dbConfig = config.DbConfig;

        if (dbConfig == null)
        {
            return;
        }

        if (_connection == null)
        {
            OpenConnection(config);

            return;
        }

        try
        {
            using var command = _connection.CreateCommand();

            command.CommandText = "SELECT name FROM duckdb_settings()";

            using var reader = command.ExecuteReader();
            var validNames = new HashSet<string>();

            while (reader.Read())
            {
                validNames.Add(reader.GetString(0));
            }

            foreach (var (key, value) in dbConfig)
            {
                if (validNames.Contains(key))
                {
                    // DuckDB doesn't support prepared statements for SET.
                    // So we quote the value as a string.
                    var quotedValue = QuoteDuckDbString(value.Value);

                    command.CommandText = $"SET {key} = {quotedValue}";
                    command.ExecuteNonQuery();
                }
            }
        }
        catch (Exception e)
            when (e is InvalidOperationException or DbException)
        {
            throw new PlotanceException(
                dbConfig.Path,
                dbConfig.Line,
                "Cannot set options.",
                e
            );
        }
    }

    /// <summary>
    /// Processes SQL queries from a plotter configuration and attaches the
    /// results to the Markdown block.
    /// </summary>
    /// <param name="block">The Markdown block containing the query.</param>
    /// <param name="config">
    /// The plotter configuration containing the query.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown if the connection is not open or cannot execute the query.
    /// </exception>
    private void ProcessQuery(Block block, Configuration config)
    {
        if (_connection == null)
        {
            try
            {
                _connection = new DuckDBConnection("Data Source=:memory:");
                _connection.Open();
                SetVariables(_connection);
            }
            catch (Exception e)
                when (e is InvalidOperationException or DbException)
            {
                throw new PlotanceException(
                    config.QueryFile?.Path ?? config.Query?.Path,
                    config.QueryFile?.Line ?? config.Query?.Line,
                    "Cannot open and initialize connection.",
                    e
                );
            }
        }

        var existingQueryResults = block
            .GetData("query_results")
            as IReadOnlyList<QueryResultSet>
            ?? [];
        var results = new List<QueryResultSet>(existingQueryResults);

        if (config.QueryFile != null)
        {
            var query = new ValueWithLocation<string>(
                config.QueryFile.Path,
                config.QueryFile.Line,
                ReadAllText(config.QueryFile)
            );

            results.AddRange(ExecuteQueries(query));
        }

        if (config.Query != null)
        {
            results.AddRange(ExecuteQueries(config.Query));
        }

        block.SetData("query_results", results);
    }

    /// <summary>
    /// Executes SQL queries and returns the results.
    /// </summary>
    /// <param name="query">The SQL queries to execute.</param>
    /// <returns>An enumerable of query result sets.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the connection is not open or cannot execute the query.
    /// </exception>
    private IEnumerable<QueryResultSet> ExecuteQueries(
        ValueWithLocation<string> query
    )
    {
        try
        {
            return ExecuteQueries(query.Value).ToList();
        }
        catch (Exception e)
            when (e is InvalidOperationException or DbException)
        {
            throw new PlotanceException(
                query.Path,
                query.Line,
                $"Cannot execute query: {e.Message}",
                e
            );
        }
    }

    /// <summary>Executes SQL queries and returns the results.</summary>
    /// <param name="query">The SQL queries to execute.</param>
    /// <returns>An enumerable of query result sets.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the connection is not open.
    /// </exception>
    /// <exception cref="DbException">
    /// Thrown if an error occurs while executing the query.
    /// </exception>
    private IEnumerable<QueryResultSet> ExecuteQueries(string query)
    {
        using var command = _connection!.CreateCommand();

        command.CommandText = query;

        using var reader = command.ExecuteReader();

        do
        {
            yield return QueryResultSet.FromDataReader(reader);
        } while (reader.NextResult());
    }

    /// <summary>Releases all resources used by the QueryProcessor.</summary>
    public void Dispose()
    {
        _connection?.Dispose();
    }
}

/// <summary>
/// Represents the result of processing a Markdown document with the
/// QueryProcessor.
/// </summary>
/// <param name="Blocks">
/// The processed Markdown blocks, including blocks from included files.
/// </param>
/// <param name="Config">The accumulated plotter configuration.</param>
/// <param name="Variables">
/// The dictionary of variables at the end of processing.
/// </param>
public record QueryProcessorResult(
    IReadOnlyList<Block> Blocks,
    Configuration Config,
    IReadOnlyDictionary<string, string> Variables
);
