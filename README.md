# Plotance

A cross-platform command-line tool converting **Markdown** + **SQL** into plain editable **PowerPoint** presentations with charts, powered by **DuckDB**.

## Example

`````markdown
## Page views of ${year}-${month}

<?plotance rows: 1:48pt ?>

```plotance
format: line
x_axis_label_format: d
x_axis_major_unit: 7 days
legend_position: none
query: |
  SELECT
    date_trunc('day', timestamp) AS date,
    COUNT(*) AS count
  FROM
    read_csv(
      'access_'
      || getvariable('year')
      || '-'
      || getvariable('month')
      || '.csv.gz'
    )
  GROUP BY ALL
  ORDER BY ALL;
```

We are growing!
`````

↓

```bash
plotance --arg year=2025 --arg month=01 --template theme.pptx pageviews.md
```

↓

[![pageviews.pptx](https://plotance.github.io/examples/pageviews/pageviews.svg)](https://plotance.github.io/examples/pageviews/pageviews.pptx)

See [other examples](https://plotance.github.io/examples.html) and [user guide](https://plotance.github.io/documentation.html) for usage.


## Building from source

- Run `dotnet build` to build the project.
- Run `dotnet publish` to create a single binary executable.
- Run `create_release_archives.ps1` in PowerShell to create release ZIP archives for all platforms.


## License

This project is released under the MIT License.

2025 Plotance contributors <https://plotance.github.io/>

See [MIT.txt](LICENSES/MIT.txt) for more details.
