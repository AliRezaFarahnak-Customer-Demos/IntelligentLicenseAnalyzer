using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Data;
using System.Text;

namespace IntelligentLicenseAnalyzer.Console;

public class Program
{
    const string PerUserData = "!Per_User_Dataset.xlsx";
    const string ConcurrentUserData = "!Concurrent_User_Raw_Data.xlsx";

    private static readonly AssetIntelligence _cleaner;

    static Program()
    {
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false)
            .Build();

        var apiKey = configuration["ApiKey"];
        _cleaner = new AssetIntelligence(apiKey);
    }

    public static async Task Main(string[] args)
    {
        // Register encoding provider for Excel
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        AnsiConsole.Write(
            new FigletText("License Analyzer")
                .LeftJustified()
                .Color(Color.Blue));

        AnsiConsole.WriteLine();

        var choice = AnsiConsole.Prompt(
            new SelectionPrompt<string>()
                .Title("\nWhat would you like to do?")
                .PageSize(5)
                .AddChoices(new[]
                {
                    "Analyze Software by User",
                    "Analyze Concurrent Users",
                    "Exit"
                }));

        await HandleMenuChoice(choice);
    }

    private static async Task HandleMenuChoice(string choice)
    {
        switch (choice)
        {
            case "Analyze Software by User":
                await EntitlementsPerUserAsync();
                break;
            case "Analyze Concurrent Users":
                await AnalyzeConcurrentUsersAsync();
                break;
            default:
                await AnsiConsole.Status()
                    .StartAsync("Exiting", async ctx =>
                    {
                        await Task.Delay(1000);
                    });
                break;
        }
    }

    private static async Task EntitlementsPerUserAsync()
    {
        await AnsiConsole.Status()
            .StartAsync("Reading Excel data...", async ctx =>
            {
                try
                {
                    // Read Excel file
                    using var stream = File.Open(PerUserData, FileMode.Open, FileAccess.Read);
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    var rawData = result.Tables[0];
                    if (rawData == null)
                    {
                        AnsiConsole.MarkupLine("[red]Error: Data sheet not found[/]");
                        return;
                    }

                    // Group original data by date
                    var originalDataByDate = rawData.AsEnumerable()
                        .GroupBy(row => new
                        {
                            Date = DateTime.Parse(row["LastModifiedDate"]?.ToString() ?? DateTime.MinValue.ToString()).Date,
                            Software = row["SoftwareName"]?.ToString() ?? "Unknown",
                            Publisher = row["Publisher"]?.ToString() ?? "Unknown",
                            Edition = row["Edition"]?.ToString() ?? "Unknown",
                            User = row["LastLoggedOnUser"]?.ToString() ?? "Unknown"
                        })
                        .Select(group => new
                        {
                            group.Key,
                            MachineCount = group.Select(row => row["MachineName"]?.ToString())
                                               .Where(m => !string.IsNullOrEmpty(m))
                                               .Distinct()
                                               .Count()
                        })
                        .OrderBy(g => g.Key.Date);

                    var originalTable = new Table()
                        .Title("[blue]Raw Software Installations by Query[/]")
                        .BorderColor(Color.Blue)
                        .AddColumn("EvaluationDate")
                        .AddColumn("SoftwareName")
                        .AddColumn("Publisher")
                        .AddColumn("Edition")
                        .AddColumn("NumberOfConsumedEntitlements")
                        .AddColumn("Username");

                    foreach (var group in originalDataByDate)
                    {
                        originalTable.AddRow(
                            group.Key.Date.ToShortDateString(),
                            group.Key.Software,
                            group.Key.Publisher,
                            group.Key.Edition,
                            group.MachineCount.ToString(),
                            group.Key.User
                        );
                    }
                    AnsiConsole.Write(originalTable);



                    // Process data concurrently
                    var processedRows = await _cleaner.AnalyzeSoftwareByUserAsync(
                        rawData,
                        rawData.Rows.Count,
                        status => ctx.Status(status)
                    );

                    var processedByDateAndSoftware = processedRows
                        .GroupBy(r => new
                        {
                            r.date.Date,
                            r.cleanedSoftware,
                            r.publisher,
                            r.edition,
                            r.username
                        })
                        .OrderBy(g => g.Key.Date);

                    var combinedTable = new Table()
                        .Title("[blue]Software Installations by Query and AI[/]")
                        .BorderColor(Color.Blue)
                        .AddColumn("EvaluationDate")
                        .AddColumn("SoftwareName")
                        .AddColumn("Publisher")
                        .AddColumn("Edition")
                        .AddColumn("NumberOfConsumedEntitlements")
                        .AddColumn("Username");

                    foreach (var group in processedByDateAndSoftware)
                    {
                        // Count distinct machines for this user-software combination on this date
                        var distinctMachines = group
                            .Select(r => r.machineName)
                            .Where(m => !string.IsNullOrEmpty(m))
                            .Distinct()
                            .Count();

                        combinedTable.AddRow(
                            group.Key.Date.ToShortDateString(),
                            group.Key.cleanedSoftware,
                            group.Key.publisher,
                            group.Key.edition,
                            distinctMachines.ToString(),
                            group.Key.username
                        );
                    }

                    AnsiConsole.Write(combinedTable);

                    // Add warning section for multiple entitlements
                    var multipleEntitlements = processedByDateAndSoftware
                        .Where(g => g.Select(r => r.machineName)
                                   .Where(m => !string.IsNullOrEmpty(m))
                                   .Distinct()
                                   .Count() > 1)
                        .ToList();

                    if (multipleEntitlements.Any())
                    {
                        AnsiConsole.WriteLine();
                        AnsiConsole.MarkupLine("[red]Warning: Multiple Entitlements Found:[/]");
                        foreach (var group in multipleEntitlements)
                        {
                            var count = group.Select(r => r.machineName)
                                           .Where(m => !string.IsNullOrEmpty(m))
                                           .Distinct()
                                           .Count();
                            AnsiConsole.MarkupLine($"[red]• {group.Key.cleanedSoftware} ({group.Key.publisher}) - User: {group.Key.username} - Entitlements: {count}[/]");
                        }
                    }

                    // Create CSV
                    var csvPath = "SoftwareAnalysisReport.csv";
                    if (File.Exists(csvPath)) File.Delete(csvPath);

                    using (var writer = new StreamWriter(File.Open(csvPath, FileMode.Create)))
                    {
                        writer.WriteLine("EvaluationDate,SoftwareName,Publisher,Edition,NumberOfConsumedEntitlements,Username");
                        foreach (var group in processedByDateAndSoftware)
                        {
                            var distinctMachines = group
                                .Select(r => r.machineName)
                                .Where(m => !string.IsNullOrEmpty(m))
                                .Distinct()
                                .Count();

                            writer.WriteLine($"{group.Key.Date.ToShortDateString()},{group.Key.cleanedSoftware},{group.Key.publisher},{group.Key.edition},{distinctMachines},{group.Key.username}");
                        }
                    }

                    var currentComplianceTable = new Table()
                        .Title("[blue]Current License Usage Summary[/]")
                        .BorderColor(Color.Blue)
                        .AddColumn("Scan Date")
                        .AddColumn("Software Title")
                        .AddColumn("Publisher")
                        .AddColumn("Edition")
                        .AddColumn("Licenses In Use")
                        .AddColumn("Active Users");

                    var historicalUsageTable = new Table()
                        .Title("[blue]License Usage Trends[/]")
                        .BorderColor(Color.Blue)
                        .AddColumn("Date")
                        .AddColumn("Normalized Software Name")
                        .AddColumn("Publisher")
                        .AddColumn("Edition")
                        .AddColumn("License Count")
                        .AddColumn("Licensed Users");
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error reading Excel file: {ex.Message}[/]");
                    AnsiConsole.WriteException(ex);
                }
            });
    }

    private static async Task AnalyzeConcurrentUsersAsync()
    {
        await AnsiConsole.Status()
            .StartAsync("Analyzing concurrent users...", async ctx =>
            {
                try
                {
                    using var stream = File.Open(ConcurrentUserData, FileMode.Open, FileAccess.Read);
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    var rawData = result.Tables[0];
                    if (rawData == null)
                    {
                        AnsiConsole.MarkupLine("[red]Error: Data sheet not found[/]");
                        return;
                    }

                    var concurrentUsage = _cleaner.AnalyzeConcurrentUsers(rawData, 
                        status => ctx.Status(status));

                    // Introduction table
                    var introTable = new Table()
                        .Title("[blue] Peak Concurrent User Analysis[/]")
                        .BorderColor(Color.LightSlateGrey)
                        .AddColumn("Software Name")
                        .AddColumn("Analysis Details");

                    introTable.AddRow(
                        "Software Applications",
                        "Analysis will track concurrent usage patterns for all monitored software"
                    );
                    introTable.AddRow(
                        "Output Metrics",
                        "- Peak concurrent users per software\n- Daily peak usage trends\n- Timestamp of maximum utilization"
                    );

                    AnsiConsole.Write(introTable);
                    AnsiConsole.WriteLine();

                    // Peak usage table
                    var peakUsage = concurrentUsage
                        .GroupBy(u => u.SoftwareName)
                        .Select(g => new { 
                            Software = g.Key, 
                            PeakUsers = g.Max(x => x.NumberOfConcurrentUsers),
                            PeakTime = g.OrderByDescending(x => x.NumberOfConcurrentUsers)
                                      .First().DateTime
                        });

                    // Daily peaks table
                    var dailyPeaksTable = new Table()
                        .Title("[blue]Daily Peak Concurrent Users[/]")
                        .BorderColor(Color.Blue)
                        .AddColumn("Date")
                        .AddColumn("Software Name")
                        .AddColumn("Peak Concurrent Users");

                    var dailyPeaks = concurrentUsage
                        .GroupBy(u => new { 
                            Date = u.DateTime.Date,
                            u.SoftwareName
                        })
                        .Select(g => new {
                            g.Key.Date,
                            g.Key.SoftwareName,
                            PeakUsers = g.Max(x => x.NumberOfConcurrentUsers)
                        })
                        .OrderBy(x => x.Date)
                        .ThenBy(x => x.SoftwareName);

                    foreach (var peak in dailyPeaks)
                    {
                        dailyPeaksTable.AddRow(
                            peak.Date.ToShortDateString(),
                            peak.SoftwareName,
                            peak.PeakUsers.ToString()
                        );
                    }

                    AnsiConsole.WriteLine();
                    AnsiConsole.Write(dailyPeaksTable);

                    // Export to CSV
                    var csvPath = "ConcurrentUsageReport.csv";
                    await File.WriteAllLinesAsync(csvPath, 
                        new[] { "Date,SoftwareName,PeakConcurrentUsers" }
                        .Concat(dailyPeaks.Select(p => 
                            $"{p.Date:dd/MM/yyyy},{p.SoftwareName},{p.PeakUsers}")));

                    AnsiConsole.MarkupLine($"\n[green]Results exported to {csvPath}[/]");
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error analyzing concurrent users: {ex.Message}[/]");
                    AnsiConsole.WriteException(ex);
                }
            });
    }
}

public class ConcurrentUsageData
{
    public DateTime DateTime { get; set; }
    public string SoftwareName { get; set; } = string.Empty;
    public int NumberOfConcurrentUsers { get; set; }
}