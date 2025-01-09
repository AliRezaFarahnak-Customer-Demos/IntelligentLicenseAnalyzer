using Azure;
using Azure.AI.OpenAI;
using Newtonsoft.Json;
using OpenAI.Chat;
using System.Collections.Concurrent;
using System.Data;

namespace IntelligentLicenseAnalyzer.Console;
public class AssetIntelligence
{
    private ChatClient _aiClient;

    public AssetIntelligence(string apiKey)
    {
        _aiClient = new AzureOpenAIClient(new Uri("https://intelligentlicenseanalyzer-ai.openai.azure.com"), new AzureKeyCredential(apiKey)).GetChatClient("gpt-4o");
    }

    public async Task<List<(DateTime date, string originalSoftware, string cleanedSoftware, string publisher, string edition, string machineName, string username)>> AnalyzeSoftwareByUserAsync(DataTable rawData, int totalRows, Action<string> ctx)
    {
        var processedRows = new ConcurrentQueue<(DateTime date, string originalSoftware, string cleanedSoftware, string publisher, string edition, string machineName, string username)>();
        var semaphore = new SemaphoreSlim(50);
        var tasks = new List<Task>();
        var errors = new ConcurrentQueue<Exception>();

        for (int i = 0; i < rawData.Rows.Count; i++)
        {
            var row = rawData.Rows[i];
            var index = i;

            tasks.Add(Task.Run(async () =>
            {
                try
                {
                    await semaphore.WaitAsync();
                    var progress = (index + 1.0) / totalRows * 100;
                    var rawSoftwareName = row["SoftwareName"]?.ToString() ?? "Unknown";

                    var cleanedSoftwareJson = await CleanSoftwarePerLicenseAsync(rawSoftwareName);
                    ctx($"[yellow]AI Processing and Data Cleansing: ({progress:F1}%)[/] {index + 1}/{totalRows}: [blue]{rawSoftwareName}[/] => [green]{cleanedSoftwareJson}[/]");

                    DateTime.TryParse(row["LastModifiedDate"]?.ToString(), out var date);

                    processedRows.Enqueue((
                        date,
                        rawSoftwareName,
                        cleanedSoftwareJson,
                        row["Publisher"]?.ToString() ?? "Unknown",
                        row["Edition"]?.ToString() ?? "Unknown",
                        row["MachineName"]?.ToString() ?? "Unknown",
                        row["LastLoggedOnUser"]?.ToString() ?? "Unknown"
                    ));
                }
                catch (Exception ex)
                {
                    errors.Enqueue(ex);
                }
                finally
                {
                    semaphore.Release();
                }
            }));

            if (tasks.Count >= 100) // Process in batches
            {
                await Task.WhenAll(tasks);
                tasks.Clear();
            }
        }

        await Task.WhenAll(tasks); // Process remaining

        if (errors.Any())
        {
            ctx($"[red]Encountered {errors.Count} errors during processing[/]");
        }

        return processedRows.OrderBy(r => r.date).ToList();
    }

    public async Task<string> CleanSoftwarePerLicenseAsync(string softwareName)
    {
        var response = await _aiClient.CompleteChatAsync(
            [new SystemChatMessage(@$"
Extract the software string and only include the version if it requires a license; otherwise, it should just be the software string. 
Don't add an explanation, as this string will be used to make a data query later. 
For example, Visual Studio 2019 becomes Visual Studio 2019 because a license is required, but Docker Desktop 4.3 becomes Docker Desktop, as there is no special license required for each version. 
Don't mention if it's Community, Professional, or Enterprise if the software string doesn't present it. Visual Studio just needs Visual Studio and the year, not the edition.
"
            ),
            new UserChatMessage($"Classify this software string: {softwareName}")
            ],
            new ChatCompletionOptions
            {
                Temperature = 0.2f,
                TopP = 0.2f,
                MaxOutputTokenCount = 500,
                ResponseFormat = ChatResponseFormat.CreateJsonSchemaFormat(
                    "software_schema",
                    BinaryData.FromString("""
                    {
                      "type": "object",
                      "properties": {
                        "softwarename": {
                          "type": "string"
                        }
                      },
                      "required": [
                        "softwarename"
                      ],
                      "additionalProperties": false
                    }
                    """),
                    "Extracted software name",
                    jsonSchemaIsStrict: true)
            });

        var resultJson = response.Value.Content.First().Text;
        var software = JsonConvert.DeserializeObject<dynamic>(resultJson).softwarename;
        return software;
    }

    public List<ConcurrentUserData> AnalyzeConcurrentUsers(DataTable rawData, Action<string> progressCallback)
    {
        var results = new List<ConcurrentUserData>();
        
        var sessions = rawData.AsEnumerable()
            .Where(row => !string.IsNullOrEmpty(row["SoftwareName"]?.ToString()))
            .Select(row => new
            {
                SoftwareName = row["SoftwareName"]?.ToString() ?? "Unknown",
                LoginTime = ParseDateTime(row["LOGIN_DATE_TIME"]?.ToString()),
                LogoutTime = ParseDateTime(row["LOGOUT_DATE_TIME"]?.ToString()),
                SessionId = row["SESSION_ID"]?.ToString()
            })
            .Where(s => s.LoginTime != DateTime.MinValue)
            .ToList();

        var timePoints = sessions
            .SelectMany(s => new[] { s.LoginTime, s.LogoutTime })
            .Where(t => t != DateTime.MinValue)
            .Distinct()
            .OrderBy(t => t)
            .ToList();

        var totalPoints = timePoints.Count;
        for (int i = 0; i < timePoints.Count; i++)
        {
            var timePoint = timePoints[i];
            var progress = ((i + 1.0) / totalPoints * 100);
            progressCallback($"Processing concurrent users: [yellow]{progress:F1}%[/] ({i + 1}/{totalPoints})");

            var concurrentUsers = sessions
                .Where(s => 
                    s.LoginTime <= timePoint && 
                    (s.LogoutTime >= timePoint || s.LogoutTime == DateTime.MinValue))
                .GroupBy(s => s.SoftwareName)
                .Select(g => new ConcurrentUserData
                {
                    DateTime = timePoint,
                    SoftwareName = g.Key,
                    NumberOfConcurrentUsers = g.Count()
                });

            results.AddRange(concurrentUsers);
        }

        progressCallback("Analysis complete!");
        return results.OrderBy(r => r.DateTime).ThenBy(r => r.SoftwareName).ToList();
    }

    private DateTime ParseDateTime(string? value)
    {
        if (string.IsNullOrEmpty(value) || value == "NULL")
            return DateTime.MinValue;
        return DateTime.TryParse(value, out DateTime result) ? result : DateTime.MinValue;
    }
}

public class ConcurrentUserData
{
    public DateTime DateTime { get; set; }
    public string SoftwareName { get; set; } = string.Empty;
    public int NumberOfConcurrentUsers { get; set; }
}