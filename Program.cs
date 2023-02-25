// Import necessary packages
using System.Text.Json;
using Octokit.GraphQL;
using Octokit.GraphQL.Core;
using Octokit.GraphQL.Model;
using Azure.Identity;
using Microsoft.Graph;
using MicrosoftGraphDotNet;

// retrive the config
var config = new Config();
// Define user agent and connection string for GitHub GraphQL API
var userAgent = new ProductHeaderValue("YOUR_PRODUCT_NAME", "1.0.0");
var connection = new Connection(userAgent, config.GitHubClientSecret!);

// Define GraphQL query to fetch repository names and their associated programming languages
var query = new Query()
.Viewer.Repositories(
isFork: false,
affiliations: new Arg<IEnumerable<RepositoryAffiliation?>>(
new RepositoryAffiliation?[] { RepositoryAffiliation.Owner })
).AllPages().Select(repo => new
{
    repo.Name,
    Languages = repo.Languages(null, null, null, null, null).AllPages().Select(language => language.Name).ToList()
}).Compile();

// Execute the GraphQL query and deserialize the result into a list of repositories
var result = await connection.Run(query);
var languages = result.SelectMany(repo => repo.Languages).Distinct().ToList();
var output = JsonSerializer.Deserialize<Repository[]>(JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));

// Define credentials and access scopes for Microsoft Graph API
var tokenCred = new ClientSecretCredential(
config.AzureTenantId!,
config.AzureClientId!,
config.AzureClientSecret!);

var scopes = new[] { "User.Read", "Files.ReadWrite.All" };
var graphClient = new GraphServiceClient(tokenCred);

// Define the file name and create a new Excel file in OneDrive
var driveItem = new DriveItem
{
    Name = config.NameOfNewFile!,
    File = new Microsoft.Graph.File
    {
    }
};
var newFile = await graphClient.Drive.Root.Children
.Request()
.AddAsync(driveItem);

// Define the address of the Excel table and create a new table in the file
var address = "Sheet1!A1:" + (char)('A' + languages.Count) + output?.Count();
var hasHeaders = true;
var table = await graphClient.Drive.Items[newFile.Id].Workbook.Tables
.Add(hasHeaders, address)
.Request()
.PostAsync();

// Define the first row of the Excel table with the column headers
var firstRow = new List<string> { "Repository Name" }.Concat(languages).ToList();

// Convert the repository data into a two-dimensional list
List<List<string>> totalRows = new List<List<string>> { firstRow };
foreach (var value in output!)
{
    var row = new List<string> { value.Name! };
    foreach (var language in languages)
    {
        row.Add(value.Languages!.Contains(language) ? "1" : "0");
    }
    totalRows.Add(row);
}

// Add a new row to the table with the total number of repositories for each language
var languageTotalRow = new List<string>();
// Add "Total" as the first item in the list
languageTotalRow.Add("Total");
// Loop through each programming language in the header row
for (var languageIndex = 1; languageIndex < totalRows[0].Count; languageIndex++)
{
    // Set the total count for this language to 0
    var languageTotal = 0;
    // Loop through each repository in the table
    for (var repoIndex = 1; repoIndex < totalRows.Count; repoIndex++)
    {
        // If the repository uses this language, increment the count
        if (totalRows[repoIndex][languageIndex] == "1")
        {
            languageTotal++;
        }
    }
    // Add the total count for this language to the languageTotalRow list
    languageTotalRow.Add(languageTotal.ToString());
}
// Add the languageTotalRow list to the bottom of the table
totalRows.Add(languageTotalRow);

// Create a new WorkbookTableRow object with the totalRows list serialized as a JSON document
var workbookTableRow = new WorkbookTableRow
{
    Values = JsonSerializer.SerializeToDocument(totalRows),
    Index = 0,
};
// Add the new row to the workbook table
await graphClient.Drive.Items[newFile.Id].Workbook.
Tables[table.Id].Rows
.Request()
.AddAsync(workbookTableRow);

// Add a new chart to the worksheet with the language totals as data
await graphClient.Drive.Items[newFile.Id].Workbook.Worksheets["Sheet1"].Charts
.Add("ColumnClustered", "Auto", JsonSerializer.SerializeToDocument($"Sheet1!B2:{(char)('A' + languages.Count)}2, Sheet1!B{output.Count() + 3}:{(char)('A' + languages.Count)}{output.Count() + 3}"))
.Request()
.PostAsync();

// Print the URL of the new file to the console
Console.WriteLine(newFile.WebUrl);

// Define a class to hold repository data
class Repository
{
    public string? Name { get; set; }
    public List<string>? Languages { get; set; }
}




