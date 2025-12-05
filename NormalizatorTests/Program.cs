using Microsoft.Extensions.Configuration;
using NormalizatorTests;

var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

var originalFilePath = config["OriginalFilePath"];
var resultFilePath = config["ResultFilePath"];
var apiUrl = config["ApiUrl"];
var probabilityThreshold = config.GetValue("ProbabilityThreshold", new decimal(0.8));
var maxParallelRequests = config.GetValue("MaxParallelRequests", 10);

await TestEngine.RunTest(originalFilePath!, resultFilePath!, apiUrl!, probabilityThreshold, maxParallelRequests);