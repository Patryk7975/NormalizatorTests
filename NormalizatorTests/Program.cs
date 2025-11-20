using Microsoft.Extensions.Configuration;
using NormalizatorTests;

var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

var originalFilePath = config["OriginalFilePath"];
var resultFilePath = config["ResultFilePath"];
var apiUrl = config["ApiUrl"];

TestEngine.RunTest(originalFilePath!, resultFilePath!, apiUrl!);