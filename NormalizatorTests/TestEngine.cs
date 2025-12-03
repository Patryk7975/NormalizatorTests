using ClosedXML.Excel;
using System.Net.Http.Json;

namespace NormalizatorTests
{
    internal class TestEngine
    {
        public static void RunTest(string originalFilePath, string resultFilePath, string apiUrl, decimal probabilityThreshold)
        {
            using var workbook = new XLWorkbook(originalFilePath);
            var sheet = workbook.Worksheet(1);

            AddResultColumns(sheet);

            var indexes = GetColumnIndexes(sheet);
            var rowsNo = sheet.Rows().Count();

            for (int i = 2; i <= rowsNo; i++)
            {
                var result = GetResultForRow(sheet, i, indexes, apiUrl);

                if (result.NormalizationMetadata?.CombinedProbability >= probabilityThreshold && result.Address != null)
                {
                    WriteRowResult(sheet, i, indexes, result.Address);
                }
                
                SetBackroundColor(sheet, i);
            }

            workbook.SaveAs(resultFilePath);
        }

        private static void SetBackroundColor(IXLWorksheet sheet, int rowNo)
        {
            var columnsNo = sheet.Columns().Count();

            for (int i = columnsNo; i > 0; i--)
            {
                var header = sheet.Cell(1, i).Value.ToString();
                if (header.Contains("EXPECTED"))
                {
                    var expectedValue = sheet.Cell(rowNo, i).Value.ToString();
                    if (expectedValue == "null")
                        expectedValue = string.Empty;

                    var resultValue = sheet.Cell(rowNo, i + 1).Value.ToString();
                    if (resultValue == "null")
                        resultValue = string.Empty;

                    var color = XLColor.Red;

                    if (string.Equals(expectedValue, resultValue, StringComparison.CurrentCultureIgnoreCase))
                    {
                        color = XLColor.Green;
                    }

                    sheet.Cell(rowNo, i + 1).Style.Fill.BackgroundColor = color;

                    var isCorrectCurrentValue = sheet.Cell(rowNo, columnsNo).Value.ToString();
                    if (isCorrectCurrentValue == "1" || isCorrectCurrentValue == string.Empty)
                    {
                        sheet.Cell(rowNo, columnsNo).Value = color == XLColor.Red ? 0 : 1;
                    }                   
                }
            }
        }

        private static void WriteRowResult(IXLWorksheet sheet, int rowNo, ColumnIndexes idx, ApiResponseAddressDto result)
        {
            sheet.Cell(rowNo, idx.ResultBuildingNoIndex).Value = result.BuildingNumber;
            sheet.Cell(rowNo, idx.ResultCityIndex).Value = result.City;
            sheet.Cell(rowNo, idx.ResultCommuneIndex).Value = result.Commune;
            sheet.Cell(rowNo, idx.ResultDistrictIndex).Value = result.District;
            sheet.Cell(rowNo, idx.ResultPostalCodeIndex).Value = result.PostalCode;
            sheet.Cell(rowNo, idx.ResultProvinceIndex).Value = result.Province;
            sheet.Cell(rowNo, idx.ResultStreetIndex).Value = result.StreetName;
            sheet.Cell(rowNo, idx.ResultPrefixIndex).Value = result.StreetPrefix;
        }

        private static NormalizationApiResponseDto GetResultForRow(IXLWorksheet sheet, int rowNo, ColumnIndexes idx, string endpoint)
        {
            try
            {
                var streetName = sheet.Cell(rowNo, idx.RequestStreetIndex).Value.ToString();
                var prefix = sheet.Cell(rowNo, idx.RequestPrefixIndex).Value.ToString();
                var buildingNo = sheet.Cell(rowNo, idx.RequestBuildingNoIndex).Value.ToString();
                var city = sheet.Cell(rowNo, idx.RequestCityIndex).Value.ToString();
                var postalCode = sheet.Cell(rowNo, idx.RequestPostalCodeIndex).Value.ToString();

                if (streetName == "null")
                    streetName = string.Empty;
                if (prefix == "null")
                    prefix = string.Empty;
                if (buildingNo == "null")
                    buildingNo = string.Empty;
                if (city == "null")
                    city = string.Empty;
                if (postalCode == "null")
                    postalCode = string.Empty;

                var body = new
                {
                    StreetName = streetName,
                    StreetPrefix = prefix,
                    BuildingNumber = buildingNo,
                    City = city,
                    PostalCode = postalCode
                };

                using (var client = new HttpClient())
                {
                    var response = client.PostAsJsonAsync(endpoint, body).Result;
                    if (!response.IsSuccessStatusCode)
                    {
                        return new NormalizationApiResponseDto();
                    }

                    var responseObject = response.Content.ReadFromJsonAsync<NormalizationApiResponseDto>().Result;
                    if (responseObject is null)
                    {
                        return new NormalizationApiResponseDto();
                    }

                    return responseObject;
                }
            }
            catch
            {
                return new NormalizationApiResponseDto();
            }
        }

        private static void AddResultColumns(IXLWorksheet sheet)
        {
            var columnsNo = sheet.Columns().Count();

            for (int i = columnsNo; i > 0; i--)
            {
                var header = sheet.Cell(1, i).Value.ToString();
                if (header.Contains("EXPECTED"))
                {
                    sheet.Column(i).InsertColumnsAfter(1);
                    sheet.Cell(1, i + 1).Value = header.Replace("EXPECTED", "RESULT");
                }
            }

            columnsNo = sheet.Columns().Count();
            sheet.Column(columnsNo).InsertColumnsAfter(1);
            sheet.Cell(1, columnsNo + 1).Value = "IsCorrect";
            sheet.Column(columnsNo + 1).SetAutoFilter();
            sheet.Column(columnsNo + 1).Width = 15;
        }

        private static ColumnIndexes GetColumnIndexes(IXLWorksheet sheet)
        {
            var result = new ColumnIndexes();

            var columnsNo = sheet.Columns().Count();

            for (int i = columnsNo; i > 0; i--)
            {
                var header = sheet.Cell(1, i).Value.ToString();

                if (header.Contains("REQUEST"))
                {
                    if (header.Contains("streetN"))
                        result.RequestStreetIndex = i;
                    if (header.Contains("streetP"))
                        result.RequestPrefixIndex = i;
                    if (header.Contains("building"))
                        result.RequestBuildingNoIndex = i;
                    if (header.Contains("city"))
                        result.RequestCityIndex = i;
                    if (header.Contains("postal"))
                        result.RequestPostalCodeIndex = i;
                }

                if (header.Contains("RESULT"))
                {
                    if (header.Contains("streetN"))
                        result.ResultStreetIndex = i;
                    if (header.Contains("streetP"))
                        result.ResultPrefixIndex = i;
                    if (header.Contains("building"))
                        result.ResultBuildingNoIndex = i;
                    if (header.Contains("city"))
                        result.ResultCityIndex = i;
                    if (header.Contains("postal"))
                        result.ResultPostalCodeIndex = i;
                    if (header.Contains("commune"))
                        result.ResultCommuneIndex = i;
                    if (header.Contains("district"))
                        result.ResultDistrictIndex = i;
                    if (header.Contains("province"))
                        result.ResultProvinceIndex = i;
                }
            }

            return result;
        }

        private class ColumnIndexes
        {
            public int RequestStreetIndex { get; set; }
            public int RequestPrefixIndex { get; set; }
            public int RequestBuildingNoIndex { get; set; }
            public int RequestCityIndex { get; set; }
            public int RequestPostalCodeIndex { get; set; }
      
            public int ResultStreetIndex { get; set; }
            public int ResultPrefixIndex { get; set; }
            public int ResultBuildingNoIndex { get; set; }
            public int ResultCityIndex { get; set; }
            public int ResultPostalCodeIndex { get; set; }
            public int ResultProvinceIndex { get; set; }
            public int ResultDistrictIndex { get; set; }
            public int ResultCommuneIndex{ get; set; }
        }

        private class NormalizationApiResponseDto
        {
            public ApiResponseMetadataDto? NormalizationMetadata { get; set; }

            public ApiResponseAddressDto? Address { get; set; }
        }

        private class ApiResponseMetadataDto
        {
            public decimal StreetProbability { get; set; }

            public decimal PostalCodeProbability { get; set; }

            public decimal CityProbability { get; set; }

            public decimal CombinedProbability { get; set; }

            public int? NormalizationId { get; set; }
        }

        private class ApiResponseAddressDto
        {
            public bool IsMultiFamily { get; set; } = false;

            public decimal? Longitude { get; set; }

            public decimal? Latitude { get; set; }

            public string? PostOfficeLocation { get; set; }

            public string? Commune { get; set; }

            public string? Province { get; set; }

            public string? District { get; set; }

            public string? StreetPrefix { get; set; }

            public string? StreetName { get; set; }

            public string? BuildingNumber { get; set; }

            public string? City { get; set; }

            public string? PostalCode { get; set; }
        }
    }
}
