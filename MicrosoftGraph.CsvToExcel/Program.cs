using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json.Linq;

namespace MicrosoftGraph.CsvToExcel
{
    public class Program
    {
        private const string ExcelFileName = "data.xlsx";
        private const string CsvFileName = "data.csv";
        private const char Delimiter = ',';

        public static void Main()
        {
            ConvertCsvToExcel().Wait();

            Console.WriteLine("CSV uploaded successfully to O365");
            Console.ReadKey();
        }

        private static async Task ConvertCsvToExcel()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            if (graphClient == null)
                return;

            var drive = graphClient.Me.Drive;

            // Create an empty workbook in root of OneDrive, file will be overriden if it already exists
            var emptyFile = await drive.Root.ItemWithPath(ExcelFileName).Content.Request()
                .PutAsync<DriveItem>(System.IO.File.OpenRead("empty.xlsx"));

            // Convert CSV to Json Array of Array
            var lines = System.IO.File.ReadAllLines(CsvFileName);
            var data = JArray.FromObject(lines.Select(x => JArray.FromObject(x.Split(Delimiter))));

            var columns = ((JArray) data[0]).Count;
            var rows = lines.Length;
            var range = $"A1:{(char)(64 + columns)}{rows}";

            // Update excel with range data
            await drive.Items[emptyFile.Id].Workbook.Worksheets["sheet1"].Range(range).Request()
                .PatchAsync(new WorkbookRange {Values = data});
        }
    }
}
