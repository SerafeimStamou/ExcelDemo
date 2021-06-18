using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace DataAccessLibrary
{
    public class DataAccess : IDataAccess
    {
        FileInfo excelFile;

        public DataAccess()
        {
            excelFile = new(@"C:\Projects\ExcelDemo\People.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public List<Person> SetMockData()
        {
            List<Person> people = new()
            {
                new() { Id = 1, FirstName = "Mary", LastName = "Watson" },
                new() { Id = 2, FirstName = "Jane", LastName = "Wilson" },
                new() { Id = 3, FirstName = "Peter", LastName = "Smith" }
            };

            return people;
        }
        public async Task CreateExcelFile(FileInfo file)
        {
            using var package = new ExcelPackage(file);

            var people = SetMockData();

            var worksheet = package.Workbook.Worksheets.Add("MainReport");

            var range = worksheet.Cells["A1"].LoadFromCollection(people, true);
            range.AutoFitColumns();

            worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

            worksheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Row(1).Style.Font.Bold = true;

            await package.SaveAsync();
        }
        public async Task CreateFile()
        {
            await CreateExcelFile(excelFile);
        }

        public async Task<List<Person>> LoadExcelFile(FileInfo file)
        {
            List<Person> people = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var worksheet = package.Workbook.Worksheets[0];

            int row = 2;
            int column = 1;

            while (string.IsNullOrWhiteSpace(worksheet.Cells[row, column].Value?.ToString()) == false)
            {
                Person person = new();
                person.Id = int.Parse(worksheet.Cells[row, column].Value.ToString());
                person.FirstName = worksheet.Cells[row, column + 1].Value.ToString();
                person.LastName = worksheet.Cells[row, column + 2].Value.ToString();

                people.Add(person);

                row += 1;
            }

            return people;
        }
        public async Task<List<Person>> LoadFile()
        {
            var result = await LoadExcelFile(excelFile);

            return result;
        }
    }
}
