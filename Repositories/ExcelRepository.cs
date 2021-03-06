using OfficeOpenXml;
using SimpleExelApp.Models;
using System.Drawing;

namespace SimpleExelApp.Repositories
{
    public interface IExcelRepository
    {
        public MemoryStream ExportToExcel();
        public List<User> BatchUserUpload(IFormFile batchUsers, Stream stream, List<User> users);
        public List<User> GetUserList();

    }
    public class ExcelRepository:IExcelRepository
    {
        public List<User> BatchUserUpload(IFormFile batchUsers,Stream stream,List<User> users)
        {
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.First();
                var rowCount = worksheet.Dimension.Rows;

                for (var row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        var name = worksheet.Cells[row, 1].Value?.ToString();
                        var email = worksheet.Cells[row, 2].Value?.ToString();
                        var age = worksheet.Cells[row, 3].Value?.ToString();
                        var phone = worksheet.Cells[row, 4].Value?.ToString();


                        var user = new User()
                        {
                            Email = email,
                            Name = name,
                            Age = age,
                            Phone = phone
                        };

                        users.Add(user);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return users;
        }

        public MemoryStream ExportToExcel()
        {
            // Getting the information from our mimic db
            var users = GetUserList();

            // Start exporting to Excel
            var stream = new MemoryStream();

            using (var xlPackage = new ExcelPackage(stream))
            {
                // Define a worksheet
                var worksheet = xlPackage.Workbook.Worksheets.Add("Users");

                // Styling
                var customStyle = xlPackage.Workbook.Styles.CreateNamedStyle("CustomStyle");
                customStyle.Style.Font.UnderLine = true;
                customStyle.Style.Font.Color.SetColor(Color.Red);

                // First row
                var startRow = 6;
                var row = startRow;

                worksheet.Cells["A1"].Value = "Sample User Export";
                using (var r = worksheet.Cells["A1:D1"])
                {
                    r.Merge = true;
                    r.Style.Font.Color.SetColor(Color.Green);
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(23, 55, 93));
                }

                worksheet.Cells["A4"].Value = "Name";
                worksheet.Cells["B4"].Value = "Email";
                worksheet.Cells["C4"].Value = "Age";
                worksheet.Cells["D4"].Value = "Phone";
                worksheet.Cells["A4:D4"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells["A4:D4"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                row = 5;
                foreach (var user in users)
                {
                    worksheet.Cells[row, 1].Value = user.Name;
                    worksheet.Cells[row, 2].Value = user.Email;
                    worksheet.Cells[row, 3].Value = user.Age;
                    worksheet.Cells[row, 4].Value = user.Phone;


                    row++; // row = row + 1;
                }

                xlPackage.Workbook.Properties.Title = "User list";
                xlPackage.Workbook.Properties.Author = "Mohamad";

                xlPackage.Save();
            }

            stream.Position = 0;
            return stream;
        }

        // Minic the database operations
        public List<User> GetUserList()
        {
            var users = new List<User>()
            {
                new User {
                    Email = "khanbala@email.com",
                    Name = "Khanbala",
                    Age = "21",
                    Phone = "111111"
                },
                new User {
                    Email = "rashidov@email.com",
                    Name = "Rashidov",
                    Age= "22",
                    Phone = "22222"
                },
                new User {
                    Email = "Khan@email.com",
                    Name = "Khan",
                    Age ="23",
                    Phone = "33333"
                }
            };

            return users;
        }
    }
}
