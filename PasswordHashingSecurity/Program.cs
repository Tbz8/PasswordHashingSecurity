using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Helpers;
using OfficeOpenXml;

namespace PasswordHashingSecurity
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start");

            var users = GetUsers(new StreamReader("C:\\Users\\Maxime\\Downloads\\Données à importer final version.xlsx"));

            Console.WriteLine("Users count : " + users.Count);

            StringBuilder stringBuilder = CryptPassword(users);

            WriteFile(stringBuilder.ToString(), "C:\\Users\\Maxime\\Downloads\\password.txt");

            Console.WriteLine("End");
        }

        private static StringBuilder CryptPassword(List<UserAccountDto> users)
        {
            StringBuilder builder = new StringBuilder();

            foreach (var userAccountDto in users)
            {
                var hash = Crypto.HashPassword(!string.IsNullOrEmpty(userAccountDto.Password) ? userAccountDto.Password : "UNKNOW");
                builder.AppendLine(hash);
            }

            return builder;
        }

        public static List<UserAccountDto> GetUsers(StreamReader file, int page = 0)
        {
            ExcelPackage package = new ExcelPackage(file.BaseStream);
            ExcelWorksheet worksheet = package.Workbook.Worksheets[page];

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            int columns = worksheet.Dimension.Columns; // 7

            var results = new List<UserAccountDto>();
            // loop through the worksheet rows and columns
            for (int i = 2; i <= rows; i++)
            {
                for (int j = 1; j <= columns; j++)
                {
                    UserAccountDto currentLine = new UserAccountDto();
                    currentLine.Email = worksheet.Cells[i, j++].Value?.ToString();
                    currentLine.FirstName = worksheet.Cells[i, j++].Value?.ToString();
                    currentLine.LastName = worksheet.Cells[i, j++].Value?.ToString();
                    currentLine.Password = worksheet.Cells[i, j++].Value?.ToString();

                    results.Add(currentLine);
                }
            }

            return results;
        }

        public static void WriteFile(string content, string path)
        {
            using (StreamWriter swriter = new StreamWriter(path))
            {
                swriter.Write(content);
            }

            Console.WriteLine("File created");
        }
    }
}
