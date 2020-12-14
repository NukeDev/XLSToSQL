using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace XLSToSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            while ( true)
            {
                Console.Write("XLS File Path: ");
                var filePath = Console.ReadLine();
                if (!File.Exists(filePath))
                {
                    Console.WriteLine("File not Exist!");
                    CleanRestart();
                    continue;
                }

                Console.Write("SQL Table Name:");
                var tableName = Console.ReadLine();
                if (string.IsNullOrEmpty(tableName))
                {
                    Console.WriteLine("Table Name is empty!");
                    CleanRestart();
                    continue;
                }

                Console.WriteLine("Loading...");

                try
                {
                    var result = GenerateSqlInsertScripts(tableName, filePath);
                    var pathToSave = @$"{Path.GetDirectoryName(filePath)}\{tableName}-{DateTime.Now.ToFileTime()}.txt";

                    File.WriteAllText(pathToSave, result);

                    Console.WriteLine($"File Saved! - ({pathToSave})");
                    CleanRestart();

                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                    CleanRestart();
                }

            }
        }

        static void CleanRestart()
        {
            Console.WriteLine("Press any key to continue...");
            Console.ReadLine();
            Console.Clear();
        }
        static string GenerateSqlInsertScripts(string tableName, string filePath)
        {

            using (OfficeOpenXml.ExcelPackage xlPackage = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
            {
                var myWorksheet = xlPackage.Workbook.Worksheets.First();
                var totalRows = myWorksheet.Dimension.End.Row;
                var totalColumns = myWorksheet.Dimension.End.Column;

                var columns = new StringBuilder();

                var columnRows = myWorksheet.Cells[1, 1, 1, totalColumns].Select(c => c.Value.ToString()).ToList();

                GenerateInsert(tableName, columns, columnRows);

                for (int rowNum = 2; rowNum <= totalRows; rowNum++)
                {
                    columns.Append("(");

                    var arrayTemp = new List<string>();

                    for (int i=1; i<= totalColumns; i++)
                    {
                        var item = myWorksheet.Cells[rowNum, i];
                        if (item.Value == null)
                            arrayTemp.Add("''");
                        else if (item.Value.ToString() == "NULL")
                            arrayTemp.Add("NULL");
                        else
                            arrayTemp.Add("'" + item.Value.ToString().Replace("'", "''") + "'");
                    }

                    columns.Append(string.Join(",", arrayTemp));

                    if (rowNum % 1000 == 0)
                    {
                        columns.AppendLine(")");

                        GenerateInsert(tableName, columns, columnRows);
                    }
                    else if (rowNum == totalRows)
                        columns.AppendLine(")");
                    else
                        columns.AppendLine("),");
                }

                return columns.ToString();
            }
        }

        private static void GenerateInsert(string tableName, StringBuilder columns, List<string> columnRows)
        {
            columns.Append("INSERT INTO " + tableName + " (");
            columns.Append(string.Join(",", columnRows.Select(r => "[" + r + "]")));
            columns.AppendLine(") VALUES ");
        }

    }
}
