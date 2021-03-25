using BenchmarkDotNet.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    [MemoryDiagnoser]
    public class EPPlus
    {
        /// <summary>
        /// Sets the directory to the /BenchmarkingExcelPackages instead of /BenchmarkingExcelPackages/bin/Release or /bin/Debug
        /// </summary>
        public string SetDirectoryPath()
        {
            // based on execution context...
            //starting would be either /bin/debug or /bin/release
            string initialDir = Directory.GetCurrentDirectory();
            // so go up one level to /bin
            string parentDir = Directory.GetParent(initialDir).ToString();
            // and up another to /BenchmarkingExcelPackages
            string dir = Directory.GetParent(parentDir).ToString();

            return dir;
        }

        /// <summary>
        /// Implements the ReadData() method as an asynchronous task
        /// </summary>
        [Benchmark]
        public async Task<DataTable> ReadDataAsync()
        {
            var task = Task.Run(() => ReadData());
            var result = await task;
            return result;
        }

        /// <summary>
        /// Implements the WriteData() method as an asynchronous task
        /// </summary>
        [Benchmark]
        public async Task<bool> WriteDataAsync()
        {
            var task = Task.Run(() => WriteData());
            var result = await task;
            return result;
        }

        /// <summary>
        /// Reads cell data from .xlsx file in the specified directory
        /// </summary>
        [Benchmark]
        public DataTable ReadData()
        {
            var targetDir = SetDirectoryPath();

            byte[] file = File.ReadAllBytes($"{targetDir}\\ExcelFiles\\SampleData.xlsx");

            var dataTable = new DataTable("Data");

            using (MemoryStream stream = new MemoryStream(file))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    var worksheet = excelPackage.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null)
                    {
                        return dataTable;
                    }

                    int startRow = worksheet.Dimension.Start.Row;
                    int endRow = worksheet.Dimension.End.Row;

                    int startCol = worksheet.Dimension.Start.Column;
                    int endCol = worksheet.Dimension.End.Column;

                    //create a list to hold the column names
                    List<string> columnNames = new List<string>();

                    int currentColumn = 1;

                    foreach (var cell in worksheet.Cells[1, 1, 1, endCol])
                    {
                        string columnName = cell.Text.Trim();

                        //check if the previous header was empty and add it if it was
                        if (cell.Start.Column != currentColumn)
                        {
                            columnNames.Add("Header_" + currentColumn);
                            dataTable.Columns.Add("Header_" + currentColumn);

                            currentColumn++;
                        }

                        columnNames.Add(columnName);

                        int occurrences = columnNames.Count(x => x.Equals(columnName));

                        if (occurrences > 1)
                        {
                            columnName = columnName + "_" + occurrences;
                        }

                        dataTable.Columns.Add(columnName);
                        currentColumn++;
                    }

                    //start adding the contents of the excel file to the datatable
                    for (int i = 2; i <= endRow; i++)
                    {
                        var row = worksheet.Cells[i, 1, i, endCol];
                        DataRow newRow = dataTable.NewRow();

                        //loop all cells in the row
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }
                        dataTable.Rows.Add(newRow);
                    }
                    return dataTable;
                }
            }
        }

        /// <summary>
        /// Writes data to a new .xlsx file using the collected DataTable. 
        /// </summary>
        [Benchmark]
        public async Task<bool> WriteData()
        {
            var data = await ReadDataAsync();

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                try
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                    ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets.Add("Sheet 2");

                    // add all the data to the excel sheet, starting at cell A1 including headers
                    worksheet.Cells["A1"].LoadFromDataTable(data, true);

                    // 1.4, 1.5 get a range of cells
                    var rangeOfCells = worksheet.Cells[2, 6, worksheet.Dimension.End.Row, 6];

                    // 1.8 style range with color
                    rangeOfCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rangeOfCells.Style.Fill.BackgroundColor.SetColor(Color.Red);

                    // 2 style range using wingdings font
                    rangeOfCells.Style.Font.Name = "WingDings";
                    rangeOfCells.Value = "ü";

                    // 1.9 add value to a specific cell 
                    var specificCell = worksheet.Cells["I1"];
                    specificCell.Value = "Success";

                    // 2.1, 2.4 style text using wrap & alignment
                    specificCell.Style.WrapText = true;
                    specificCell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                    specificCell.Style.TextRotation = 90; //degrees

                    // 2.3 bold font
                    specificCell.Style.Font.Bold = true;

                    // 2.4 merge cells
                    specificCell = worksheet.Cells["I1:K1"];
                    specificCell.Merge = true;

                    // 2.2 style cell border types
                    specificCell.Style.Border.Top.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Right.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Left.Style = ExcelBorderStyle.Double;


                    var targetDir = SetDirectoryPath();

                    FileInfo fileInfo = new FileInfo($"{targetDir}\\ExcelFiles\\EPPlusGeneratedFile.xlsx");
                    excelPackage.SaveAs(fileInfo);

                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw;
                }
            }

        }
    }
}