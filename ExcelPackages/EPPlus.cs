using BenchmarkDotNet.Attributes;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    [MemoryDiagnoser]
    public class EPPlus
    {
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
            string path = "";
            string actualPath = path.SetDirectoryPath();

            byte[] file = File.ReadAllBytes($@"{actualPath}\ExcelFiles\SampleData.xlsx");

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
                            newRow[cell.Start.Column - 1] = cell.Value;
                        }
                        dataTable.Rows.Add(newRow);
                    }

                    // clone it and update this one with column data types defined
                    DataTable finalDataTable = dataTable.Clone();
                    finalDataTable.Columns[0].DataType = typeof(Double);
                    finalDataTable.Columns[1].DataType = typeof(String);
                    finalDataTable.Columns[2].DataType = typeof(Boolean);
                    finalDataTable.Columns[3].DataType = typeof(String);
                    finalDataTable.Columns[4].DataType = typeof(DateTime);

                    foreach (DataRow row in dataTable.Rows)
                    {
                        finalDataTable.ImportRow(row);

                    }

                    return finalDataTable;
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

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    if (worksheet.Cells[2, 3, worksheet.Dimension.End.Row, 3].Value.ToString() == "1")
                    {
                        worksheet.Cells.Value = "true";
                    }

                    // 1.4, 1.5 get a range of cells
                    var rangeOfCells = worksheet.Cells[2, 6, worksheet.Dimension.End.Row, 6];
                    var dateCells = worksheet.Cells[2, 5, worksheet.Dimension.End.Row, 5];

                    dateCells.Style.Numberformat.Format = "mm-dd-yyyy";

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
                    specificCell = worksheet.Cells["G1:I1"];
                    specificCell.Merge = true;

                    // 2.2 style cell border types
                    specificCell.Style.Border.Top.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Right.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                    specificCell.Style.Border.Left.Style = ExcelBorderStyle.Double;


                    string path = "";
                    string actualPath = path.SetDirectoryPath();

                    FileInfo fileInfo = new FileInfo($@"{actualPath}\ExcelFiles\EPPlusGeneratedFile.xlsx");
                    excelPackage.SaveAs(fileInfo);
                    data.Dispose();
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