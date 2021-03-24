using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace BenchmarkingExcelPackages
{
    public class ExcelDataReader
    {
        //[Benchmark]
        public DataTableCollection ReadDataFromFile()

        {
            var filePath = @"C:\Users\FKANE\source\repos\ExcelDataReader\SampleDataFK.xlsx";
            //var dataTableCollection = new DataTableCollection;

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = true,

                        // Gets or sets a callback to determine whether to include the current sheet
                        // in the DataSet. Called once per sheet before ConfigureDataTable.
                        FilterSheet = (tableReader, sheetIndex) => true,

                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            // Gets or sets a value indicating the prefix of generated column names.
                            EmptyColumnNamePrefix = "Column",

                            // Gets or sets a value indicating whether to use a row from the 
                            // data as column names.
                            UseHeaderRow = true,

                            // Gets or sets a callback to determine which row is the header row. 
                            // Only called when UseHeaderRow = true.
                            //ReadHeaderRow = (rowReader) =>
                            //{
                            //    // F.ex skip the first row and use the 2nd row as column headers:
                            //    rowReader.read();
                            //},

                            // Gets or sets a callback to determine whether to include the 
                            // current row in the DataTable.
                            FilterRow = (rowReader) =>
                            {
                                return true;
                            },

                            // Gets or sets a callback to determine whether to include the specific
                            // column in the DataTable. Called once per column after reading the 
                            // headers.
                            //FilterColumn = (rowReader, columnIndex) =>
                            //{
                            //    return true;
                            //}
                        }
                    });
                    Console.WriteLine("table read and configured");
                    DataTableCollection resultFromSpreadsheet = result.Tables;
                    return resultFromSpreadsheet;
                }
            }
        }

        //public ActionResult WriteDataToExcel()
        //{
        //    DataTable dt = getData();
        //    //Name of File  
        //    string fileName = "Sample.xlsx";
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        //Add DataTable in worksheet  
        //        wb.Worksheets.Add(dt);
        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            wb.SaveAs(stream);
        //            //Return xlsx Excel File  
        //            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        //        }
        //    }
        public void WriteDataToFile()

        {

            DataTable dt = ReadDataFromFile();
            String newlyCreatedFilePath = @"C:\Users\FKANE\source\repos\ExcelPackages\BenchmarkingExcelPackages\ExcelFiles\WriteToFile.xlsx";

            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Primary", 1);
            IXLWorksheet ws2 = wb.Worksheets.Add("Secondary", 2);

            //foreach (var ws in Enumerable.Range(1, 2))
            //{

            //}

            ws.Cell(2, 3).Value = "Hello data";

            ws.Column(1).SetDataType(XLDataType.Number);
            ws.Column(2).SetDataType(XLDataType.Text);
            ws.Column(3).SetDataType(XLDataType.Boolean);
            ws.Column(4).SetDataType(XLDataType.Text);
            ws.Column(5).SetDataType(XLDataType.TimeSpan);


            ws2.Column(1).SetDataType(XLDataType.Number);
            ws2.Column(2).SetDataType(XLDataType.Text);
            ws2.Column(3).SetDataType(XLDataType.Boolean);
            ws2.Column(4).SetDataType(XLDataType.Text);
            ws2.Column(5).SetDataType(XLDataType.TimeSpan);


            IXLRange range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(100001, 5).Address);

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

            //Adjust column widths to their content

            ws.Columns(1, 5).AdjustToContents();

            wb.SaveAs(newlyCreatedFilePath);
        }
    }
}



