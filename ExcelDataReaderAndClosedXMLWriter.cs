using ExcelDataReader;
using System;
using System.Data;
using System.IO;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;

namespace BenchmarkingExcelPackages
{
    public class ExcelDataReaderAndClosedXMLWriter
    {
        [Benchmark]
        public DataTable ReadDataFromFile()

        {
            string path = "";
            string actualPath = path.SetDirectoryPath();
            string filePath = $@"{actualPath}\ExcelFiles\SampleData.xlsx";

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
                            //    rowReader.Read();
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
                            FilterColumn = (rowReader, columnIndex) =>
                            {
                                return true;
                            }
                        }

                    });

                    DataTableCollection resultFromSpreadsheet = result.Tables;

                    DataTable resultTable = resultFromSpreadsheet[0];
                    return resultTable;

                }
            }

        }


        [Benchmark]
        public void WriteDataToFile()

        {
            string path = "";
            string actualPath = path.SetDirectoryPath();
            string newlyCreatedFilePath = $@"{actualPath}\ExcelFiles\ClosedXMLGeneratedFile.xlsx";

            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Primary", 1);
            IXLWorksheet ws2 = wb.Worksheets.Add("Secondary", 2);

            var dataTable = ReadDataFromFile();


            ws.Range(1, 1, 1, 5).Merge().AddToNamed("Titles");
            ws2.Range(1, 1, 1, 5).Merge().AddToNamed("Workbook");
            var rangeWithData = ws.Cell(2, 1).InsertData(dataTable.AsEnumerable());
            var rangeWithData2 = ws2.Cell(2, 1).InsertData(dataTable.AsEnumerable());

            ws.Column(1).SetDataType(XLDataType.Number);
            ws.Column(2).SetDataType(XLDataType.Text);
            ws.Column(3).SetDataType(XLDataType.Boolean);
            ws.Column(4).SetDataType(XLDataType.Text);
            ws.Column(5).Style.NumberFormat.Format = "mm/dd/yyyy";
            ws2.Column(5).Style.NumberFormat.Format = "mm/dd/yyyy";

            //Adjust column widths to their content
            ws.Columns(1, 5).AdjustToContents();
            ws2.Columns(1, 5).AdjustToContents();

            // Prepare the style 

            var dataStyle = ws.Style;
            dataStyle.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // wingdings column
            var rangeForWingDings = ws.Range(2, 6, 100000, 6).AddToNamed("wingdings");
            rangeForWingDings.Value = char.ConvertFromUtf32(0x00002713);
            rangeForWingDings.Style.Fill.BackgroundColor = XLColor.Red;

            // Merge cells
            var mergedCellrange = ws.Range(2, 10, 2, 12);
            mergedCellrange.Cell(1, 1).Value = "merged";
            mergedCellrange.Merge();

            // outside border
            var outsideBorderRange = ws.Range(2, 10, 2, 12);
            outsideBorderRange.Style.Border.OutsideBorder = XLBorderStyleValues.Double;


            wb.SaveAs(newlyCreatedFilePath);
        }

    }
}



