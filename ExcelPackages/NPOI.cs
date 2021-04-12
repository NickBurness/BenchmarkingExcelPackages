using BenchmarkDotNet.Attributes;
using System;
using System.IO;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Linq;
using System.Drawing;
using NPOI.SS.Util;
using System.Windows;
using NPOI.HSSF.Util;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    [MemoryDiagnoser]
    public class NPOI
    {
        private IWorkbook workbook;

        public async Task<DataTable> ImportDataAsync()
        {
            var task = Task.Run(() => ImportData());
            var result = await task;

            return result;
        }


        [Benchmark]
        public DataTable ImportData()
        {
            string path = "";
            string actualPath = path.SetDirectoryPath();

            using (var stream = new FileStream($@"{actualPath}\ExcelFiles\SampleData.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(stream);
            }

            var sheet = workbook.GetSheetAt(0);
            var dataTable = new DataTable(sheet.SheetName);
            var headerRow = sheet.GetRow(0);

            foreach (var cell in headerRow)
            {
                dataTable.Columns.Add(cell.ToString());
            }

            for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)
            {
                var sheetRow = sheet.GetRow(i);
                var dataTableRow = dataTable.NewRow();
                dataTableRow.ItemArray = dataTable.Columns.Cast<DataColumn>()
                    .Select(c => sheetRow.GetCell(c.Ordinal, MissingCellPolicy.CREATE_NULL_AS_BLANK)
                    .ToString())
                    .ToArray();
                dataTable.Rows.Add(dataTableRow);
            }

            return dataTable;
        }

        [Benchmark]
        public async Task<bool> WriteDataAsync()
        {
            var task = Task.Run(() => WriteData());
            var result = await task;
            return result;
        }


        [Benchmark]
        public async Task<bool> WriteData()
        {
            DataTable table = await ImportDataAsync();

            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("sheet 1");
                ISheet sheet2 = workbook.CreateSheet("sheet 2");

                List<String> columns = new List<string>();
                IRow sheetRow = sheet.CreateRow(0);
                int columnIndex = 0;

                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);
                    sheetRow.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow row in table.Rows)
                {
                    sheetRow = sheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        sheetRow.CreateCell(cellIndex).SetCellValue(row[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }

                // get a row in this case 1 ... Create a cell at column 6 (no data persists here currently hence create)

                ICellStyle style = workbook.CreateCellStyle();
                IFont font = workbook.CreateFont();
                //Loops styling
                for (var i = 1; i <= sheet.LastRowNum; i++)
                {
                    var cell = sheet.GetRow(i).CreateCell(5);
                    if (cell == null)
                    {
                        continue;
                    }

                    font.IsBold = true;
                    font.FontName = "WingDings";
                    style.SetFont(font);

                    style.FillForegroundColor = IndexedColors.Red.Index;
                    style.FillPattern = FillPattern.SolidForeground;

                    cell.SetCellValue("ü");
                    cell.CellStyle = style;
                }

                //Creates single border around merged cells.
                sheet.GetRow(0);
                CellRangeAddress region = new CellRangeAddress(0, 0, 6, 8);
                sheet.AddMergedRegion(region);

                //Note: The first parameter 1 indicates the thickness of the border  
                RegionUtil.SetBorderBottom(1, region, sheet);//Bottom border  
                RegionUtil.SetBorderLeft(1, region, sheet);//Left border  
                RegionUtil.SetBorderRight(1, region, sheet);//Right border  
                RegionUtil.SetBorderTop(1, region, sheet);//top border


                string path = "";
                string actualPath = path.SetDirectoryPath();
                using (FileStream fileStream = new FileStream($@"{actualPath}\ExcelFiles\NPOIGeneratedFile.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    workbook.Write(fileStream);
                }
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