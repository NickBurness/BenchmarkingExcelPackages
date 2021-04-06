using System;
using System.IO;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using System.Data;
using System.Linq;
using NPOI.SS;
using System.Drawing;

namespace BenchmarkingExcelPackages
{
    public class NPOI
    {

        public DataTable ImportData()
        {
            string path = "";
            string actualPath = path.SetDirectoryPath();
            string newlyCreatedFilePath = $@"{actualPath}\ExcelFiles\NPOIGeneratedFile.xlsx";

            IWorkbook workbook;
            using (var stream = new FileStream(@"C:\Users\aashraf1\source\repos\BenchmarkingExcelPackages\ExcelFiles\SampleData.xlsx", FileMode.Open, FileAccess.Read))
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
        // Write excel
        public void WriteData()
        {
            DataTable table = ImportData();
            Console.WriteLine("Datatable created");
            // start try
            //IWorkbook workbook = new XSSFWorkbook();
            Console.WriteLine("Workbook created");
           // ISheet sheet = workbook.CreateSheet("sheet 1");
            //ISheet sheet2 = workbook.CreateSheet("sheet 2");
            Console.WriteLine("Worksheets created");

            // Create styling 1
            //  var getSheet = workbook.GetSheetAt(0);
            // Get a range of cells
            //var range = "A1:A6";
            //var cellRange = CellRangeAddress.ValueOf(range);
            //var cell = workbook.CreateCellStyle();

            //for (var i = cellRange.FirstRow; i <= cellRange.LastRow; i++)
            //{
            //    var row = sheet.GetRow(i);
            //    for (var j = cellRange.FirstColumn; j <= cellRange.LastColumn; j++)
            //    {
            //        skip cell with column index 5(column F)
            //        if (j == 5) continue;

            //        do your work here
            //        Console.Write("{0}\t", row.GetCell(j));
            //    }

            //    Console.WriteLine();

            //}
            //create styling 2
            //gets first worksheet
            // XSSFWorkbook ws = (XSSFWorkbook)workbook.GetSheet("sheet 1");
            var workbook = new XSSFWorkbook(); 
            ISheet sheet = ((XSSFWorkbook)workbook).CreateSheet("sheetOne");
            ICell cell = sheet.CreateRow(1).CreateCell(3);


            for (int i = 0; i < 5; i++)
            {
                IRow row = sheet.CreateRow(i);
                Console.WriteLine("created row");
                for (int j = 0; j < 4; j++)
                {
                    cell = row.CreateCell(j);
                    Console.WriteLine("cell created");
                    cell.SetCellValue("test");
                    setCellStyle(workbook, cell);
                    Console.WriteLine("style set");
                }
            }

           



            // IRow gRow = sheet.GetRow(0);
            //ICellStyle colorStyle = workbook.CreateCellStyle();
            //colorStyle.FillForegroundColor = IndexedColors.Red.Index;
            //colorStyle.FillPattern = FillPattern.SolidForeground;


            //ICell cell = gRow.CreateCell(5);
            //cell.SetCellValue("test");
            //cell.CellStyle = colorStyle;



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
            
            using (FileStream fs = new FileStream(@"C:\Users\aashraf1\source\repos\BenchmarkingExcelPackages\ExcelFiles\NPOIGeneratedFile.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }

            

        }
        public void setCellStyle(XSSFWorkbook workbook, ICell cell)
        {
            XSSFCellStyle fCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();

            //fCellStyle.FillForegroundColor = XSSFColor.ToXSSFColor(color);
            //fCellStyle.FillPattern = FillPattern.SolidForeground;
            //fCellStyle = (XSSFCellStyle)cell.getCellStyle();
            XSSFColor myColor = new XSSFColor(Color.Red);
            fCellStyle.SetFillBackgroundColor(myColor);
            


            XSSFFont ffont = (XSSFFont)workbook.CreateFont();
            ffont.FontHeight = 20 * 20;
            //ffont.Color = XSSFColor.Red.Index;
            fCellStyle.SetFont(ffont);

            fCellStyle.VerticalAlignment = VerticalAlignment.Center;
            fCellStyle.Alignment = HorizontalAlignment.Center;

            cell.CellStyle = fCellStyle;
        }

    }
    


}










//class XXX
//{
//    public int Cell { get; set; }
//    public string Value { get; set; }
//}

//private static void ImportExcel()
//{

//    var newFile = "newbook2.core.xlsx";
//    var celldata = new List<XXX>{
//                new XXX{ Cell =0,Value="00000"},
//                new XXX{ Cell = 1,Value = "1111111"  }
//            };

//    using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
//    {
//        //excelPath
//        IWorkbook wb = new XSSFWorkbook("newbook.core.xlsx");

//        ISheet sheet1 = wb.GetSheetAt(0);

//        //celldata        
//        foreach (var x in celldata)
//        {
//            IRow row = sheet1.GetRow(x.Cell);
//            row.GetCell(x.Cell).SetCellValue(x.Value);
//        }
//        wb.Write(fs);
//    }
//}

//private static void ExportExcelHSSF()
//{
//    var newFile = @"newbook.core.xls";

//    using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
//    {
//        IWorkbook workbook = new HSSFWorkbook();
//        ISheet sheet1 = workbook.CreateSheet("Sheet1");
//        sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
//        //ICreationHelper cH = wb.GetCreationHelper();
//        var rowIndex = 0;
//        IRow row = sheet1.CreateRow(rowIndex);
//        row.Height = 30 * 80;
//        var cell = row.CreateCell(0);
//        var font = workbook.CreateFont();
//        font.IsBold = true;
//        font.Color = HSSFColor.DarkBlue.Index2;
//        cell.CellStyle.SetFont(font);

//        cell.SetCellValue("A very long piece of text that I want to auto-fit innit, yeah. Although if it gets really, really long it'll probably start messing up more.");
//        sheet1.AutoSizeColumn(0);
//        rowIndex++;

//        // create sheet
//        var sheet2 = workbook.CreateSheet("My Sheet");
//        // create cell styles?
//        var style1 = workbook.CreateCellStyle();
//        style1.FillForegroundColor = HSSFColor.Blue.Index2;
//        style1.FillPattern = FillPattern.SolidForeground;

//        var style2 = workbook.CreateCellStyle();
//        style2.FillForegroundColor = HSSFColor.Yellow.Index2;
//        style2.FillPattern = FillPattern.SolidForeground;

//        // format cells?
//        var cell2 = sheet2.CreateRow(0).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(0);

//        cell2 = sheet2.CreateRow(1).CreateCell(0);
//        cell2.CellStyle = style2;
//        cell2.SetCellValue(1);

//        cell2 = sheet2.CreateRow(2).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(2);

//        cell2 = sheet2.CreateRow(3).CreateCell(0);
//        cell2.CellStyle = style2;
//        cell2.SetCellValue(3);

//        cell2 = sheet2.CreateRow(4).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(4);

//        workbook.Write(fs);
//    }
//    Console.WriteLine("Excel  Done");
//}
//private static void ExportExcel()
//{
//    var newFile = @"newbook.core.xlsx";

//    using (var fs = new FileStream(newFile, FileMode.Create, FileAccess.Write))
//    {
//        IWorkbook workbook = new XSSFWorkbook();
//        ISheet sheet1 = workbook.CreateSheet("Sheet1");
//        sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
//        //ICreationHelper cH = wb.GetCreationHelper();
//        var rowIndex = 0;
//        IRow row = sheet1.CreateRow(rowIndex);
//        row.Height = 30 * 80;
//        var cell = row.CreateCell(0);
//        var font = workbook.CreateFont();
//        font.IsBold = true;
//        font.Color = HSSFColor.DarkBlue.Index2;
//        cell.CellStyle.SetFont(font);

//        cell.SetCellValue("A very long piece of text that I want to auto-fit innit, yeah. Although if it gets really, really long it'll probably start messing up more.");
//        sheet1.AutoSizeColumn(0);
//        rowIndex++;

//        // 新增試算表。
//        var sheet2 = workbook.CreateSheet("My Sheet");
//        // 建立儲存格樣式。
//        var style1 = workbook.CreateCellStyle();
//        style1.FillForegroundColor = HSSFColor.Blue.Index2;
//        style1.FillPattern = FillPattern.SolidForeground;

//        var style2 = workbook.CreateCellStyle();
//        style2.FillForegroundColor = HSSFColor.Yellow.Index2;
//        style2.FillPattern = FillPattern.SolidForeground;

//        // 設定儲存格樣式與資料。
//        var cell2 = sheet2.CreateRow(0).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(0);

//        cell2 = sheet2.CreateRow(1).CreateCell(0);
//        cell2.CellStyle = style2;
//        cell2.SetCellValue(1);

//        cell2 = sheet2.CreateRow(2).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(2);

//        cell2 = sheet2.CreateRow(3).CreateCell(0);
//        cell2.CellStyle = style2;
//        cell2.SetCellValue(3);

//        cell2 = sheet2.CreateRow(4).CreateCell(0);
//        cell2.CellStyle = style1;
//        cell2.SetCellValue(4);

//        workbook.Write(fs);
//    }
//    Console.WriteLine("Excel  Done");

//}
//    }
//}






