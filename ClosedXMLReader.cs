using System;
using System.Data;
using System.IO;
using ClosedXML.Excel;

namespace BenchmarkingExcelPackages

{
    public class ClosedXMLReader
    {
        public DataTable GetDataFromExcel()
        {
            string path = "";
            string actualPath = path.SetDirectoryPath();
            string file = $@"{actualPath}\ExcelFiles\SmallerSampleData.xlsx";

            // Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(file))
            {
                IXLWorksheet workSheet = workBook.Worksheet(1);
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }

                }
                return dt;
            }
        }
    }
}





