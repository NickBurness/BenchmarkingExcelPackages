using System;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using System.Collections.Generic;


namespace BenchmarkingExcelPackages

//    using (var wb = new XLWorkbook(fileName, XLEventTracking.Disabled))
//{
//    var ws = wb.Worksheet(1);
//DataTable dataTable = ws.RangeUsed().AsTable().AsNativeDataTable();
///* Process data table as you wish */
//}



{
    public class ClosedXML
    {
        public DataTable GetDataFromExcel()
            {
            //Save the uploaded Excel file.
            var filePath = @"C:\Users\aashraf1\source\repos\\BenchmarkingExcelPackages\ExcelFiles\samp.xlsx";

            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(filePath))
                {
                    //Read the first Sheet from Excel file.
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    //Create a new DataTable.
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
                            foreach (IXLCell cell in row.Cells())
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

        //        public DataTable GetDataFromExcel2()
        //        {
        //            Console.WriteLine("start method");
        //            var filePath = @"C:\Users\aashraf1\source\repos\\BenchmarkingExcelPackages\ExcelFiles\SampleData.xlsx";

        //            Open the Excel file using ClosedXML.
        //            Keep in mind the Excel file cannot be open when trying to read it
        //            using (XLWorkbook workBook = new XLWorkbook(filePath))
        //            {
        //                Read the first Sheet from Excel file.
        //                IXLWorksheet workSheet = workBook.Worksheet(1);
        //                Console.WriteLine("worksheet 1");
        //                Create a new DataTable.
        //                DataTable dt = new DataTable();
        //                Console.WriteLine("created dt");
        //                Loop through the Worksheet rows.
        //                bool firstRow = true;
        //                foreach (IXLRow row in workSheet.Rows())
        //                {
        //                    Use the first row to add columns to DataTable.
        //                    if (firstRow)
        //                    {
        //                        foreach (IXLCell cell in row.Cells())
        //                        {
        //                            dt.Columns.Add(cell.Value.ToString());
        //                        }
        //                        firstRow = false;
        //                    }
        //                    else
        //                    {
        //                        Add rows to DataTable.
        //                        dt.Rows.Add();
        //                        int i = 0;
        //                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
        //                        {
        //                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
        //                            i++;
        //                        }
        //                    }
        //                    Console.WriteLine("data added");

        //                }
        //                return dt;
        //            }
        //        }

        //    }
        //}

    }

