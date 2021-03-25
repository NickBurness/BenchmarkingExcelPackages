using BenchmarkDotNet.Running;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
            //EPPlus
            var EPPlus = new EPPlus();
            await EPPlus.ReadDataAsync();
            await EPPlus.WriteDataAsync();
            Console.WriteLine("EPPlus Read/Write complete...");

            //NPOI

            //ExcelDataReader

            var ExcelDR = new ExcelDataReader();
            ExcelDR.ReadDataFromFile();
            Console.WriteLine("ExcelDataRead read data");
            ExcelDR.WriteDataToFile();
            Console.WriteLine("ClosedXML written data");

            //BenchmarkDotNet

#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif

            return;
        }
    }
}
