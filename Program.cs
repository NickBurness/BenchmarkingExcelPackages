using BenchmarkDotNet.Running;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
            ////EPPlus
            //var epplus = new EPPlus();
            //await epplus.ReadDataAsync();
            //await epplus.WriteDataAsync();
            //Console.WriteLine("epplus read/write complete...");

            ////npoi
            //var npoi = new npoi();
            //npoi.importdata();
            //npoi.writedata();
            //console.writeline("npoi read/write complete...");

            //ExcelDataReader

            // ClosedXML

            var ClosedXML = new ClosedXML();
            ClosedXML.GetDataFromExcel();
            Console.WriteLine("ClosedXML read data");



            //BenchmarkDotNet
#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif
            return;
        }
    }
}
