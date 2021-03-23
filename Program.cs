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
            //EPPlus
            var EPPlus = new EPPlus();
            await EPPlus.ReadDataAsync();
            await EPPlus.WriteDataAsync();
            Console.WriteLine("EPPlus Read/Write complete...");

            //NPOI

            //ExcelDataReader


            //BenchmarkDotNet
#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif
            return;
        }
    }
}
