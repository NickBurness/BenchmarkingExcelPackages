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
            var epplus = new EPPlus();
            await epplus.ReadDataAsync();
            await epplus.WriteDataAsync();
            Console.WriteLine("epplus read/write complete...");

            //NPOI
            var npoi = new NPOI();
            npoi.ImportData();
           // await npoi.writedataasync();
            Console.WriteLine("epplus read/write complete...");

            //ExcelDataReader


            //BenchmarkDotNet
#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif
            return;
        }
    }
}
