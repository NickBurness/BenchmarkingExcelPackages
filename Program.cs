using BenchmarkDotNet.Running;
using System;
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
