using BenchmarkDotNet.Running;
using System;
using System.Threading.Tasks;
using System.Diagnostics;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
            string memoryUsage = "";
            Console.WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            // set up a stopwatch
            var watch = Stopwatch.StartNew();

            //EPPlus
            var EPPlus = new EPPlus();


            watch.Start();
            Console.WriteLine("EPPlus Processes Started...");
            Console.WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            Console.WriteLine("Read Method Started...");
            await EPPlus.ReadDataAsync();
            Console.WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            Console.WriteLine("Write Method Started...");
            await EPPlus.WriteDataAsync();
            Console.WriteLine("Write Method Complete...");
            watch.Stop();
            Console.WriteLine("EPPlus Read/Write Complete...");
            Console.WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.ElapsedMilliseconds / 1000} seconds");


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
