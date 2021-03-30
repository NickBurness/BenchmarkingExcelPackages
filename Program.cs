using BenchmarkDotNet.Running;
using System.Threading.Tasks;
using System.Diagnostics;
using static System.Console;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
            string memoryUsage = "";
            Console.WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            //set up a stopwatch
            var watch = Stopwatch.StartNew();

            //EPPlus
            var EPPlus = new EPPlus();
            WriteLine("EPPlus Processes Started...");
            
            WriteLine("Read Method Started...");
            await EPPlus.ReadDataAsync();
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            WriteLine("Write Method Started...");
            await EPPlus.WriteDataAsync();
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            WriteLine("Write Method Complete...");

            watch.Stop();

            WriteLine("EPPlus Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds");

            ////NPOI

            ////ExcelDataReader and ClosedXML Writer

            ExcelDR.ReadDataFromFile();
            ExcelDR.WriteDataToFile();


            // ClosedXML Reader only

            var ClosedXML = new ClosedXMLReader();
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
