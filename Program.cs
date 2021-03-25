using BenchmarkDotNet.Running;
using System.Threading.Tasks;
using System.Diagnostics;
using static System.Console;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
            var memoryUsage = "";
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
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

            //NPOI

            //ExcelDataReader

            var ExcelDR = new ExcelDataReaderAndClosedXMLWriter();
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
