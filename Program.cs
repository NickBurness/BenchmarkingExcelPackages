using BenchmarkDotNet.Running;
using System.Threading.Tasks;
using System.Diagnostics;
using static System.Console;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Validators;
using BenchmarkDotNet.Loggers;
using BenchmarkDotNet.Columns;

namespace BenchmarkingExcelPackages
{
    class Program
    {
        static async Task Main()
        {
#if (Debug)
            string memoryUsage = "";
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            var watch = Stopwatch.StartNew();

            #region EPPlus

            var epplus = new EPPlus();
            WriteLine("EPPlus Processes Starting...");

            WriteLine("Read Method Started...");
            await epplus.ReadDataAsync();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            WriteLine("Write Method Started...");
            await epplus.WriteDataAsync();
            WriteLine("Write Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            watch.Stop();
            WriteLine("EPPlus Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds. \n");
            #endregion

            #region NPOI
            watch.Reset();
            watch.Start();

            var NPOI = new NPOI();
            WriteLine("NPOI Processes Starting...");

            WriteLine("Read Method Started...");
            await NPOI.ImportDataAsync();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            WriteLine("Write Method Started...");
            await NPOI.WriteDataAsync();
            WriteLine("Write Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            watch.Stop();
            WriteLine("NPOI Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds. \n");
            #endregion

            #region ExcelDataReader and ClosedXML Writer
            watch.Reset();
            watch.Start();

            var ExcelDR = new ExcelDataReaderAndClosedXMLWriter();
            WriteLine("ExcelDataReader / ClosedXML Writer Processes Started...");

            WriteLine("Read Method Started...");
            await ExcelDR.ReadExcelDataAsync();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            WriteLine("Write Method Started...");
            await ExcelDR.WriteClosedXMLDataAsync();
            WriteLine("Write Complete Method...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            watch.Stop();
            WriteLine("ExcelDataReader / ClosedXML Writer Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds. \n");
            #endregion

            #region ClosedXML Reader only
            watch.Reset();
            watch.Start();

            var ClosedXML = new ClosedXMLReader();
            WriteLine("ClosedXML Read Data Process Started...");
            ClosedXML.GetDataFromExcel();
            WriteLine("ClosedXML Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage() + "\n");

            watch.Stop();
            WriteLine($"Read Process Only - Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds. \n");
            #endregion
#endif

#if (!Debug)
            var config = new ManualConfig()
                .WithOptions(ConfigOptions.DisableOptimizationsValidator)
                .AddValidator(JitOptimizationsValidator.DontFailOnError)
                .AddLogger(ConsoleLogger.Default)
                .AddColumnProvider(DefaultColumnProviders.Instance);

            BenchmarkRunner.Run(typeof(Program).Assembly, config);
#endif

            return;
        }
    }
}
