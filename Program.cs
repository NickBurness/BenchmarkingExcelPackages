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
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());
            var watch = Stopwatch.StartNew();

            #region EPPlus

            var EPPlus = new EPPlus();
            WriteLine("EPPlus Processes Starting...");

            WriteLine("Read Method Started...");
            await EPPlus.ReadDataAsync();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            WriteLine("Write Method Started...");
            await EPPlus.WriteDataAsync();
            WriteLine("Write Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            watch.Stop();
            WriteLine("EPPlus Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds");
            #endregion

            #region NPOI
            watch.Start();

            var NPOI = new NPOI();
            WriteLine("NPOI Processes Starting...");

            WriteLine("Read Method Started...");
            NPOI.ImportData();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            WriteLine("Write Method Started...");
            NPOI.WriteData();
            WriteLine("Write Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            watch.Stop();
            WriteLine("NPOI Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds");
            #endregion

            #region ExcelDataReader and ClosedXML Writer
            watch.Start();

            var ExcelDR = new ExcelDataReaderAndClosedXMLWriter();
            WriteLine("ExcelDataReader / ClosedXML Writer Processes Started...");

            WriteLine("Read Method Started...");
            ExcelDR.ReadDataFromFile();
            WriteLine("Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            WriteLine("Write Method Started...");
            ExcelDR.WriteDataToFile();
            WriteLine("Write Complete Method...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            watch.Stop();
            WriteLine("ExcelDataReader / ClosedXML Writer Processes Complete...");
            WriteLine($"Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds");
            #endregion

            #region ClosedXML Reader only
            watch.Start();
            var ClosedXML = new ClosedXMLReader();
            WriteLine("ClosedXML Read Data Process Started...");
            ClosedXML.GetDataFromExcel();
            WriteLine("ClosedXML Read Method Complete...");
            WriteLine(memoryUsage.GetLowDetailAboutMemoryUsage());

            watch.Stop();
            WriteLine($"Read Process Only - Execution Time: {watch.ElapsedMilliseconds} milliseconds or around {watch.Elapsed.TotalSeconds} seconds");
            #endregion

            #region BenchmarkDotNet
#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif
            #endregion

            return;
        }
    }
}
