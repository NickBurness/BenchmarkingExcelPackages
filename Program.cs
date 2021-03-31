﻿using BenchmarkDotNet.Running;
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
            WriteLine("EPPlus Processes Started...");
            
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
            ////npoi
            //var npoi = new npoi();
            //npoi.importdata();
            //npoi.writedata();
            //console.writeline("npoi read/write complete...");

            #region ClosedXML Reader only
            var ClosedXML = new ClosedXMLReader();
            WriteLine("ClosedXML Read Data Process Started...");
            ClosedXML.GetDataFromExcel();
            WriteLine("ClosedXML Read Method Complete...");


            //BenchmarkDotNet
#if (!Debug)
            var summary = BenchmarkRunner.Run(typeof(Program).Assembly);
#endif
            return;
        }
    }
}
