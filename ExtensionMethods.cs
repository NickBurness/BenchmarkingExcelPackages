using System;
using System.IO;

namespace BenchmarkingExcelPackages
{
    public static class ExtensionMethods
    {
        public static string SetDirectoryPath(this string path)
        {
            // based on execution context...
            //starting would be either /bin/debug or /bin/release
            string initialDir = Directory.GetCurrentDirectory();
            // so go up one level to /bin
            string parentDir = Directory.GetParent(initialDir).ToString();
            // and up another to /BenchmarkingExcelPackages
            string dir = Directory.GetParent(parentDir).ToString();

            return dir;
        }

        public static string GetLowDetailAboutMemoryUsage(this string method)
        {
            var megabytesOfMemory = (GC.GetTotalMemory(false) / 1000000);
            var stringified = megabytesOfMemory.ToString();
            var result = $"{stringified} megabytes of data thought to be currently allocated to memory";
            return result;
        }
    }
}
