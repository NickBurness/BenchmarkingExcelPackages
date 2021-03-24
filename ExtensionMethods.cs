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
    }
}
