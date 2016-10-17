using System;
using System.IO;
using System.Reflection;

namespace Runner
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileName = "revert-and-two-matches-same-line.docx";

            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
            var inputPath = new Uri(Path.Combine(assemblyFolder, "template", fileName)).LocalPath;

            var directoryName = Path.GetDirectoryName(inputPath);
            var outputPath = Path.Combine(directoryName, string.Format("out-{0}.docx", Guid.NewGuid()));

            MergeFieldMigrator.Migrate(inputPath, outputPath);
        }
    }
}
