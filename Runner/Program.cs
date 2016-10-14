using System;
using System.IO;
using System.Reflection;

namespace Runner
{
    class Program
    {
        static void Main(string[] args)
        {
            const string fileName = "repro.docx";

            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);
            var inputPath = new Uri(Path.Combine(assemblyFolder, "template", fileName)).LocalPath;

            var directoryName = Path.GetDirectoryName(inputPath);
            var outputPath = Path.Combine(directoryName, string.Format("out-{0}.docx", Guid.NewGuid()));

            MergeFieldMigrator.Migrate(inputPath, outputPath);
        }
    }
}
