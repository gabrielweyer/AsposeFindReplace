using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace Runner
{
    public class MergeFieldMigrator
    {
        public static void Migrate(string inputPath, string outputPath)
        {
            var document = new Document(inputPath);

            var mergeFieldReplacer = new MergeFieldReplacer();
            var options = new FindReplaceOptions
            {
                ReplacingCallback = mergeFieldReplacer
            };

            document.Range.Replace(new Regex("~[A-Za-z0-9]+~"), string.Empty, options);

            document.Save(outputPath, SaveFormat.Docx);
        }

        private class MergeFieldReplacer : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                var mergeFieldName = args.Match.ToString().Replace("~", "");

                var builder = new DocumentBuilder((Document)args.MatchNode.Document);

                builder.MoveTo(args.MatchNode);
                builder.InsertField(string.Format("MERGEFIELD {0} \\* MERGEFORMAT", mergeFieldName));

                return ReplaceAction.Replace;
            }
        }
    }
}