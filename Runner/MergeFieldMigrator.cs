using System.Collections;
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
            public ReplaceAction Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match,
                // in this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further removing.
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength = remainingLength - currentNode.GetText().Length;

                    // Select the next Run node.
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }


                var mergeFieldName = e.Match.ToString().Replace("~", "");

                DocumentBuilder builder = new DocumentBuilder((Document)e.MatchNode.Document);
                builder.MoveTo((Run)runs[runs.Count - 1]);
                builder.InsertField(string.Format("MERGEFIELD {0} \\* MERGEFORMAT", mergeFieldName));

                //Now remove all runs in the sequence.
                foreach (Run run in runs)
                    run.Remove();

                return ReplaceAction.Skip;
            }

            private static Run SplitRun(Run run, int position)
            {
                Run afterRun = (Run)run.Clone(true);
                afterRun.Text = run.Text.Substring(position);
                run.Text = run.Text.Substring(0, position);
                run.ParentNode.InsertAfter(afterRun, run);
                return afterRun;
            }
        }
    }
}