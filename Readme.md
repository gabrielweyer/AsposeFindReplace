# Find and replace sometimes modifies the order of the text

I'm trying to replace placeholders such as `~FieldName~` by MERGEFIELDs so that we can leverage the MailMerge feature.

I've created a template with the following content:

```txt

Dear ~ClientDearName~,



```

After running my code I get

```txt
Evaluation Only. Created with Aspose.Words. Copyright 2003-2016 Aspose Pty Ltd.

«ClientDearName»Dear ,



```

Notice that the `ClientDearName` MERGEFIELD is in front of `Dear` whereas the placeholder `~FieldName~` was behind `Dear `.

**Note**: I'm actually working on a more complex template but all the other placeholders are being replaced correctly.

## Configuration

- Windows 7 Enterprise SP1 64-bit
  - Location: Australia
  - Format: English (Australia)
  - Display language: English
  - Keyboard: English (Australia) - US
- Visual Studio Premium 2013
- `Aspose.Words.dll` version `16.10.0.0` (trial version)
- .NET 4.5

## Code

The code is located in the `MergeFieldMigrator` class.

```csharp
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
```

## Sample document

The issue can be reproduced by using the document located here: `Runner\template\repro.docx`