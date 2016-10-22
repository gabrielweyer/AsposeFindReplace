# Find and replace sometimes modifies the order of the text

I'm trying to replace placeholders such as `~FieldName~` by MERGEFIELDs so that I can leverage the MailMerge feature. [Aspose's support][aspose-support] provided me with the right way of doing it and I updated this repository accordingly. If I find other issues in complex templates I'll update this repository again.

I've created a template with the following content:

```txt

Dear ~ClientDearName~,

~HelloOne~ ~HelloTwo~

```

After running my code I get

```txt
Evaluation Only. Created with Aspose.Words. Copyright 2003-2016 Aspose Pty Ltd.

Dear «ClientDearName»,

«HelloOne» «HelloTwo»

```

Which is the expected behavior.

## Configuration

- Windows 7 Enterprise SP1 64-bit
  - Location: Australia
  - Format: English (Australia)
  - Display language: English
  - Keyboard: English (Australia) - US
- Visual Studio Premium 2013
- `Aspose.Words.dll` version `16.10.0.0` (trial version)
- .NET 4.5

## Running the code

Open the solution in Visual Studio, press F5. The output document will be written to `$(TargetDir) of the Runner project\template`. The template will also be written to the same directory.

## Code

The code is located in the [MergeFieldMigrator][code] class.

## Sample document

The issue can be reproduced by using [this document][document].

[code]: Runner/MergeFieldMigrator.cs
[document]: Runner/template/revert-and-two-matches-same-line.docx
[aspose-support]: https://www.aspose.com/community/forums/thread/795097.aspx