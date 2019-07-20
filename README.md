# convert-docx-to-pdf

Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs

## Use

The details below are produced from PowerShell using the following command. It removes the path prefix to the script.

```powershell
@"
(Get-Command .\Convert-DocxToPdf.ps1).Source | Split-Path |
ForEach-Object {
    (Get-Help .\Convert-DocxToPdf.ps1 -Detailed |
    Out-String -Stream).replace((Join-Path `$_ '\'), '')
}
"@ | Invoke-Expression
```

<br/>

```
NAME
    Convert-DocxToPdf.ps1

SYNOPSIS
    Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs


SYNTAX
    Convert-DocxToPdf.ps1 [-docxDir] <FileInfo> [<CommonParameters>]


DESCRIPTION
    Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs.
    The resulting pdf files will have the same filename as the Word documents.


PARAMETERS
    -docxDir <FileInfo>
        The full path to the directory containing the Word documents

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https://go.microsoft.com/fwlink/?LinkID=113216).

    -------------------------- EXAMPLE 1 --------------------------

    PS > Convert-DocxToPdf -docxDir "/Users/myusername/Documents/Word"






REMARKS
    To see the examples, type: "get-help Convert-DocxToPdf.ps1 -examples".
    For more information, type: "get-help Convert-DocxToPdf.ps1 -detailed".
    For technical information, type: "get-help Convert-DocxToPdf.ps1 -full".
    For online help, type: "get-help Convert-DocxToPdf.ps1 -online"
```
