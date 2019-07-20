<#
.SYNOPSIS
Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs

.DESCRIPTION
Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs.
The resulting pdf files will have the same filename as the Word documents.

.PARAMETER docxDir
The full path to the directory containing the Word documents

.EXAMPLE
Convert-DocxToPdf -docxDir "/Users/myusername/Documents/Word"

.LINK
https://github.com/curtisalexander/convert-docx-to-pdf

.NOTES
Author: Curtis Alexander
#>
Param(
    [Parameter(Mandatory = $true, HelpMessage = "The full path to the directory containing the Word documents")]
    [ValidateScript( {
            if (-not (Test-Path -Path $_)) {
                throw "The docxDir directory does not exist."
            }
            if (-not (Test-Path -Path $_ -PathType Container) ) {
                throw "The docxDir argument must be a directory."
            }
            return $true
        })]
    [System.IO.FileInfo]$docxDir
)

function convertToPdf {
    Param(
        [Parameter(Mandatory = $true)]
        [System.__ComObject]$wordApp,
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$wordFile
    )
    $pdfFile = Join-Path $wordFile.DirectoryName "$($wordFile.Basename).pdf" 
    $wordDoc = $wordApp.Documents.Open($wordFile.FullName)
    Write-Host "Converting $wordFile to $pdfFile"
    try {
        $wordDoc.SaveAs($pdfFile, 17)
    }
    finally {
        $wordDoc.Close()
    }
}

$wordApp = New-Object -ComObject Word.Application
Get-ChildItem -Path "${docxDir}" -Filter "*.docx" | ForEach-Object {
    convertToPdf -wordApp $wordApp -wordFile $_
}
$wordApp.Quit()
