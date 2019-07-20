<#
.SYNOPSIS
Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs
.DESCRIPTION
Convert-DocxToPdf converts a directory of Word documents (docx) into pdfs
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

$wordObject = New-Object -ComObject Word.Application
$wordObject | Get-Member