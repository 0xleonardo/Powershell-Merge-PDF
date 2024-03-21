<#
.SYNOPSIS
    Merges multiple PDF files into one.

    .DESCRIPTION
        Merges multiple PDF files into a single PDF document. Optionally allows for replacing an existing output file.

    .PARAMETER Force
        If specified, enables replacing the output file if it already exists.

    .INPUTS
        PDF files. Multiple PDF files can be selected using the OpenFileDialog.

    .OUTPUTS
        Merged PDF file. The merged PDF file is saved at the chosen location.

    .EXAMPLE
        PS> Merge-PDF
        Opens a dialog to select multiple PDF files, merges them, and saves the merged PDF as "Merged.pdf" in the chosen directory.

    .EXAMPLE
        PS> Merge-PDF -Force
        Opens a dialog to select multiple PDF files, merges them, and replaces the existing file "Merged.pdf" if it already exists in the chosen directory.

    .LINK
        PdfSharp package: https://www.nuget.org/packages/PdfSharp/ - PDFsharp is an open-source .NET library that easily creates and processes PDF documents on the fly from any .NET language.

    .NOTES
        Version:        1.0
        Author:         Leonardo Štavalj-Ladišić
        Creation Date:  15/03/2024
#>

param(
    [switch]$Force
)
begin {
    $PSDefaultParameterValues['*:Encoding'] = 'utf8BOM'

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -Path "PdfSharp.dll"

    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Choose folder where you want you PDF file to be saved"
    $dialogResult = $folderBrowser.ShowDialog()

    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Output "`nChoosen folder: $($folderBrowser.SelectedPath)"
        $OutputPath = Join-Path $folderBrowser.SelectedPath Merged.pdf
    } else {
        $scriptDir = $PSScriptRoot
        Write-Output "`nFolder not selected, using current one: $scriptDir"
        $OutputPath = Join-Path $scriptDir Merged.pdf
    }

    if (Test-Path $OutputPath) {
        if (-not $Force.IsPresent) {
            throw "`n$OutputPath exists. Use -Force parameter to override the file!"
        } else {
            Remove-Item $OutputPath -Force
            Write-Output "`nFile exists on the path! Deleted: $OutputPath"
        }
    }

    $outputDocument = New-Object PdfSharp.Pdf.PdfDocument
}
process {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Multiselect = $true
    $openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf"
    Write-Host "`nChoose PDF files to merge"

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $selectedFiles = $openFileDialog.FileNames
        Write-Host "`nChoosen PDF files:"
        foreach ($file in $selectedFiles) {
            Write-Host $file
        }
    } else {
        Write-Host "`nNo file was selected! Exiting the script..."
        exit
    }

    try {
        $progress = 0
        foreach ($file in $selectedFiles) {
            $inputPdf = [PdfSharp.Pdf.IO.PdfReader]::Open($file, [PdfSharp.Pdf.IO.PdfDocumentOpenMode]::Import)
            $pages = $inputPdf.Pages

            foreach($page in $pages) {
                $outputDocument.AddPage($page) | Out-Null
            }

            $progress++
            $percentComplete = ($progress / $selectedFiles.Count) * 100

            Write-Progress -Activity "Merging PDFs..." -Status "$percentComplete% completed" -PercentComplete $percentComplete
        }
    } catch {
        Write-Warning "`nExiting Script. An error occured while processing the files: $($_.Exception.Message)"
        exit
    }
}
end {
    try {
        if ($outputDocument.PageCount -eq 0) {
            Write-Warning "`nThere are no pages to create new PDF file! Exiting the script..."
            exit
        }

        $outputDocument.Save($outputPath) | Out-Null
        Write-Host "`nPDFs Merged! You can find your merged file here: $OutputPath"
        Invoke-Item -Path $OutputPath

        $outputDocument.Close()
    }
    finally {
        if ($null -ne $outputDocument) { 
            $outputDocument.Dispose()
        }
    }
}
