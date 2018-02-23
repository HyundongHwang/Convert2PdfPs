# $global:G_STORAGE_CONTEXT = $null



<#
.SYNOPSIS
.EXAMPLE
#>
function convert2pdf
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelinebyPropertyName=$true)]
        [System.IO.FileInfo]
        $FILE_PATH
    )

    process {
        $pdfPath = "$($FILE_PATH.DirectoryName)\$($FILE_PATH.BaseName).pdf"

        if (Test-Path $pdfPath) {
            Write-Host "$pdfPath already exist !!!"
            Write-Output $pdfPath
            return
        }



        if(($FILE_PATH -like "*.doc") -or ($FILE_PATH -like "*.docx")) {
            $wordCom = New-Object -ComObject Word.Application
            $doc = $wordCom.Documents.Open($FILE_PATH.FullName)
            Write-Host "$($FILE_PATH.FullName) start ..."
            $doc.SaveAs($pdfPath, 17)
            $doc.Close()
            Write-Host "$pdfPath done ..."
            $wordCom.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordCom)
            Write-Output $pdfPath
        }

        if(($FILE_PATH -like "*.ppt") -or ($FILE_PATH -like "*.pptx")) {
            $pptCom = New-Object -ComObject PowerPoint.Application
            $doc = $pptCom.Presentations.Open($FILE_PATH.FullName)
            Write-Host "$($FILE_PATH.FullName) start ..."
            $opt= [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
            $doc.SaveAs($pdfPath, $opt)
            $doc.Close()
            Write-Host "$pdfPath done ..."
            $pptCom.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptCom)
            Write-Output $pdfPath
        }
    }
}