param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

$resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
$resolvedOutput = [System.IO.Path]::GetFullPath($OutputPath)
$extension = [System.IO.Path]::GetExtension($resolvedInput).ToLowerInvariant()

switch ($extension) {
    ".ppt" { $kind = "powerpoint" }
    ".pptx" { $kind = "powerpoint" }
    ".doc" { $kind = "word" }
    ".docx" { $kind = "word" }
    ".xls" { $kind = "excel" }
    ".xlsx" { $kind = "excel" }
    default { throw "Unsupported Office input: $extension" }
}

try {
    if ($kind -eq "powerpoint") {
        $app = New-Object -ComObject PowerPoint.Application
        $presentation = $app.Presentations.Open($resolvedInput, $true, $false, $false)
        $presentation.SaveAs($resolvedOutput, 32)
        $presentation.Close()
        $app.Quit()
        Release-ComObject $presentation
        Release-ComObject $app
    }

    if ($kind -eq "word") {
        $wdExportFormatPdf = 17
        $app = New-Object -ComObject Word.Application
        $app.Visible = $false
        $document = $app.Documents.Open($resolvedInput, $false, $true)
        $document.ExportAsFixedFormat($resolvedOutput, $wdExportFormatPdf)
        $document.Close([ref]0)
        $app.Quit()
        Release-ComObject $document
        Release-ComObject $app
    }

    if ($kind -eq "excel") {
        $xlTypePdf = 0
        $app = New-Object -ComObject Excel.Application
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $workbook = $app.Workbooks.Open($resolvedInput, $null, $true)
        $workbook.ExportAsFixedFormat($xlTypePdf, $resolvedOutput)
        $workbook.Close($false)
        $app.Quit()
        Release-ComObject $workbook
        Release-ComObject $app
    }
}
finally {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
