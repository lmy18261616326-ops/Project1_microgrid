param(
    [Parameter(Mandatory=$true)][string]$InputDocx,
    [Parameter(Mandatory=$true)][string]$OutputPdf
)

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($InputDocx)
    $document.SaveAs2($OutputPdf, 17)
    $document.Close($false)
}
finally {
    $word.Quit()
}
