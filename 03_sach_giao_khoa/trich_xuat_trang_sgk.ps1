param(
    [Parameter(Mandatory=$true)]
    [string]$PdfPath,

    [Parameter(Mandatory=$true)]
    [int]$StartPage,

    [Parameter(Mandatory=$true)]
    [int]$EndPage,

    [Parameter(Mandatory=$false)]
    [string]$OutputDir = "$PSScriptRoot\extracted_pages"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFolder, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$null = [Windows.Storage.CreationCollisionOption, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.FileAccessMode, Windows.Storage, ContentType = WindowsRuntime]

function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}

function AwaitAction($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and !$_.IsGenericMethod
    })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}

$resolvedPdf = (Resolve-Path $PdfPath).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$resolvedOut = (Resolve-Path $OutputDir).Path

$storageFile = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($resolvedPdf)) ([Windows.Storage.StorageFile])
$pdfDoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])
$outFolder = Await ([Windows.Storage.StorageFolder]::GetFolderFromPathAsync($resolvedOut)) ([Windows.Storage.StorageFolder])

$total = [int]$pdfDoc.PageCount
if ($StartPage -lt 1) { $StartPage = 1 }
if ($EndPage -gt $total) { $EndPage = $total }

Write-Host "Đang trích xuất từ trang $StartPage đến $EndPage (Tổng số trang PDF: $total)..."
for ($p = $StartPage; $p -le $EndPage; $p++) {
    $page = $pdfDoc.GetPage([uint32]($p - 1))
    $fileName = "page_{0:D3}.png" -f $p
    $imgFile = Await ($outFolder.CreateFileAsync($fileName, [Windows.Storage.CreationCollisionOption]::ReplaceExisting)) ([Windows.Storage.StorageFile])
    $stream = Await ($imgFile.OpenAsync([Windows.Storage.FileAccessMode]::ReadWrite)) ([Windows.Storage.Streams.IRandomAccessStream])
    AwaitAction ($page.RenderToStreamAsync($stream))
    $stream.Dispose()
    $page.Dispose()
    Write-Host "Đã xuất: $resolvedOut\$fileName"
}
