# DropPdfToPowerPoint.ps1
# Drag & drop PDF files into PowerPoint as high-quality pictures, like on macOS.
#
# Two ways to use it:
#   1. Run with no arguments  -> a small always-on-top drop window appears.
#      Drag PDFs from Explorer onto it; pages are inserted into the active slide.
#   2. Run with file paths as arguments (e.g. drop a PDF onto the .bat launcher)
#      -> pages are inserted immediately, no window.
#
# Requires: Windows 10/11 (built-in Windows.Data.Pdf renderer) + PowerPoint.
# No third-party software needed.

param([string[]]$Files)

$ErrorActionPreference = 'Stop'
$script:RenderDpi = 300

# ---------------------------------------------------------------------------
# WinRT PDF renderer setup (built into Windows)
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Runtime.WindowsRuntime
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

$script:AsTaskOp = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    })[0]
$script:AsTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.FullName -eq 'Windows.Foundation.IAsyncAction'
    })[0]

function Wait-WinRtOperation([object]$Operation, [Type]$ResultType) {
    $task = $script:AsTaskOp.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    $null = $task.Wait(-1)
    return $task.Result
}

function Wait-WinRtAction([object]$Operation) {
    $task = $script:AsTaskAction.Invoke($null, @($Operation))
    $null = $task.Wait(-1)
}

# ---------------------------------------------------------------------------
# PDF -> PNG rendering
# ---------------------------------------------------------------------------
function Convert-PdfToPngs {
    param([string]$PdfPath, [switch]$FirstPageOnly)

    $outDir = Join-Path $env:TEMP ('DropPdf2Ppt\' + [Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Force -Path $outDir

    $storageFile = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)) ([Windows.Storage.StorageFile])
    $pdfDoc = Wait-WinRtOperation ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($storageFile)) ([Windows.Data.Pdf.PdfDocument])

    $pageCount = [int]$pdfDoc.PageCount
    if ($FirstPageOnly) { $pageCount = 1 }

    $pngs = @()
    for ($i = 0; $i -lt $pageCount; $i++) {
        $page = $pdfDoc.GetPage($i)
        try {
            $options = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            # Page.Size is in device-independent units (96/inch); scale to target DPI.
            $options.DestinationWidth = [uint32][Math]::Round($page.Size.Width * ($script:RenderDpi / 96.0))

            $raStream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            try {
                Wait-WinRtAction ($page.RenderToStreamAsync($raStream, $options))
                $pngPath = Join-Path $outDir ('page{0:d3}.png' -f ($i + 1))
                $inStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($raStream.GetInputStreamAt(0))
                $fileStream = [System.IO.File]::Create($pngPath)
                try { $inStream.CopyTo($fileStream) } finally { $fileStream.Dispose(); $inStream.Dispose() }
                $pngs += $pngPath
            }
            finally { $raStream.Dispose() }
        }
        finally { $page.Dispose() }
    }
    return , $pngs
}

# ---------------------------------------------------------------------------
# PowerPoint automation
# ---------------------------------------------------------------------------
function Get-PowerPoint {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
    catch {
        $ppt = New-Object -ComObject PowerPoint.Application
        return $ppt
    }
}

function Get-TargetSlide([object]$Ppt) {
    if ($Ppt.Presentations.Count -eq 0) {
        $null = $Ppt.Presentations.Add(-1)   # msoTrue: with window
    }
    $pres = $Ppt.ActivePresentation
    if ($pres.Slides.Count -eq 0) {
        $null = $pres.Slides.Add(1, 12)      # 12 = ppLayoutBlank
    }
    try {
        return $Ppt.ActiveWindow.View.Slide  # slide currently being edited
    }
    catch {
        return $pres.Slides.Item($pres.Slides.Count)
    }
}

function Add-PictureToSlide {
    param([object]$Slide, [string]$PngPath, [int]$StackOffset = 0)

    $pres = $Slide.Parent
    $slideW = [double]$pres.PageSetup.SlideWidth
    $slideH = [double]$pres.PageSetup.SlideHeight

    $img = [System.Drawing.Image]::FromFile($PngPath)
    try { $pxW = $img.Width; $pxH = $img.Height } finally { $img.Dispose() }

    # Natural size in points (72/inch), capped at 95% of the slide.
    $natW = $pxW / $script:RenderDpi * 72.0
    $natH = $pxH / $script:RenderDpi * 72.0
    $scale = [Math]::Min(1.0, [Math]::Min(($slideW * 0.95) / $natW, ($slideH * 0.95) / $natH))
    $w = $natW * $scale
    $h = $natH * $scale
    $left = ($slideW - $w) / 2 + $StackOffset * 15
    $top = ($slideH - $h) / 2 + $StackOffset * 15

    # AddPicture(FileName, LinkToFile=msoFalse, SaveWithDocument=msoTrue, L, T, W, H)
    $shape = $Slide.Shapes.AddPicture($PngPath, 0, -1, $left, $top, $w, $h)
    $null = $shape.Select()
    return $shape
}

function Insert-PdfIntoPowerPoint {
    param(
        [string[]]$PdfPaths,
        [bool]$OnePagePerSlide = $true,
        [scriptblock]$Status = { param($msg) Write-Host $msg }
    )

    $ppt = Get-PowerPoint
    $slide = Get-TargetSlide $ppt
    $pres = $slide.Parent
    try { $ppt.Activate() } catch { }

    $imageIndex = 0
    $tempDirs = @()
    foreach ($pdf in $PdfPaths) {
        $name = [System.IO.Path]::GetFileName($pdf)
        & $Status "Rendering $name ..."
        $pngs = Convert-PdfToPngs -PdfPath $pdf
        if ($pngs.Count -gt 0) { $tempDirs += (Split-Path $pngs[0] -Parent) }

        $pageNum = 0
        foreach ($png in $pngs) {
            $pageNum++
            & $Status ("Inserting {0}  (page {1}/{2})" -f $name, $pageNum, $pngs.Count)
            if ($imageIndex -eq 0) {
                $null = Add-PictureToSlide -Slide $slide -PngPath $png
            }
            elseif ($OnePagePerSlide) {
                $newIndex = $slide.SlideIndex + 1
                $slide = $pres.Slides.Add($newIndex, 12)   # blank layout
                try { $ppt.ActiveWindow.View.GotoSlide($newIndex) } catch { }
                $null = Add-PictureToSlide -Slide $slide -PngPath $png
            }
            else {
                $null = Add-PictureToSlide -Slide $slide -PngPath $png -StackOffset $imageIndex
            }
            $imageIndex++
        }
    }

    foreach ($dir in $tempDirs) {
        Remove-Item -Recurse -Force -Path $dir -ErrorAction SilentlyContinue
    }
    return $imageIndex
}

# ---------------------------------------------------------------------------
# CLI mode: files passed as arguments (e.g. PDF dropped onto the .bat)
# ---------------------------------------------------------------------------
if ($Files -and $Files.Count -gt 0) {
    $pdfs = @($Files | Where-Object { $_ -and (Test-Path $_) -and ([System.IO.Path]::GetExtension($_) -ieq '.pdf') } |
        ForEach-Object { (Resolve-Path $_).Path })
    if ($pdfs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No PDF files found in the dropped items.', 'Drop PDF to PowerPoint',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        exit 1
    }
    try {
        $count = Insert-PdfIntoPowerPoint -PdfPaths $pdfs -OnePagePerSlide $true
        Write-Host "Inserted $count page(s) into PowerPoint."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Could not insert PDF:`n$($_.Exception.Message)", 'Drop PDF to PowerPoint',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        exit 1
    }
    exit 0
}

# ---------------------------------------------------------------------------
# GUI mode: always-on-top drop window
# ---------------------------------------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'PDF -> PowerPoint'
$form.Size = New-Object System.Drawing.Size(300, 230)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.TopMost = $true
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form.Location = New-Object System.Drawing.Point(($screen.Right - $form.Width - 24), ($screen.Bottom - $form.Height - 24))
$form.AllowDrop = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 246, 250)

$dropZone = New-Object System.Windows.Forms.Label
$dropZone.Text = "Drop PDF here`n`nPages are inserted into`nthe current slide as pictures"
$dropZone.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$dropZone.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$dropZone.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 80)
$dropZone.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$dropZone.Location = New-Object System.Drawing.Point(12, 12)
$dropZone.Size = New-Object System.Drawing.Size(260, 110)
$dropZone.AllowDrop = $true
$form.Controls.Add($dropZone)

$chkPerSlide = New-Object System.Windows.Forms.CheckBox
$chkPerSlide.Text = 'Each page on its own slide'
$chkPerSlide.Checked = $true
$chkPerSlide.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$chkPerSlide.Location = New-Object System.Drawing.Point(14, 130)
$chkPerSlide.Size = New-Object System.Drawing.Size(260, 22)
$form.Controls.Add($chkPerSlide)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = 'Ready'
$statusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 8.5)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(110, 110, 120)
$statusLabel.Location = New-Object System.Drawing.Point(14, 158)
$statusLabel.Size = New-Object System.Drawing.Size(260, 30)
$form.Controls.Add($statusLabel)

function Set-Status([string]$Message) {
    $statusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}

$onDragEnter = {
    param($s, $e)
    $hasPdf = $false
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        $hasPdf = @($paths | Where-Object { [System.IO.Path]::GetExtension($_) -ieq '.pdf' }).Count -gt 0
    }
    if ($hasPdf) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy }
    else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
}

$onDragDrop = {
    param($s, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    $pdfs = @($paths | Where-Object { [System.IO.Path]::GetExtension($_) -ieq '.pdf' })
    if ($pdfs.Count -eq 0) { return }
    try {
        $dropZone.BackColor = [System.Drawing.Color]::FromArgb(222, 235, 250)
        $count = Insert-PdfIntoPowerPoint -PdfPaths $pdfs -OnePagePerSlide $chkPerSlide.Checked -Status ${function:Set-Status}
        Set-Status "Done - inserted $count page(s)."
    }
    catch {
        Set-Status 'Failed - see message.'
        [System.Windows.Forms.MessageBox]::Show("Could not insert PDF:`n$($_.Exception.Message)", 'Drop PDF to PowerPoint',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
    finally {
        $dropZone.BackColor = $form.BackColor
    }
}

$form.Add_DragEnter($onDragEnter)
$form.Add_DragDrop($onDragDrop)
$dropZone.Add_DragEnter($onDragEnter)
$dropZone.Add_DragDrop($onDragDrop)

$null = $form.ShowDialog()
