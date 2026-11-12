# Update the contents page numbers in Word, then save the .docx and a PDF copy.
# Run: powershell -File report\to_pdf.ps1
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $dir "Tideline_Project_Report.docx"
$dst = Join-Path $dir "Tideline_Project_Report.pdf"
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
try {
  # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ..., Visible, OpenAndRepair, ..., NoEncodingDialog)
  $d = $w.Documents.Open($src, $false, $false, $false)
  $d.Repaginate()
  [void]$d.Fields.Update()
  $d.Save()
  $d.ExportAsFixedFormat($dst, 17)
  "pages: " + $d.ComputeStatistics(2)
  $d.Close($false)
} finally { $w.Quit() }
