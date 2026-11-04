param([string[]]$Files, [string]$OutDir)
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$w.Options.CheckSpellingAsYouType = $false
$w.Options.CheckGrammarAsYouType = $false
try {
  foreach ($f in $Files) {
    $t0 = Get-Date
    $full = (Resolve-Path $f).Path
    $name = [IO.Path]::GetFileNameWithoutExtension($full)
    $out = Join-Path $OutDir ($name + ".pdf")
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $d = $w.Documents.Open($full, $false, $true, $false)
    "{0}: opened in {1:n1}s" -f $f, ((Get-Date) - $t0).TotalSeconds
    $d.SaveAs2([string]$out, 17)
    "{0}: saved pdf in {1:n1}s -> {2}" -f $f, ((Get-Date) - $t0).TotalSeconds, $out
    $d.Close(0)
  }
} finally { $w.Quit() }
