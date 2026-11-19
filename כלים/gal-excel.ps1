# gal-excel.ps1 - the only supported way for Gal to read or write his opportunities workbook.
# ASCII only on purpose: Windows PowerShell 5.1 misreads non-ASCII text inside script files.
# Hebrew text arrives through a UTF-8 JSON file (-Rows) and the workbook path (-Path) typed on the command line.
#
# Modes:
#   status  : prints FREE, IN-USE or MISSING (exit 0, 2, 3).
#   links   : writes every logged row to a UTF-8 tab-separated file (-Out, default %TEMP%\gal-excel-rows.tsv)
#             and prints "ROWS <n> FILE <path>". Works while the workbook is open in Excel (read-only access).
#   append  : appends rows to sheet 1 and an optional summary row to sheet 2.
#             Holds the file exclusively for the whole operation, backs it up first to
#             %LOCALAPPDATA%\gal-backup, validates the result, and restores the original on any failure.
#             Prints "OK ..." (exit 0), IN-USE (exit 2), MISSING (exit 3) or "FAIL <reason>" (exit 1).
#
# JSON for append (UTF-8):
#   { "rows": [ { "A":"2026-11-12", "B":"...", ..., "I":"https://...", "J":72, "K":"...", "L":"...", "M":"..." } ],
#     "summary": { "date":"2026-11-12", "scanned":30, "checked":12, "passed":1, "failed":8, "clarify":3, "notes":"..." } }
#   Keys A..M are the sheet-1 columns. N and O belong to Orit and are never written. "summary" is optional.

param(
  [Parameter(Mandatory = $true)][ValidateSet('status', 'links', 'append')][string]$Mode,
  [Parameter(Mandatory = $true)][string]$Path,
  [string]$Rows,
  [string]$Out
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName WindowsBase

$Utf8 = New-Object System.Text.UTF8Encoding($false)
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$PartSheet1 = 'xl/worksheets/sheet1.xml'
$PartSheet2 = 'xl/worksheets/sheet2.xml'
$PartRels1 = 'xl/worksheets/_rels/sheet1.xml.rels'
$HyperlinkType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

function Read-Entry($zip, [string]$name) {
  $e = $zip.GetEntry($name)
  if (-not $e) { return $null }
  $sr = New-Object System.IO.StreamReader($e.Open(), $Utf8)
  try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Write-Entry($zip, [string]$name, [string]$text) {
  $old = $zip.GetEntry($name)
  if ($old) { $old.Delete() }
  $e = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
  $sw = New-Object System.IO.StreamWriter($e.Open(), $Utf8)
  try { $sw.Write($text) } finally { $sw.Dispose() }
}

function Clean([object]$value) {
  if ($null -eq $value) { return '' }
  $s = [string]$value
  $s = $s -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
  if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) }
  if ($s -match '^[=+\-@]') { $s = ' ' + $s }
  return [System.Security.SecurityElement]::Escape($s)
}

function Date-Serial([string]$d) {
  $dt = [datetime]::ParseExact($d.Trim(), 'yyyy-MM-dd', $Inv)
  return [int]($dt - (New-Object datetime 1899, 12, 30)).TotalDays
}

function Text-Cell([string]$ref, [object]$text, [string]$style) {
  $s = ''
  if ($style) { $s = ' s="' + $style + '"' }
  return '<c r="' + $ref + '"' + $s + ' t="inlineStr"><is><t xml:space="preserve">' + (Clean $text) + '</t></is></c>'
}

function Num-Cell([string]$ref, [double]$num, [string]$style) {
  $s = ''
  if ($style) { $s = ' s="' + $style + '"' }
  return '<c r="' + $ref + '"' + $s + '><v>' + $num.ToString($Inv) + '</v></c>'
}

function Has-Value([object]$v) { return ($null -ne $v) -and ([string]$v).Trim() -ne '' }

function Get-Prop($obj, [string]$name) {
  $p = $obj.PSObject.Properties[$name]
  if ($p) { return $p.Value }
  return $null
}

function Max-Row([string]$xml) {
  $m = [regex]::Matches($xml, '<row r="(\d+)"')
  $max = 0
  foreach ($x in $m) { $n = [int]$x.Groups[1].Value; if ($n -gt $max) { $max = $n } }
  return $max
}

function Count-Rows([string]$xml) { return ([regex]::Matches($xml, '<row ')).Count }

if (-not (Test-Path -LiteralPath $Path)) { 'MISSING'; exit 3 }

# ---------------- status ----------------
if ($Mode -eq 'status') {
  try { $f = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None'); $f.Dispose(); 'FREE'; exit 0 }
  catch { 'IN-USE'; exit 2 }
}

# ---------------- links ----------------
if ($Mode -eq 'links') {
  try {
    if (-not $Out) { $Out = Join-Path $env:TEMP 'gal-excel-rows.tsv' }
    $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    try { $mem = New-Object System.IO.MemoryStream; $fs.CopyTo($mem) } finally { $fs.Dispose() }
    $mem.Position = 0
    $zip = New-Object System.IO.Compression.ZipArchive($mem, [System.IO.Compression.ZipArchiveMode]::Read)
    $shared = New-Object System.Collections.Generic.List[string]
    $sstText = Read-Entry $zip 'xl/sharedStrings.xml'
    if ($sstText) {
      [xml]$sst = $sstText
      $nm = New-Object System.Xml.XmlNamespaceManager($sst.NameTable)
      $nm.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
      foreach ($si in $sst.SelectNodes('//m:si', $nm)) {
        $parts = @(); foreach ($t in $si.SelectNodes('.//m:t', $nm)) { $parts += $t.InnerText }
        $shared.Add(($parts -join ''))
      }
    }
    [xml]$sheet = Read-Entry $zip $PartSheet1
    $zip.Dispose()
    $nm = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
    $nm.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("row`tA`tB`tC`tI`tK`tM`tN`tO")
    $count = 0
    foreach ($row in $sheet.SelectNodes('//m:sheetData/m:row', $nm)) {
      $rn = [int]$row.GetAttribute('r')
      if ($rn -lt 2) { continue }
      $cells = @{}
      foreach ($c in $row.SelectNodes('m:c', $nm)) {
        $col = $c.GetAttribute('r') -replace '\d', ''
        $t = $c.GetAttribute('t')
        $v = $c.SelectSingleNode('m:v', $nm)
        $val = ''
        if ($t -eq 's' -and $v) { $val = $shared[[int]$v.InnerText] }
        elseif ($t -eq 'inlineStr') { $p = @(); foreach ($tn in $c.SelectNodes('.//m:t', $nm)) { $p += $tn.InnerText }; $val = ($p -join '') }
        elseif ($v) { $val = $v.InnerText }
        if ($col -eq 'A' -and $val -match '^\d+(\.\d+)?$') { $val = (New-Object datetime 1899, 12, 30).AddDays([double]::Parse($val, $Inv)).ToString('yyyy-MM-dd') }
        $cells[$col] = ($val -replace '[\t\r\n]+', ' ')
      }
      $fields = @($rn)
      foreach ($col in 'A', 'B', 'C', 'I', 'K', 'M', 'N', 'O') { $fields += $cells[$col] }
      $lines.Add(($fields -join "`t"))
      $count++
    }
    [System.IO.File]::WriteAllLines($Out, $lines, $Utf8)
    'ROWS ' + $count + ' FILE ' + $Out
    exit 0
  } catch { 'FAIL ' + $_.Exception.Message; exit 1 }
}

# ---------------- append ----------------
if (-not $Rows -or -not (Test-Path -LiteralPath $Rows)) { 'FAIL rows file not found'; exit 1 }
try { $data = Get-Content -LiteralPath $Rows -Raw -Encoding UTF8 | ConvertFrom-Json }
catch { 'FAIL rows file is not valid JSON: ' + $_.Exception.Message; exit 1 }
$newRows = @()
$rowsProp = Get-Prop $data 'rows'
if ($rowsProp) { $newRows = @($rowsProp) }
$summary = Get-Prop $data 'summary'
if ($newRows.Count -eq 0 -and -not $summary) { 'FAIL nothing to append'; exit 1 }

$fs = $null
try { $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None') }
catch { 'IN-USE'; exit 2 }

$origBytes = $null
$writing = $false
try {
  $orig = New-Object System.IO.MemoryStream
  $fs.CopyTo($orig)
  $origBytes = $orig.ToArray()

  $bdir = Join-Path $env:LOCALAPPDATA 'gal-backup'
  New-Item -ItemType Directory -Force -Path $bdir | Out-Null
  $bfile = Join-Path $bdir ('opps_' + (Get-Date -Format 'yyyy-MM-dd_HHmmss_fff') + '.xlsx')
  [System.IO.File]::WriteAllBytes($bfile, $origBytes)
  if ((Get-Item -LiteralPath $bfile).Length -ne $origBytes.Length) { throw 'backup size mismatch' }
  Get-ChildItem -LiteralPath $bdir -Filter 'opps_*.xlsx' | Sort-Object Name -Descending | Select-Object -Skip 30 | Remove-Item -Force

  $ms = New-Object System.IO.MemoryStream
  $ms.Write($origBytes, 0, $origBytes.Length)
  $zip = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Update, $true)

  $s1 = Read-Entry $zip $PartSheet1
  if (-not $s1 -or $s1.IndexOf('</sheetData>') -lt 0) { throw 'sheet 1 not found or empty' }
  $rowsBefore1 = Count-Rows $s1
  $last = Max-Row $s1

  $rels = Read-Entry $zip $PartRels1
  if (-not $rels) { $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>' }
  $rid = 0
  foreach ($x in [regex]::Matches($rels, 'Id="rId(\d+)"')) { $n = [int]$x.Groups[1].Value; if ($n -gt $rid) { $rid = $n } }

  $rowXml = New-Object System.Text.StringBuilder
  $linkXml = New-Object System.Text.StringBuilder
  $relXml = New-Object System.Text.StringBuilder
  foreach ($r in $newRows) {
    $last++
    [void]$rowXml.Append('<row r="' + $last + '" spans="1:15">')
    $a = Get-Prop $r 'A'
    if (Has-Value $a) { [void]$rowXml.Append((Num-Cell ('A' + $last) (Date-Serial ([string]$a)) '4')) }
    foreach ($col in 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M') {
      $v = Get-Prop $r $col
      if (-not (Has-Value $v)) { continue }
      $ref = $col + $last
      if ($col -eq 'J') {
        $num = 0.0
        if ([double]::TryParse(([string]$v).Trim(), [System.Globalization.NumberStyles]::Float, $Inv, [ref]$num)) { [void]$rowXml.Append((Num-Cell $ref $num '')) }
        else { [void]$rowXml.Append((Text-Cell $ref $v '')) }
      } elseif ($col -eq 'I') {
        [void]$rowXml.Append((Text-Cell $ref $v '3'))
        $url = ([string]$v).Trim()
        if ($url -match '^https?://\S+$') {
          $rid++
          $target = [System.Security.SecurityElement]::Escape($url)
          [void]$relXml.Append('<Relationship Id="rId' + $rid + '" Type="' + $HyperlinkType + '" Target="' + $target + '" TargetMode="External"/>')
          [void]$linkXml.Append('<hyperlink ref="' + $ref + '" r:id="rId' + $rid + '"/>')
        }
      } else {
        [void]$rowXml.Append((Text-Cell $ref $v ''))
      }
    }
    [void]$rowXml.Append('</row>')
  }

  if ($newRows.Count -gt 0) {
    $s1 = $s1.Replace('</sheetData>', $rowXml.ToString() + '</sheetData>')
    $s1 = [regex]::Replace($s1, '<dimension ref="[A-Z]+\d+(:[A-Z]+\d+)?"/>', ('<dimension ref="A1:O' + $last + '"/>'))
    if ($linkXml.Length -gt 0) {
      if ($s1.IndexOf('</hyperlinks>') -ge 0) { $s1 = $s1.Replace('</hyperlinks>', $linkXml.ToString() + '</hyperlinks>') }
      elseif ($s1.IndexOf('<printOptions') -ge 0) { $s1 = $s1.Replace('<printOptions', '<hyperlinks>' + $linkXml.ToString() + '</hyperlinks><printOptions') }
      elseif ($s1.IndexOf('<pageMargins') -ge 0) { $s1 = $s1.Replace('<pageMargins', '<hyperlinks>' + $linkXml.ToString() + '</hyperlinks><pageMargins') }
      else { $s1 = $s1.Replace('</worksheet>', '<hyperlinks>' + $linkXml.ToString() + '</hyperlinks></worksheet>') }
      $rels = $rels.Replace('</Relationships>', $relXml.ToString() + '</Relationships>')
      Write-Entry $zip $PartRels1 $rels
    }
    Write-Entry $zip $PartSheet1 $s1
  }

  $rowsBefore2 = 0
  if ($summary) {
    $s2 = Read-Entry $zip $PartSheet2
    if (-not $s2 -or $s2.IndexOf('</sheetData>') -lt 0) { throw 'sheet 2 not found or empty' }
    $rowsBefore2 = Count-Rows $s2
    $n2 = (Max-Row $s2) + 1
    $sx = '<row r="' + $n2 + '" spans="1:7">'
    $sx += Num-Cell ('A' + $n2) (Date-Serial ([string](Get-Prop $summary 'date'))) '2'
    $i = 0
    foreach ($k in 'scanned', 'checked', 'passed', 'failed', 'clarify') {
      $col = [string][char](66 + $i)
      $v = Get-Prop $summary $k
      if (Has-Value $v) { $sx += Num-Cell ($col + $n2) ([double]$v) '' }
      $i++
    }
    $notes = Get-Prop $summary 'notes'
    if (Has-Value $notes) { $sx += Text-Cell ('G' + $n2) $notes '' }
    $sx += '</row>'
    $s2 = $s2.Replace('</sheetData>', $sx + '</sheetData>')
    $s2 = [regex]::Replace($s2, '<dimension ref="[A-Z]+\d+(:[A-Z]+\d+)?"/>', ('<dimension ref="A1:G' + $n2 + '"/>'))
    Write-Entry $zip $PartSheet2 $s2
  }

  $zip.Dispose()

  # ---- validate the new package before touching the real file ----
  $ms.Position = 0
  $check = New-Object System.IO.Compression.ZipArchive($ms, [System.IO.Compression.ZipArchiveMode]::Read, $true)
  try {
    foreach ($name in '[Content_Types].xml', 'xl/workbook.xml', $PartSheet1, $PartSheet2, $PartRels1) {
      $txt = Read-Entry $check $name
      if ($null -eq $txt) { if ($name -eq $PartRels1 -and $linkXml.Length -eq 0) { continue } else { throw ('missing part ' + $name) } }
      [xml]$null = $txt
    }
    $after1 = Count-Rows (Read-Entry $check $PartSheet1)
    if ($after1 -ne $rowsBefore1 + $newRows.Count) { throw ('sheet 1 row count ' + $after1 + ', expected ' + ($rowsBefore1 + $newRows.Count)) }
    if ($summary) {
      $after2 = Count-Rows (Read-Entry $check $PartSheet2)
      if ($after2 -ne $rowsBefore2 + 1) { throw ('sheet 2 row count ' + $after2 + ', expected ' + ($rowsBefore2 + 1)) }
    }
  } finally { $check.Dispose() }
  $ms.Position = 0
  $pkg = [System.IO.Packaging.Package]::Open($ms, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
  try { $parts = @($pkg.GetParts()).Count; if ($parts -lt 5) { throw 'package has too few parts' } } finally { $pkg.Close() }

  # ---- write back through the same exclusive handle ----
  $newBytes = $ms.ToArray()
  $writing = $true
  $fs.Position = 0
  $fs.SetLength(0)
  $fs.Write($newBytes, 0, $newBytes.Length)
  $fs.Flush()
  $writing = $false
  'OK rows=' + $newRows.Count + ' summary=' + [bool]$summary + ' lastRow=' + $last + ' backup=' + $bfile
  exit 0
} catch {
  $msg = $_.Exception.Message
  if ($writing -and $origBytes) {
    try { $fs.Position = 0; $fs.SetLength(0); $fs.Write($origBytes, 0, $origBytes.Length); $fs.Flush(); $msg += ' (original restored)' }
    catch { $msg += ' (RESTORE FAILED - copy the latest file from ' + (Join-Path $env:LOCALAPPDATA 'gal-backup') + ')' }
  }
  'FAIL ' + $msg
  exit 1
} finally {
  if ($fs) { $fs.Dispose() }
}
