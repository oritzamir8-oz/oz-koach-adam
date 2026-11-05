param([string]$In, [string]$Out)
# Minimal right-to-left Markdown -> .docx (headings, bullets, **bold**)
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

function Esc($s) { [System.Security.SecurityElement]::Escape($s) }
function Runs($text, $size) {
  $sz = if ($size) { "<w:sz w:val=`"$size`"/><w:szCs w:val=`"$size`"/>" } else { "" }
  $parts = [regex]::Split($text, '(\*\*[^*]+\*\*)')
  ($parts | Where-Object { $_ -ne '' } | ForEach-Object {
    if ($_ -match '^\*\*(.+)\*\*$') { $t = $Matches[1]; $b = '<w:b/><w:bCs/>' } else { $t = $_; $b = '' }
    "<w:r><w:rPr><w:rFonts w:ascii=`"Arial`" w:hAnsi=`"Arial`" w:cs=`"Arial`"/>$b$sz<w:rtl/></w:rPr><w:t xml:space=`"preserve`">$(Esc $t)</w:t></w:r>"
  }) -join ''
}

$body = foreach ($line in (Get-Content -LiteralPath $In -Encoding UTF8)) {
  if ($line -match '^\s*$') { continue }
  if ($line -match '^# (.+)') { "<w:p><w:pPr><w:pStyle w:val=`"Title`"/><w:bidi/><w:jc w:val=`"right`"/></w:pPr>$(Runs $Matches[1] 36)</w:p>"; continue }
  if ($line -match '^## (.+)') { "<w:p><w:pPr><w:pStyle w:val=`"Heading1`"/><w:keepNext/><w:spacing w:before=`"240`" w:after=`"120`"/><w:bidi/><w:jc w:val=`"right`"/></w:pPr>$(Runs $Matches[1] 28)</w:p>"; continue }
  if ($line -match '^(\s*)- (.+)') {
    $lvl = [math]::Min([int]($Matches[1].Length / 2), 1)
    "<w:p><w:pPr><w:numPr><w:ilvl w:val=`"$lvl`"/><w:numId w:val=`"1`"/></w:numPr><w:spacing w:after=`"60`"/><w:bidi/><w:jc w:val=`"right`"/></w:pPr>$(Runs $Matches[2] 22)</w:p>"; continue
  }
  "<w:p><w:pPr><w:bidi/><w:jc w:val=`"right`"/></w:pPr>$(Runs $line 22)</w:p>"
}

$W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
$files = @{
  '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
  '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
  'word/_rels/document.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
  'word/styles.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:styles $W><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=`"Arial`" w:hAnsi=`"Arial`" w:cs=`"Arial`"/><w:lang w:bidi=`"he-IL`"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/><w:spacing w:after=`"120`"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=`"paragraph`" w:default=`"1`" w:styleId=`"Normal`"><w:name w:val=`"Normal`"/></w:style><w:style w:type=`"paragraph`" w:styleId=`"Title`"><w:name w:val=`"Title`"/><w:basedOn w:val=`"Normal`"/><w:pPr><w:spacing w:before=`"0`" w:after=`"240`"/><w:keepNext/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val=`"1F3864`"/></w:rPr></w:style><w:style w:type=`"paragraph`" w:styleId=`"Heading1`"><w:name w:val=`"heading 1`"/><w:basedOn w:val=`"Normal`"/><w:pPr><w:outlineLvl w:val=`"0`"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val=`"2E74B5`"/></w:rPr></w:style></w:styles>"
  'word/numbering.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:numbering $W><w:abstractNum w:abstractNumId=`"0`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"&#8226;`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:bidi/><w:ind w:right=`"720`" w:hanging=`"360`"/></w:pPr></w:lvl><w:lvl w:ilvl=`"1`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"o`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:bidi/><w:ind w:right=`"1440`" w:hanging=`"360`"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=`"1`"><w:abstractNumId w:val=`"0`"/></w:num></w:numbering>"
  'word/document.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:document $W><w:body>$($body -join '')<w:sectPr><w:pgSz w:w=`"11906`" w:h=`"16838`"/><w:pgMar w:top=`"1440`" w:right=`"1440`" w:bottom=`"1440`" w:left=`"1440`" w:header=`"708`" w:footer=`"708`" w:gutter=`"0`"/><w:bidi/></w:sectPr></w:body></w:document>"
}

if (Test-Path -LiteralPath $Out) { Remove-Item -LiteralPath $Out }
$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Create')
$enc = New-Object System.Text.UTF8Encoding($false)
foreach ($k in $files.Keys) {
  $e = $zip.CreateEntry($k)
  $w = New-Object System.IO.StreamWriter($e.Open(), $enc)
  $w.Write($files[$k]); $w.Close()
}
$zip.Dispose()
"written $Out"
