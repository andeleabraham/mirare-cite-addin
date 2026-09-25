$ErrorActionPreference = 'Stop'
$bin = 'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\src\MirareCiteAddIn\bin\x64\Release'
Get-ChildItem $bin -Filter *.dll | ForEach-Object { [void][Reflection.Assembly]::LoadFrom($_.FullName) }

$raw = Get-Content 'C:\Users\admin\3D Objects\Ref Manager\styles\acm-sig-proceedings.csl' -Raw -Encoding UTF8
$i = $raw.IndexOf('<style')
Write-Host ('ROOT: ' + $raw.Substring($i, 220))

$c = New-Object MirareCiteAddIn.Models.Citation
$c.Id = 'x'; $c.Title = 'Some title'; $c.Authors = @('Kaur, Tarandeep', 'Newell, Samantha')
$c.Year = 2024; $c.Journal = 'J'; $c.Volume = '9'; $c.Issue = '5'; $c.Pages = '147'
$c.Type = 'journal-article'

try {
    $f = New-Object MirareCiteAddIn.Services.CitationFormatter('C:\Users\admin\3D Objects\Ref Manager\styles\acm-sig-proceedings.csl')
    Write-Host ('formatter loaded, IsNumericStyle: ' + $f.IsNumericStyle)
    Write-Host ('InTextCore(c): ' + $f.InTextCore($c))
    Write-Host ('InTextCore(c,5): ' + $f.InTextCore($c, 5))
    $l = New-Object 'System.Collections.Generic.List[MirareCiteAddIn.Models.Citation]'
    $l.Add($c)
    Write-Host ('bib: ' + $f.BuildBibliography($l).Trim())
} catch {
    Write-Host ('THREW: ' + $_.Exception.ToString().Substring(0, 300))
}
