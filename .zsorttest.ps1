$ErrorActionPreference = 'Stop'
$bin = 'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\src\MirareCiteAddIn\bin\x64\Release'
Get-ChildItem $bin -Filter *.dll | ForEach-Object { [void][Reflection.Assembly]::LoadFrom($_.FullName) }

$styles = 'C:\Users\admin\3D Objects\Ref Manager\styles'

# cited in this order: Liu (L) first, Kaur (K) second — APA should sort K before L
$cL = New-Object MirareCiteAddIn.Models.Citation
$cL.Id = 'b'; $cL.Title = 'Reproduction of educational disadvantage'; $cL.Authors = @('Liu, Jiajun', 'Pascarella, Ernest')
$cL.Year = 2022; $cL.Journal = 'Journal of Language, Identity & Education'; $cL.Doi = '10.1000/y'

$cK = New-Object MirareCiteAddIn.Models.Citation
$cK.Id = 'a'; $cK.Title = 'The silent struggle'; $cK.Authors = @('Kaur, Tarandeep', 'Newell, Samantha')
$cK.Year = 2024; $cK.Journal = 'Australian Journal of Psychology'; $cK.Doi = '10.1000/x'

$cA = New-Object MirareCiteAddIn.Models.Citation
$cA.Id = 'c'; $cA.Title = 'Zeta effects on learning'; $cA.Authors = @('Alaofi, Suad')
$cA.Year = 2020; $cA.Journal = 'ITiCSE'; $cA.Doi = '10.1000/z'

foreach ($t in @('apa.csl', 'chicago-author-date.csl', 'ieee.csl')) {
    $f = New-Object MirareCiteAddIn.Services.CitationFormatter((Join-Path $styles $t))
    $l = New-Object 'System.Collections.Generic.List[MirareCiteAddIn.Models.Citation]'
    $l.Add($cL); $l.Add($cK); $l.Add($cA)     # document order: L, K, A
    Write-Host ('=== ' + $t + ' (cited order: Liu, Kaur, Alaofi)')
    Write-Host ($f.BuildBibliography($l).Trim() -replace [char]13, [char]10)
    Write-Host ''
}
