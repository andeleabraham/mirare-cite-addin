$ErrorActionPreference = 'Stop'
$bin = 'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\src\MirareCiteAddIn\bin\x64\Release'
Get-ChildItem $bin -Filter *.dll | ForEach-Object { [void][Reflection.Assembly]::LoadFrom($_.FullName) }

$logger = New-Object MirareCiteAddIn.Services.Logger
$lib    = New-Object MirareCiteAddIn.Services.RefManagerLoader($logger)

$cits = $lib.Load('C:\Users\admin\3D Objects\Ref Manager\refmanager.json')
Write-Host ("LIBRARY: {0} citations loaded" -f $cits.Count)
foreach ($c in $cits[0..2]) {
    Write-Host ("  id={0}" -f $c.Id)
    Write-Host ("    title : {0}" -f $c.Title)
    Write-Host ("    year  : {0}  journal: {1}" -f $c.Year, $c.Journal)
    Write-Host ("    authors ({0}): {1}" -f $c.Authors.Count, ($c.Authors -join ' | '))
    Write-Host ("    doi   : {0}  pages: {1}" -f $c.Doi, $c.Pages)
}

$proj = New-Object MirareCiteAddIn.Services.ProjectLoader($logger)
$p = $proj.Load('E:\citation_manager_app\test.mrrcite')
Write-Host ("PROJECT: {0} citations loaded" -f $p.Count)
foreach ($c in $p[0..2]) {
    Write-Host ("  id={0}" -f $c.Id)
    Write-Host ("    title : {0}" -f $c.Title)
    Write-Host ("    year  : {0}  journal: {1}" -f $c.Year, $c.Journal)
    Write-Host ("    authors ({0}): {1}" -f $c.Authors.Count, ($c.Authors -join ' | '))
}
