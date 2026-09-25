$ErrorActionPreference = 'Continue'
$targets = @(
    'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\src\MirareCiteAddIn\bin\x64\Release',
    'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\src\MirareCiteAddIn\bin\x64\Debug',
    'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\dist\MirareCite-AddIn-1.0.1'
)
$patterns = 'System.Text.*','System.Runtime.CompilerServices.Unsafe.dll','System.Buffers.dll',
            'System.Numerics.Vectors.dll','System.Threading.Tasks.Extensions.dll',
            'System.ValueTuple.dll','Microsoft.Bcl.AsyncInterfaces.dll'
foreach ($t in $targets) {
    if (-not (Test-Path $t)) { continue }
    foreach ($p in $patterns) {
        Get-ChildItem $t -Filter $p -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item $_.FullName -Force
            Write-Host ('removed ' + $_.FullName)
        }
    }
}
# Refresh the zip with the cleaned payload
$zip = 'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\dist\MirareCite-AddIn-1.0.1.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path 'C:\Users\admin\Downloads\mirare-cite-addin\mirare-cite-addin\dist\MirareCite-AddIn-1.0.1' -DestinationPath $zip
Write-Host ('zip refreshed: ' + (Get-Item $zip).Length + ' bytes')
