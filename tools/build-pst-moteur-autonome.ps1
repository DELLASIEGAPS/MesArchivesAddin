param([Parameter(Mandatory = $true)][string]$SourceDirectory)

$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $SourceDirectory).ProviderPath
$revision = 'da80989623f60685beaf5132948d32a3da48fd60'
$head = & git -C $source rev-parse HEAD
if ($LASTEXITCODE -ne 0 -or $head -ne $revision) { throw "Révision XstReader inattendue." }
$changes = & git -C $source status --porcelain
if ($LASTEXITCODE -ne 0 -or $changes) { throw "Le dossier XstReader doit être propre avant correction." }
$file = Join-Path $source 'src\XstReader.Api\XstFile.cs'
$project = Join-Path $source 'src\XstReader.Api\XstReader.Api.csproj'
$code = [System.IO.File]::ReadAllText($file)
$classPattern = '(public class XstFile : XstElement, IDisposable\s*\{)'
$oldClose = '_RootFolder.ClearContents();'
$closePattern = 'ClearStream\(\);\s*ClearRootFolder\(\);\s*_Ndb = null;\s*_Ltp = null;'
if ([regex]::Matches($code, $classPattern).Count -ne 1 -or -not $code.Contains($oldClose) -or
    [regex]::Matches($code, $closePattern).Count -ne 1) { throw 'Code XstReader inattendu ; correction refusée.' }
$code = [regex]::Replace($code, $classPattern, ('$1' + "`n        /// <summary>MesArchives synchronous close patch version.</summary>`n        public const string MesArchivesSafeCloseVersion = `"1`";"))
$code = $code.Replace($oldClose, '_RootFolder.ClearContentsInternal();')
$close = @'
try
            {
                ClearRootFolder();
            }
            finally
            {
                ClearStream();
                _Ndb = null;
                _Ltp = null;
            }
'@
$code = [regex]::Replace($code, $closePattern, $close)
$xml = [System.IO.File]::ReadAllText($project)
$old = 'Include="System.Security.Cryptography.Pkcs" Version="7.0.1"'
if (-not $xml.Contains($old)) { throw 'Dépendance XstReader inattendue ; correction refusée.' }
$xml = $xml.Replace($old, 'Include="System.Security.Cryptography.Pkcs" Version="8.0.2"')
[System.IO.File]::WriteAllText($file, $code)
[System.IO.File]::WriteAllText($project, $xml)
& dotnet restore $project --source 'https://api.nuget.org/v3/index.json'
if ($LASTEXITCODE -ne 0) { throw 'Restauration NuGet impossible.' }
& dotnet build $project --no-restore -c Release
if ($LASTEXITCODE -ne 0) { throw 'Compilation du moteur PST impossible.' }
Write-Host "Moteur corrigé : $(Join-Path $source 'src\XstReader.Api\bin\Release\netstandard2.0\XstReader.Api.dll')"
