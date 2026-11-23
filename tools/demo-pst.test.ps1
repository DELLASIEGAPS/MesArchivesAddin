param([Parameter(Mandatory = $true)][string]$PstPath)
$ErrorActionPreference = "Stop"
$PstPath = (Resolve-Path -LiteralPath $PstPath).ProviderPath
$source = Join-Path $PSScriptRoot "local-pst-agent.ps1"
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw $errors[0] }
foreach ($name in @("Add-CorsHeaders", "Get-ArchiveFolder", "Get-ArchiveMail", "Get-ArchiveFolderContents", "Get-ArchiveAddress", "Get-ArchiveMessage", "Write-ArchiveAttachment")) {
  $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
  if (-not $definition) { throw "Fonction absente : $name" }
  Invoke-Expression $definition.Extent.Text
}
$outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$namespace = $outlook.GetNamespace("MAPI")
foreach ($store in $namespace.Stores) {
  try { $storePath = $store.FilePath } catch { continue }
  if ($storePath -and $storePath -ieq $PstPath) { throw "Le PST de démonstration est déjà monté ; le test ne le modifiera pas." }
}
$root = $null
try {
  $namespace.AddStoreEx($PstPath, 3)
  foreach ($store in $namespace.Stores) {
    try { $storePath = $store.FilePath } catch { continue }
    if ($storePath -and $storePath -ieq $PstPath) { $root = $store.GetRootFolder(); break }
  }
  if (-not $root) { throw "PST de démonstration introuvable dans Outlook." }
  $archive = @{ namespace = $namespace; store = $store }
  $inbox = (Get-ArchiveFolderContents $archive "" 0).folders | Where-Object name -eq "Boite de reception"
  if (-not $inbox) { throw "Boite de reception absente." }
  $folder = (Get-ArchiveFolderContents $archive $inbox.id 0).folders | Where-Object name -eq "MADAIR"
  if (-not $folder) { throw "Dossier MADAIR absent." }
  $listing = Get-ArchiveFolderContents $archive $folder.id 0
  if ($listing.messages.Count -ne 1) { throw "Le message de démonstration est absent." }
  $id = $listing.messages[0].id
  $message = Get-ArchiveMessage $archive $id
  if ($message.body -notmatch "fictifs" -or $message.attachments.Count -ne 1) { throw "Message ou pièce jointe inattendu." }
  $response = [pscustomobject]@{ Headers = @{}; StatusCode = 0; ContentType = ""; ContentLength64 = 0; OutputStream = [System.IO.MemoryStream]::new() }
  $response | Add-Member ScriptMethod Close { }
  Write-ArchiveAttachment $response $archive $id $message.attachments[0].index
  if ($response.StatusCode -ne 200 -or $response.ContentLength64 -eq 0) { throw "Téléchargement de la pièce jointe échoué." }
  Write-Output "PST de démonstration : navigation, message et pièce jointe vérifiés."
} finally {
  if ($root) { $namespace.RemoveStore($root) }
}
