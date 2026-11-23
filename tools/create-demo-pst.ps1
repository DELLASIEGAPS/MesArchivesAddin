param(
  [Parameter(Mandatory = $true)]
  [string]$DestinationDirectory
)

$ErrorActionPreference = "Stop"
if (-not (Test-Path -LiteralPath $DestinationDirectory -PathType Container)) {
  throw "Le dossier de destination n'existe pas : $DestinationDirectory"
}
$pstPath = Join-Path (Resolve-Path -LiteralPath $DestinationDirectory).ProviderPath "ADR_2021.pst"
if (Test-Path -LiteralPath $pstPath) { throw "Le fichier existe déjà et ne sera pas remplacé : $pstPath" }
$outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
$namespace = $outlook.GetNamespace("MAPI")
$tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("MesArchives-demo-" + [guid]::NewGuid().ToString() + ".txt")
$root = $null
try {
  [System.IO.File]::WriteAllText($tempPath, "Piece jointe fictive pour la demonstration Mes Archives.", [System.Text.Encoding]::UTF8)
  $namespace.AddStoreEx($pstPath, 3)
  foreach ($store in $namespace.Stores) {
    try { $storePath = $store.FilePath } catch { continue }
    if ($storePath -and $storePath -ieq $pstPath) {
      $root = $store.GetRootFolder()
      break
    }
  }
  if (-not $root) { throw "Le nouveau PST n'a pas été trouvé dans Outlook." }
  $root.Name = "Archives 2021 - demonstration"
  $inbox = $root.Folders.Add("Boite de reception", 6)
  $folder = $inbox.Folders.Add("MADAIR", 6)
  $mail = $folder.Items.Add("IPM.Note")
  $mail.Subject = "[DEMO] Message d'archive avec piece jointe"
  $mail.To = "destinataire@example.test"
  $mail.Body = "Bonjour,`r`n`r`nCe message et sa piece jointe sont fictifs.`r`n`r`nDemonstration Mes Archives."
  $null = $mail.Attachments.Add($tempPath, 1, 1, "demonstration.txt")
  $mail.Save()
  $savedMail = $mail.Move($folder)
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)
  $mail = $null
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($savedMail)
  $savedMail = $null
  if ($folder.Items.Count -ne 1) { throw "Le message n'a pas ete enregistre dans le PST." }
  Write-Output "PST de demonstration cree : $pstPath"
} finally {
  if ($mail) {
    try { $mail.Delete() } finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
  }
  if ($savedMail) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($savedMail) }
  if ($folder) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($folder) }
  if ($inbox) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($inbox) }
  if ($root) { $namespace.RemoveStore($root) }
  if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
}
