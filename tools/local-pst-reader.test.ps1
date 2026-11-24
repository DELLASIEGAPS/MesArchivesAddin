$ErrorActionPreference = "Stop"
$source = Join-Path $PSScriptRoot "local-pst-agent.ps1"
$tokens = $null
$errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw $errors[0] }
$names = @("Add-CorsHeaders", "Get-ArchiveFolder", "Get-ArchiveMail", "Get-ArchiveFolderContents", "Get-ArchiveAddress", "Get-ArchiveMessage", "Write-ArchiveAttachment")
foreach ($name in $names) {
  $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
  if (-not $definition) { throw "Missing function $name" }
  Invoke-Expression $definition.Extent.Text
}

function Assert($condition, [string]$message) {
  if (-not $condition) { throw $message }
}
function Collection($records) {
  $collection = [pscustomobject]@{ records = @($records); Count = @($records).Count }
  $collection | Add-Member ScriptMethod Item { param($index) $this.records[$index - 1] }
  return $collection
}

$attachment = [pscustomobject]@{ FileName = "document.txt"; Size = 7; SavedPath = "" }
$attachment | Add-Member ScriptMethod SaveAsFile {
  param($path)
  $this.SavedPath = $path
  [System.IO.File]::WriteAllText($path, "fixture", [System.Text.Encoding]::ASCII)
}
$item = [pscustomobject]@{
  Class = 43; EntryID = "mail-1"; Subject = "Sujet"; SenderName = "Expéditeur"
  Sender = [pscustomobject]@{ Type = "SMTP"; Address = "sender@example.test" }
  Recipients = @([pscustomobject]@{ Type = 1; Name = "Destinataire"; AddressEntry = [pscustomobject]@{ Type = "SMTP"; Address = "recipient@example.test" } })
  To = "Destinataire"; CC = ""; BCC = "Copie cachée"; ReceivedTime = [datetime]"2026-01-02T03:04:05"
  Body = "Contenu de test"; Parent = [pscustomobject]@{ StoreID = "store-1" }
  Attachments = (Collection @($attachment))
}
$otherItem = [pscustomobject]@{ Class = 26; EntryID = "other-1" }
$child = [pscustomobject]@{ EntryID = "folder-1"; StoreID = "store-1"; Name = "Sous-dossier"; Items = (Collection @()); Folders = @() }
$folder = [pscustomobject]@{ EntryID = "root"; StoreID = "store-1"; Name = "Racine"; Items = (Collection (@($item, $otherItem) * 26)); Folders = @($child) }
$store = [pscustomobject]@{ StoreID = "store-1"; Root = $folder }
$store | Add-Member ScriptMethod GetRootFolder { $this.Root }
$namespace = [pscustomobject]@{ Folder = $child; Mail = $item }
$namespace | Add-Member ScriptMethod GetFolderFromID { param($id, $storeId) $this.Folder }
$namespace | Add-Member ScriptMethod GetItemFromID { param($id, $storeId) $this.Mail }
$archive = @{ store = $store; namespace = $namespace }

$first = Get-ArchiveFolderContents $archive "" 0
Assert ($first.ok -and $first.totalItems -eq 52 -and $first.nextOffset -eq 50) "Pagination incorrecte"
Assert ($first.folders.Count -eq 1 -and $first.messages.Count -eq 25) "Dossiers ou filtre de messages incorrects"
$last = Get-ArchiveFolderContents $archive "" 50
Assert ($last.messages.Count -eq 1 -and $null -eq $last.nextOffset) "Dernière page incorrecte"
$message = Get-ArchiveMessage $archive "mail-1"
Assert ($message.body -eq "Contenu de test" -and $message.attachments.Count -eq 1) "Détail du message incorrect"
Assert ($message.bcc -eq "Copie cachée") "Destinataire Cci indisponible pour la recherche"
Assert ($message.senderContact.address -eq "sender@example.test" -and $message.recipients[0].contact.address -eq "recipient@example.test") "Adresses SMTP incorrectes"
$exchangeEntry = [pscustomobject]@{ Type = "EX" }
$exchangeEntry | Add-Member ScriptMethod GetExchangeUser { [pscustomobject]@{ PrimarySmtpAddress = "exchange@example.test" } }
Assert ((Get-ArchiveAddress $exchangeEntry "Exchange").address -eq "exchange@example.test") "Adresse Exchange incorrecte"
$unknownEntry = [pscustomobject]@{ Type = "UNKNOWN" }
Assert ((Get-ArchiveAddress $unknownEntry "Inconnu").address -eq "") "Adresse inconnue inventee"

$response = [pscustomobject]@{ Headers = @{}; StatusCode = 0; ContentType = ""; ContentLength64 = 0; OutputStream = [System.IO.MemoryStream]::new() }
$response | Add-Member ScriptMethod Close { }
Write-ArchiveAttachment $response $archive "mail-1" 1
Assert ($response.StatusCode -eq 200 -and $response.ContentLength64 -eq 7) "Téléchargement incorrect"
Assert ([System.Text.Encoding]::ASCII.GetString($response.OutputStream.ToArray()) -eq "fixture") "Contenu téléchargé incorrect"
Assert (-not (Test-Path -LiteralPath $attachment.SavedPath)) "Fichier temporaire non supprimé"

$item.Parent.StoreID = "other-store"
$rejected = $false
try { $null = Get-ArchiveMessage $archive "mail-1" } catch { $rejected = $true }
Assert $rejected "Un message appartenant à un autre PST a été accepté"
$child.StoreID = "other-store"
$rejected = $false
try { $null = Get-ArchiveFolder $archive "folder-1" } catch { $rejected = $true }
Assert $rejected "Un dossier appartenant à un autre PST a été accepté"
Write-Output "Reader fixture tests passed."
