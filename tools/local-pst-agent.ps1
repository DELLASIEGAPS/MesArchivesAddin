param(
  [string]$ListenPrefix = "http://127.0.0.1:45731/",
  [string]$PstSourceDirectory = ""
)

$ErrorActionPreference = "Stop"

# Outlook COM est détecté par GetActiveObject, indisponible dans PowerShell 7.
if (-not ([System.Runtime.InteropServices.Marshal].GetMethods() | Where-Object { $_.Name -eq "GetActiveObject" })) {
  throw "Cet agent nécessite Windows PowerShell 5.1 pour détecter Outlook classique. Lancez-le avec powershell.exe -NoProfile -File."
}

if ($PstSourceDirectory) {
  if (-not (Test-Path -LiteralPath $PstSourceDirectory -PathType Container)) {
    throw "Le dossier source PST n'existe pas : $PstSourceDirectory"
  }
  $PstSourceDirectory = (Resolve-Path -LiteralPath $PstSourceDirectory).ProviderPath
}

function Get-OneDriveRoots {
  $roots = @()
  if ($env:OneDriveCommercial) { $roots += $env:OneDriveCommercial }
  if ($env:OneDrive) { $roots += $env:OneDrive }
  if ($env:OneDriveConsumer) { $roots += $env:OneDriveConsumer }
  $roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
}

function Find-PstFile([string]$fileName) {
  if ([string]::IsNullOrWhiteSpace($fileName)) { return $null }
  if ($fileName.Contains("\") -or $fileName.Contains("/") -or $fileName.Contains("..")) { return $null }
  if ($fileName -notlike "*.pst") { return $null }
  # Un dossier explicite évite de choisir une autre copie du même PST dans OneDrive.
  if ($PstSourceDirectory) {
    $candidate = Join-Path $PstSourceDirectory $fileName
    if (Test-Path -LiteralPath $candidate -PathType Leaf) { return $candidate }
    return $null
  }
  foreach ($root in Get-OneDriveRoots) {
    $match = Get-ChildItem -Path $root -Recurse -File -Filter $fileName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($match) { return $match.FullName }
  }
  return $null
}

function Get-WorkingPstPath([string]$fileName) {
  if ([string]::IsNullOrWhiteSpace($fileName) -or
      $fileName.Contains("\") -or $fileName.Contains("/") -or
      $fileName.Contains("..") -or $fileName -notlike "*.pst") {
    throw "Invalid PST filename."
  }
  if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    throw "LOCALAPPDATA is not configured."
  }
  return Join-Path (Join-Path $env:LOCALAPPDATA "MesArchives\PST") $fileName
}

function Get-OrCreateWorkingPst([string]$fileName) {
  $workingPath = Get-WorkingPstPath $fileName
  if (Test-Path -LiteralPath $workingPath -PathType Leaf) { return $workingPath }

  $sourcePath = Find-PstFile $fileName
  if (-not $sourcePath) { return $null }
  if (Test-CloudOnlyPlaceholder $sourcePath) {
    throw "Le fichier PST est uniquement dans le cloud. Rendez-le disponible sur ce PC dans OneDrive : $fileName"
  }

  $directory = Split-Path -Parent $workingPath
  New-Item -ItemType Directory -Path $directory -Force | Out-Null
  $temporaryPath = Join-Path $directory ([System.Guid]::NewGuid().ToString() + ".tmp")
  try {
    Copy-Item -LiteralPath $sourcePath -Destination $temporaryPath
    [System.IO.File]::Move($temporaryPath, $workingPath)
  } finally {
    if (Test-Path -LiteralPath $temporaryPath) {
      Remove-Item -LiteralPath $temporaryPath -Force
    }
  }
  return $workingPath
}

function Add-CorsHeaders($res) {
  $res.Headers["Access-Control-Allow-Origin"] = "https://dellasiegaps.github.io"
  $res.Headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
  $res.Headers["Access-Control-Allow-Headers"] = "Content-Type"
  $res.Headers["Access-Control-Allow-Private-Network"] = "true"
}

function Write-Json($res, [int]$status, $obj) {
  $res.StatusCode = $status
  $res.ContentType = "application/json; charset=utf-8"
  $res.Headers["Cache-Control"] = "no-store"
  $res.Headers["X-Content-Type-Options"] = "nosniff"
  Add-CorsHeaders $res
  $json = ($obj | ConvertTo-Json -Depth 16 -Compress)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
  $res.Close()
}

# Attache uniquement une instance déjà ouverte : New-Object relancerait Outlook en arrière-plan.
function Get-RunningOutlook {
  try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { return $null }
}

function Test-PstMounted($outlook, [string]$pstPath) {
  foreach ($store in $outlook.GetNamespace("MAPI").Stores) {
    try { if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) { return $true } } catch {}
  }
  return $false
}

function Set-PstDisplayName($outlook, [string]$pstPath, [string]$displayName) {
  if ([string]::IsNullOrWhiteSpace($displayName)) { return }
  foreach ($store in $outlook.GetNamespace("MAPI").Stores) {
    try {
      if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) {
        $root = $store.GetRootFolder()
        if ($root.Name -ne $displayName) { $root.Name = $displayName }
        return
      }
    } catch {}
  }
}

function Mount-Pst([string]$pstPath, [string]$displayName) {
  $outlook = Get-RunningOutlook
  if (-not $outlook) { throw [System.InvalidOperationException]::new("OUTLOOK_NOT_RUNNING") }
  $already = Test-PstMounted $outlook $pstPath
  if (-not $already) { $outlook.GetNamespace("MAPI").AddStoreEx($pstPath, 3) } # olStoreUnicode
  Set-PstDisplayName $outlook $pstPath $displayName
  if ($already) { return @{ ok = $true; alreadyMounted = $true; path = $pstPath } }
  return @{ ok = $true; mounted = $true; path = $pstPath }
}

function Unmount-Pst([string]$pstPath) {
  $outlook = Get-RunningOutlook
  if (-not $outlook) { return @{ ok = $true; alreadyUnmounted = $true; outlookRunning = $false; path = $pstPath } }
  $ns = $outlook.GetNamespace("MAPI")
  foreach ($store in $ns.Stores) {
    $storePath = $null
    try { $storePath = $store.FilePath } catch { continue }
    if ($storePath -and ($storePath -ieq $pstPath)) {
      $ns.RemoveStore($store.GetRootFolder())
      return @{ ok = $true; unmounted = $true; path = $pstPath }
    }
  }
  return @{ ok = $true; alreadyUnmounted = $true; path = $pstPath }
}

# Fichier OneDrive « en ligne uniquement » : attributs RECALL_ON_DATA_ACCESS, RECALL_ON_OPEN ou OFFLINE.
function Test-CloudOnlyPlaceholder([string]$path) {
  $attrs = [int](Get-Item -LiteralPath $path -Force).Attributes
  return (($attrs -band (0x400000 -bor 0x40000 -bor 0x1000)) -ne 0)
}

function Test-WithinOneDrive([string]$path) {
  $fullPath = [System.IO.Path]::GetFullPath($path)
  foreach ($root in Get-OneDriveRoots) {
    $prefix = [System.IO.Path]::GetFullPath($root).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if ($fullPath.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  }
  return $false
}

# Recalculé à chaque appel : copie de travail, présence dans le dossier OneDrive, état de montage.
function Get-PstState([string]$fileName) {
  $pstPath = Get-WorkingPstPath $fileName
  $workingCopy = Test-Path -LiteralPath $pstPath -PathType Leaf
  $sourcePath = Find-PstFile $fileName
  $sourceExists = [bool]$sourcePath
  $sourceLocal = $sourceExists -and -not (Test-CloudOnlyPlaceholder $sourcePath)
  $inOneDrive = $sourceExists -and (Test-WithinOneDrive $sourcePath)
  $onedriveLocal = $inOneDrive -and -not (Test-CloudOnlyPlaceholder $sourcePath)
  $state = @{
    ok = $true; fileName = $fileName
    exists = ($workingCopy -or $sourceExists)
    local = ($workingCopy -or $sourceLocal)
    workingCopy = $workingCopy
    onedriveLocal = $onedriveLocal
    # Un fichier visible dans le dossier de synchronisation existe dans le cloud ; son absence ne prouve rien.
    cloud = $(if ($inOneDrive) { $true } else { $null })
    mounted = $false
    outlookRunning = $false
  }
  $outlook = Get-RunningOutlook
  if ($outlook) {
    $state.outlookRunning = $true
    if ($workingCopy) { $state.mounted = (Test-PstMounted $outlook $pstPath) }
  }
  return $state
}

# Arborescence du PST monté (dossiers et nombre d'éléments) : lecture locale seule, aucun contenu de message.
function Get-PstTree([string]$pstPath) {
  $outlook = Get-RunningOutlook
  if (-not $outlook) { throw [System.InvalidOperationException]::new("OUTLOOK_NOT_RUNNING") }
  foreach ($store in $outlook.GetNamespace("MAPI").Stores) {
    $storePath = $null
    try { $storePath = $store.FilePath } catch { continue }
    if ($storePath -and ($storePath -ieq $pstPath)) {
      return @{ ok = $true; path = $pstPath; root = (ConvertTo-FolderNode $store.GetRootFolder() 0) }
    }
  }
  return @{ ok = $false; code = "NOT_MOUNTED"; error = "PST non monté dans Outlook" }
}

function ConvertTo-FolderNode($folder, [int]$depth) {
  $node = @{ name = [string]$folder.Name; count = [int]$folder.Items.Count; children = @() }
  if ($depth -ge 6) { return $node }
  foreach ($child in $folder.Folders) {
    $node.children += ConvertTo-FolderNode $child ($depth + 1)
  }
  return $node
}

function Get-MountedStore([string]$fileName) {
  $path = Get-WorkingPstPath $fileName
  $outlook = Get-RunningOutlook
  if (-not $outlook) { throw [System.InvalidOperationException]::new("OUTLOOK_NOT_RUNNING") }
  $namespace = $outlook.GetNamespace("MAPI")
  foreach ($store in $namespace.Stores) {
    $storePath = $null
    try { $storePath = $store.FilePath } catch { continue }
    if ($storePath -and ($storePath -ieq $path)) {
      return @{ namespace = $namespace; store = $store }
    }
  }
  throw "PST non monté dans Outlook : $fileName"
}

function Get-ArchiveFolder($archive, [string]$folderId) {
  if ([string]::IsNullOrWhiteSpace($folderId)) {
    return $archive.store.GetRootFolder()
  }
  $folder = $archive.namespace.GetFolderFromID($folderId, $archive.store.StoreID)
  if (-not $folder -or $folder.StoreID -ne $archive.store.StoreID) {
    throw "Dossier introuvable dans cette archive."
  }
  return $folder
}

function Get-ArchiveMail($archive, [string]$itemId) {
  if ([string]::IsNullOrWhiteSpace($itemId)) { throw "Identifiant de message manquant." }
  $item = $archive.namespace.GetItemFromID($itemId, $archive.store.StoreID)
  if (-not $item -or $item.Class -ne 43 -or $item.Parent.StoreID -ne $archive.store.StoreID) {
    throw "Message introuvable dans cette archive."
  }
  return $item
}

function Get-ArchiveFolderContents($archive, [string]$folderId, [int]$offset) {
  if ($offset -lt 0) { throw "Position de page invalide." }
  $folder = Get-ArchiveFolder $archive $folderId
  $folders = @()
  foreach ($child in $folder.Folders) {
    $folders += @{ id = [string]$child.EntryID; name = [string]$child.Name; count = [int]$child.Items.Count }
  }
  $items = $folder.Items
  $total = [int]$items.Count
  $messages = @()
  $end = [Math]::Min([long]$total, [long]$offset + 50)
  for ($i = [long]$offset + 1; $i -le $end; $i++) {
    $item = $items.Item([int]$i)
    if ($item.Class -eq 43) {
      $messages += @{
        id = [string]$item.EntryID
        subject = [string]$item.Subject
        sender = [string]$item.SenderName
        date = ([datetime]$item.ReceivedTime).ToString("s")
      }
    }
  }
  return @{
    ok = $true; folder = @{ id = [string]$folder.EntryID; name = [string]$folder.Name }
    folders = $folders; messages = $messages; offset = $offset
    nextOffset = $(if ($end -lt $total) { [int]$end } else { $null }); totalItems = $total
  }
}

function Get-ArchiveAddress($entry, [string]$name) {
  $address = ""
  try {
    if ($entry -and $entry.Type -eq "SMTP") {
      $address = [string]$entry.Address
    } elseif ($entry -and $entry.Type -eq "EX") {
      $exchangeUser = $entry.GetExchangeUser()
      if ($exchangeUser) {
        $address = [string]$exchangeUser.PrimarySmtpAddress
      } else {
        $exchangeList = $entry.GetExchangeDistributionList()
        if ($exchangeList) { $address = [string]$exchangeList.PrimarySmtpAddress }
      }
    }
  } catch {
    Write-Warning "Adresse SMTP non résolue pour un participant du message."
  }
  return @{ name = $name; address = $address }
}

function Get-ArchiveMessage($archive, [string]$itemId) {
  $item = Get-ArchiveMail $archive $itemId
  $attachments = @()
  for ($i = 1; $i -le $item.Attachments.Count; $i++) {
    $attachment = $item.Attachments.Item($i)
    $attachments += @{ index = $i; name = [string]$attachment.FileName; size = [long]$attachment.Size }
  }
  $senderAddress = Get-ArchiveAddress $item.Sender ([string]$item.SenderName)
  $recipients = @()
  foreach ($recipient in $item.Recipients) {
    if ($recipient.Type -in @(1, 2)) {
      $entry = $null
      try { $entry = $recipient.AddressEntry } catch { Write-Warning "Adresse du destinataire indisponible." }
      $recipients += @{ type = [int]$recipient.Type; contact = (Get-ArchiveAddress $entry ([string]$recipient.Name)) }
    }
  }
  return @{
    ok = $true; subject = [string]$item.Subject; sender = [string]$item.SenderName
    to = [string]$item.To; cc = [string]$item.CC; bcc = [string]$item.BCC
    senderContact = $senderAddress; recipients = $recipients
    date = ([datetime]$item.ReceivedTime).ToString("s")
    body = [string]$item.Body; attachments = $attachments
  }
}

function Write-ArchiveAttachment($res, $archive, [string]$itemId, [int]$index) {
  $item = Get-ArchiveMail $archive $itemId
  if ($index -lt 1 -or $index -gt $item.Attachments.Count) { throw "Pièce jointe introuvable." }
  $attachment = $item.Attachments.Item($index)
  $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("MesArchives-" + [guid]::NewGuid().ToString() + ".tmp")
  try {
    $attachment.SaveAsFile($tempPath)
    $stream = [System.IO.File]::OpenRead($tempPath)
    try {
      $res.StatusCode = 200
      $res.ContentType = "application/octet-stream"
      $res.Headers["Content-Disposition"] = 'attachment; filename="attachment.bin"'
      $res.Headers["X-Content-Type-Options"] = "nosniff"
      $res.Headers["Cache-Control"] = "no-store"
      Add-CorsHeaders $res
      $res.ContentLength64 = $stream.Length
      $stream.CopyTo($res.OutputStream)
    } finally { $stream.Dispose() }
    $res.Close()
  } finally {
    if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force }
  }
}

function Test-ViewerRequest($req) {
  $origin = $req.Headers["Origin"]
  $localOrigin = ([uri]$ListenPrefix).GetLeftPart([System.UriPartial]::Authority)
  if ($origin -ne $localOrigin -and $origin -ne "https://dellasiegaps.github.io") { return $false }
  return $req.ContentType -match '^application/json(?:\s*;|$)'
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add($ListenPrefix)
$listener.Start()
Write-Host "PST agent listening on $ListenPrefix"

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $req = $ctx.Request
  $res = $ctx.Response
  try {
    if ($req.HttpMethod -eq "GET" -and $req.Url.AbsolutePath -eq "/view") {
      $viewerFile = Join-Path $PSScriptRoot "local-pst-view.html"
      if (-not (Test-Path -LiteralPath $viewerFile -PathType Leaf)) { throw "Le fichier du visualiseur local est introuvable : $viewerFile" }
      $bytes = [System.IO.File]::ReadAllBytes($viewerFile)
      $res.StatusCode = 200
      $res.ContentType = "text/html; charset=utf-8"
      $res.Headers["Cache-Control"] = "no-store"
      $res.Headers["Content-Security-Policy"] = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; frame-ancestors 'self' https://dellasiegaps.github.io; base-uri 'none'; form-action 'none'"
      $res.Headers["X-Content-Type-Options"] = "nosniff"
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
      $res.Close()
      continue
    }
    if ($req.HttpMethod -eq "OPTIONS") { Add-CorsHeaders $res; $res.StatusCode = 204; $res.Close(); continue }
    if ($req.HttpMethod -ne "POST") { Write-Json $res 405 @{ ok = $false; error = "Method not allowed" }; continue }
    if ($req.Url.AbsolutePath -in @("/folder", "/message", "/attachment") -and -not (Test-ViewerRequest $req)) {
      Write-Json $res 403 @{ ok = $false; error = "Origine ou type de requête refusé." }
      continue
    }

    $reader = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
    $bodyText = $reader.ReadToEnd()
    $body = @{}
    if ($bodyText) { $body = $bodyText | ConvertFrom-Json }

    switch ($req.Url.AbsolutePath) {
      "/mount-pst" {
        $fileName = [string]$body.fileName
        $pstPath = Get-OrCreateWorkingPst $fileName
        if (-not $pstPath) { Write-Json $res 404 @{ ok = $false; code = "PST_NOT_FOUND"; error = "PST not found locally"; fileName = $fileName }; continue }
        Write-Json $res 200 (Mount-Pst $pstPath ([string]$body.displayName))
      }
      "/unmount-pst" {
        $fileName = [string]$body.fileName
        $pstPath = Get-WorkingPstPath $fileName
        if (-not (Test-Path -LiteralPath $pstPath -PathType Leaf)) {
          Write-Json $res 200 @{ ok = $true; alreadyUnmounted = $true; fileName = $fileName }
          continue
        }
        Write-Json $res 200 (Unmount-Pst $pstPath)
      }
      "/state" { Write-Json $res 200 (Get-PstState ([string]$body.fileName)) }
      "/tree" {
        $pstPath = Get-WorkingPstPath ([string]$body.fileName)
        $tree = Get-PstTree $pstPath
        if (-not $tree.ok) { Write-Json $res 409 $tree; continue }
        Write-Json $res 200 $tree
      }
      "/folder" {
        $archive = Get-MountedStore ([string]$body.fileName)
        $offset = 0
        if ($null -ne $body.offset -and -not [int]::TryParse([string]$body.offset, [ref]$offset)) { throw "Position de page invalide." }
        Write-Json $res 200 (Get-ArchiveFolderContents $archive ([string]$body.folderId) $offset)
      }
      "/message" {
        $archive = Get-MountedStore ([string]$body.fileName)
        Write-Json $res 200 (Get-ArchiveMessage $archive ([string]$body.itemId))
      }
      "/attachment" {
        $archive = Get-MountedStore ([string]$body.fileName)
        $index = 0
        if (-not [int]::TryParse([string]$body.index, [ref]$index)) { throw "Index de pièce jointe invalide." }
        Write-ArchiveAttachment $res $archive ([string]$body.itemId) $index
      }
      "/status" { Write-Json $res 200 @{ ok = $true; running = $true; roots = @(Get-OneDriveRoots) } }
      default { Write-Json $res 404 @{ ok = $false; error = "Route not found" } }
    }
  } catch {
    if ($_.Exception.Message -eq "OUTLOOK_NOT_RUNNING") {
      Write-Json $res 409 @{ ok = $false; code = "OUTLOOK_NOT_RUNNING"; error = "Outlook n'est pas ouvert : lancez Outlook puis réessayez." }
      continue
    }
    Write-Json $res 500 @{ ok = $false; error = $_.Exception.Message }
  }
}
