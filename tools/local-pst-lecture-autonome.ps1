param(
  [Parameter(Mandatory = $true)][string]$PstSourceDirectory,
  [Parameter(Mandatory = $true)][string]$LibraryPath,
  [int]$Port = 45732
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.Major -lt 7) { throw "Ce lecteur demande PowerShell 7." }
if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port invalide." }
$source = (Resolve-Path -LiteralPath $PstSourceDirectory).ProviderPath
if (-not (Test-Path -LiteralPath $source -PathType Container)) { throw "Dossier source introuvable." }
Add-Type -Path (Resolve-Path -LiteralPath $LibraryPath).ProviderPath
if (-not [XstReader.XstFolder].GetProperty('HasErrorInFolders') -or
    -not [XstReader.XstFolder].GetProperty('HasErrorInMessages') -or
    -not [XstReader.XstMessage].GetProperty('HasErrorInAttachments') -or
    -not [XstReader.XstFile].GetField('MesArchivesSafeCloseVersion') -or
    [XstReader.XstFile].GetField('MesArchivesSafeCloseVersion').GetRawConstantValue() -ne '1') {
  throw "Bibliothèque PST incompatible : utilisez la version corrigée décrite dans la documentation."
}
$files = @('ADR_2021.pst', 'ADR_2022.pst', 'ADR_2023.pst', 'ADR.pst')
$loadedArchives = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
$token = [Convert]::ToHexString([System.Security.Cryptography.RandomNumberGenerator]::GetBytes(32))
$authority = "127.0.0.1:$Port"
$origin = "http://$authority"

function Write-Json($response, [int]$status, $value) {
  $bytes = [System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $value -Depth 20 -Compress))
  $response.StatusCode = $status
  $response.ContentType = 'application/json; charset=utf-8'
  $response.Headers['Cache-Control'] = 'no-store'
  $response.Headers['X-Content-Type-Options'] = 'nosniff'
  $response.ContentLength64 = $bytes.Length
  $response.OutputStream.Write($bytes, 0, $bytes.Length)
  $response.Close()
}

function Get-ArchivePath([string]$name) {
  if ($name -cnotin $files) { throw 'Nom de fichier PST non autorisé.' }
  $path = Join-Path $source $name
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Archive absente : $name" }
  $attributes = [int](Get-Item -LiteralPath $path -Force).Attributes
  if (($attributes -band (0x400000 -bor 0x40000 -bor 0x1000)) -ne 0) {
    throw "Archive non disponible localement : $name"
  }
  return $path
}

function Get-Folders($folder) {
  $result = @($folder.Folders | ForEach-Object { $_ })
  if ($folder.HasErrorInFolders) { throw "Lecture des dossiers impossible : $($folder.ErrorInFolders)" }
  return ,$result
}

function Get-Messages($folder) {
  $result = @($folder.Messages | ForEach-Object { $_ })
  if ($folder.HasErrorInMessages) { throw "Lecture des messages impossible : $($folder.ErrorInMessages)" }
  return ,$result
}

function Find-Folder($file, [string]$id) {
  $folder = $file.RootFolder
  if (-not $folder) { throw "Racine PST illisible." }
  if ($id -eq '') { return $folder }
  if ($id -notmatch '^\d+(?:\.\d+)*$' -or $id.Length -gt 300) { throw 'Identifiant de dossier invalide.' }
  foreach ($part in $id.Split('.')) {
    $number = 0
    if (-not [int]::TryParse($part, [ref]$number)) { throw 'Identifiant de dossier invalide.' }
    $children = Get-Folders $folder
    if ($number -ge $children.Count) { throw 'Dossier introuvable.' }
    $folder = $children[$number]
  }
  return $folder
}

function Find-Message($file, [string]$id, [string]$version) {
  if (-not $id.StartsWith("$version|", [System.StringComparison]::Ordinal)) { throw "Archive modifiée : rechargez le dossier." }
  $id = $id.Substring($version.Length + 1)
  if ($id -notmatch '^(\d+(?:\.\d+)*)?:(\d+)$') { throw 'Identifiant de message invalide.' }
  $folderId = $Matches[1]
  $position = 0
  if (-not [int]::TryParse($Matches[2], [ref]$position)) { throw 'Identifiant de message invalide.' }
  $folder = Find-Folder $file $folderId
  $items = Get-Messages $folder
  if ($position -ge $items.Count) { throw 'Message introuvable.' }
  return $items[$position]
}

function Get-Contact($contact, [string]$fallback) {
  if (-not $contact) { return @{ name = $fallback; address = '' } }
  return @{ name = [string]$contact.DisplayName; address = [string]$contact.Address }
}

function Invoke-Archive($response, [string]$route, $body) {
  $name = [string]$body.fileName
  if ($route -eq '/state') {
    if ($name -cnotin $files) { throw 'Nom de fichier PST non autorisé.' }
    $exists = Test-Path -LiteralPath (Join-Path $source $name) -PathType Leaf
    $available = $false
    if ($exists) {
      $attributes = [int](Get-Item -LiteralPath (Join-Path $source $name) -Force).Attributes
      $available = (($attributes -band (0x400000 -bor 0x40000 -bor 0x1000)) -eq 0)
    }
    Write-Json $response 200 @{ ok = $true; mode = 'standalone'; available = $available; loaded = $loadedArchives.Contains($name) }
    return
  }
  if ($route -eq '/unload') {
    if ($name -cnotin $files) { throw 'Nom de fichier PST non autorisé.' }
    $null = $loadedArchives.Remove($name)
    Write-Json $response 200 @{ ok = $true; loaded = $false }
    return
  }
  $path = Get-ArchivePath $name
  if ($route -eq '/load') {
    $file = [XstReader.XstFile]::new($path)
    try {
      if (-not $file.RootFolder) { throw 'Racine PST illisible.' }
    } finally { $file.Dispose() }
    $null = $loadedArchives.Add($name)
    Write-Json $response 200 @{ ok = $true; loaded = $true }
    return
  }
  if (-not $loadedArchives.Contains($name)) { throw "Archive non chargée : $name" }
  $metadata = Get-Item -LiteralPath $path
  $version = "$name-$($metadata.Length)x$($metadata.LastWriteTimeUtc.Ticks)"
  $file = [XstReader.XstFile]::new($path)
  try {
    if ($route -eq '/folder') {
      $folderId = [string]$body.folderId
      if ($folderId) {
        if (-not $folderId.StartsWith("$version|", [System.StringComparison]::Ordinal)) { throw "Archive modifiée : rechargez le dossier." }
        $folderId = $folderId.Substring($version.Length + 1)
      }
      $offset = 0
      if ($null -ne $body.offset -and (-not [int]::TryParse([string]$body.offset, [ref]$offset) -or $offset -lt 0)) { throw 'Position invalide.' }
      $folder = Find-Folder $file $folderId
      $children = Get-Folders $folder
      $items = Get-Messages $folder
      $folders = @()
      for ($i = 0; $i -lt $children.Count; $i++) {
        $id = if ($folderId) { "$folderId.$i" } else { "$i" }
        $folders += @{ id = "$version|$id"; name = [string]$children[$i].DisplayName; count = [int]$children[$i].ContentCount }
      }
      $messages = @()
      $end = [Math]::Min([long]$items.Count, [long]$offset + 50)
      for ($i = $offset; $i -lt $end; $i++) {
        $item = $items[$i]
        $messages += @{ id = "$version|${folderId}:$i"; subject = [string]$item.Subject; sender = [string]$item.From; date = $(if ($item.Date) { $item.Date.ToString('s') } else { '' }) }
      }
      Write-Json $response 200 @{ ok = $true; folder = @{ id = "$version|$folderId"; name = [string]$folder.DisplayName }; folders = $folders; messages = $messages; offset = $offset; nextOffset = $(if ($end -lt $items.Count) { [int]$end } else { $null }); totalItems = $items.Count }
      return
    }
    $message = Find-Message $file ([string]$body.itemId) $version
    if ($route -eq '/message') {
      $attachments = @($message.Attachments | ForEach-Object { $_ })
      if ($message.HasErrorInAttachments) { throw "Lecture des pièces jointes impossible : $($message.ErrorInAttachments)" }
      $attachmentInfo = @()
      for ($i = 0; $i -lt $attachments.Count; $i++) {
        $attachmentInfo += @{ index = $i + 1; name = $(if ($attachments[$i].FileNameForSaving) { [string]$attachments[$i].FileNameForSaving } else { "piece-jointe-$($i + 1)" }); size = $null }
      }
      $recipients = @()
      foreach ($r in $message.Recipients.To) { $recipients += @{ type = 1; contact = (Get-Contact $r '') } }
      foreach ($r in $message.Recipients.Cc) { $recipients += @{ type = 2; contact = (Get-Contact $r '') } }
      $content = $message.Body
      $text = if (-not $content) { '' } elseif ($content.Format -eq [XstReader.XstMessageBodyFormat]::PlainText) { [string]$content.Text } else { "Format $($content.Format) non pris en charge en texte brut." }
      Write-Json $response 200 @{ ok = $true; subject = [string]$message.Subject; sender = [string]$message.From; to = [string]$message.To; cc = [string]$message.Cc; bcc = [string]$message.Bcc; senderContact = (Get-Contact $message.Recipients.Sender ([string]$message.From)); recipients = $recipients; date = $(if ($message.Date) { $message.Date.ToString('s') } else { '' }); body = $text; attachments = $attachmentInfo }
      return
    }
    if ($route -eq '/attachment') {
      $index = 0
      if (-not [int]::TryParse([string]$body.index, [ref]$index)) { throw 'Index de pièce jointe invalide.' }
      $attachments = @($message.Attachments | ForEach-Object { $_ })
      if ($message.HasErrorInAttachments) { throw "Lecture des pièces jointes impossible : $($message.ErrorInAttachments)" }
      if ($index -lt 1 -or $index -gt $attachments.Count) { throw 'Pièce jointe introuvable.' }
      if (-not $attachments[$index - 1].IsFile) { throw "Ce type de pièce jointe n'est pas pris en charge." }
      $buffer = [System.IO.MemoryStream]::new()
      try {
        $attachments[$index - 1].SaveToStream($buffer)
        $response.StatusCode = 200
        $response.ContentType = 'application/octet-stream'
        $response.Headers['Content-Disposition'] = 'attachment; filename="attachment.bin"'
        $response.Headers['X-Content-Type-Options'] = 'nosniff'
        $response.Headers['Cache-Control'] = 'no-store'
        $response.ContentLength64 = $buffer.Length
        $buffer.Position = 0
        $buffer.CopyTo($response.OutputStream)
        $response.Close()
      } finally { $buffer.Dispose() }
      return
    }
    throw 'Route inconnue.'
  } finally { $file.Dispose() }
}

$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("$origin/")
$listener.Start()
Write-Host "Lecteur autonome : $origin/view"
try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
    $response.KeepAlive = $false
    try {
      if ($request.Headers['Host'] -ne $authority) { Write-Json $response 403 @{ ok = $false; error = 'Hôte refusé.' }; continue }
      if ($request.HttpMethod -eq 'GET' -and $request.Url.AbsolutePath -eq '/view') {
        $html = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'local-pst-view.html'))
        $html = $html.Replace('<head>', "<head>`n<script>window.MES_ARCHIVES_STANDALONE={token:'$token'};</script>")
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        $response.ContentType = 'text/html; charset=utf-8'
        $response.Headers['Cache-Control'] = 'no-store'
        $response.Headers['Content-Security-Policy'] = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; frame-ancestors 'self' https://dellasiegaps.github.io; base-uri 'none'; form-action 'none'"
        $response.Headers['X-Content-Type-Options'] = 'nosniff'
        $response.ContentLength64 = $bytes.Length
        $response.OutputStream.Write($bytes, 0, $bytes.Length)
        $response.Close()
        continue
      }
      if ($request.HttpMethod -ne 'POST' -or $request.Url.AbsolutePath -notin @('/state', '/load', '/unload', '/folder', '/message', '/attachment')) { Write-Json $response 404 @{ ok = $false; error = 'Route inconnue.' }; continue }
      if ($request.Headers['Origin'] -ne $origin -or $request.Headers['X-MesArchives-Token'] -cne $token -or $request.ContentType -notmatch '^application/json(?:\s*;|$)') { Write-Json $response 403 @{ ok = $false; error = 'Requête refusée.' }; continue }
      if ($request.ContentLength64 -lt 0 -or $request.ContentLength64 -gt 16384) { Write-Json $response 413 @{ ok = $false; error = 'Requête trop grande.' }; continue }
      $reader = [System.IO.StreamReader]::new($request.InputStream, $request.ContentEncoding)
      try { $body = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
      Invoke-Archive $response $request.Url.AbsolutePath $body
    } catch {
      Write-Warning $_.Exception.Message
      if ($response.OutputStream.CanWrite) { Write-Json $response 500 @{ ok = $false; error = $_.Exception.Message } }
    }
  }
} finally { $listener.Stop(); $listener.Close() }
