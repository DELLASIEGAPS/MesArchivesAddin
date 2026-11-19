param(
  [string]$ListenPrefix = "http://127.0.0.1:45731/"
)

$ErrorActionPreference = "Stop"

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

# Recalculé à chaque appel : copie de travail, présence dans le dossier OneDrive, état de montage.
function Get-PstState([string]$fileName) {
  $pstPath = Get-WorkingPstPath $fileName
  $workingCopy = Test-Path -LiteralPath $pstPath -PathType Leaf
  $sourcePath = Find-PstFile $fileName
  $inOneDrive = [bool]$sourcePath
  $onedriveLocal = $inOneDrive -and -not (Test-CloudOnlyPlaceholder $sourcePath)
  $state = @{
    ok = $true; fileName = $fileName
    exists = ($workingCopy -or $inOneDrive)
    local = ($workingCopy -or $onedriveLocal)
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
    try {
      if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) {
        return @{ ok = $true; path = $pstPath; root = (ConvertTo-FolderNode $store.GetRootFolder() 0) }
      }
    } catch {}
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

$viewPage = @'
<!doctype html>
<html lang="fr"><meta charset="utf-8"><title>MesArchives - visualiseur local</title>
<style>body{font-family:Segoe UI,sans-serif;margin:16px}ul{list-style:none;padding-left:18px}li{margin:2px 0}.n{color:#555}</style>
<h2>Archives montées (local)</h2>
<p>Aucune donnée ne quitte ce poste.</p>
<label>Fichier PST : <input id="f" value="ADR_2022.pst"></label>
<button id="go">Afficher</button>
<div id="out"></div>
<script>
function render(node, parent) {
  const li = document.createElement("li");
  const name = document.createElement("span");
  name.textContent = node.name;
  const c = document.createElement("span");
  c.className = "n";
  c.textContent = " (" + node.count + ")";
  li.append(name, c);
  if (node.children && node.children.length) {
    const ul = document.createElement("ul");
    node.children.forEach((ch) => render(ch, ul));
    li.appendChild(ul);
  }
  parent.appendChild(li);
}
document.getElementById("go").addEventListener("click", async () => {
  const out = document.getElementById("out");
  out.textContent = "Chargement…";
  const r = await fetch("/tree", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fileName: document.getElementById("f").value }) });
  const d = await r.json();
  out.textContent = "";
  if (!r.ok || d.ok !== true) { out.textContent = d.error || "Erreur"; return; }
  const ul = document.createElement("ul");
  render(d.root, ul);
  out.appendChild(ul);
});
</script></html>
'@

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
      $bytes = [System.Text.Encoding]::UTF8.GetBytes($viewPage)
      $res.StatusCode = 200
      $res.ContentType = "text/html; charset=utf-8"
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
      $res.Close()
      continue
    }
    if ($req.HttpMethod -eq "OPTIONS") { Add-CorsHeaders $res; $res.StatusCode = 204; $res.Close(); continue }
    if ($req.HttpMethod -ne "POST") { Write-Json $res 405 @{ ok = $false; error = "Method not allowed" }; continue }

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
