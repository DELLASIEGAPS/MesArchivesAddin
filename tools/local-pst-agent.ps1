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
  $json = ($obj | ConvertTo-Json -Depth 6 -Compress)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
  $res.Close()
}

function Mount-Pst([string]$pstPath) {
  $outlook = New-Object -ComObject Outlook.Application
  $ns = $outlook.GetNamespace("MAPI")
  foreach ($store in $ns.Stores) {
    try { if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) { return @{ ok = $true; alreadyMounted = $true; path = $pstPath } } } catch {}
  }
  $ns.AddStoreEx($pstPath, 3) # olStoreUnicode
  return @{ ok = $true; mounted = $true; path = $pstPath }
}

function Unmount-Pst([string]$pstPath) {
  $outlook = New-Object -ComObject Outlook.Application
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

function Get-PstState([string]$fileName) {
  $pstPath = Get-WorkingPstPath $fileName
  $copyExists = Test-Path -LiteralPath $pstPath -PathType Leaf
  if (-not $copyExists -and -not (Find-PstFile $fileName)) {
    return @{ ok = $true; exists = $false; mounted = $false; fileName = $fileName }
  }
  if (-not $copyExists) {
    return @{ ok = $true; exists = $true; mounted = $false; fileName = $fileName }
  }
  $outlook = New-Object -ComObject Outlook.Application
  $mounted = $false
  foreach ($store in $outlook.GetNamespace("MAPI").Stores) {
    try { if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) { $mounted = $true; break } } catch {}
  }
  return @{ ok = $true; exists = $true; mounted = $mounted; path = $pstPath }
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
        Write-Json $res 200 (Mount-Pst $pstPath)
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
      "/status" { Write-Json $res 200 @{ ok = $true; running = $true; roots = @(Get-OneDriveRoots) } }
      default { Write-Json $res 404 @{ ok = $false; error = "Route not found" } }
    }
  } catch {
    Write-Json $res 500 @{ ok = $false; error = $_.Exception.Message }
  }
}
