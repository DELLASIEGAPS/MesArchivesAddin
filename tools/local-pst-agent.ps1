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
    try {
      if ($store.FilePath -and ($store.FilePath -ieq $pstPath)) {
        $ns.RemoveStore($store.GetRootFolder())
        return @{ ok = $true; unmounted = $true; path = $pstPath }
      }
    } catch {}
  }
  return @{ ok = $true; alreadyUnmounted = $true; path = $pstPath }
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
        $pstPath = Find-PstFile $fileName
        if (-not $pstPath) { Write-Json $res 404 @{ ok = $false; code = "PST_NOT_FOUND"; error = "PST not found locally"; fileName = $fileName }; continue }
        Write-Json $res 200 (Mount-Pst $pstPath)
      }
      "/unmount-pst" {
        $fileName = [string]$body.fileName
        $pstPath = Find-PstFile $fileName
        if (-not $pstPath) { Write-Json $res 404 @{ ok = $false; code = "PST_NOT_FOUND"; error = "PST not found locally"; fileName = $fileName }; continue }
        Write-Json $res 200 (Unmount-Pst $pstPath)
      }
      "/status" { Write-Json $res 200 @{ ok = $true; running = $true; roots = @(Get-OneDriveRoots) } }
      default { Write-Json $res 404 @{ ok = $false; error = "Route not found" } }
    }
  } catch {
    Write-Json $res 500 @{ ok = $false; error = $_.Exception.Message }
  }
}
