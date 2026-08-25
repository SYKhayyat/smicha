param(
    [string]$OutDir = "$env:TEMP\opencode\smicha-verify\ahheasid",
    [string]$Api = 'https://wiki.jewishbooks.org.il/mediawiki/api.php'
)
$ErrorActionPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$prefix = [Uri]::EscapeDataString('ערוך השולחן העתיד/')

Write-Host 'enumerating all pages...'
$titles = New-Object System.Collections.Generic.List[string]
$apfrom = $null
while ($true) {
    $u = "$Api" + "?action=query&list=allpages&apprefix=$prefix&aplimit=500&format=json&formatversion=2"
    if ($apfrom) { $u += "&apfrom=" + [Uri]::EscapeDataString($apfrom) }
    try {
        $r = Invoke-WebRequest -Uri $u -TimeoutSec 60 -UseBasicParsing
        $j = $r.Content | ConvertFrom-Json
        foreach ($p in $j.query.allpages) { $titles.Add($p.title) }
        if ($j.continue -and $j.continue.apcontinue) { $apfrom = $j.continue.apcontinue } else { break }
    } catch {
        Write-Host ("enumerate error: " + $_.Exception.Message)
        Start-Sleep -Seconds 3
    }
}
Write-Output ("total pages found: " + $titles.Count)

$done = 0
for ($i = 0; $i -lt $titles.Count; $i += 10) {
    $batch = @($titles[$i..([Math]::Min($i+9, $titles.Count-1))])
    $tq = ($batch | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '|'
    $u = "$Api" + "?action=query&prop=revisions&rvprop=content&format=json&formatversion=2&titles=" + $tq
    try {
        $r = Invoke-WebRequest -Uri $u -TimeoutSec 90 -UseBasicParsing
        $j = $r.Content | ConvertFrom-Json
        foreach ($pg in $j.query.pages) {
            if ($pg.missing) { continue }
            $content = $null
            if ($pg.revisions -and $pg.revisions.Count -gt 0) {
                $rev = $pg.revisions[0]
                $content = $rev.content
                if (-not $content) { $content = $rev.PSObject.Properties['*'].Value }
            }
            if (-not $content) { continue }
            $safe = ($pg.title -replace '^[^/]+/', '' -replace '/', '_') + '.txt'
            $dest = Join-Path $OutDir $safe
            [IO.File]::WriteAllText($dest, $content, (New-Object System.Text.UTF8Encoding $true))
            $done++
        }
        Write-Host ("batch {0}: cumulative {1}/{2}" -f ([Math]::Floor($i/10)+1), $done, $titles.Count)
    } catch {
        Write-Host ("fetch error at batch starting $i : " + $_.Exception.Message)
    }
    Start-Sleep -Milliseconds 300
}
Write-Output ("saved files: " + $done)
Write-Output ("corpus files now: " + @(Get-ChildItem -LiteralPath $OutDir -Filter '*.txt').Count)
