param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('extract', 'fetch', 'report')]
    [string]$Mode,
    [string]$RepoRoot,
    [string]$WorkDir
)

if (-not $RepoRoot) {
    $scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $RepoRoot = Split-Path -Parent $scriptPath
}
if (-not $WorkDir) { $WorkDir = Join-Path $env:TEMP 'opencode\smicha-verify' }

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = New-Object System.Text.UTF8Encoding $false
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ManifestPath = Join-Path $WorkDir 'citations.json'
$IndexMapPath = Join-Path $WorkDir 'sefaria-index.json'
$ResultsPath = Join-Path $WorkDir 'results.json'

function Initialize-WorkDir {
    if (-not (Test-Path -LiteralPath $WorkDir)) {
        New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
    }
}

function ConvertFrom-HebrewNumeral {
    param([string]$Text)
    $values = @{
        [char]'א' = 1; [char]'ב' = 2; [char]'ג' = 3; [char]'ד' = 4; [char]'ה' = 5
        [char]'ו' = 6; [char]'ז' = 7; [char]'ח' = 8; [char]'ט' = 9
        [char]'י' = 10; [char]'כ' = 20; [char]'ל' = 30; [char]'מ' = 40; [char]'נ' = 50
        [char]'ס' = 60; [char]'ע' = 70; [char]'פ' = 80; [char]'צ' = 90
        [char]'ק' = 100; [char]'ר' = 200; [char]'ש' = 300; [char]'ת' = 400
    }
    $total = 0
    foreach ($ch in $Text.ToCharArray()) {
        if ($values.ContainsKey($ch)) { $total += $values[$ch] }
    }
    return $total
}

function Get-OrgFiles {
    Get-ChildItem -LiteralPath $RepoRoot -Recurse -Filter '*.org' |
        Where-Object { $_.FullName -notlike "$RepoRoot\tools*" }
}

function Normalize-HebrewPunctuation {
    param([string]$Text)
    return $Text.Replace([char]0x05F4, '"').Replace([char]0x05F3, "'")
}

function Find-Citations {
    param([string]$RawText, [string]$RelPath)

    $found = @()
    $lines = $RawText -split "`r?`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = Normalize-HebrewPunctuation $lines[$i]
        $lineNo = $i + 1

        $authorPattern = '(שו"ע|שולחן ערוך|רמ"א|הרמ"א|מ"ב|משנה ברורה|מג"א|מגן אברהם|ערוך השלחן|ערוה"ש|ט"ז|ש"ך|פמ"ג|אליה רבה|פרי מגדים|פרי חדש|כף החיים|בן איש חי|יביע אומר|ילקוט יוסף|חזון עובדיה|אגרות משה|מנחת יצחק|שבט הלוי)'
        $saMatches = [regex]::Matches($line, $authorPattern)
        foreach ($m in $saMatches) {
            $windowEnd = [Math]::Min($line.Length, $m.Index + 160)
            $window = $line.Substring($m.Index, $windowEnd - $m.Index)
            $simanMatch = [regex]::Match($window, "סי['′]?\s*([\p{IsHebrew}""']{1,8})")
            if (-not $simanMatch.Success) {
                $simanMatch = [regex]::Match($window, 'סימן\s+([\p{IsHebrew}"]{1,8})')
            }
            if (-not $simanMatch.Success) { continue }
            $siman = ConvertFrom-HebrewNumeral $simanMatch.Groups[1].Value
            if ($siman -le 0 -or $siman -gt 1200) { continue }

            $seif = $null
            $seifKind = $null
            $seifMatch = [regex]::Match($window, 'ס"(?:ק)?["′]?[:：]?\s*([\p{IsHebrew}"]{1,4})')
            if ($seifMatch.Success) { $seifKind = 'seif-or-sk' }
            if (-not $seifMatch.Success) {
                $seifMatch = [regex]::Match($window, 'סעיף\s+([\p{IsHebrew}"]{1,4})')
                if ($seifMatch.Success) { $seifKind = 'seif' }
            }
            if ($seifMatch.Success) {
                $seifVal = ConvertFrom-HebrewNumeral $seifMatch.Groups[1].Value
                if ($seifVal -gt 0 -and $seifVal -lt 300) { $seif = $seifVal } else { $seif = $null; $seifKind = $null }
            }

            $branch = $null
            $ctxStart = [Math]::Max(0, $m.Index - 140)
            $ctxLen = [Math]::Min($line.Length - $ctxStart, 320)
            $ctx = $line.Substring($ctxStart, $ctxLen)
            if ($ctx -match 'או"ח|א"ח|אורח חיים') { $branch = 'O.C.' }
            elseif ($ctx -match 'יו"ד|יורה דעה') { $branch = 'Y.D.' }
            elseif ($ctx -match 'אה"ע|אבן העזר') { $branch = 'E.H.' }
            elseif ($ctx -match 'חו"מ|חושן משפט') { $branch = 'C.M.' }
            if (-not $branch) {
                if ($m.Groups[1].Value -match '^(מ"ב|משנה ברורה|מג"א|מגן אברהם|ערוך השלחן|ערוה"ש|אליה רבה)$') { $branch = 'O.C.' }
            }

            $found += [pscustomobject]@{
                type = 'sa'; author = $m.Groups[1].Value; branch = $branch
                siman = $siman; seif = $seif; seifKind = $seifKind
                file = $RelPath; line = $lineNo; text = $line.Trim()
            }
        }

        $rambamPattern = 'רמב"ם\s+(?:הלכות\s+)?([\p{IsHebrew}"][\p{IsHebrew}" ]*?)\s*(?:פ"?\s*([\p{IsHebrew}"]{1,8}))?(?:\s+ה"?\s*([\p{IsHebrew}"]{1,6}))?(?=\s*[;,.):]|$)'
        $rambamMatch = [regex]::Match($line, $rambamPattern)
        if ($rambamMatch.Success) {
            $perek = $null; $halacha = $null
            if ($rambamMatch.Groups[2].Success) { $perek = ConvertFrom-HebrewNumeral $rambamMatch.Groups[2].Value }
            if ($rambamMatch.Groups[3].Success) { $halacha = ConvertFrom-HebrewNumeral $rambamMatch.Groups[3].Value }
            $found += [pscustomobject]@{
                type = 'rambam'; title = $rambamMatch.Groups[1].Value.Trim()
                perek = $perek; halacha = $halacha
                file = $RelPath; line = $lineNo; text = $line.Trim()
            }
        }

        $gemaraMatches = [regex]::Matches($line, "([\p{IsHebrew}""'][\p{IsHebrew}""' ]{2,18})\s+([\p{IsHebrew}""']{2,5})\s+ע""([אב])")
        foreach ($g in $gemaraMatches) {
            $daf = ConvertFrom-HebrewNumeral $g.Groups[2].Value
            if ($daf -le 1 -or $daf -gt 2000) { continue }
            $found += [pscustomobject]@{
                type = 'gemara'; tractate = $g.Groups[1].Value.Trim()
                daf = $daf; amud = $g.Groups[3].Value
                file = $RelPath; line = $lineNo; text = $line.Trim()
            }
        }
    }
    return ,$found
}

function Get-SefariaMaps {
    if (Test-Path -LiteralPath $IndexMapPath) {
        return Get-Content -LiteralPath $IndexMapPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    Write-Host 'fetching Sefaria index...'
    $idx = Invoke-RestMethod -Uri 'https://www.sefaria.org/api/index/' -TimeoutSec 180
    $tractateMap = @{}
    $rambamMap = @{}
    function Walk-Tree {
        param($nodes, [string]$Mode)
        foreach ($c in $nodes) {
            $m2 = $Mode
            if ($c.category -eq 'Bavli' -or $c.title -in @('Bavli', 'Babylonian Talmud')) { $m2 = 'bavli' }
            elseif ($c.category -eq 'Mishneh Torah' -or $c.title -in @('Mishneh Torah') -or $c.heTitle -in @('משנה תורה', 'יד החזקה')) { $m2 = 'rambam' }

            if ($m2 -eq 'bavli' -and $c.heTitle -and $c.title -and $c.heTitle -notmatch '^סדר' -and $c.title -notmatch '^Seder' -and $c.heTitle -notmatch ' .. ' -and $c.title -notmatch ' on ') {
                $script:tmap[$c.heTitle] = $c.title
            }
            if ($m2 -eq 'rambam' -and $c.heTitle -match 'הלכות' -and $c.heTitle -notmatch ' על ' -and $c.title -notmatch ' on ') {
                $core = ($c.heTitle -split ',' | Where-Object { $_ -match 'הלכות' } | Select-Object -First 1)
                if ($core) { $script:rmap[$core.Trim()] = ($c.title -replace '^Mishneh Torah,\s*', '') }
            }

            if ($c.contents) {
                $childMode = $m2
                if ($c.category -match 'Rishonim|Acharonim|Commentary' -or $c.title -match '^Rashi|^Tosafot|^Rif |^Meiri|^Chiddushei|on ') { $childMode = '' }
                Walk-Tree $c.contents $childMode
            }
        }
    }
    $script:tmap = $tractateMap
    $script:rmap = $rambamMap
    Walk-Tree $idx ''
    $maps = [pscustomobject]@{
        tractates   = $script:tmap
        rambamBooks = $script:rmap
    }
    $maps | ConvertTo-Json -Depth 4 | Out-File -LiteralPath $IndexMapPath -Encoding utf8
    Write-Host ("index maps built: {0} tractates, {1} rambam halachos-books" -f $script:tmap.Count, $script:rmap.Count)
    return $maps
}

function Resolve-Lookup {
    param($MapObject, [string]$Key)
    if ($MapObject -is [hashtable]) {
        if ($MapObject.ContainsKey($Key)) { return $MapObject[$Key] }
        $bare = $Key -replace '^הלכות ', ''
        foreach ($k in $MapObject.Keys) { if ($k -like "*$bare*") { return $MapObject[$k] } }
        return $null
    }
    $prop = $MapObject.PSObject.Properties[$Key]
    if ($prop) { return $prop.Value }
    $bare = $Key -replace '^הלכות ', ''
    $hit = $MapObject.PSObject.Properties | Where-Object { $_.Name -like "*$bare*" } | Select-Object -First 1
    if ($hit) { return $hit.Value }
    return $null
}

switch ($Mode) {

    'extract' {
        Initialize-WorkDir
        $all = @()
        foreach ($f in (Get-OrgFiles)) {
            $rel = $f.FullName.Substring($RepoRoot.Length).TrimStart('\', '/').Replace('\', '/')
            $raw = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
            $all += Find-Citations -RawText $raw -RelPath $rel
        }
        $all | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $ManifestPath -Encoding utf8
        Write-Output ("extracted {0} citations -> {1}" -f $all.Count, $ManifestPath)
        $all | Group-Object type | ForEach-Object { Write-Output ("  {0}: {1}" -f $_.Name, $_.Count) }
        Write-Output ("files scanned: {0}" -f @(Get-OrgFiles).Count)
    }

    'fetch' {
        Initialize-WorkDir
        if (-not (Test-Path -LiteralPath $ManifestPath)) { throw 'run extract first' }
        $citations = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($citations.PSObject.Properties['value']) { $citations = @($citations.value) }
        $maps = Get-SefariaMaps

        $targets = @{}
        foreach ($c in $citations) {
            $key = $null
            if ($c.type -eq 'sa') {
                if (-not $c.branch) { continue }
                $branchSlug = @{
                    'O.C.' = 'Orach%20Chayim'; 'Y.D.' = 'Yoreh%20De%27ah'
                    'E.H.' = 'Even%20HaEzer'; 'C.M.' = 'Choshen%20Mishpat'
                }[$c.branch]
                $key = 'sa|Shulchan%20Arukh%2C%20{0}|{1}' -f $branchSlug, $c.siman
            }
            elseif ($c.type -eq 'rambam') {
                $en = Resolve-Lookup $maps.rambamBooks ('הלכות ' + $c.title)
                if ($en -and $c.perek) {
                    $enEnc = [uri]::EscapeDataString("Mishneh Torah, $en")
                    $key = 'rambam|{0}|{1}' -f $enEnc, $c.perek
                    $c | Add-Member -NotePropertyName resolvedBook -NotePropertyValue $en -Force
                }
                elseif ($en) {
                    $c | Add-Member -NotePropertyName resolvedBook -NotePropertyValue $en -Force
                    $script:bookOkCount++
                }
            }
            elseif ($c.type -eq 'gemara') {
                $en = Resolve-Lookup $maps.tractates $c.tractate
                if ($en) {
                    $enEnc = [uri]::EscapeDataString($en)
                    $key = 'gemara|{0}|{1}{2}' -f $enEnc, $c.daf, $c.amud
                    $c | Add-Member -NotePropertyName resolvedTractate -NotePropertyValue $en -Force
                }
            }
            if ($key) {
                $c | Add-Member -NotePropertyName targetKey -NotePropertyValue $key -Force
                $targets[$key] = $true
            }
        }

        Write-Output ("unique fetch targets: {0} (of {1} citations)" -f $targets.Keys.Count, @($citations).Count)
        $script:bookOkCount = 0
        $results = [ordered]@{}
        $n = 0
        foreach ($key in $targets.Keys) {
            $n++
            $parts = $key -split '\|'
            $status = 'error'
            $snippet = ''
            try {
                $uri = 'https://www.sefaria.org/api/texts/{0}.{1}?context=0&commentary=0' -f $parts[1], $parts[2]
                $resp = Invoke-RestMethod -Uri $uri -TimeoutSec 45
                if ($resp.text -is [array]) {
                    $flat = @($resp.text | ForEach-Object { if ($_ -is [array]) { ($_ -join ' ') } else { [string]$_ } })
                    $snippet = ($flat -join ' ')
                } else {
                    $snippet = [string]$resp.text
                }
                if ($snippet.Length -gt 3500) { $snippet = $snippet.Substring(0, 3500) }
                $status = 'ok'
            } catch {
                $code = $null
                if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                if ($code -eq 404) { $status = 'not-found' } else { $status = "error:$code" }
            }
            $results[$key] = [pscustomobject]@{ status = $status; snippet = $snippet }
            if ($n % 50 -eq 0) { Write-Output ("  progress {0}/{1}" -f $n, $targets.Keys.Count) }
            Start-Sleep -Milliseconds 150
        }

        $errored = @($results.Keys | Where-Object { $results[$_].status -like 'error*' })
        if ($errored.Count -gt 0) {
            Write-Output ("retrying {0} errored targets..." -f $errored.Count)
            foreach ($key in $errored) {
                $parts = $key -split '\|'
                try {
                    $uri = 'https://www.sefaria.org/api/texts/{0}.{1}?context=0&commentary=0' -f $parts[1], $parts[2]
                    $resp = Invoke-RestMethod -Uri $uri -TimeoutSec 90
                    if ($resp.text -is [array]) {
                        $flat = @($resp.text | ForEach-Object { if ($_ -is [array]) { ($_ -join ' ') } else { [string]$_ } })
                        $snippet = ($flat -join ' ')
                    } else {
                        $snippet = [string]$resp.text
                    }
                    if ($snippet.Length -gt 3500) { $snippet = $snippet.Substring(0, 3500) }
                    $results[$key] = [pscustomobject]@{ status = 'ok'; snippet = $snippet }
                } catch {
                    $code = $null
                    if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                    $results[$key] = [pscustomobject]@{ status = "error:$code"; snippet = '' }
                }
                Start-Sleep -Milliseconds 400
            }
        }
        $out = [pscustomobject]@{ citations = $citations; fetchResults = $results }
        $out | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $ResultsPath -Encoding utf8
        $ok = @($results.Values | Where-Object { $_.status -eq 'ok' }).Count
        $nf = @($results.Values | Where-Object { $_.status -eq 'not-found' }).Count
        Write-Output ("done -> {0}" -f $ResultsPath)
        Write-Output ("targets ok={0} not-found={1} error={2}" -f $ok, $nf, ($targets.Keys.Count - $ok - $nf))
    }

    'report' {
        if (-not (Test-Path -LiteralPath $ResultsPath)) { throw 'run fetch first' }
        $data = Get-Content -LiteralPath $ResultsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($data.citations.PSObject.Properties['value']) {
            $citationList = @($data.citations.value)
        } else {
            $citationList = @($data.citations)
        }
        $resolvedCount = 0
        $bookOkCount = 0
        $unresolved = @()
        foreach ($c in $citationList) {
            if ($c.targetKey) {
                $r = $data.fetchResults.PSObject.Properties[$c.targetKey]
                if ($r -and $r.Value.status -eq 'ok') { $resolvedCount++ } else { $unresolved += $c }
            } elseif ($c.type -eq 'rambam' -and $c.resolvedBook -and -not $c.perek) {
                $bookOkCount++
            } else {
                $unresolved += $c
            }
        }
        Write-Output ("citations: {0} total | {1} resolved | {2} book-level (Rambam, no perek cited) | {3} unresolved" -f $citationList.Count, $resolvedCount, $bookOkCount, @($unresolved).Count)
        $unresolved | Group-Object type | ForEach-Object { Write-Output ("  {0}: {1}" -f $_.Name, $_.Count) }
        Write-Output ''
        Write-Output '--- unresolved detail (first 150) ---'
        foreach ($c in ($unresolved | Select-Object -First 150)) {
            $what = switch ($c.type) {
                'sa'     { ('SA [{0}] {1} siman {2}{3}' -f $c.author, $c.branch, $c.siman, $(if ($c.seif) { ':' + $c.seif } else { '' })) }
                'rambam' { ('Rambam "{0}" perek {1} halacha {2} -> {3}' -f $c.title, $c.perek, $c.halacha, $c.resolvedBook) }
                'gemara' { ('Gemara "{0}" daf {1}{2} -> {3}' -f $c.tractate, $c.daf, $c.amud, $c.resolvedTractate) }
            }
            Write-Output ("{0}:{1}  {2}" -f $c.file, $c.line, $what)
        }
    }
}
