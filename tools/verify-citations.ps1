param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('extract', 'fetch', 'report', 'match', 'matchlocal')]
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

function Get-DefaultBranchFromPath {
    param([string]$RelPath)
    if ($RelPath -match '^(taaroves|issur-veheter|basar-bechalav|yoreh-deah|nidda)/') { return 'Y.D.' }
    if ($RelPath -match '^choshen-mishpat/') { return 'C.M.' }
    if ($RelPath -match '^even-haezer/') { return 'E.H.' }
    return 'O.C.'
}

$script:RambamAliases = @{
    'יום טוב'            = 'שביתת יו"ט'
    'יו"ט'               = 'שביתת יו"ט'
    'נדרים ושבועות'      = 'נדרים'
    'סנהדרין ועדות'      = 'סנהדרין'
    'נזקי ממון וחובל ומזיק' = 'נזקי ממון'
    'גזלה ואבדה'         = 'גניבה'
    'ציצית ותפילה ומזוזה וספר תורה' = 'ציצית'
    'ציצית ות'           = 'ציצית'
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
            $simanMatch = [regex]::Match($window, "(?:סי['\u05F3\x22\u201D]\s*|סימן\s+)([\p{IsHebrew}\x22']{1,8})")
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
                else { $branch = Get-DefaultBranchFromPath $RelPath }
            }

            $found += [pscustomobject]@{
                type = 'sa'; author = $m.Groups[1].Value; branch = $branch
                siman = $siman; seif = $seif; seifKind = $seifKind
                file = $RelPath; line = $lineNo; text = $line.Trim()
            }
        }

        $rambamPattern = 'רמב"ם\s+(?:הלכות\s+)?([\p{IsHebrew}"][\p{IsHebrew}" ]*?)\s*(?:פ"?\s*([\p{IsHebrew}"]{1,8}))?(?:\s+ה"?\s*([\p{IsHebrew}"]{1,6}))?(?=\s*[;,.):]|$)'
        $rambamMatch = [regex]::Match($line, $rambamPattern)
        if ($rambamMatch.Success -and $rambamMatch.Value -match 'הלכות|פ"') {
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

function Get-NormalizedHebrew {
    param([string]$s)
    return ($s -replace '[\u05D5\u05D9\u05F4\u05F3"'' ]', '')
}

function Resolve-Lookup {
    param($MapObject, [string]$Key)
    $normKey = Get-NormalizedHebrew $Key
    $bareKey = $Key -replace '^הלכות ', ''
    $normBare = Get-NormalizedHebrew $bareKey

    if ($MapObject -is [hashtable]) {
        foreach ($k in $MapObject.Keys) {
            if ((Get-NormalizedHebrew $k) -eq $normKey) { return $MapObject[$k] }
        }
        foreach ($k in $MapObject.Keys) {
            if ((Get-NormalizedHebrew $k).Contains($normBare)) { return $MapObject[$k] }
        }
        return $null
    }

    $prop = $MapObject.PSObject.Properties[$Key]
    if ($prop) { return $prop.Value }
    $best = $null
    foreach ($p in $MapObject.PSObject.Properties) {
        $nk = Get-NormalizedHebrew $p.Name
        if ($nk -eq $normKey) { return $p.Value }
        if (-not $best -and $nk.Contains($normBare)) { $best = $p }
    }
    if ($best) { return $best.Value }
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
                $lookupTitle = $c.title
                if ($script:RambamAliases.ContainsKey($lookupTitle)) { $lookupTitle = $script:RambamAliases[$lookupTitle] }
                $en = Resolve-Lookup $maps.rambamBooks ('הלכות ' + $lookupTitle)
                if (-not $en) { $en = Resolve-Lookup $maps.rambamBooks ('הלכות ' + $c.title) }
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
                    $amudAscii = if ($c.amud -eq [char]0x05D0) { 'a' } else { 'b' }
                    $enEnc = [uri]::EscapeDataString($en)
                    $key = 'gemara|{0}|{1}{2}' -f $enEnc, $c.daf, $amudAscii
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
                $srcText = if ($resp.PSObject.Properties['he'] -and $resp.he) { $resp.he } else { $resp.text }
                if ($srcText -is [array]) {
                    $flat = @($srcText | ForEach-Object { if ($_ -is [array]) { ($_ -join ' ') } else { [string]$_ } })
                    $snippet = ($flat -join ' ')
                } else {
                    $snippet = [string]$srcText
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
        $noiseCount = 0
        $unresolved = @()
        foreach ($c in $citationList) {
            if ($c.targetKey) {
                $r = $data.fetchResults.PSObject.Properties[$c.targetKey]
                if ($r -and $r.Value.status -eq 'ok') { $resolvedCount++ } else { $unresolved += $c }
            } elseif ($c.type -eq 'rambam' -and $c.resolvedBook -and -not $c.perek) {
                $bookOkCount++
            } elseif ($c.type -eq 'rambam' -and $c.title -match '^ו') {
                $noiseCount++
            } else {
                $unresolved += $c
            }
        }
        Write-Output ("citations: {0} total | {1} resolved | {2} book-level (Rambam, no perek cited) | {3} prose-noise | {4} unresolved" -f $citationList.Count, $resolvedCount, $bookOkCount, $noiseCount, @($unresolved).Count)
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

    'match' {
        if (-not (Test-Path -LiteralPath $ResultsPath)) { throw 'run fetch first' }
        $data = Get-Content -LiteralPath $ResultsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($data.citations.PSObject.Properties['value']) { $citationList = @($data.citations.value) } else { $citationList = @($data.citations) }

        function Fold-Hebrew {
            param([string]$s)
            $s = $s -replace '[\u0591-\u05C7\u05F4\u05F3"''<>/|*\[\]]', ''
            $s = $s -replace '\u05DA', '\u05DB' -replace '\u05DD', '\u05DE' -replace '\u05DF', '\u05E0' -replace '\u05E3', '\u05E4' -replace '\u05E5', '\u05E6'
            return $s
        }
        function DeMatres {
            param([string]$s)
            return ($s -replace '[\u05D5\u05D9]', '')
        }

        $stop = @('של', 'את', 'על', 'אין', 'אם', 'כל', 'רמב', 'הלכות', 'מקורות', 'שו', 'רמ', 'סי', 'עי', 'וכן', 'אלא', 'שה', 'מה', 'או', 'זה', 'כגון', 'לכן', 'מכל')
        $fileCache = @{}
        $rows = @()
        foreach ($c in $citationList) {
            if (-not $c.targetKey) { continue }
            $r = $data.fetchResults.PSObject.Properties[$c.targetKey]
            if (-not $r -or $r.Value.status -ne 'ok') { continue }
            $snippet = $r.Value.snippet
            if (-not $snippet) { continue }

            $cacheKey = "$($c.file)"
            if (-not $fileCache.ContainsKey($cacheKey)) {
                $fp = Join-Path $RepoRoot ($c.file -replace '/', '\')
                if (Test-Path -LiteralPath $fp) {
                    $fileCache[$cacheKey] = Get-Content -LiteralPath $fp -Encoding UTF8
                } else {
                    $fileCache[$cacheKey] = @()
                }
            }
            $lines = $fileCache[$cacheKey]
            $lo = [Math]::Max(0, $c.line - 3)
            $hi = [Math]::Min($lines.Count - 1, $c.line + 1)
            $context = ($lines[$lo..$hi] -join ' ')
            if ($context -match 'מקורות:') { continue }

            $tokens = [regex]::Matches($context, '[\p{IsHebrew}]{3,}') |
                ForEach-Object { $_.Value.Trim([char]0x05F4, [char]0x05F3, '"', "'") } |
                Where-Object { $_.Length -ge 3 -and $stop -notcontains $_ } |
                Select-Object -Unique -First 25
            if (@($tokens).Count -lt 4) { continue }

            $normSnippet = Fold-Hebrew (($snippet -replace '<[^>]+>', ' '))
            $normSnippetDeM = DeMatres $normSnippet
            $hits = 0
            foreach ($t in $tokens) {
                $pat = Fold-Hebrew $t
                if ($normSnippet.Contains($t) -or $normSnippet.Contains($pat)) { $hits++; continue }
                if ($normSnippetDeM.Contains((DeMatres $pat))) { $hits++ }
            }
            $score = [Math]::Round(100 * $hits / @($tokens).Count)
            $rows += [pscustomobject]@{ score = $score; file = $c.file; line = $c.line; key = $c.targetKey; text = $c.text }
        }

        $sorted = @($rows | Sort-Object score)
        $out = @()
        foreach ($row in $sorted) {
            $out += ("{0}%  {1}:{2}  [{3}]  {4}" -f $row.score, $row.file, $row.line, ($row.key -split '\|')[0], $row.text)
        }
        [IO.File]::WriteAllLines((Join-Path $WorkDir 'match-report.txt'), $out, (New-Object System.Text.UTF8Encoding $true))
        $low = @($sorted | Where-Object { $_.score -lt 25 }).Count
        Write-Output ("matched {0} citations; below-25%: {1}" -f @($rows).Count, $low)
        Write-Output ("full list -> {0}\match-report.txt" -f $WorkDir)
    }

    'matchlocal' {
        if (-not (Test-Path -LiteralPath $ResultsPath)) { throw 'run fetch first' }
        $corpusDir = Join-Path $WorkDir 'corpus'
        if (-not (Test-Path -LiteralPath $corpusDir)) { throw "no corpus at $corpusDir" }
        $data = Get-Content -LiteralPath $ResultsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($data.citations.PSObject.Properties['value']) { $citationList = @($data.citations.value) } else { $citationList = @($data.citations) }

        function Fold-Hebrew {
            param([string]$s)
            $s = $s -replace '[\u0591-\u05C7\u05F4\u05F3"''<>/|*\[\]]', ''
            $s = $s -replace '\u05DA', '\u05DB' -replace '\u05DD', '\u05DE' -replace '\u05DF', '\u05E0' -replace '\u05E3', '\u05E4' -replace '\u05E5', '\u05E6'
            return $s
        }
        function DeMatres {
            param([string]$s)
            return ($s -replace '[\u05D5\u05D9]', '')
        }

        $branchFile = @{
            'O.C.' = 'Shulchan_Arukh,_Orach_Chayim.txt'
            'Y.D.' = 'Shulchan_Arukh,_Yoreh_Deah.txt'
            'E.H.' = 'Shulchan_Arukh,_Even_HaEzer.txt'
            'C.M.' = 'Shulchan_Arukh,_Choshen_Mishpat.txt'
        }

        $sectionCache = @{}
        function Get-LocalSection {
            param([string]$FilePath, [string]$StartMarker, [string]$EndPattern)
            $ck = "$FilePath::$StartMarker"
            if ($sectionCache.ContainsKey($ck)) { return $sectionCache[$ck] }
            $full = Join-Path $corpusDir $FilePath
            $result = ''
            if (Test-Path -LiteralPath $full) {
                $lines = [IO.File]::ReadAllLines($full, [Text.Encoding]::UTF8)
                $startIdx = -1
                for ($i = 0; $i -lt $lines.Count; $i++) {
                    if ($lines[$i] -match $StartMarker) { $startIdx = $i; break }
                }
                if ($startIdx -ge 0) {
                    $endIdx = $lines.Count
                    for ($j = $startIdx + 1; $j -lt $lines.Count; $j++) {
                        if ($lines[$j] -match $EndPattern) { $endIdx = $j; break }
                    }
                    $chunk = $lines[($startIdx + 1)..($endIdx - 1)] -join ' '
                    if ($chunk.Length -gt 400000) { $chunk = $chunk.Substring(0, 400000) }
                    $result = $chunk
                }
            }
            $sectionCache[$ck] = $result
            return $result
        }

        $stop = @('של', 'את', 'על', 'אין', 'אם', 'כל', 'רמב', 'הלכות', 'מקורות', 'שו', 'רמ', 'סי', 'עי', 'וכן', 'אלא', 'שה', 'מה', 'או', 'זה', 'כגון', 'לכן', 'מכל')
        $fileCtx = @{}
        $rows = @()
        foreach ($c in $citationList) {
            if (-not $c.targetKey) { continue }
            $parts = $c.targetKey -split '\|'

            $localText = ''
            $localHow = 'missing'
            try {
                if ($c.type -eq 'sa') {
                    $bf = $branchFile[$c.branch]
                    $localText = Get-LocalSection -FilePath $bf -StartMarker ("^Siman\s+{0}\s*$" -f $c.siman) -EndPattern '^Siman\s+\d+\s*$'
                    if ($localText) { $localHow = "siman $($c.siman)" }
                } elseif ($c.type -eq 'gemara' -and $c.resolvedTractate) {
                    $tf = ($c.resolvedTractate -replace '[^a-zA-Z0-9 ,.\-]', '') -replace '\s+', '_'
                    $amudAscii = if ("$($c.amud)" -eq [string][char]0x05D0) { 'a' } else { 'b' }
                    $localText = Get-LocalSection -FilePath "$tf.txt" -StartMarker ("^Daf\s+{0}{1}\s*$" -f $c.daf, $amudAscii) -EndPattern '^Daf\s+\d+[ab]\s*$'
                    if ($localText) { $localHow = "daf $($c.daf)$($c.amud)" }
                } elseif ($c.type -eq 'rambam' -and $c.resolvedBook -and $c.perek) {
                    $rf = ($c.resolvedBook -replace '[^a-zA-Z0-9 ,.\-]', '') -replace '\s+', '_'
                    $localText = Get-LocalSection -FilePath "Mishneh_Torah,_$rf.txt" -StartMarker ("^Chapter\s+{0}\s*$" -f $c.perek) -EndPattern '^Chapter\s+\d+\s*$'
                    if ($localText) { $localHow = "perek $($c.perek)" }
                }
            } catch { $localHow = "error:$($_.Exception.Message)" }

            if (-not $localText) {
                $r = $data.fetchResults.PSObject.Properties[$c.targetKey]
                if ($r -and $r.Value.status -eq 'ok') { $localText = $r.Value.snippet; $localHow += '+snippet' } else { continue }
            }

            $cacheKey = "$($c.file)"
            if (-not $fileCtx.ContainsKey($cacheKey)) {
                $fp = Join-Path $RepoRoot ($c.file -replace '/', '\')
                if (Test-Path -LiteralPath $fp) { $fileCtx[$cacheKey] = Get-Content -LiteralPath $fp -Encoding UTF8 } else { $fileCtx[$cacheKey] = @() }
            }
            $lines = $fileCtx[$cacheKey]
            $lo = [Math]::Max(0, $c.line - 3)
            $hi = [Math]::Min($lines.Count - 1, $c.line + 1)
            $context = ($lines[$lo..$hi] -join ' ')
            if ($context -match 'מקורות:') { continue }

            $tokens = [regex]::Matches($context, '[\p{IsHebrew}]{3,}') |
                ForEach-Object { $_.Value.Trim([char]0x05F4, [char]0x05F3, '"', "'") } |
                Where-Object { $_.Length -ge 3 -and $stop -notcontains $_ } |
                Select-Object -Unique -First 25
            if (@($tokens).Count -lt 4) { continue }

            $secFolded = Fold-Hebrew $localText
            $secDeM = DeMatres $secFolded
            $hits = 0
            foreach ($t in $tokens) {
                $pat = Fold-Hebrew $t
                if ($secFolded.Contains($t) -or $secFolded.Contains($pat)) { $hits++; continue }
                if ($secDeM.Contains((DeMatres $pat))) { $hits++ }
            }
            $score = [Math]::Round(100 * $hits / @($tokens).Count)
            $rows += [pscustomobject]@{ score = $score; file = $c.file; line = $c.line; how = $localHow; text = $c.text }
        }

        $sorted = @($rows | Sort-Object score)
        $out = @()
        foreach ($row in $sorted) {
            $out += ("{0}%  [{1}]  {2}:{3}  {4}" -f $row.score, $row.how, $row.file, $row.line, $row.text)
        }
        [IO.File]::WriteAllLines((Join-Path $WorkDir 'match-local-report.txt'), $out, (New-Object System.Text.UTF8Encoding $true))
        $low = @($sorted | Where-Object { $_.score -lt 25 }).Count
        Write-Output ("matchlocal: {0} scored; below-25%: {1}" -f @($rows).Count, $low)
    }
}
