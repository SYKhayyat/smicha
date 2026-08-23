param(
    [string]$OutDir = (Join-Path $env:TEMP 'opencode\smicha-verify\corpus'),
    [string]$BooksJsonUrl = 'https://raw.githubusercontent.com/Sefaria/Sefaria-Export/master/books.json'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = New-Object System.Text.UTF8Encoding $false
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }

$commentaryAuthors = @(
    'Siftei Kohen', 'Turei Zahav', 'Mishnah Berurah', 'Aruch HaShulchan',
    'Magen Avraham', 'Eliyah Rabbah', 'Ba''er Hetev',
    'Be''er HaGolah', 'Pri Megadim', 'Pri Chadash', 'Kaf Hachayim',
    'Pitchei Teshuva', 'Darkei Moshe', 'Beit Yosef', 'Shaarei Teshuvah',
    'Eshel Avraham', 'Chelkat Mechokek', 'Beit Shmuel'
)

function Test-Wanted {
    param($book)
    $cats = $book.categories
    $isBavli = ($cats -contains 'Talmud') -and ($cats -contains 'Bavli')
    $isMt = ($cats -contains 'Mishneh Torah')
    $isSaBase = ($cats -contains 'Shulchan Arukh') -and ($cats.Count -le 4)
    $isCommentary = $false
    foreach ($a in $commentaryAuthors) {
        if ($book.title -like "*$a*") { $isCommentary = $true; break }
    }
    $isSa = $cats -contains 'Shulchan Arukh'
    return ($isBavli -or $isMt -or $isSa -or $isCommentary)
}

Write-Host 'downloading books.json...'
$tmp = Join-Path $env:TEMP 'opencode\books.json'
Invoke-WebRequest -Uri $BooksJsonUrl -OutFile $tmp -TimeoutSec 120
$books = Get-Content -LiteralPath $tmp -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Output ("total books in index: " + @($books.books).Count)

$wanted = @($books.books | Where-Object {
    ($_.language -eq 'he') -and (Test-Wanted $_)
})
Write-Output ("wanted hebrew books: " + $wanted.Count)

$n = 0
foreach ($b in $wanted) {
    $n++
    $safe = ($b.title -replace '[^a-zA-Z0-9 ,.\-\u0590-\u05FF]', '' ) -replace '\s+', '_'
    $dest = Join-Path $OutDir ($safe + '.txt')
    if (Test-Path -LiteralPath $dest) { continue }
    try {
        Invoke-WebRequest -Uri $b.txt_url -OutFile $dest -TimeoutSec 300
        Write-Host ("[{0}/{1}] OK  {2}" -f $n, $wanted.Count, $b.title)
    } catch {
        Write-Host ("[{0}/{1}] FAIL {2} : {3}" -f $n, $wanted.Count, $b.title, $_.Exception.Message)
    }
}
Write-Output ("corpus files now: " + @(Get-ChildItem -LiteralPath $OutDir -Filter '*.txt').Count)
