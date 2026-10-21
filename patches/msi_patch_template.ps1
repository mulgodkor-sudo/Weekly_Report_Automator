# ============================================================================
#  Weekly Report Automator — MSI 설치본 패치 (make_patch.py msi 가 .bat 안에 넣어 배포)
#  Windows PowerShell 5.1 호환 문법만 사용할 것.
#
#  동작
#    1) MSI 고정 설치 경로에서 프로그램 찾기
#         C:\Program Files\Autotools_Mechanical\Weekly Report Automator\
#    2) 설치된 버전 확인 (패치보다 높으면 중단)
#    3) src\ 의 .py / .json 교체 → 복사 결과 해시 검증 → src\__pycache__ 정리
#    4) 결과와 패치 내용을 한글로 표시
#  .bat 이 마지막에 "아무 키나 누르면" 창을 닫는다.
# ============================================================================
param(
    [Parameter(Mandatory = $true)] [string] $PayloadDir
)

$PatchVersion = '@VERSION@'
$PatchNotes   = @'
@NOTES@
'@

$ErrorActionPreference = 'Stop'
$InstallSub = 'Autotools_Mechanical\Weekly Report Automator'

function Write-Line([string] $text = '', [string] $color = 'Gray') {
    Write-Host $text -ForegroundColor $color
}

function Get-VersionKey([string] $v) {
    # 1.1 < 1.11 < 1.2 < 1.22 < 1.3 순서가 되도록 소수부를 두 자리로 맞춰 비교
    $parts = $v.Split('.')
    $major = [int] $parts[0]
    $frac  = if ($parts.Length -gt 1) { $parts[1] } else { '0' }
    $frac  = $frac.PadRight(2, '0').Substring(0, 2)
    return ($major * 100) + [int] $frac
}

function Get-InstalledVersion([string] $srcDir) {
    $cfg = Join-Path $srcDir 'config.py'
    if (-not (Test-Path -LiteralPath $cfg)) { return '' }
    $text = [IO.File]::ReadAllText($cfg, [Text.Encoding]::UTF8)
    $m = [regex]::Match($text, '(?m)^APP_VERSION\s*=\s*"([0-9.]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($text, '"version"\s*:\s*"([0-9.]+)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

function Stop-WithError([string[]] $lines, [int] $code) {
    Write-Line
    Write-Line '============================================================' 'Red'
    foreach ($l in $lines) { Write-Line $l 'Red' }
    Write-Line '============================================================' 'Red'
    Write-Line
    Write-Line '창을 닫으려면 아무 키나 누르세요...'
    exit $code
}

Write-Line '============================================================' 'Cyan'
Write-Line " Weekly Report Automator 패치 (v$PatchVersion, MSI 설치본용)" 'Cyan'
Write-Line '============================================================' 'Cyan'
Write-Line

# ── 1) 설치 경로 찾기 ─────────────────────────────────────────────────
$roots = @()
foreach ($name in @('ProgramFiles', 'ProgramW6432', 'ProgramFiles(x86)')) {
    $val = [Environment]::GetEnvironmentVariable($name)
    if ($val -and ($roots -notcontains $val)) { $roots += $val }
}
$candidates = @($roots | ForEach-Object { Join-Path $_ $InstallSub })

$installDir = $null
foreach ($c in $candidates) {
    if ((Test-Path -LiteralPath (Join-Path $c 'src')) -and
        (Get-ChildItem -LiteralPath $c -Filter '*.exe' -ErrorAction SilentlyContinue)) {
        $installDir = $c
        break
    }
}
if (-not $installDir) {
    $msg = @('[오류] 다음 경로에서 Weekly Report Automator 설치 파일을 찾을 수 없습니다:')
    foreach ($c in $candidates) { $msg += "        $c" }
    $msg += ''
    $msg += '        MSI 설치 파일(Weekly_Report_Automator_Vx.x.msi)로 먼저 설치한 뒤 다시 실행하세요.'
    Stop-WithError $msg 2
}
$srcDir = Join-Path $installDir 'src'
Write-Line "설치 경로 : $installDir"

# ── 2) 버전 확인 ──────────────────────────────────────────────────────
$installed = Get-InstalledVersion $srcDir
if ($installed) {
    Write-Line "설치 버전 : v$installed  →  패치 버전 : v$PatchVersion"
    if ((Get-VersionKey $installed) -gt (Get-VersionKey $PatchVersion)) {
        Stop-WithError @("[오류] 이미 더 높은 버전(v$installed)이 설치되어 있습니다.",
                         "        이 패치(v$PatchVersion)는 적용하지 않았습니다.") 3
    }
} else {
    Write-Line "설치 버전 : (확인 불가)  →  패치 버전 : v$PatchVersion"
}

$running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    try { $_.Path -and $_.Path.StartsWith($installDir, [StringComparison]::OrdinalIgnoreCase) }
    catch { $false }
})
if ($running.Count -gt 0) {
    Write-Line '※ 프로그램이 실행 중입니다. 패치 후 프로그램을 다시 시작해야 적용됩니다.' 'Yellow'
}
Write-Line

# ── 3) 파일 교체 + 검증 ───────────────────────────────────────────────
$files = @(Get-ChildItem -LiteralPath $PayloadDir -File | Where-Object { $_.Extension -in '.py', '.json' })
if ($files.Count -eq 0) {
    Stop-WithError @('[오류] 패치 데이터가 비어 있습니다. 패치 파일을 다시 받아 실행하세요.') 4
}
try {
    foreach ($f in $files) {
        Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $srcDir $f.Name) -Force
    }
} catch [UnauthorizedAccessException] {
    Stop-WithError @("[오류] 다음 경로에 파일을 쓸 수 있는 권한이 없습니다:",
                     "        $srcDir",
                     "        패치 파일을 마우스 오른쪽 클릭 → [관리자 권한으로 실행] 하세요.") 5
} catch {
    Stop-WithError @("[오류] 파일을 교체하지 못했습니다:", "        $($_.Exception.Message)",
                     "        권한 문제라면 패치 파일을 마우스 오른쪽 클릭 → [관리자 권한으로 실행] 하세요.") 5
}

$bad = @()
foreach ($f in $files) {
    $dst = Join-Path $srcDir $f.Name
    if (-not (Test-Path -LiteralPath $dst) -or
        (Get-FileHash -LiteralPath $dst -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash) {
        $bad += $f.Name
    }
}
if ($bad.Count -gt 0) {
    Stop-WithError (@('[오류] 다음 파일이 제대로 교체되지 않았습니다:') + ($bad | ForEach-Object { "        $_" })) 6
}

# 이전 버전 컴파일 캐시 정리 (실패해도 무시 — 파이썬이 다시 만듦)
$cache = Join-Path $srcDir '__pycache__'
if (Test-Path -LiteralPath $cache) {
    Remove-Item -LiteralPath $cache -Recurse -Force -ErrorAction SilentlyContinue
}

# ── 4) 결과 ──────────────────────────────────────────────────────────
Write-Line '============================================================' 'Green'
Write-Line " 패치가 성공적으로 완료되었습니다.  (v$PatchVersion, 파일 $($files.Count)개 교체·검증 완료)" 'Green'
Write-Line '============================================================' 'Green'
Write-Line
Write-Line '패치 내용' 'White'
foreach ($l in ($PatchNotes -split "`r?`n")) { Write-Line "  $l" }
Write-Line
Write-Line '※ 프로그램을 다시 시작하면 적용됩니다. (프로그램 화면에 새 버전이 표시됩니다)' 'Yellow'
Write-Line '※ 시작 메뉴 / 앱 목록의 버전 표시는 다음에 MSI로 설치할 때 바뀝니다.' 'DarkGray'
Write-Line
Write-Line '창을 닫으려면 아무 키나 누르세요...'
exit 0
