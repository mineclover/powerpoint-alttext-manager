# PowerPoint 대체 텍스트 관리자 빌드 스크립트

Write-Host "=== PowerPoint 대체 텍스트 관리자 빌드 ===" -ForegroundColor Cyan
Write-Host ""

# MSBuild 경로 찾기
$msbuildPaths = @(
    "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files\Microsoft Visual Studio\2022\Preview\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        Write-Host "MSBuild 발견: $msbuild" -ForegroundColor Green
        break
    }
}

if ($null -eq $msbuild) {
    # PATH에서 msbuild 찾기 시도
    $msbuild = Get-Command msbuild -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source

    if ($null -eq $msbuild) {
        Write-Host "오류: MSBuild를 찾을 수 없습니다!" -ForegroundColor Red
        Write-Host ""
        Write-Host "해결 방법:" -ForegroundColor Yellow
        Write-Host "1. Visual Studio 2022 설치 (권장)" -ForegroundColor Yellow
        Write-Host "2. Visual Studio Build Tools 2022 설치" -ForegroundColor Yellow
        Write-Host "   다운로드: https://visualstudio.microsoft.com/ko/downloads/" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "또는 Developer Command Prompt for VS 2022에서 다음 명령 실행:" -ForegroundColor Yellow
        Write-Host "   msbuild AltTextManager.csproj /p:Configuration=Debug" -ForegroundColor Cyan
        exit 1
    }
}

Write-Host ""

# 프로젝트 경로
$projectPath = Join-Path $PSScriptRoot "AltTextManager.csproj"

if (-not (Test-Path $projectPath)) {
    Write-Host "오류: 프로젝트 파일을 찾을 수 없습니다: $projectPath" -ForegroundColor Red
    exit 1
}

Write-Host "프로젝트: $projectPath" -ForegroundColor Gray
Write-Host ""

# 빌드 실행
Write-Host "빌드 시작..." -ForegroundColor Yellow
& $msbuild $projectPath /p:Configuration=Debug /verbosity:minimal /nologo

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "✓ 빌드 성공!" -ForegroundColor Green
    Write-Host ""

    $outputPath = Join-Path $PSScriptRoot "bin\Debug"
    $dllPath = Join-Path $outputPath "AltTextManager.dll"
    $vstoPath = Join-Path $outputPath "AltTextManager.vsto"

    if (Test-Path $dllPath) {
        Write-Host "출력 파일:" -ForegroundColor Cyan
        Write-Host "  DLL: $dllPath" -ForegroundColor Gray

        if (Test-Path $vstoPath) {
            Write-Host "  설치 파일: $vstoPath" -ForegroundColor Gray
            Write-Host ""
            Write-Host "설치 방법:" -ForegroundColor Yellow
            Write-Host "  $vstoPath 파일을 더블클릭하여 설치" -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "다음 단계:" -ForegroundColor Cyan
    Write-Host "1. 생성된 .vsto 파일을 더블클릭하여 설치" -ForegroundColor Gray
    Write-Host "2. PowerPoint를 실행하고 애드인이 로드되는지 확인" -ForegroundColor Gray
    Write-Host "3. PowerPoint 리본 메뉴에서 '대체 텍스트 관리자' 찾기" -ForegroundColor Gray
} else {
    Write-Host ""
    Write-Host "✗ 빌드 실패!" -ForegroundColor Red
    Write-Host "위의 오류 메시지를 확인하세요." -ForegroundColor Yellow
    exit 1
}
