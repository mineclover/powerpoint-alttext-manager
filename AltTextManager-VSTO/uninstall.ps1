# PowerPoint 대체 텍스트 관리자 - 제거 스크립트

Write-Host "=== PowerPoint 대체 텍스트 관리자 제거 ===" -ForegroundColor Cyan
Write-Host ""

# 레지스트리 경로
$registryPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AltTextManager"

# 애드인이 설치되어 있는지 확인
if (Test-Path $registryPath) {
    Write-Host "애드인 설치 확인됨: $registryPath" -ForegroundColor Yellow
    Write-Host ""

    # 현재 설정 표시
    Write-Host "현재 설정:" -ForegroundColor Cyan
    Get-ItemProperty -Path $registryPath | Format-List Manifest, FriendlyName, Description, LoadBehavior

    Write-Host ""
    $confirm = Read-Host "정말로 제거하시겠습니까? (Y/N)"

    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        try {
            Remove-Item -Path $registryPath -Recurse -Force
            Write-Host ""
            Write-Host "✓ 애드인이 제거되었습니다!" -ForegroundColor Green
            Write-Host ""
            Write-Host "PowerPoint를 재시작하면 애드인이 더 이상 로드되지 않습니다." -ForegroundColor Gray
        } catch {
            Write-Host ""
            Write-Host "❌ 제거 중 오류 발생!" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            exit 1
        }
    } else {
        Write-Host ""
        Write-Host "제거가 취소되었습니다." -ForegroundColor Yellow
    }
} else {
    Write-Host "애드인이 설치되어 있지 않습니다." -ForegroundColor Gray
    Write-Host "레지스트리 경로: $registryPath" -ForegroundColor Gray
}
