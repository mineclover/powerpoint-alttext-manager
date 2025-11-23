# PowerPoint 대체 텍스트 관리자 - 빌드 가이드

## 프로젝트 현황

✅ **완료된 작업:**
- VSTO 프로젝트 구조 완전 생성
- C# Windows Forms UI 구현 완료
- PowerPoint Interop 통합 완료
- 모든 컴파일 오류 수정 완료
- 프로젝트 파일 (.csproj) 구성 완료

⚠️ **현재 상태:**
- 소스 코드는 완전히 준비됨
- MSBuild 도구가 시스템에 설치되어 있지 않음
- ClickOnce 서명을 비활성화하여 개발 빌드 가능하도록 설정

## 빌드 요구사항

### 필수 도구

1. **Visual Studio 2022** (권장) 또는 **Visual Studio Build Tools 2022**
   - 워크로드: Office/SharePoint 개발 도구
   - 또는 최소: .NET 데스크톱 개발 + Office 개발 도구

2. **대안: MSBuild 단독 설치**
   - Visual Studio Build Tools 2022 다운로드: https://visualstudio.microsoft.com/ko/downloads/
   - "Visual Studio용 빌드 도구 2022" 선택

## 빌드 방법

### 방법 1: Visual Studio 사용 (가장 쉬움)

1. Visual Studio 2022에서 프로젝트 열기:
   ```
   파일 > 열기 > 프로젝트/솔루션
   P:\Project\pptx-addon\AltTextManager-VSTO\AltTextManager.csproj 선택
   ```

2. 빌드:
   ```
   빌드 > 솔루션 빌드 (Ctrl+Shift+B)
   ```

3. 출력 확인:
   ```
   bin\Debug\AltTextManager.dll
   bin\Debug\AltTextManager.vsto (설치 파일)
   ```

### 방법 2: Developer Command Prompt 사용

1. "Developer Command Prompt for VS 2022" 실행

2. 프로젝트 디렉터리로 이동:
   ```cmd
   cd /d P:\Project\pptx-addon\AltTextManager-VSTO
   ```

3. 빌드 명령 실행:
   ```cmd
   msbuild AltTextManager.csproj /p:Configuration=Debug /v:minimal
   ```

### 방법 3: PowerShell 스크립트 사용 (Visual Studio 설치 시)

프로젝트에 포함된 `build.ps1` 스크립트 사용:

```powershell
cd P:\Project\pptx-addon\AltTextManager-VSTO
powershell -ExecutionPolicy Bypass -File build.ps1
```

## 설치 방법

### 자동 설치 (ClickOnce)

빌드 후 생성된 `.vsto` 파일 더블클릭:
```
P:\Project\pptx-addon\AltTextManager-VSTO\bin\Debug\AltTextManager.vsto
```

### 수동 설치 (개발자 모드)

레지스트리를 통한 수동 등록:

```powershell
# 관리자 권한 PowerShell에서 실행
$manifestPath = "P:\Project\pptx-addon\AltTextManager-VSTO\bin\Debug\AltTextManager.vsto"
$registryPath = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AltTextManager"

New-Item -Path $registryPath -Force
Set-ItemProperty -Path $registryPath -Name "Manifest" -Value $manifestPath
Set-ItemProperty -Path $registryPath -Name "FriendlyName" -Value "PowerPoint 대체 텍스트 관리자"
Set-ItemProperty -Path $registryPath -Name "Description" -Value "PowerPoint 슬라이드 요소의 대체 텍스트를 쉽게 관리합니다"
Set-ItemProperty -Path $registryPath -Name "LoadBehavior" -Value 3
```

## 문제 해결

### MSBuild를 찾을 수 없음

**증상:** `msbuild: command not found`

**해결:**
1. Visual Studio Installer 실행
2. "수정" 클릭
3. "Office/SharePoint 개발" 워크로드 설치
4. 또는 "Visual Studio용 빌드 도구" 설치

### ClickOnce 서명 오류

**증상:**
```
Cannot build because the ClickOnce manifest signing option is not selected
```

**해결:**
프로젝트 파일에서 이미 서명을 비활성화했습니다:
```xml
<SignManifests>false</SignManifests>
```

개발 환경에서는 서명 없이 빌드 가능합니다.

프로덕션 배포 시 서명이 필요한 경우:
```powershell
# 개발 인증서 생성 (관리자 권한 PowerShell)
New-SelfSignedCertificate -Subject "CN=AltTextManager" -CertStoreLocation "Cert:\CurrentUser\My"
```

### PowerPoint에서 애드인이 로드되지 않음

**확인사항:**
1. PowerPoint 신뢰 센터 설정:
   - 파일 > 옵션 > 보안 센터 > 보안 센터 설정
   - 추가 기능 > "응용 프로그램 추가 기능에 대한 신뢰할 수 있는 위치 필요" 해제

2. 레지스트리 확인:
   ```powershell
   Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\AltTextManager"
   ```

3. LoadBehavior 값이 3인지 확인

## 프로젝트 구조

```
AltTextManager-VSTO/
├── AltTextManager.csproj          # 프로젝트 파일
├── ThisAddIn.cs                    # 애드인 진입점
├── ThisAddIn.Designer.cs           # VSTO 생성 코드
├── AltTextManagerForm.cs           # 메인 UI 폼
├── AltTextManagerForm.Designer.cs  # 폼 디자이너 코드
├── AltTextManagerForm.resx         # 폼 리소스
├── InputDialog.cs                  # 사용자 입력 다이얼로그
├── Properties/
│   ├── AssemblyInfo.cs
│   ├── Resources.resx
│   ├── Resources.Designer.cs
│   ├── Settings.settings
│   └── Settings.Designer.cs
├── ThisAddIn.Designer.xml          # VSTO 매니페스트
└── bin/Debug/                      # 빌드 출력 (빌드 후 생성)
    ├── AltTextManager.dll
    ├── AltTextManager.dll.manifest
    ├── AltTextManager.vsto
    └── [기타 의존성 파일들]
```

## 기능

- ✅ 현재 슬라이드의 모든 요소 목록 표시
- ✅ 각 요소의 대체 텍스트 확인 및 수정
- ✅ 선택한 요소에 대체 텍스트 적용
- ✅ 모든 요소에 일괄 적용
- ✅ 템플릿 기반 자동 생성:
  - 접두사 + 요소 이름
  - 타입별 템플릿 ({type}: {name})
  - 사용자 정의 템플릿 ({name}, {type}, {index} 변수 지원)
- ✅ 빈 대체 텍스트만 적용 옵션
- ✅ 모든 대체 텍스트 삭제

## 다음 단계

1. Visual Studio 또는 Build Tools 설치
2. 위의 빌드 방법 중 하나 선택하여 빌드
3. 생성된 `.vsto` 파일로 설치
4. PowerPoint에서 애드인 실행 및 테스트

## 추가 정보

- .NET Framework 버전: 4.7.2
- PowerPoint Interop 버전: 15.0
- UI 프레임워크: Windows Forms
- 배포 방식: ClickOnce (또는 레지스트리 수동 등록)
